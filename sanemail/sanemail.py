#!/usr/bin/env python3
"""
SaneMail: SaneBox-style sorting for IMAP mailboxes (iCloud, custom domains, any host)
and Office 365 / Exchange Online (via Microsoft Graph).

Sorting happens on the mail server, so every device and every mail app sees the same
folders:

  Inbox          people you correspond with, replies to your mail, VIPs, your own domain
  SaneLater      senders you've never written to
  SaneNews       newsletters and bulk mail
  SaneBlackHole  drop one message here and that sender goes straight to Trash from then on
  SaneTomorrow / SaneNextWeek
                 snooze: the message comes back to the Inbox, unread, at the wake-up hour

Training: move a message into SaneLater / SaneNews / SaneBlackHole, or from those folders
back to the Inbox, and the sender is remembered (in rules.json, shared by all accounts).

A daily digest of what was filed is delivered to the Inbox of one account.

Run:  python sanemail.py            (normal)
      SANEMAIL_DRY_RUN=1 python sanemail.py   (log only; moves nothing, saves nothing)
"""
from __future__ import annotations

import datetime as dt
import email
import email.header
import email.policy
import hashlib
import html
import imaplib
import json
import logging
import os
import re
import ssl
import sys
import time
from dataclasses import dataclass, field
from email.message import EmailMessage
from email.utils import format_datetime, getaddresses, make_msgid, parseaddr, parsedate_to_datetime
from zoneinfo import ZoneInfo

import yaml

ROOT = os.path.dirname(os.path.abspath(__file__))
STATE_DIR = os.environ.get('SANEMAIL_STATE_DIR', os.path.join(ROOT, '.state'))
RULES_PATH = os.environ.get('SANEMAIL_RULES', os.path.join(ROOT, 'rules.json'))
CONFIG_PATH = os.environ.get('SANEMAIL_CONFIG', os.path.join(ROOT, 'config.yaml'))
DRY_RUN = os.environ.get('SANEMAIL_DRY_RUN', '').lower() in ('1', 'true', 'yes')
DEBUG = os.environ.get('SANEMAIL_DEBUG', '').lower() in ('1', 'true', 'yes')

FOLDER_NAMES = {
    'later': 'SaneLater', 'news': 'SaneNews', 'blackhole': 'SaneBlackHole',
    'tomorrow': 'SaneTomorrow', 'nextweek': 'SaneNextWeek',
}
WINDOW_DAYS = 30          # how far back folders are compared for training
PUBLIC_DOMAINS = {'gmail.com', 'googlemail.com', 'icloud.com', 'me.com', 'mac.com', 'outlook.com',
                  'hotmail.com', 'live.com', 'msn.com', 'yahoo.com', 'aol.com', 'proton.me', 'protonmail.com'}
BULK_LOCALPART = re.compile(
    r'^(no-?reply|do-?not-?reply|donotreply|newsletters?|news|marketing|notifications?|notify|updates?|'
    r'info|hello|team|mailer|digest|offers?|deals?|promo(tions)?|bounce[s]?|alerts?)([@+._-]|$)', re.I)
BULK_MAILER = re.compile(r'mailchimp|sendgrid|mailgun|constant ?contact|hubspot|klaviyo|substack|'
                         r'campaign|marketo|salesforce|braze|customer\.io|sendinblue|brevo', re.I)

log = logging.getLogger('sanemail')
UTC = dt.timezone.utc


def now_utc() -> dt.datetime:
    return dt.datetime.now(UTC)


def iso(d: dt.datetime) -> str:
    return d.astimezone(UTC).isoformat(timespec='seconds')


def parse_iso(s: str) -> dt.datetime:
    d = dt.datetime.fromisoformat(s.replace('Z', '+00:00'))
    return d if d.tzinfo else d.replace(tzinfo=UTC)


# ───────────────────────────── Secrets ─────────────────────────────

def secret(name: str, required: bool = True) -> str:
    """Read a secret from the environment, or from SECRETS_JSON (all GitHub secrets as JSON)."""
    val = os.environ.get(name)
    if not val and os.environ.get('SECRETS_JSON'):
        try:
            val = json.loads(os.environ['SECRETS_JSON']).get(name)
        except ValueError:
            val = None
    if required and not val:
        raise RuntimeError(f'secret {name} is not set')
    return val or ''


# ───────────────────────────── Messages ─────────────────────────────

@dataclass
class Msg:
    key: str                         # stable id across folders (Message-ID)
    sender: str
    name: str
    subject: str
    date: dt.datetime
    folder: str
    ref: object                      # backend handle (IMAP UID or Graph id)
    thread_keys: set = field(default_factory=set)
    recipients: list = field(default_factory=list)
    headers: dict = field(default_factory=dict)   # lower-case header name → value
    flagged: bool = False


def decode_header(v) -> str:
    if v is None:
        return ''
    try:
        return str(email.header.make_header(email.header.decode_header(str(v))))
    except Exception:
        return str(v)


def fallback_key(*parts) -> str:
    return 'h:' + hashlib.sha1('|'.join(map(str, parts)).encode()).hexdigest()[:20]


# ───────────────────────────── IMAP backend ─────────────────────────────

MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def imap_quote(name: str) -> str:
    return '"' + name.replace('\\', '\\\\').replace('"', '\\"') + '"'


class ImapBackend:
    HDRS = ('MESSAGE-ID FROM TO CC SUBJECT DATE IN-REPLY-TO REFERENCES LIST-UNSUBSCRIBE LIST-ID '
            'PRECEDENCE X-MAILER AUTO-SUBMITTED')

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.c: imaplib.IMAP4_SSL | None = None
        self.selected = None

    # connection / folders
    def connect(self):
        self.c = imaplib.IMAP4_SSL(self.cfg['host'], int(self.cfg.get('port', 993)),
                                   ssl_context=ssl.create_default_context(), timeout=60)
        self.c.login(self.cfg['username'], secret(self.cfg['password_secret']))
        typ, data = self.c.capability()  # capabilities can grow after login (e.g. MOVE)
        self.caps = set(b' '.join(data).decode().upper().split()) | {c.upper() for c in self.c.capabilities}
        self._discover()

    def close(self):
        try:
            if self.c:
                self.c.logout()
        except Exception:
            pass

    def _discover(self):
        typ, data = self.c.list()
        existing, roles, delim, names = {}, {'inbox': 'INBOX'}, '/', []
        for raw in data or []:
            if raw is None:
                continue
            if isinstance(raw, tuple):  # literal folder name
                line = raw[0].decode(errors='replace') + imap_quote(raw[1].decode(errors='replace'))
            else:
                line = raw.decode(errors='replace')
            m = re.match(r'\((?P<flags>[^)]*)\)\s+(?P<delim>"(?:[^"\\]|\\.)*"|NIL)\s+(?P<name>.+)$', line)
            if not m:
                continue
            name = m['name'].strip()
            if name.startswith('"') and name.endswith('"'):
                name = re.sub(r'\\(.)', r'\1', name[1:-1])
            if m['delim'] != 'NIL':
                delim = re.sub(r'\\(.)', r'\1', m['delim'][1:-1])
            flags = m['flags'].lower()
            existing[name.lower()] = name
            names.append(name)
            if '\\sent' in flags:
                roles.setdefault('sent', name)
            if '\\trash' in flags:
                roles.setdefault('trash', name)
        fallbacks = {'sent': ['Sent Messages', 'Sent', 'Sent Items', 'Sent Mail', 'INBOX.Sent'],
                     'trash': ['Deleted Messages', 'Trash', 'Deleted Items', 'Bin', 'INBOX.Trash']}
        for role, cands in fallbacks.items():
            if role not in roles:
                roles[role] = next((existing[c.lower()] for c in cands if c.lower() in existing), None)
        # Some servers (cPanel/Courier) keep every folder under "INBOX."
        others = [n for n in names if n.upper() != 'INBOX']
        prefix = self.cfg.get('folder_prefix')
        if prefix is None:
            prefix = 'INBOX' + delim if others and all(n.upper().startswith('INBOX' + delim) for n in others) else ''
        self.prefix, self.existing = prefix, existing
        self.roles = {**roles, **self.cfg.get('folders', {})}
        for role in ('sent', 'trash'):
            if not self.roles.get(role):
                raise RuntimeError(f'could not find the {role} folder; set folders: {{{role}: "<name>"}} in config.yaml')

    def resolve(self, folder: str) -> str:
        return self.roles.get(folder) or self.prefix + folder

    def ensure_folders(self):
        for f in FOLDER_NAMES.values():
            name = self.resolve(f)
            if name.lower() not in self.existing:
                log.info('  creating folder %s', name)
                typ, data = self.c.create(imap_quote(name))
                if typ != 'OK':
                    raise RuntimeError(f'could not create folder {name}: {data}')
                self.c.subscribe(imap_quote(name))
                self.existing[name.lower()] = name

    def addresses(self) -> set:
        return {a.lower() for a in [self.cfg['username'], *self.cfg.get('addresses', [])]}

    def _select(self, folder: str):
        if self.selected != folder:
            typ, data = self.c.select(imap_quote(self.resolve(folder)))
            if typ != 'OK':
                raise RuntimeError(f'cannot open folder {self.resolve(folder)}: {data}')
            self.selected = folder

    # reading
    def list(self, folder: str, since: dt.datetime | None, headers: bool = False) -> list[Msg]:
        self._select(folder)
        crit = ['ALL'] if since is None else ['SINCE', f'{since.day}-{MONTHS[since.month - 1]}-{since.year}']
        typ, data = self.c.uid('SEARCH', None, *crit)
        uids = data[0].split() if data and data[0] else []
        out = []
        for i in range(0, len(uids), 200):
            chunk = b','.join(uids[i:i + 200]).decode()
            typ, resp = self.c.uid('FETCH', chunk, f'(UID FLAGS BODY.PEEK[HEADER.FIELDS ({self.HDRS})])')
            for j, item in enumerate(resp or []):
                if not isinstance(item, tuple):
                    continue
                meta = item[0].decode(errors='replace')
                nxt = resp[j + 1] if j + 1 < len(resp) and isinstance(resp[j + 1], bytes) else b''
                meta_all = meta + ' ' + nxt.decode(errors='replace')
                um = re.search(r'UID (\d+)', meta_all)
                if not um:
                    continue
                fm = re.search(r'FLAGS \(([^)]*)\)', meta_all)
                flags = fm.group(1) if fm else ''
                msg = self._msg(item[1], folder, um.group(1), '\\Flagged' in flags)
                if msg:
                    out.append(msg)
        return out

    def _msg(self, raw: bytes, folder: str, uid: str, flagged: bool) -> Msg | None:
        h = email.message_from_bytes(raw, policy=email.policy.compat32)
        name, sender = parseaddr(decode_header(h.get('From')))
        try:
            date = parsedate_to_datetime(h.get('Date'))
            date = date if date.tzinfo else date.replace(tzinfo=UTC)
        except Exception:
            date = now_utc()
        subject = decode_header(h.get('Subject'))
        mid = (h.get('Message-ID') or '').strip()
        key = mid or fallback_key(sender, subject, h.get('Date'))
        refs = set(re.findall(r'<[^>]+>', f"{h.get('In-Reply-To') or ''} {h.get('References') or ''}"))
        recips = [a.lower() for _, a in getaddresses([decode_header(h.get('To')), decode_header(h.get('Cc'))]) if a]
        hdrs = {k.lower(): str(v) for k, v in h.items()}
        return Msg(key=key, sender=sender.lower(), name=name, subject=subject, date=date, folder=folder,
                   ref=uid, thread_keys=refs | ({mid} if mid else set()), recipients=recips,
                   headers=hdrs, flagged=flagged)

    # writing
    def move(self, m: Msg, dest: str):
        self._select(m.folder)
        target = imap_quote(self.resolve(dest))
        if 'MOVE' in self.caps:
            typ, data = self.c.uid('MOVE', m.ref, target)
        else:
            typ, data = self.c.uid('COPY', m.ref, target)
            if typ == 'OK':
                self.c.uid('STORE', m.ref, '+FLAGS', '(\\Deleted)')
                if 'UIDPLUS' in self.caps:
                    self.c.uid('EXPUNGE', m.ref)
                else:
                    self.c.expunge()
        if typ != 'OK':
            raise RuntimeError(f'move failed: {data}')

    def mark_unread(self, m: Msg):
        self._select(m.folder)
        self.c.uid('STORE', m.ref, '-FLAGS', '(\\Seen)')

    def deliver(self, subject: str, html_body: str, text_body: str):
        me = self.cfg['username']
        msg = EmailMessage()
        msg['From'] = f'SaneMail <{me}>'
        msg['To'] = me
        msg['Subject'] = subject
        msg['Date'] = format_datetime(now_utc())
        msg['Message-ID'] = make_msgid(domain='sanemail.local')
        msg.set_content(text_body)
        msg.add_alternative(html_body, subtype='html')
        typ, data = self.c.append('INBOX', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        if typ != 'OK':
            raise RuntimeError(f'could not deliver digest: {data}')


# ───────────────────────────── Office 365 (Microsoft Graph) backend ─────────────────────────────

GRAPH_SCOPES = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/User.Read']
WELL_KNOWN = {'inbox': 'inbox', 'sent': 'sentitems', 'trash': 'deleteditems'}


def o365_token_cache_path(account_name: str) -> str:
    return os.path.join(STATE_DIR, f'o365-token-{re.sub(r"[^A-Za-z0-9_-]", "_", account_name)}.json')


class GraphBackend:
    BASE = 'https://graph.microsoft.com/v1.0'

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.folder_ids: dict[str, str] = {}

    def connect(self):
        import msal
        import requests
        cache = msal.SerializableTokenCache()
        path = o365_token_cache_path(self.cfg['name'])
        if os.path.exists(path):
            with open(path) as fh:
                cache.deserialize(fh.read())
        else:
            cache.deserialize(secret(self.cfg['token_secret']))
        app = msal.PublicClientApplication(
            self.cfg['client_id'], authority=f"https://login.microsoftonline.com/{self.cfg['tenant']}",
            token_cache=cache)
        accounts = app.get_accounts()
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0]) if accounts else None
        if not result or 'access_token' not in result:
            raise RuntimeError('Office 365 sign-in has expired or is missing; run o365_login.py again '
                               f"and update the {self.cfg['token_secret']} secret")
        if cache.has_state_changed:
            os.makedirs(STATE_DIR, exist_ok=True)
            with open(path, 'w') as fh:
                fh.write(cache.serialize())
            os.chmod(path, 0o600)
        self.s = requests.Session()
        self.s.headers['Authorization'] = 'Bearer ' + result['access_token']
        me = self._req('GET', '/me', params={'$select': 'mail,userPrincipalName'})
        self.me = {a.lower() for a in [me.get('mail'), me.get('userPrincipalName'),
                                        *self.cfg.get('addresses', [])] if a}

    def close(self):
        pass

    def _req(self, method: str, path: str, **kw):
        url = path if path.startswith('http') else self.BASE + path
        for attempt in range(5):
            r = self.s.request(method, url, timeout=60, **kw)
            if r.status_code in (429, 503, 504):
                time.sleep(min(60, int(r.headers.get('Retry-After', 2 ** attempt))))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f'Graph {method} {path.split("?")[0]} → {r.status_code}: {r.text[:300]}')
            return r.json() if r.content else {}
        raise RuntimeError(f'Graph {method} {path} kept being throttled')

    def _pages(self, path: str, params: dict):
        data = self._req('GET', path, params=params)
        while True:
            yield from data.get('value', [])
            nxt = data.get('@odata.nextLink')
            if not nxt:
                return
            data = self._req('GET', nxt)

    def ensure_folders(self):
        existing = {f['displayName'].lower(): f['id'] for f in
                    self._pages('/me/mailFolders', {'$top': 250, '$select': 'id,displayName'})}
        for f in FOLDER_NAMES.values():
            name = self.cfg.get('folders', {}).get(f, f)
            if name.lower() not in existing:
                log.info('  creating folder %s', name)
                existing[name.lower()] = self._req('POST', '/me/mailFolders', json={'displayName': name})['id']
            self.folder_ids[f] = existing[name.lower()]

    def _fid(self, folder: str) -> str:
        return WELL_KNOWN.get(folder) or self.folder_ids[folder]

    def addresses(self) -> set:
        return self.me

    def list(self, folder: str, since: dt.datetime | None, headers: bool = False) -> list[Msg]:
        select = 'id,internetMessageId,from,subject,receivedDateTime,conversationId,flag,toRecipients,ccRecipients'
        if headers:
            select += ',internetMessageHeaders'
        params = {'$select': select, '$top': 100}
        if since is not None:
            params['$filter'] = f"receivedDateTime ge {since.astimezone(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')}"
        out = []
        for it in self._pages(f'/me/mailFolders/{self._fid(folder)}/messages', params):
            frm = (it.get('from') or {}).get('emailAddress') or {}
            recips = [(r.get('emailAddress') or {}).get('address', '').lower()
                      for r in (it.get('toRecipients') or []) + (it.get('ccRecipients') or [])]
            hdrs = {h['name'].lower(): h['value'] for h in it.get('internetMessageHeaders') or []}
            out.append(Msg(
                key=it.get('internetMessageId') or 'g:' + it['id'],
                sender=(frm.get('address') or '').lower(), name=frm.get('name') or '',
                subject=it.get('subject') or '', date=parse_iso(it['receivedDateTime']),
                folder=folder, ref=it['id'],
                thread_keys={it['conversationId']} if it.get('conversationId') else set(),
                recipients=[r for r in recips if r], headers=hdrs,
                flagged=(it.get('flag') or {}).get('flagStatus') == 'flagged'))
        return out

    def move(self, m: Msg, dest: str):
        self._req('POST', f'/me/messages/{m.ref}/move', json={'destinationId': self._fid(dest)})

    def mark_unread(self, m: Msg):
        self._req('PATCH', f'/me/messages/{m.ref}', json={'isRead': False})

    def deliver(self, subject: str, html_body: str, text_body: str):
        me = sorted(self.me)[0]
        self._req('POST', '/me/mailFolders/inbox/messages', json={
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'from': {'emailAddress': {'name': 'SaneMail', 'address': me}},
            'toRecipients': [{'emailAddress': {'address': me}}],
            # PR_MESSAGE_FLAGS = read, not a draft, so it shows as a normal received message.
            'singleValueExtendedProperties': [{'id': 'Integer 0x0E07', 'value': '1'}],
        })


BACKENDS = {'imap': ImapBackend, 'o365': GraphBackend}


# ───────────────────────────── Rules (rules.json) ─────────────────────────────

RANK = {'sent': 1, 'replied': 1, 'trained': 2}


class Rules:
    """Sender → folder. Learned automatically, editable by hand. Keys: an address or '@domain'."""

    def __init__(self, path: str):
        self.path = path
        try:
            with open(path) as fh:
                self.map = json.load(fh)
        except FileNotFoundError:
            self.map = {}
        self.dirty = False

    def lookup(self, sender: str):
        if sender in self.map:
            return self.map[sender]['folder']
        domain = sender.rpartition('@')[2]
        parts = domain.split('.')
        for i in range(len(parts) - 1):
            r = self.map.get('@' + '.'.join(parts[i:]))
            if r:
                return r['folder']
        return None

    def learn(self, sender: str, folder: str, source: str) -> bool:
        if not sender or '@' not in sender:
            return False
        cur = self.map.get(sender)
        if cur and RANK.get(cur.get('source'), 2) > RANK[source]:
            return False
        if cur and cur.get('folder') == folder and (cur.get('source') == source or source != 'trained'):
            return False
        self.map[sender] = {'folder': folder, 'source': source, 'updated': now_utc().strftime('%Y-%m-%d')}
        self.dirty = True
        if DEBUG:
            log.info('    rule: %s → %s (%s)', sender, folder, source)
        return True

    def save(self):
        if self.dirty:
            with open(self.path, 'w') as fh:
                json.dump(dict(sorted(self.map.items())), fh, indent=1)
                fh.write('\n')


# ───────────────────────────── Sorting logic ─────────────────────────────

def looks_bulk(m: Msg) -> bool:
    h = m.headers
    local = m.sender.split('@')[0]
    if BULK_LOCALPART.match(local):
        return True
    if h.get('list-unsubscribe') or h.get('list-id'):
        return True
    if re.search(r'bulk|list|junk', h.get('precedence', ''), re.I):
        return True
    if h.get('auto-submitted', 'no').lower() not in ('', 'no'):
        return True
    if BULK_MAILER.search(h.get('x-mailer', '')):
        return True
    return False


def classify(m: Msg, rules: Rules, me: set, sent_keys: dict, settings: dict) -> str:
    # 1. Replies to conversations you started or took part in.
    if m.thread_keys & sent_keys.keys():
        rules.learn(m.sender, 'inbox', 'replied')
        return 'inbox'
    # 2. Learned or hand-written rule.
    rule = rules.lookup(m.sender)
    if rule:
        return rule
    # 3. VIPs and your own (non-public) domains.
    domain = m.sender.rpartition('@')[2]
    vip_senders = {s.lower() for s in settings.get('vip_senders', [])}
    vip_domains = {d.lower().lstrip('@') for d in settings.get('vip_domains', [])}
    vip_domains |= {a.rpartition('@')[2] for a in me} - PUBLIC_DOMAINS
    if m.sender in vip_senders or any(domain == d or domain.endswith('.' + d) for d in vip_domains):
        return 'inbox'
    # 4. Newsletters / bulk vs. unknown people.
    if looks_bulk(m):
        return 'later' if settings.get('news_into_later') else 'news'
    return 'later'


def wake_time(kind: str, tz: ZoneInfo, hour: int) -> dt.datetime:
    local = now_utc().astimezone(tz)
    days = 1 if kind == 'tomorrow' else ((7 - local.weekday()) % 7 or 7)  # next Monday
    target = (local + dt.timedelta(days=days)).replace(hour=hour, minute=0, second=0, microsecond=0)
    return target.astimezone(UTC)


def process_account(b, acct: dict, st: dict, rules: Rules, settings: dict, stats: dict, digest: list):
    now = now_utc()
    tz = ZoneInfo(settings.get('timezone', 'UTC'))
    window = now - dt.timedelta(days=WINDOW_DAYS)
    keep_after = now - dt.timedelta(days=WINDOW_DAYS - 2)
    first = not st.get('initialized')
    counts = {'trained': 0, 'woke': 0, 'snoozed': 0, 'inbox': 0, 'later': 0, 'news': 0, 'blackhole': 0}

    b.connect()
    b.ensure_folders()  # creating empty folders is harmless, even in a dry run
    me = set(b.addresses()) | set(st.get('me', []))

    def move(m: Msg, dest: str):
        if DRY_RUN:
            log.info('    [dry run] %s → %s', m.sender if DEBUG else '…', dest)
        else:
            b.move(m, dest)

    # 1. Learn who you write to (people you email are important).
    learn_days = int(settings.get('learn_from_sent_days', 365)) if first else 2
    sent_keys: dict = st.get('sent_keys', {})
    sent = b.list('sent', now - dt.timedelta(days=learn_days))
    me |= {m.sender for m in sent if m.sender}  # every address you send from (aliases, custom domains)
    for m in sent:
        for r in m.recipients:
            if r not in me:
                rules.learn(r, 'inbox', 'sent')
        if m.date >= window:
            for k in m.thread_keys:
                sent_keys[k] = iso(m.date)

    inbox = b.list('inbox', window, headers=True)
    in_inbox = {m.key: m for m in inbox}
    seen: dict = st.get('seen', {})
    filed: dict = st.get('filed', {'later': {}, 'news': {}})
    snoozes: dict = st.get('snoozes', {})
    if first:
        lookback = now - dt.timedelta(days=float(settings.get('initial_lookback_days', 1)))
        seen = {m.key: iso(m.date) for m in inbox if m.date < lookback}

    # 2. Training: what you moved into or out of SaneLater / SaneNews since last time.
    for f in ('later', 'news'):
        current = b.list(FOLDER_NAMES[f], window)
        cur_keys = {m.key for m in current}
        if first:
            filed[f] = {m.key: iso(m.date) for m in current}
            continue
        for m in current:
            if m.key not in filed[f]:
                filed[f][m.key] = iso(m.date)
                counts['trained'] += rules.learn(m.sender, f, 'trained')
        for k in list(filed[f]):
            if k in cur_keys:
                continue
            if parse_iso(filed[f][k]) >= keep_after and k in in_inbox:
                counts['trained'] += rules.learn(in_inbox[k].sender, 'inbox', 'trained')
                seen[k] = filed[f][k]
            del filed[f][k]

    # 3. Black hole: anything here → sender blocked, message to Trash.
    for m in b.list(FOLDER_NAMES['blackhole'], None):
        counts['trained'] += rules.learn(m.sender, 'blackhole', 'trained')
        move(m, 'trash')

    # 4. Snoozes.
    live = set()
    for kind in ('tomorrow', 'nextweek'):
        for m in b.list(FOLDER_NAMES[kind], None):
            live.add(m.key)
            if m.key not in snoozes:
                snoozes[m.key] = iso(wake_time(kind, tz, int(settings.get('wake_hour', 7))))
                counts['snoozed'] += 1
            elif parse_iso(snoozes[m.key]) <= now:
                if not DRY_RUN:
                    b.mark_unread(m)
                move(m, 'inbox')
                seen[m.key] = iso(m.date)
                del snoozes[m.key]
                live.discard(m.key)
                counts['woke'] += 1
    snoozes = {k: v for k, v in snoozes.items() if k in live}

    # 5. Sort new Inbox mail.
    for m in sorted(inbox, key=lambda x: x.date):
        if m.key in seen:
            continue
        seen[m.key] = iso(m.date)
        if m.flagged or m.sender in me or not m.sender:
            continue
        folder = classify(m, rules, me, sent_keys, settings)
        counts[folder] += 1
        stats[folder] = stats.get(folder, 0) + 1
        if folder == 'inbox':
            continue
        move(m, 'trash' if folder == 'blackhole' else FOLDER_NAMES[folder])
        if folder in filed:
            filed[folder][m.key] = iso(m.date)
        if folder != 'blackhole':
            digest.append({'account': acct['name'], 'folder': folder, 'name': m.name or m.sender,
                           'sender': m.sender, 'subject': m.subject[:200], 'date': iso(m.date)})

    # 6. Prune and store.
    cut = iso(window)
    st.update({
        'initialized': True,
        'seen': {k: v for k, v in seen.items() if v >= cut},
        'filed': {f: {k: v for k, v in d.items() if v >= cut} for f, d in filed.items()},
        'sent_keys': {k: v for k, v in sent_keys.items() if v >= cut},
        'snoozes': snoozes,
        'last_run': iso(now),
        'me': sorted(me),
    })
    log.info('  %s: %s', acct['name'], ', '.join(f'{k} {v}' for k, v in counts.items() if v) or 'nothing new')


# ───────────────────────────── Digest ─────────────────────────────

def build_digest(items: list, stats: dict, today: str) -> tuple[str, str, str]:
    label = {'later': 'SaneLater', 'news': 'SaneNews'}
    week = [stats.get(d, {}) for d in sorted(stats)[-7:]]
    tot = {k: sum(d.get(k, 0) for d in week) for k in ('inbox', 'later', 'news', 'blackhole')}
    total = sum(tot.values()) or 1
    kept_out = round((tot['later'] + tot['news'] + tot['blackhole']) / total * 100)

    parts_html, parts_text = [], []
    for f in ('later', 'news'):
        rows = [i for i in items if i['folder'] == f]
        if not rows:
            continue
        parts_text.append(f"\n{label[f]} ({len(rows)})")
        trs = []
        for i in rows:
            parts_text.append(f"  {i['name']} — {i['subject']}  [{i['account']}]")
            trs.append(
                '<tr><td style="padding:6px 8px;border-bottom:1px solid #e5e5e5;vertical-align:top">'
                f'<b>{html.escape(i["name"])}</b><br><span style="color:#666">{html.escape(i["subject"] or "(no subject)")}</span></td>'
                f'<td style="padding:6px 8px;border-bottom:1px solid #e5e5e5;color:#888;font-size:12px;white-space:nowrap;vertical-align:top">{html.escape(i["account"])}</td></tr>')
        parts_html.append(f'<h3 style="margin:22px 0 6px">{label[f]} ({len(rows)})</h3>'
                          f'<table style="border-collapse:collapse;width:100%">{"".join(trs)}</table>')

    summary = (f"Last 7 days: {tot['inbox']} kept in Inbox, {tot['later']} to SaneLater, "
               f"{tot['news']} to SaneNews, {tot['blackhole']} black-holed — {kept_out}% kept out of your Inbox.")
    howto = ('To teach SaneMail, move a message into SaneLater, SaneNews or SaneBlackHole — '
             'or move it from those folders back to your Inbox. Senders are remembered on every account.')
    body_html = ('<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:720px;color:#222">'
                 f'<p style="color:#555">{summary}</p>{"".join(parts_html)}'
                 f'<p style="color:#999;font-size:12px;margin-top:24px">{howto}</p></div>')
    body_text = summary + '\n' + '\n'.join(parts_text) + '\n\n' + howto
    subject = f'SaneMail digest — {len(items)} filed — {today}'
    return subject, body_html, body_text


# ───────────────────────────── Main ─────────────────────────────

def load_json(path: str, default):
    try:
        with open(path) as fh:
            return json.load(fh)
    except (FileNotFoundError, ValueError):
        return default


def main() -> int:
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    with open(CONFIG_PATH) as fh:
        cfg = yaml.safe_load(fh)
    settings = cfg.get('settings') or {}
    tz = ZoneInfo(settings.get('timezone', 'UTC'))
    os.makedirs(STATE_DIR, exist_ok=True)
    state_path = os.path.join(STATE_DIR, 'state.json')
    state = load_json(state_path, {})
    rules = Rules(RULES_PATH)
    today = now_utc().astimezone(tz).strftime('%Y-%m-%d')
    stats_all = state.setdefault('stats', {})
    stats = stats_all.setdefault(today, {})
    digest = state.setdefault('digest', [])
    log.info('SaneMail%s — %s', ' (dry run)' if DRY_RUN else '', now_utc().astimezone(tz).strftime('%Y-%m-%d %H:%M %Z'))

    backends, failures = {}, 0
    for acct in cfg['accounts']:
        b = BACKENDS[acct['type']](acct)
        backends[acct['name']] = b
        try:
            process_account(b, acct, state.setdefault('accounts', {}).setdefault(acct['name'], {}),
                            rules, settings, stats, digest)
        except Exception as e:
            failures += 1
            log.error('  %s: FAILED — %s: %s', acct['name'], type(e).__name__, e)
            if DEBUG:
                raise

    # Daily digest.
    local = now_utc().astimezone(tz)
    if (settings.get('digest', True) and digest and state.get('digest_sent') != today
            and local.hour >= int(settings.get('digest_hour', 7))):
        target = settings.get('digest_account') or cfg['accounts'][0]['name']
        subject, body_html, body_text = build_digest(digest, stats_all, today)
        try:
            if DRY_RUN:
                log.info('  [dry run] would deliver digest with %d items to %s', len(digest), target)
            else:
                backends[target].deliver(subject, body_html, body_text)
                log.info('  digest delivered to %s (%d items)', target, len(digest))
                state['digest'] = []
                state['digest_sent'] = today
        except Exception as e:
            failures += 1
            log.error('  digest: FAILED — %s: %s', type(e).__name__, e)

    for b in backends.values():
        b.close()

    if DRY_RUN:
        log.info('Dry run: nothing moved, nothing saved.')
        return 0
    state['digest'] = state.get('digest', [])[-500:]
    for d in sorted(stats_all)[:-60]:
        del stats_all[d]
    tmp = state_path + '.tmp'
    with open(tmp, 'w') as fh:
        json.dump(state, fh)
    os.replace(tmp, state_path)
    rules.save()
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
