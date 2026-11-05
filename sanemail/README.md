# SaneMail

SaneMail does the same job as SaneBox, but you run it yourself. It works with:

- **your own-domain mail** (randlefamily.com) through IMAP: iCloud+ custom domains, Fastmail, cPanel hosts and other IMAP providers
- **Office 365 / Outlook** through Microsoft Graph

Sorting happens **on the mail server**, so the folders look the same in Apple Mail on every device, in Outlook, and in webmail. It runs every 30 minutes on GitHub Actions in a **private** repository. There is no server to look after and no third party sees your mail.

| Folder | What goes there |
|---|---|
| **Inbox** | People you've emailed, replies to your threads, your own domain, VIPs, flagged mail |
| **SaneLater** | Senders you've never written to |
| **SaneNews** | Newsletters and bulk mail, detected from List-Unsubscribe, bulk mailer and no-reply addresses |
| **SaneBlackHole** | Drop one message here and that sender goes straight to Trash from then on |
| **SaneTomorrow / SaneNextWeek** | Snooze: the message comes back to the Inbox unread at 7am |

**Training.** To teach SaneMail, drag a message into SaneLater, SaneNews or SaneBlackHole, or drag it from those folders back to the Inbox. SaneMail remembers the sender in `rules.json`. Those rules apply to **all** your accounts, and you can edit the file on GitHub. You can also add whole domains, for example `"@amazon.com": {"folder": "news"}`.

**Daily digest.** Each morning, a "SaneMail digest" message in your Inbox lists what was filed, with weekly totals.

---

## Setup

### 1. Create the private repository
1. On GitHub, create a **private** repository, for example `sanemail`.
2. Put the contents of this folder at the top level of that repository. The workflow must end up at `.github/workflows/sanemail.yml`.

### 2. Own-domain mail (Apple Mail)
First find your server. On a Mac, open **Mail → Settings → Accounts → (your randlefamily.com account) → Server Settings** and note the **Incoming Mail Server**.

- **iCloud+ custom domain** (server `imap.mail.me.com`):
  1. Go to account.apple.com → Sign-In and Security → **App-Specific Passwords** and create one called `SaneMail`.
  2. The username is your **iCloud email address**, not the custom-domain address.
  3. Add your `@randlefamily.com` addresses under `addresses:`.
- **Any other host:** use that host's IMAP server and your mailbox password, or an app password if the host offers them.

Edit the `home` account in `config.yaml`. Then in the repository go to **Settings → Secrets and variables → Actions → New repository secret** and add the secret `HOME_IMAP_PASSWORD`.

### 3. Office 365 (needs one-time IT approval)
Send your IT team this request:

> Please create an app registration in Entra ID for a personal mail-sorting tool:
> - Name: **SaneMail**. Supported account types: *this organization only*
> - Authentication → **Allow public client flows: Yes** (it signs in with the device-code flow; there is no client secret)
> - API permissions (**delegated** Microsoft Graph): `Mail.ReadWrite`, `User.Read`, `offline_access`. Grant admin consent if users can't consent themselves.
> - Send me the **Application (client) ID** and the **Directory (tenant) ID**.
>
> It only acts as me, on my own mailbox (delegated, not application permissions). It moves messages between folders and creates a few folders (SaneLater, SaneNews, SaneBlackHole, SaneTomorrow, SaneNextWeek). It never sends email.

When IT replies:
1. Put the `client_id` and `tenant` into the `work` account in `config.yaml`.
2. On your own computer, in this folder, run:
   ```
   pip install -r requirements.txt
   python o365_login.py work
   ```
   Sign in with the code it shows.
3. Copy the contents of the file it creates into a repository secret named `O365_TOKEN_CACHE`, then delete the file.

If you don't have an Office 365 account (or haven't got approval yet), delete the `work` block from `config.yaml` for now.

### 4. Set your timezone and do a test run
1. Set `timezone:` in `config.yaml`, for example `America/New_York`.
2. Go to **Actions → SaneMail → Run workflow** and tick **Dry run**. The log shows how many messages would go where, and nothing is moved.
3. Run it again with Dry run unticked. After that it runs by itself every 30 minutes.

The first real run:
- learns your contacts from the last year of Sent mail
- sorts only the last day of your Inbox; older mail is left alone

---

## Good to know
- **Flagged/starred mail is never moved.** Neither is mail you sent yourself.
- **State.** Message tracking and the Office 365 sign-in are kept in the GitHub Actions cache. Sender rules are committed to `rules.json`. If the cache is ever lost, SaneMail starts tracking afresh. It won't re-sort your existing Inbox and you don't lose your training.
- **Failures.** If a run fails, for example after a password change or an expired Office 365 sign-in, GitHub emails you and the log says what to fix.
- **Cost.** 48 runs a day use about 1,500 of the 2,000 free GitHub Actions minutes a month for private repositories. To use fewer, change the schedule in the workflow, for example `'7 * * * *'` for hourly.
- **Pausing.** To pause, go to **Actions → SaneMail → ⋯ → Disable workflow**.
- **Running locally.** You can also run it on a Mac: `pip install -r requirements.txt`, set `HOME_IMAP_PASSWORD` in the environment, then run `python sanemail.py`. Add `SANEMAIL_DRY_RUN=1` to only preview.
