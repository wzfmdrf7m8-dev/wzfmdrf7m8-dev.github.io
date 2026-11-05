# SaneMail: a SaneBox-style filter you host yourself

SaneMail is a Google Apps Script that runs inside your own Gmail or Google Workspace account. It works like SaneBox, without a third party reading your mail.

| SaneBox | SaneMail |
|---|---|
| Important mail stays in the Inbox | Mail from people you have written to or replied to, your own domain, and VIPs stays in the Inbox |
| SaneLater | `@SaneLater`: unknown senders, archived out of the Inbox |
| SaneNews | `@SaneNews`: newsletters and bulk mail, detected from List-Unsubscribe, Precedence and similar headers |
| SaneBlackHole | `@SaneBlackHole`: put one thread here and that sender's mail goes to Trash from then on |
| Training | Move a thread into or out of a Sane folder and SaneMail remembers the sender |
| SaneTomorrow / SaneNextWeek | `@SaneTomorrow` and `@SaneNextWeek`: the thread is hidden, then comes back unread at 7am |
| SaneNoReplies | `@SaneNoReplies`: mail you sent that has had no reply for 3 days |
| Daily digest | Morning email listing what was filed, with **→ Inbox** and **Black hole** links |
| Dashboard | A web page with daily counts, the percentage of mail kept out of the Inbox, and your sender rules |

## Install (about 5 minutes per inbox)

1. Go to **https://script.google.com** while signed in to the inbox you want to filter, and click **New project**. Name it `SaneMail`.
2. Delete the sample code, paste in all of `Code.gs`, and save.
3. Optional: change `CONFIG` at the top of the file, for example by adding `VIP_DOMAINS` such as your employer or your kids' school. For a first test you can set `DRY_RUN: true`; SaneMail will then only log what it would do (see **Executions**).
4. Choose **setup** in the function dropdown and click **Run**. Approve the permissions. Google will show "unverified app". This is expected because it is your own script: click **Advanced → Go to SaneMail**.
   setup does the following:
   - creates the labels
   - creates a **SaneMail rules** Google Sheet in your Drive
   - learns your contacts from your last ~1,500 sent threads
   - sorts the last 2 days of your Inbox
   - installs triggers to run every 10 minutes and to send the digest at 7am
5. Optional dashboard and one-click digest links: click **Deploy → New deployment → Web app**, set *Execute as: **Me*** and *Who has access: **Only myself***, then click Deploy. Bookmark the URL. Digest emails will now include action links.

**Several inboxes:** repeat these steps once in each Google account. Each one keeps its own rules.

## Using it

- **Teach it:** in Gmail, use *Move to* `@SaneLater`, `@SaneNews` or `@SaneBlackHole`. To rescue a sender, open the thread in the label and click *Move to Inbox*. Training takes effect on the next run (within 10 minutes).
- **Edit rules directly** in the *SaneMail rules* sheet. Set `folder` to `inbox`, `later`, `news` or `blackhole` and `source` to `trained`.
- **Snooze:** apply `@SaneTomorrow` or `@SaneNextWeek`.
- **Starred threads** are never moved.
- The `_sane/…` labels are internal markers. In Gmail settings → Labels you can set them to *Hide* in the label list. Don't delete them.
- **Stop it:** run `uninstall()`. This removes the triggers and leaves your labels and rules in place.

## Notes and limits

- This version supports Gmail and Google Workspace only. Outlook, iCloud and generic IMAP accounts need a different approach.
- It uses Google's free Apps Script quotas, which are well above what a personal inbox needs.
- The first sort of a sender costs one Gmail search. After that the result is cached in the sheet.
