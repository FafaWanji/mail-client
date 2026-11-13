# Mail client: design

Date: 2026-11-13. Status: waiting for approval.

## Intent

A personal mail client for Gmail and Outlook accounts, built around `+alias` addresses
(`name+spotify@gmail.com`). The owner gives each service its own alias, then wants to:

- sort mail by alias into self-made folders ("Subscriptions" = `+spotify`, `+uber`, `+netflix`),
- keep that mail out of the inbox,
- see when an alias receives mail from a sender it was never given to (a leak),
- turn any of these into a rule with a few clicks,
- delete and clean up mail quickly.

Gmail's and Outlook's own filter screens are too clumsy for this.

## Goals

1. Read, search, open, trash, mark read/unread, and move mail.
2. Alias sidebar: every `+alias` found in the mailbox, with mail count. Click shows its mail.
3. Folders: self-made groups holding one or more aliases. Mail in a folder skips the inbox.
4. Leak view: per alias, the sender domains that wrote to it. An alias with more than one
   domain is flagged until the owner marks the extra domains as trusted.
5. Rule builder: condition + exception + action, saved as a real server-side rule
   (Gmail filter, Outlook inbox rule), so it runs on the phone and while the app is closed.
6. "Apply to existing mail" per rule, because server rules only act on new mail.
7. Several accounts: Gmail first, Outlook second, one account switcher.

## Non-goals

- Composing and sending mail (not asked for; add later if wanted).
- A combined inbox across accounts.
- Mobile layout, offline use, notifications.
- Work/school Outlook accounts that need admin consent.

## Anti-goals

- Storing passwords or long-lived tokens on disk.
- Rules that only work while the app is open.
- Permanent deletion without a confirmation that names what is deleted.

## Approaches

| | What | Trade-off |
|---|---|---|
| **A (chosen)** | Local web page in the browser, started by `Start.bat`, calls the Gmail API and Microsoft Graph directly | Least code, no stored token. Sign-in click once per session. |
| B | Desktop app with own window | Stays signed in. About three times the code, installer, token on disk to protect. |
| C | Add-on inside Gmail's page | Not an own client, breaks when Google changes the page, no Outlook. |

A was the assistant's recommendation; the owner did not object. Choice is the assistant's call.

## Decided by the owner

- Aliases are used for both sorting and leak detection.
- Groups are folders. Mail in them skips the inbox and shows only in the folder and in All Mail.
- These behaviours must be settable as rules.
- Outlook accounts must be supported.
- The rule builder offers every rule possibility Outlook has.

## Rule builder

The builder shows the full Outlook rule set as defined by Microsoft Graph `messageRule`:
all `messageRulePredicates` as conditions and as exceptions, all `messageRuleActions` as actions.
The builder reads the current Graph documentation and implements the whole list, not a subset.

On a Gmail account the same builder is shown, but Gmail filters can do less. Options Gmail
cannot run server-side are greyed out with a one-line reason. Gmail supports roughly:

- Conditions: from, to, subject, contains words (covers body), does not contain,
  has attachment, size larger/smaller than.
- Actions: move to folder (label + skip inbox), copy to folder (label, keep in inbox),
  mark read, star, mark important, trash, never spam, forward to a verified address.

Not available on Gmail: importance, sensitivity, meeting/receipt/auto-reply type checks,
header contains, categories, permanent delete, redirect, stop processing more rules.
(Assistant's call: grey out rather than emulate in the app, because an emulated rule
would only run while the app is open.)

Forward, redirect and permanent-delete actions ask for confirmation before saving.

## How it maps

| App | Gmail | Outlook |
|---|---|---|
| Folder | label; rule adds label and removes `INBOX` | mail folder; rule moves the message |
| Rule | `users.settings.filters` | `mailFolders/inbox/messageRules` |
| Trash | add `TRASH` label (30 days recoverable) | move to Deleted Items |
| Alias source | `Delivered-To` / `To` headers | `toRecipients` |

## Constraints

- One-time setup by the owner: Google Cloud project with OAuth client ID (own address added
  as test user), and a Microsoft app registration. Step-by-step instructions are part of the build.
- Gmail limit: 1,000 filters per account.
- Sign-in uses the browser token flow from Google Identity Services and MSAL.js. Tokens live
  in memory only.
- Trusted-sender marks are stored in the browser's `localStorage`; clearing browser data loses them.
- Lives in `Desktop\MailClient`, separate from the Enshrouded mods.

## Left to the builder

- Layout and styling. Three panes (sidebar, list, message) is the default.
- How `Start.bat` serves the page: use what is installed, add no new runtime if avoidable.
- How far back the alias scan reads, and how it is cached.
- Exact OAuth scopes: the narrowest that allow reading, labelling, trashing and managing rules.
- Rendering HTML mail safely: sandboxed frame, remote images blocked by default.

## Order

1. Gmail: sign-in, mail list, read, trash.
2. Alias sidebar and folders.
3. Rule builder on Gmail, with apply-to-existing.
4. Leak view.
5. Outlook account on the same screens, full rule set.
