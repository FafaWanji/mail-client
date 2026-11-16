# One-time Microsoft setup for Outlook accounts

Microsoft only lets an app read mail if the app is registered. You register your own private copy.
Steps and labels checked against Microsoft's documentation on 2026-11-14. Menu names can change.

## Before you start

Microsoft requires an **Azure account with an active subscription** to register an app.
A free Azure account is enough, and registering an app costs nothing, but the sign-up at
<https://azure.microsoft.com/free> asks for a phone number and a payment card for identity checks.
If you do not want that, stop here: there is no way around it for an app of your own.

The app is for personal Outlook accounts (outlook.com, hotmail.com, live.com). A work or school
account can refuse the sign-in until its administrator approves the app.

## Steps

1. Open <https://entra.microsoft.com> and sign in with the account that has the Azure subscription.
2. **Entra ID > App registrations > New registration**.
3. Name: `Mail Client`.
4. Supported account types: **Personal accounts only** (or **Any Entra ID Tenant + Personal Microsoft accounts**; both work).
5. **Register**.
6. On the Overview page copy the **Application (client) ID**.
7. Left menu, under Manage: **Authentication > Add Redirect URI**.
   - Pick the **Single-page application** tile.
   - Redirect URI, exactly: `http://localhost:8080/auth.html`
   - **Configure**.
8. Open `config.js` in this folder and paste the ID into the second line:

   ```js
   export const MICROSOFT_CLIENT_ID = '00000000-0000-0000-0000-000000000000';
   ```

9. Reload the app. Click **Add an Outlook account** (or **+ Outlook** in the sidebar).
   A small window opens. Sign in and accept the permissions:
   - read and write your mail (move, mark, trash),
   - read and write mailbox settings (inbox rules),
   - read your profile (to show your address),
   - send mail as you (only when you press Send or confirm an unsubscribe mail).

No client secret is needed. Do not create one.

## Good to know

- The browser must allow popups for `http://localhost:8080`.
- You sign in again each time you open the app. Nothing is stored on disk.
- Error `AADSTS50011` (redirect URI mismatch): step 7 has a typo, or the tile was "Web" instead of
  "Single-page application".
- Error `AADSTS9002326`: same cause, the URI is registered under "Web". Remove it there and add it
  under "Single-page application".
- Error `unauthorized_client` for a personal account: step 4 was set to a tenant-only option.
  Change it under **Authentication > Supported accounts**.

## Differences to Gmail

- Folders are real folders: a mail is in one folder only.
- All Outlook rule options work. The two Gmail-only actions (Star, Never spam) are greyed out.
- "Apply to existing mail" is only offered for simple rules (from, to, subject with move, trash or
  mark read). Outlook has no function to run a saved rule on old mail.
- Top-level folders only. Built-in folders such as Drafts and Sent Items also show in the list.
