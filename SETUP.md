# One-time Google setup (about 10 minutes)

Google only lets an app read mail if the app is registered. You register your own private copy.
It is free. Nobody else can use it. Google's menu names change now and then; if a name differs
slightly, pick the closest one.

Try the app first without any setup: run `Start.bat` and click "Open the demo with sample mail".

## Steps

1. Open <https://console.cloud.google.com/> and sign in with the Gmail account you want to use.
2. Top bar: project picker, **New project**. Name: `Mail Client`. **Create**, then select it.
3. Menu **APIs & Services > Library**. Search `Gmail API`, open it, **Enable**.
4. Menu **Google Auth Platform** (older name: "OAuth consent screen"), **Get started**.
   - App name: `Mail Client`. Support email: your address.
   - Audience: **External**.
   - Contact email: your address. Accept the policy, **Create**.
5. **Audience > Test users > Add users**: add your own Gmail address. Save.
   Only test users can sign in. Leave the app in "Testing".
6. **Data Access > Add or remove scopes**. In "Manually add scopes" paste both lines, **Add to table**, **Update**, **Save**:

   ```
   https://www.googleapis.com/auth/gmail.modify
   https://www.googleapis.com/auth/gmail.settings.basic
   ```

   The first one lets the app read, label and trash mail, and send mail (only when you press Send
   or unsubscribe from a sender by mail). The second one lets it manage filters.

   For **Empty for good** (Trash and Spam) also add this line. Google asks for it only when you press that button:

   ```
   https://mail.google.com/
   ```
7. **Clients > Create client**.
   - Application type: **Web application**.
   - Authorized JavaScript origins: **Add URI**, enter exactly `http://localhost:8080`
   - Leave "Authorized redirect URIs" empty. **Create**.
8. Copy the **Client ID** (ends with `.apps.googleusercontent.com`).
9. Open `config.js` in this folder and paste it between the quotes:

   ```js
   export const GOOGLE_CLIENT_ID = '1234-abc.apps.googleusercontent.com';
   ```

10. Run `Start.bat`. The browser opens `http://localhost:8080`. Click **Add a Gmail account**.
    - Google warns "Google hasn't verified this app". That is your own app: **Continue**.
    - Tick both permissions, **Continue**.

## Good to know

- Keep the black `Start.bat` window open while you use the app. Close it to stop.
- You sign in again each time you open the app, and after about an hour. Nothing is stored on disk.
- A new origin can take a few minutes before Google accepts it. Error `origin_mismatch` or
  `redirect_uri_mismatch` right after setup: wait five minutes, check step 7 for typos.
- Error `access_denied`: your address is missing in step 5.
- Forward rules only work to an address you already verified in Gmail under
  Settings > Forwarding and POP/IMAP.

## More Gmail accounts

One Google setup covers all your Gmail accounts. For each extra account:

1. In the Google console: **Google Auth Platform > Audience > Test users > Add users**, add that address.
2. In the app: **+ Gmail** in the sidebar, pick that account, tick both permissions.

The account list at the top of the sidebar switches between them. Each account has its own
folders, rules, alias scan and leak view. For Outlook accounts see `SETUP-OUTLOOK.md`.
