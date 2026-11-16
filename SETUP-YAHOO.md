# Yahoo accounts

Yahoo needs no app registration. You make one **app password** and paste it when you add the account.

## Steps (once)

1. Open <https://login.yahoo.com/account/security> and sign in.
2. **Generate app password** (sometimes under "Other ways to sign in"). Name it `Mail Client`, **Generate**.
3. Copy the password Yahoo shows. Spaces in it do not matter.
4. In the app: **Add a Yahoo account** (or **+ Yahoo** in the sidebar). Enter your Yahoo address and that password.

After an update of the app, close the black `Start.bat` window and start it again: Yahoo runs through the local server.

## Good to know

- The password stays in memory only (in the page and in the `Start.bat` window), never on disk. You enter it again each time you open the app.
- Yahoo gives apps no access to its filters. Rules, "Trash + block", "Move to folder + rule" and "Rule from this mail" are hidden for Yahoo.
  "Trash all" for senders still works on mail already there. For new mail, make a filter in Yahoo Mail under **Settings > More settings > Filters**.
- Yahoo has no `name+shop@yahoo.com` addresses, so the alias list and the leak view stay empty for Yahoo.
- Search text must be plain letters (no umlauts or emoji).
- To stop the app's access, delete the app password on the same Yahoo page.
