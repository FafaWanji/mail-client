# Mail Client

A small mail cleaner for Gmail and Outlook that runs on your own computer. Built around `+alias` addresses
(`name+shop@gmail.com`): see which alias a sender writes to, spot leaked aliases, and clean up fast.

- Senders list with mail count and size; trash, block, move or unsubscribe many senders at once
- Cleanup by age, size or "never opened"; empty Trash and Spam
- Rules saved in your mail account (Gmail filters, Outlook inbox rules), editable and pausable
- Undo after bulk actions, several accounts, one combined inbox
- Read, reply and write plain-text mail

Nothing is sent anywhere except to Google or Microsoft. Sign-in tokens stay in the browser's memory and are never written to disk.

## Tutorial

### 1. Install Python

The app is plain web files. Python only serves them on your own computer.

- Windows: install Python 3 from <https://www.python.org/downloads/>.
  On the first installer screen, tick **Add python.exe to PATH**.
- macOS and Linux: Python 3 is usually there already. Check with `python3 --version`.

### 2. Download the app

On this page click **Code > Download ZIP** and unzip it anywhere. Or run
`git clone https://github.com/FafaWanji/mail-client.git`.

### 3. Start it

- Windows: double-click `Start.bat`. A black window opens, then the browser opens
  <http://localhost:8080>. Keep the black window open while you use the app.
- macOS and Linux: open a terminal in the folder, run `python3 serve.py`, then open
  <http://localhost:8080>.

To stop the app, close the black window (or press Ctrl+C in the terminal).

### 4. Try the demo

Click **Open the demo with sample mail**. It needs no account and changes no real mail.

### 5. Connect your own mail

Google and Microsoft only let a registered app read mail. Each person registers their own
private copy once. It is free and takes about 10 minutes:

- Gmail: follow [SETUP.md](SETUP.md), then paste your client ID into `config.js`.
- Outlook (personal accounts): follow [SETUP-OUTLOOK.md](SETUP-OUTLOOK.md), then paste your
  client ID into `config.js`.

Reload the page and click **Add a Gmail account** or **Add an Outlook account**.

### Problems

- `'python' is not recognized`: Python is not on PATH. Reinstall it with **Add python.exe to PATH**
  ticked, or edit `Start.bat` and replace `python` with `py`.
- `Address already in use` / error 10048: the app already runs in another window. Close that window.
- Only the demo link shows: `config.js` has no client ID yet, or it has a typo. Check the quotes.

## Tests

`npm test` (Node 20 or newer, no dependencies).

## License

MIT, see [LICENSE](LICENSE). No warranty: bulk actions change your real mailbox, try the demo first.
