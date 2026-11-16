# Mail Client

A small mail cleaner for Gmail and Outlook that runs on your own computer. Built around `+alias` addresses
(`name+shop@gmail.com`): see which alias a sender writes to, spot leaked aliases, and clean up fast.

- Senders list with mail count and size; trash, block, move or unsubscribe many senders at once
- Cleanup by age, size or "never opened"; empty Trash and Spam
- Rules saved in your mail account (Gmail filters, Outlook inbox rules), editable and pausable
- Undo after bulk actions, several accounts, one combined inbox
- Read, reply and write plain-text mail

Nothing is sent anywhere except to Google or Microsoft. Sign-in tokens stay in the browser's memory and are never written to disk.

## Run it

You need [Python 3](https://www.python.org/downloads/) (only to serve the files) and a browser.

1. Download: **Code > Download ZIP** on this page, then unzip. Or `git clone` this repository.
2. Try it without any account: run `Start.bat` (Windows) or `python serve.py`, open <http://localhost:8080>,
   click **Open the demo with sample mail**.
3. For your own mail, register your own private copy once and paste its ID into `config.js`:
   - Gmail: [SETUP.md](SETUP.md)
   - Outlook (personal accounts): [SETUP-OUTLOOK.md](SETUP-OUTLOOK.md)

Each person registers their own copy. That is free; Google and Microsoft require it for apps that read mail.

## Tests

`npm test` (Node 20 or newer, no dependencies).

## License

MIT, see [LICENSE](LICENSE). No warranty: bulk actions change your real mailbox, try the demo first.
