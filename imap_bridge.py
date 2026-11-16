"""Bridge from the page to Yahoo Mail's IMAP and SMTP servers. A browser cannot open these connections itself.

The page sends {user, password, op, ...} as JSON on every call. The password stays in this process's memory
(inside the open IMAP connection) and is never written anywhere. Hosts are fixed here, never taken from a request.
Message ids are "uid:uidvalidity:folder", so a cached id from before a folder reset misses instead of hitting other mail.
"""
import email
import email.policy
import hashlib
import imaplib
import re
import smtplib
import threading
from datetime import datetime, timezone
from email.utils import formatdate

IMAP_HOST, IMAP_PORT = 'imap.mail.yahoo.com', 993
SMTP_HOST, SMTP_PORT = 'smtp.mail.yahoo.com', 465
HEADERS = 'FROM TO DELIVERED-TO SUBJECT DATE LIST-UNSUBSCRIBE LIST-UNSUBSCRIBE-POST MESSAGE-ID REFERENCES'
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']  # IMAP wants English
CHUNK = 500  # uids per IMAP command, keeps command lines short


class LoginFailed(Exception):
    pass


# ---------- pure helpers (tested) ----------

def quote(value):
    """One IMAP quoted string. Line breaks would end the command, so they go; non-ASCII is refused."""
    text = re.sub(r'[\r\n]+', ' ', str(value)).strip()
    if not text.isascii():
        # shortcut: no CHARSET UTF-8 literals, so only plain-letter search text; add literals if non-ASCII search is wanted
        raise ValueError('Yahoo search here takes plain letters only (no umlauts or emoji).')
    return '"' + text.replace('\\', '\\\\').replace('"', '\\"') + '"'


def imap_date(ms):
    d = datetime.fromtimestamp(ms / 1000, timezone.utc)
    return f'{d.day}-{MONTHS[d.month - 1]}-{d.year}'


def search_terms(criteria):
    """IMAP SEARCH keys for the app's filter. Several senders match if any matches."""
    terms = []
    senders = [s for s in (criteria.get('from') or []) if s]
    if senders:
        keys = [f'FROM {quote(s)}' for s in senders]
        expr = keys[-1]
        for key in reversed(keys[:-1]):
            expr = f'OR {key} {expr}'
        terms.append(expr)
    for name, key in (('to', 'TO'), ('subject', 'SUBJECT'), ('text', 'TEXT')):
        if criteria.get(name):
            terms.append(f'{key} {quote(criteria[name])}')
    if criteria.get('before'):
        terms.append(f'BEFORE {imap_date(criteria["before"])}')
    if criteria.get('larger'):
        terms.append(f'LARGER {int(criteria["larger"])}')
    if criteria.get('unseen'):
        terms.append('UNSEEN')
    return ' '.join(terms) or 'ALL'


def uid_set(text):
    """'1:3,7' -> [1, 2, 3, 7]"""
    out = []
    for part in text.split(','):
        lo, _, hi = part.partition(':')
        out.extend(range(int(lo), int(hi or lo) + 1))
    return out


def make_id(uid, validity, folder):
    return f'{uid}:{validity}:{folder}'


def split_id(message_id):
    uid, validity, folder = message_id.split(':', 2)
    return int(uid), validity, folder


def list_line(line):
    """One LIST answer -> (name, flags)."""
    m = re.match(rb'\((?P<flags>[^)]*)\) (?:"[^"]*"|NIL) (?P<name>.+)$', line)
    name = m['name'].decode()
    if name.startswith('"'):
        name = name[1:-1].replace('\\"', '"').replace('\\\\', '\\')
    return name, m['flags'].decode().split()


def mailbox(name):
    """A folder name as an IMAP argument."""
    return '"' + name.replace('\\', '\\\\').replace('"', '\\"') + '"'


def fetch_meta(meta):
    """UID, size, flags and date (ms) from one FETCH answer line, or None for lines without a UID."""
    if not isinstance(meta, bytes) or not (m := re.search(rb'UID (\d+)', meta)):
        return None
    rec = {'uid': int(m[1])}
    if m := re.search(rb'RFC822\.SIZE (\d+)', meta):
        rec['size'] = int(m[1])
    if m := re.search(rb'FLAGS \(([^)]*)\)', meta):
        rec['flags'] = m[1].decode().split()
    if m := re.search(rb'INTERNALDATE "([^"]+)"', meta):
        rec['date'] = int(datetime.strptime(m[1].decode().strip(), '%d-%b-%Y %H:%M:%S %z').timestamp() * 1000)
    return rec


def header_record(raw_headers, size, flags, internal):
    msg = email.message_from_bytes(raw_headers, policy=email.policy.default)
    get = lambda name: str(msg.get(name, '') or '')
    return {
        'from': get('From'), 'to': get('To'), 'deliveredTo': get('Delivered-To'), 'subject': get('Subject'),
        'date': internal, 'size': size, 'unread': '\\Seen' not in flags,
        'listUnsubscribe': get('List-Unsubscribe'), 'listUnsubscribePost': get('List-Unsubscribe-Post'),
        'messageId': get('Message-ID'), 'references': get('References'), 'snippet': '',
    }


def body_parts(raw):
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    html = msg.get_body(('html',))
    text = msg.get_body(('plain',))
    return {'html': html.get_content() if html else None, 'text': text.get_content() if text and not html else None}


# ---------- one account's connection ----------

class Account:
    def __init__(self, user, password):
        self.user, self.password = user, password
        self.lock = threading.Lock()  # imaplib is not thread-safe, the server is
        self.imap = None

    def connect(self):
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=60)
        try:
            imap.login(self.user, self.password)
        except imaplib.IMAP4.error as error:
            raise LoginFailed('Yahoo refused the login. Check the address and the app password.') from error
        self.imap = imap

    def run(self, op, args):
        with self.lock:
            for attempt in (0, 1):
                if self.imap is None:
                    self.connect()
                try:
                    return getattr(self, f'op_{op}')(**args)
                except (imaplib.IMAP4.abort, OSError):
                    self.imap = None  # dropped connection: reconnect once
                    if attempt:
                        raise

    def check(self, typ, data):
        if typ != 'OK':
            raise RuntimeError((data[0] or b'IMAP error').decode(errors='replace'))
        return data

    def select(self, folder, readonly=True):
        self.check(*self.imap.select(mailbox(folder), readonly=readonly))
        return self.imap.response('UIDVALIDITY')[1][0].decode()

    def fetch_parts(self, uids, items):
        """UID FETCH in chunks. Returns {uid: {item name: value}}."""
        out = {}
        for i in range(0, len(uids), CHUNK):
            typ, data = self.imap.uid('FETCH', ','.join(map(str, uids[i:i + CHUNK])), items)
            for entry in self.check(typ, data):
                # imaplib gives (meta, literal) for items with a body, plain bytes for the rest, and b')' closers
                meta, payload = entry if isinstance(entry, tuple) else (entry, None)
                rec = fetch_meta(meta)
                if rec is not None:
                    out.setdefault(rec.pop('uid'), {}).update(rec, **({'payload': payload} if payload is not None else {}))
        return out

    def by_folder(self, ids):
        groups = {}
        for message_id in ids:
            uid, validity, folder = split_id(message_id)
            groups.setdefault(folder, []).append((uid, validity, message_id))
        return groups

    # ---------- ops ----------

    def op_folders(self):
        typ, data = self.imap.list('""', '*')
        out = []
        for line in self.check(typ, data):
            name, flags = list_line(line)
            if '\\Noselect' in flags or '\\NonExistent' in flags:
                continue
            role = next((f[1:] for f in flags if f in ('\\Trash', '\\Junk', '\\Sent', '\\Drafts', '\\Archive')), None)
            out.append({'name': name, 'role': 'Inbox' if name.upper() == 'INBOX' else role})
        return out

    def op_search(self, folders, criteria, limit=None):
        """Newest first over the given folders. limit: at most this many, newest kept."""
        terms = search_terms(criteria)
        hits = []
        for folder in folders:
            validity = self.select(folder)
            typ, data = self.imap.uid('SEARCH', terms)
            uids = sorted(int(u) for u in self.check(typ, data)[0].split())
            if limit:
                uids = uids[-limit:]  # higher uid = newer in one folder
            dates = self.fetch_parts(uids, '(INTERNALDATE)')
            hits += [{'id': make_id(u, validity, folder), 'date': dates.get(u, {}).get('date', 0)} for u in uids]
        hits.sort(key=lambda h: -h['date'])
        more = bool(limit) and len(hits) > limit
        return {'hits': hits[:limit] if limit else hits, 'more': more}

    def op_fetch(self, ids, body=False):
        out = {}
        for folder, items in self.by_folder(ids).items():
            validity = self.select(folder)
            live = [uid for uid, v, _ in items if v == validity]
            # BODY.PEEK: reading must never mark mail as read
            parts = self.fetch_parts(live, '(UID RFC822.SIZE FLAGS INTERNALDATE BODY.PEEK[])' if body
                                     else f'(UID RFC822.SIZE FLAGS INTERNALDATE BODY.PEEK[HEADER.FIELDS ({HEADERS})])')
            for uid, v, message_id in items:
                p = parts.get(uid) if v == validity else None
                if not p or 'payload' not in p:
                    continue
                rec = header_record(p['payload'], p.get('size', 0), p.get('flags', []), p.get('date', 0))
                if body:
                    rec.update(body_parts(p['payload']))
                out[message_id] = {**rec, 'id': message_id}
        return [out[i] for i in ids if i in out]

    def op_move(self, ids, to):
        """Moves and returns [{old, new}] so the page can move them back. new is null when the server gave no uid."""
        moved = []
        to_validity = None
        for folder, items in self.by_folder(ids).items():
            if folder == to:
                continue
            validity = self.select(folder, readonly=False)
            uids = [uid for uid, v, _ in items if v == validity]
            old_by_uid = {uid: message_id for uid, v, message_id in items if v == validity}
            for i in range(0, len(uids), CHUNK):
                part = uids[i:i + CHUNK]
                self.check(*self.imap.uid('MOVE', ','.join(map(str, part)), mailbox(to)))
                code = self.imap.response('COPYUID')[1]
                pairs = {}
                if code and code[0]:
                    to_validity, src, dst = code[0].decode().split()
                    pairs = dict(zip(uid_set(src), uid_set(dst)))
                for uid in part:
                    new = pairs.get(uid)
                    moved.append({'old': old_by_uid[uid], 'new': make_id(new, to_validity, to) if new else None})
        return moved

    def op_flag(self, ids, seen):
        for folder, items in self.by_folder(ids).items():
            validity = self.select(folder, readonly=False)
            uids = [str(uid) for uid, v, _ in items if v == validity]
            for i in range(0, len(uids), CHUNK):
                self.check(*self.imap.uid('STORE', ','.join(uids[i:i + CHUNK]), '+FLAGS' if seen else '-FLAGS', '(\\Seen)'))
        return True

    def op_count(self, folder):
        typ, data = self.imap.status(mailbox(folder), '(MESSAGES)')
        return int(re.search(rb'MESSAGES (\d+)', self.check(typ, data)[0])[1])

    def op_empty(self, folder):
        """Deletes every mail in one folder for good. UID EXPUNGE, so mail other clients flagged elsewhere is untouched."""
        self.select(folder, readonly=False)
        typ, data = self.imap.uid('SEARCH', 'ALL')
        uids = self.check(typ, data)[0].split()
        for i in range(0, len(uids), CHUNK):
            part = b','.join(uids[i:i + CHUNK]).decode()
            self.check(*self.imap.uid('STORE', part, '+FLAGS.SILENT', '(\\Deleted)'))
            self.check(*self.imap.uid('EXPUNGE', part))
        return len(uids)

    def op_create(self, name):
        self.check(*self.imap.create(mailbox(name)))
        return {'id': name, 'name': name}

    def op_send(self, to, raw):
        """raw: the message without From and Date; both are set here to the signed-in user."""
        message = f'From: {self.user}\r\nDate: {formatdate(localtime=True)}\r\n'.encode() + raw.encode()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=60) as smtp:
            try:
                smtp.login(self.user, self.password)
            except smtplib.SMTPAuthenticationError as error:
                raise LoginFailed('Yahoo refused the login for sending. Check the app password.') from error
            smtp.sendmail(self.user, to, message)  # envelope sender is always the signed-in user
        return True


_accounts = {}
_accounts_lock = threading.Lock()


def handle(request):
    """One call from the page. Returns (status, body)."""
    user, password, op = request.get('user'), request.get('password'), request.get('op')
    if not user or not password or not hasattr(Account, f'op_{op}'):
        return 400, {'error': 'Bad request.'}
    key = (user.lower(), hashlib.sha256(password.encode()).hexdigest())
    with _accounts_lock:
        account = _accounts.setdefault(key, Account(user, password))
    args = {k: v for k, v in request.items() if k not in ('user', 'password', 'op')}
    try:
        return 200, {'result': account.run(op, args)}
    except LoginFailed as error:
        with _accounts_lock:
            _accounts.pop(key, None)
        return 401, {'error': str(error)}
    except Exception as error:  # the message names the IMAP problem, never the password
        return 502, {'error': f'Yahoo: {error}'}
