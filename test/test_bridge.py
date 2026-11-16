import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import imap_bridge as b  # noqa: E402


class Pure(unittest.TestCase):
    def test_quote_strips_line_breaks_and_escapes(self):
        self.assertEqual(b.quote('a"b\\c\r\nd'), '"a\\"b\\\\c d"')

    def test_quote_refuses_non_ascii(self):
        with self.assertRaises(ValueError):
            b.quote('Grüße')

    def test_search_terms(self):
        self.assertEqual(b.search_terms({}), 'ALL')
        self.assertEqual(b.search_terms({'from': ['a.com', 'b.com', 'c.com']}), 'OR FROM "a.com" OR FROM "b.com" FROM "c.com"')
        self.assertEqual(b.search_terms({'from': ['a.com'], 'before': 1731628800000, 'larger': 5242880, 'unseen': True}),
                         'FROM "a.com" BEFORE 15-Nov-2024 LARGER 5242880 UNSEEN')

    def test_uid_set_and_ids(self):
        self.assertEqual(b.uid_set('1:3,7'), [1, 2, 3, 7])
        self.assertEqual(b.split_id(b.make_id(5, '99', 'Archive:Old')), (5, '99', 'Archive:Old'))

    def test_list_line(self):
        self.assertEqual(b.list_line(b'(\\HasNoChildren \\Trash) "/" "Trash"'), ('Trash', ['\\HasNoChildren', '\\Trash']))
        self.assertEqual(b.list_line(b'(\\HasNoChildren) "/" Inbox'), ('Inbox', ['\\HasNoChildren']))

    def test_fetch_meta(self):
        rec = b.fetch_meta(b'3 (UID 42 RFC822.SIZE 1200 FLAGS (\\Seen) INTERNALDATE " 5-Nov-2026 10:00:00 +0000" BODY[] {9}')
        self.assertEqual(rec, {'uid': 42, 'size': 1200, 'flags': ['\\Seen'], 'date': 1793872800000})
        self.assertIsNone(b.fetch_meta(b')'))

    def test_header_record(self):
        rec = b.header_record(b'From: Shop <a@shop.com>\r\nSubject: Hi\r\nList-Unsubscribe: <https://u.com>\r\n\r\n', 10, [], 1)
        self.assertEqual((rec['from'], rec['subject'], rec['unread'], rec['listUnsubscribe']), ('Shop <a@shop.com>', 'Hi', True, '<https://u.com>'))

    def test_handle_refuses_unknown_ops(self):
        self.assertEqual(b.handle({'user': 'a', 'password': 'p', 'op': 'connect'})[0], 400)
        self.assertEqual(b.handle({'user': 'a', 'op': 'folders'})[0], 400)


class FakeImap:
    """Just enough IMAP for move and fetch."""
    def __init__(self):
        self.commands = []

    def select(self, name, readonly=True):
        self.commands.append(('SELECT', name, readonly))
        return 'OK', [b'1']

    def response(self, code):
        return code, [b'7' if code == 'UIDVALIDITY' else b'8 10,11 200:201']

    def uid(self, command, *args):
        self.commands.append((command,) + args)
        if command == 'FETCH':
            return 'OK', [(b'1 (UID 10 RFC822.SIZE 5 FLAGS () INTERNALDATE "05-Nov-2026 10:00:00 +0000" BODY[HEADER.FIELDS (FROM)] {20}',
                           b'From: a@b.com\r\n\r\n'), b')']
        return 'OK', [None]


class Ops(unittest.TestCase):
    def account(self):
        a = b.Account('me@yahoo.com', 'secret')
        a.imap = FakeImap()
        return a

    def test_move_returns_new_ids_from_copyuid(self):
        a = self.account()
        moved = a.op_move(['10:7:Inbox', '11:7:Inbox'], 'Trash')
        self.assertEqual(moved, [{'old': '10:7:Inbox', 'new': '200:8:Trash'}, {'old': '11:7:Inbox', 'new': '201:8:Trash'}])
        self.assertIn(('MOVE', '10,11', '"Trash"'), a.imap.commands)

    def test_fetch_never_marks_read_and_skips_stale_ids(self):
        a = self.account()
        out = a.op_fetch(['10:7:Inbox', '10:6:Inbox'])
        self.assertEqual([m['id'] for m in out], ['10:7:Inbox'])
        fetches = [c for c in a.imap.commands if c[0] == 'FETCH']
        self.assertTrue(all('BODY.PEEK' in c[2] for c in fetches))
        self.assertEqual(fetches[0][1], '10')  # the stale id (old UIDVALIDITY 6) is not even asked for


if __name__ == '__main__':
    unittest.main()
