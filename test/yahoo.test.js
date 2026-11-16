import test from 'node:test';
import assert from 'node:assert/strict';
import { createYahoo } from '../src/yahoo.js';
import { AuthExpired } from '../src/gmail.js';

const FOLDERS = [{ name: 'INBOX', role: 'Inbox' }, { name: 'Trash', role: 'Trash' }, { name: 'Bulk', role: 'Junk' },
  { name: 'Draft', role: 'Drafts' }, { name: 'Sent', role: 'Sent' }, { name: 'Shops', role: null }];

// answer(op, body) -> result, or { status, error }
function setup(answer) {
  const calls = [];
  const yahoo = createYahoo({
    user: 'me@yahoo.com', getToken: () => 'app-pass',
    fetch: async (url, init) => {
      const body = JSON.parse(init.body);
      calls.push(body);
      const out = body.op === 'folders' ? FOLDERS : answer(body.op, body);
      const status = out?.status || 200;
      return { status, ok: status === 200, json: async () => (status === 200 ? { result: out } : { error: out.error }) };
    },
  });
  return { yahoo, calls };
}

test('every call carries the address and app password, never in the URL', async () => {
  const { yahoo, calls } = setup(() => ({ hits: [], more: false }));
  await yahoo.listMessages({ folderId: 'ALL' });
  assert.deepEqual(calls.map(c => [c.op, c.user, c.password]), [['folders', 'me@yahoo.com', 'app-pass'], ['search', 'me@yahoo.com', 'app-pass']]);
});

test('a refused login becomes AuthExpired with the bridge message', async () => {
  const { yahoo } = setup(() => ({ status: 401, error: 'Yahoo refused the login.' }));
  await assert.rejects(yahoo.listMessages({ folderId: 'INBOX' }), e => e instanceof AuthExpired && /refused/.test(e.message));
});

test('inbox by default, every folder but Trash, Spam and Drafts for a search', async () => {
  const { yahoo, calls } = setup(() => ({ hits: [], more: false }));
  await yahoo.listMessages({});
  await yahoo.listMessages({ filter: { from: 'shop.com' } });
  await yahoo.listMessages({ filter: { trash: true } });
  const searches = calls.filter(c => c.op === 'search');
  assert.deepEqual(searches.map(s => s.folders), [['INBOX'], ['INBOX', 'Sent', 'Shops'], ['Trash']]);
  assert.deepEqual(searches[1].criteria.from, ['shop.com']);
});

test('paging by offset', async () => {
  const hits = Array.from({ length: 60 }, (_, i) => ({ id: `${i}:1:INBOX`, date: 1000 - i }));
  const { yahoo, calls } = setup((op, body) => (op === 'search' ? { hits: hits.slice(0, body.limit), more: body.limit < 60 } : body.ids.map(id => ({ id }))));
  const first = await yahoo.listMessages({});
  assert.equal(first.messages.length, 50);
  assert.equal(first.nextPageToken, '50');
  const second = await yahoo.listMessages({ pageToken: first.nextPageToken });
  assert.deepEqual(second.messages.map(m => m.id), hits.slice(50).map(h => h.id));
  assert.equal(second.nextPageToken, null);
  assert.equal(calls.filter(c => c.op === 'fetch').at(-1).ids.length, 10);
});

test('trash then undo moves each mail back to the folder it came from', async () => {
  const { yahoo, calls } = setup((op, body) => (op === 'move'
    ? body.ids.map((old, i) => ({ old, new: i === 2 ? null : `${100 + i}:9:${body.to}` })) : true));
  const receipt = await yahoo.trash(['1:1:INBOX', '2:1:Shops', '3:1:INBOX']);
  assert.equal(receipt.count, 3);
  assert.equal(receipt.moves.length, 2); // the server gave no new id for the third
  await yahoo.undo(receipt);
  const back = calls.filter(c => c.op === 'move').slice(1).map(c => [c.ids, c.to]);
  assert.deepEqual(back, [[['100:9:Trash'], 'INBOX'], [['101:9:Trash'], 'Shops']]);
});

test('trash all from senders: search, then move; mark-read undo only for mail that was unread', async () => {
  const { yahoo, calls } = setup((op, body) => {
    if (op === 'search') return { hits: body.criteria.unseen ? [{ id: '2:1:INBOX' }] : [{ id: '1:1:INBOX' }, { id: '2:1:INBOX' }], more: false };
    if (op === 'move') return body.ids.map(old => ({ old, new: `n${old}` }));
    return true;
  });
  const rule = { conditions: { senderContains: ['a.com', 'b.com'] }, exceptions: {}, actions: { moveToFolder: 'Shops', markAsRead: true } };
  assert.deepEqual(await yahoo.countMatching(rule), { count: 2, more: false });
  const receipt = await yahoo.applyRuleToExisting(rule);
  assert.deepEqual(receipt.moves.map(m => m.unread), [false, true]);
  assert.deepEqual(calls.find(c => c.op === 'search').criteria.from, ['a.com', 'b.com']);
});

test('rules are refused with a pointer to Yahoo settings, and a cleanup needs a condition', async () => {
  const { yahoo } = setup(() => true);
  await assert.rejects(yahoo.createRule({}), /Settings > Filters/);
  assert.equal(yahoo.rules, false);
  await assert.rejects(yahoo.trashMessages({}), /at least one/);
});

test('scan fetches only new mail and records its folder', async () => {
  const { yahoo, calls } = setup((op, body) => (op === 'search'
    ? { hits: [{ id: '1:1:INBOX' }, { id: '2:1:Shops' }], more: false }
    : body.ids.map(id => ({ id, from: 'Shop <a@news.shop.com>', to: 'me@yahoo.com', deliveredTo: '', size: 5 }))));
  const known = new Map([['1:1:INBOX', { id: '1:1:INBOX', size: 1 }]]);
  const records = await yahoo.scanAliases(known);
  assert.deepEqual(calls.find(c => c.op === 'fetch').ids, ['2:1:Shops']);
  assert.deepEqual(records[1], { id: '2:1:Shops', alias: null, domain: 'shop.com', from: 'a@news.shop.com', size: 5, folders: ['Shops'] });
});

test('sendMail passes the parsed recipients and a message without From', async () => {
  const { yahoo, calls } = setup(() => true);
  await yahoo.sendMail({ to: 'a@b.com, c@d.com', subject: 'Hi', text: 'x' });
  const send = calls.find(c => c.op === 'send');
  assert.deepEqual(send.to, ['a@b.com', 'c@d.com']);
  assert.doesNotMatch(send.raw, /^From:/m);
});
