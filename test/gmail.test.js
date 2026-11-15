import test from 'node:test';
import assert from 'node:assert/strict';
import { createGmail, AuthExpired, gmailSearch } from '../src/gmail.js';

const reply = (status, body) => ({
  status, ok: status >= 200 && status < 300,
  text: async () => (body === undefined ? '' : JSON.stringify(body)),
});

// handler(url: URL, init, callIndex) -> reply
function setup(handler) {
  const calls = [], sleeps = [];
  const gmail = createGmail({
    getToken: () => 'T',
    fetch: async (url, init) => {
      const call = { url: new URL(url), init, body: init.body ? JSON.parse(init.body) : undefined };
      calls.push(call);
      return handler(call.url, call, calls.length - 1);
    },
    sleep: async ms => { sleeps.push(ms); },
  });
  return { gmail, calls, sleeps };
}
const path = url => url.pathname.replace('/gmail/v1/users/me', '');
const meta = (id, headers, labelIds = []) => ({
  id, labelIds, snippet: `s${id}`, internalDate: '1700000000000',
  payload: { headers: Object.entries(headers).map(([name, value]) => ({ name, value })) },
});

test('sends bearer token', async () => {
  const { gmail, calls } = setup(() => reply(200, { labels: [] }));
  await gmail.listFolders();
  assert.equal(calls[0].init.headers.Authorization, 'Bearer T');
});

test('401 throws AuthExpired', async () => {
  const { gmail } = setup(() => reply(401, { error: { message: 'x' } }));
  await assert.rejects(gmail.listFolders(), AuthExpired);
});

test('429 retries then succeeds', async () => {
  const { gmail, sleeps } = setup((u, c, i) => (i < 2 ? reply(429) : reply(200, { labels: [{ id: 'L1', name: 'A', type: 'user' }] })));
  assert.deepEqual(await gmail.listFolders(), [{ id: 'L1', name: 'A' }]);
  assert.deepEqual(sleeps, [500, 1000]);
});

test('429 gives up after 9 retries, about two minutes of waiting', async () => {
  const { gmail, calls, sleeps } = setup(() => reply(429, { error: { message: 'slow down' } }));
  await assert.rejects(gmail.listFolders(), /429.*slow down/);
  assert.equal(calls.length, 10);
  assert.ok(sleeps.reduce((a, b) => a + b) >= 120000);
  assert.ok(Math.max(...sleeps) <= 32000);
});

test('scanAliases spaces its requests to stay under the quota', async () => {
  const { gmail, sleeps } = setup(u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/messages') return reply(200, { messages: [{ id: '1' }, { id: '2' }, { id: '3' }, { id: '4' }] });
    return reply(200, meta('x', { To: 'me@gmail.com', From: 'a@b.c' }));
  });
  await gmail.scanAliases(new Map([['4', { id: '4', alias: null, domain: null, size: 0 }]]));
  assert.equal(sleeps.length, 3);              // the id list goes at once, the 3 fetches wait, the cached one never waits
  assert.ok(sleeps.every(ms => ms > 0 && ms <= 400));
  assert.ok(sleeps[1] > 100);                  // slots are 100 ms apart
});

const QUOTA = "Quota exceeded for quota metric 'Total Query Cost' and limit 'Units per minute per user'";
test('403 quota error is retried', async () => {
  const { gmail, sleeps } = setup((u, c, i) => (i < 1 ? reply(403, { error: { message: QUOTA } }) : reply(200, { labels: [] })));
  assert.deepEqual(await gmail.listFolders(), []);
  assert.deepEqual(sleeps, [500]);
});

test('403 without quota is not retried', async () => {
  const { gmail, calls } = setup(() => reply(403, { error: { message: 'Insufficient Permission' } }));
  await assert.rejects(gmail.listFolders(), /403.*Insufficient/);
  assert.equal(calls.length, 1);
});

test('scanAliases keeps finished records in known when it fails', async () => {
  const { gmail } = setup(u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/messages') return reply(200, { messages: [{ id: '2' }, { id: '9' }] });
    if (p.endsWith('/2')) return reply(200, meta('2', { To: 'me+uber@gmail.com', From: 'a@uber.com' }));
    return reply(400, { error: { message: 'boom' } });
  });
  const known = new Map();
  await assert.rejects(gmail.scanAliases(known), /boom/);
  assert.equal(known.get('2').alias, 'uber');
});

test('other errors carry status and message', async () => {
  const { gmail } = setup(() => reply(400, { error: { message: 'bad filter' } }));
  await assert.rejects(gmail.deleteRule('x'), /400.*bad filter/);
});

test('trash chunks by 1000', async () => {
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204) : reply(200, { labelIds: ['INBOX'] })));
  await gmail.trash(Array.from({ length: 2500 }, (_, i) => `m${i}`));
  const posts = calls.filter(c => c.init.method === 'POST');
  assert.deepEqual(posts.map(c => c.body.ids.length), [1000, 1000, 500]);
  for (const c of posts) {
    assert.equal(path(c.url), '/messages/batchModify');
    assert.deepEqual(c.body.addLabelIds, ['TRASH']);
  }
});

test('trash with no ids makes no request', async () => {
  const { gmail, calls } = setup(() => reply(204));
  await gmail.trash([]);
  assert.equal(calls.length, 0);
});

test('moveToFolder body', async () => {
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204) : reply(200, { labelIds: ['INBOX'] })));
  await gmail.moveToFolder(['a'], 'L1');
  assert.deepEqual(calls.find(c => c.init.method === 'POST').body, { ids: ['a'], addLabelIds: ['L1'], removeLabelIds: ['INBOX'] });
});

test('undo of a move puts back only the inbox mail and leaves labels that were there', async () => {
  const labels = { a: ['INBOX'], b: ['L2'], c: ['INBOX', 'L1'] };
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204) : reply(200, { labelIds: labels[path(u).split('/').pop()] })));
  const receipt = await gmail.moveToFolder(['a', 'b', 'c'], 'L1');
  assert.equal(receipt.count, 3);
  const before = calls.length;
  await gmail.undo(receipt);
  assert.deepEqual(calls.slice(before).map(c => c.body), [
    { ids: ['a', 'b'], removeLabelIds: ['L1'] },
    { ids: ['a', 'c'], addLabelIds: ['INBOX'] }]);
});

test('undo of trash leaves mail that was in Trash already', async () => {
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204) : reply(200, { labelIds: path(u).endsWith('/old') ? ['TRASH'] : ['INBOX'] })));
  const receipt = await gmail.trash(['new', 'old']);
  const before = calls.length;
  await gmail.undo(receipt);
  assert.deepEqual(calls.slice(before).map(c => c.body), [{ ids: ['new'], removeLabelIds: ['TRASH'] }]);
});

test('undo of a rule run marks unread only the mail that was unread', async () => {
  const { gmail, calls } = setup((u, c) => {
    if (c.init.method === 'POST') return reply(204);
    const label = u.searchParams.get('labelIds');
    if (label === 'UNREAD') return reply(200, { messages: [{ id: '2' }] });
    if (label === 'INBOX') return reply(200, { messages: [{ id: '1' }, { id: '2' }] });
    if (label) return reply(200, {});
    return reply(200, { messages: [{ id: '1' }, { id: '2' }, { id: '3' }] });
  });
  const receipt = await gmail.applyRuleToExisting({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { moveToFolder: 'L1', markAsRead: true } });
  assert.equal(receipt.count, 3);
  const before = calls.length;
  await gmail.undo(receipt);
  assert.deepEqual(calls.slice(before).map(c => c.body), [
    { ids: ['1', '2', '3'], removeLabelIds: ['L1'] },
    { ids: ['1', '2'], addLabelIds: ['INBOX'] },
    { ids: ['2'], addLabelIds: ['UNREAD'] }]);
});

test('markRead', async () => {
  const { gmail, calls } = setup(() => reply(204));
  await gmail.markRead(['a'], true);
  await gmail.markRead(['a'], false);
  assert.deepEqual(calls[0].body, { ids: ['a'], removeLabelIds: ['UNREAD'] });
  assert.deepEqual(calls[1].body, { ids: ['a'], addLabelIds: ['UNREAD'] });
});

test('createFolder returns existing', async () => {
  const { gmail, calls } = setup(() => reply(200, { labels: [
    { id: 'INBOX', name: 'INBOX', type: 'system' }, { id: 'L9', name: 'Subscriptions', type: 'user' }] }));
  assert.deepEqual(await gmail.createFolder('subscriptions'), { id: 'L9', name: 'Subscriptions' });
  assert.ok(calls.every(c => c.init.method === 'GET'));
});

test('createFolder creates when missing', async () => {
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(200, { id: 'L2', name: 'New' }) : reply(200, { labels: [] })));
  assert.deepEqual(await gmail.createFolder(' New '), { id: 'L2', name: 'New' });
  assert.equal(calls.at(-1).body.name, 'New');
});

test('listMessages maps headers', async () => {
  const { gmail, calls } = setup(u => (path(u) === '/messages'
    ? reply(200, { messages: [{ id: '1' }, { id: '2' }], nextPageToken: 'N' })
    : reply(200, meta(path(u).split('/').pop(), { From: 'A <a@x.com>', To: 'me@gmail.com', Subject: 'Hi' }, path(u).endsWith('1') ? ['UNREAD'] : []))));
  const page = await gmail.listMessages({});
  assert.equal(calls[0].url.searchParams.get('labelIds'), 'INBOX');
  assert.equal(calls[0].url.searchParams.get('maxResults'), '50');
  assert.equal(page.nextPageToken, 'N');
  assert.deepEqual(page.messages.map(m => m.id), ['1', '2']);
  assert.deepEqual(page.messages[0], { id: '1', from: 'A <a@x.com>', to: 'me@gmail.com', deliveredTo: '', subject: 'Hi', date: 1700000000000, snippet: 's1', unread: true });
  assert.equal(page.messages[1].unread, false);
});

test('listMessages empty mailbox', async () => {
  const { gmail } = setup(() => reply(200, { resultSizeEstimate: 0 }));
  assert.deepEqual(await gmail.listMessages({ query: 'x' }), { messages: [], nextPageToken: null });
});

test('listMessages with query sends no label', async () => {
  const { gmail, calls } = setup(() => reply(200, {}));
  await gmail.listMessages({ query: 'from:x' });
  assert.equal(calls[0].url.searchParams.get('q'), 'from:x');
  assert.equal(calls[0].url.searchParams.get('labelIds'), null);
});

const scanHandler = u => {
  const p = path(u);
  if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
  if (p === '/messages') return reply(200, { messages: [{ id: '1' }, { id: '2' }, { id: '3' }] });
  if (p.endsWith('/2')) return reply(200, meta('2', { 'Delivered-To': 'me+uber@gmail.com', To: 'list@x.com', From: 'Uber <no-reply@mail.uber.com>' }));
  return reply(200, meta('3', { To: 'me@gmail.com', From: 'b@y.com' }));
};

test('scanAliases skips known ids', async () => {
  const { gmail, calls } = setup(scanHandler);
  const known = new Map([['1', { id: '1', alias: 'spotify', domain: 'spotify.com', size: 9 }]]);
  const recs = await gmail.scanAliases(known);
  assert.equal(calls.filter(c => /\/messages\/\w+$/.test(path(c.url))).length, 2);
  assert.deepEqual(recs, [
    { id: '1', alias: 'spotify', domain: 'spotify.com', size: 9, folders: [] },
    { id: '2', alias: 'uber', domain: 'uber.com', from: 'no-reply@mail.uber.com', size: 0, folders: [] },
    { id: '3', alias: null, domain: 'y.com', from: 'b@y.com', size: 0, folders: [] }]);
});

test('scanAliases stops at 2000 ids', async () => {
  let n = 0;
  const { gmail, calls } = setup(u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/messages') return reply(200, { messages: Array.from({ length: 500 }, () => ({ id: `k${n++}` })), nextPageToken: 'more' });
    return reply(200, meta('x', {}));
  });
  const known = new Map(Array.from({ length: 5000 }, (_, i) => [`k${i}`, { id: `k${i}`, alias: null, domain: null }]));
  assert.equal((await gmail.scanAliases(known)).length, 2000);
  assert.equal(calls.filter(c => path(c.url) === '/messages').length, 4);
});

const rule = { conditions: { sentToAddresses: ['me+a@gmail.com'] }, exceptions: { senderContains: ['a.com'] }, actions: { delete: true, forwardTo: ['x@y.z'] } };

test('applyRuleToExisting pages and does not forward', async () => {
  const { gmail, calls } = setup((u, c) => {
    if (c.init.method === 'POST') return reply(204);
    return u.searchParams.get('pageToken')
      ? reply(200, { messages: [{ id: '3' }] })
      : reply(200, { messages: [{ id: '1' }, { id: '2' }], nextPageToken: 'P' });
  });
  assert.equal((await gmail.applyRuleToExisting(rule)).count, 3);
  assert.equal(calls[0].url.searchParams.get('q'), 'to:("me+a@gmail.com") -from:("a.com")');
  const posts = calls.filter(c => c.init.method === 'POST');
  assert.equal(posts.length, 1);
  assert.deepEqual(posts[0].body, { ids: ['1', '2', '3'], addLabelIds: ['TRASH'] });
});

test('applyRuleToExisting with nothing matching', async () => {
  const { gmail, calls } = setup(() => reply(200, {}));
  assert.equal((await gmail.applyRuleToExisting(rule)).count, 0);
  assert.equal(calls.length, 1);
});

test('createRule posts toGmailFilter output', async () => {
  const { gmail, calls } = setup(() => reply(200, { id: 'F1' }));
  assert.deepEqual(await gmail.createRule({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { delete: true } }), { id: 'F1' });
  assert.equal(path(calls[0].url), '/settings/filters');
  assert.deepEqual(calls[0].body, { criteria: { query: 'subject:("x")' }, action: { addLabelIds: ['TRASH'] } });
});

test('listRules summarises with folder names', async () => {
  const { gmail } = setup(u => (path(u) === '/labels'
    ? reply(200, { labels: [{ id: 'L1', name: 'Subs', type: 'user' }] })
    : reply(200, { filter: [{ id: 'F1', criteria: { query: 'q' }, action: { addLabelIds: ['L1'], removeLabelIds: ['INBOX'] } }] })));
  const rules = await gmail.listRules();
  assert.equal(rules[0].id, 'F1');
  assert.equal(rules[0].summary, 'If q: add Subs, skip inbox');
});

test('listRules with no filters', async () => {
  const { gmail } = setup(() => reply(200, {}));
  assert.deepEqual(await gmail.listRules(), []);
});

test('updateRule makes the new filter before deleting the old one', async () => {
  const { gmail, calls } = setup((u, c) => reply(200, c.init.method === 'POST' ? { id: 'NEW' } : {}));
  assert.deepEqual(await gmail.updateRule('OLD', { conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { delete: true } }), { id: 'NEW' });
  assert.deepEqual(calls.map(c => [c.init.method, path(c.url)]), [['POST', '/settings/filters'], ['DELETE', '/settings/filters/OLD']]);
});

test('recreateRule posts the kept definition without its old id', async () => {
  const { gmail, calls } = setup(() => reply(200, { id: 'N' }));
  await gmail.recreateRule({ id: 'OLD', criteria: { from: 'a' }, action: { addLabelIds: ['TRASH'] } });
  assert.deepEqual(calls[0].body, { criteria: { from: 'a' }, action: { addLabelIds: ['TRASH'] } });
});

test('cleanup filter becomes a Gmail search', () => {
  assert.equal(gmailSearch('', { before: Date.UTC(2024, 10, 15), larger: 5 * 1048576, unread: true }), 'before:2024/11/15 larger:5242880 is:unread');
});

test('trashMessages trashes every match and undo takes all of them out of Trash', async () => {
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204)
    : u.searchParams.get('pageToken') ? reply(200, { messages: [{ id: '3' }] }) : reply(200, { messages: [{ id: '1' }, { id: '2' }], nextPageToken: 'P' })));
  const receipt = await gmail.trashMessages({ larger: 1000 });
  assert.equal(calls[0].url.searchParams.get('q'), 'larger:1000');
  assert.equal(receipt.count, 3);
  await gmail.undo(receipt);
  assert.deepEqual(calls.at(-1).body, { ids: ['1', '2', '3'], removeLabelIds: ['TRASH'] });
});

test('emptyFolder lists Trash and deletes it in batches of 1000', async () => {
  const ids = Array.from({ length: 1200 }, (_, i) => ({ id: `m${i}` }));
  const { gmail, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(204)
    : u.searchParams.get('pageToken') ? reply(200, { messages: ids.slice(500) }) : reply(200, { messages: ids.slice(0, 500), nextPageToken: 'P' })));
  assert.equal(await gmail.emptyFolder('trash'), 1200);
  assert.equal(calls[0].url.searchParams.get('labelIds'), 'TRASH');
  assert.equal(calls[0].url.searchParams.get('includeSpamTrash'), 'true');
  const posts = calls.filter(c => c.init.method === 'POST');
  assert.deepEqual(posts.map(c => [path(c.url), c.body.ids.length]), [['/messages/batchDelete', 1000], ['/messages/batchDelete', 200]]);
});

test('scanAliases with no limit reads the whole mailbox', async () => {
  let n = 0;
  const { gmail } = setup(u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/messages') return reply(200, { messages: Array.from({ length: 500 }, () => ({ id: `k${n++}` })), nextPageToken: n < 3000 ? 'more' : undefined });
    return reply(200, meta('x', {}));
  });
  const known = new Map(Array.from({ length: 3000 }, (_, i) => [`k${i}`, { id: `k${i}`, size: 1 }]));
  assert.equal((await gmail.scanAliases(known, undefined, Infinity)).length, 3000);
});

test('sendMail posts a base64url message in the parent thread', async () => {
  const { gmail, calls } = setup(() => reply(200, { id: 'S' }));
  await gmail.sendMail({ to: 'a@b.com', subject: 'Re: x', text: 'hi', replyTo: { threadId: 'T1', messageId: '<1@b>' } });
  assert.equal(path(calls[0].url), '/messages/send');
  assert.equal(calls[0].body.threadId, 'T1');
  assert.doesNotMatch(calls[0].body.raw, /[+/=]/);
  const mime = Buffer.from(calls[0].body.raw, 'base64url').toString('utf8');
  assert.match(mime, /^To: a@b\.com\r\nSubject: Re: x\r\n/);
  assert.match(mime, /In-Reply-To: <1@b>/);
});

test('scanAliases brings folder membership up to date for cached records', async () => {
  const { gmail } = setup(u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/labels') return reply(200, { labels: [{ id: 'L1', name: 'Music', type: 'user' }] });
    if (p === '/messages') return reply(200, { messages: u.searchParams.get('labelIds') === 'L1' ? [{ id: 'b' }] : [{ id: 'a' }, { id: 'b' }] });
    return reply(400, {});
  });
  const known = new Map([['a', { id: 'a', size: 1, folders: ['L1'] }], ['b', { id: 'b', size: 1, folders: [] }]]);
  const recs = await gmail.scanAliases(known);
  assert.deepEqual(recs.map(r => r.folders), [[], ['L1']]);
});

test('cleanup with no condition is refused, not the whole mailbox', async () => {
  const { gmail, calls } = setup(() => reply(200, {}));
  await assert.rejects(gmail.trashMessages({}), /at least one/);
  assert.equal(calls.length, 0);
});

test('scanAliases lists folders one at a time, paced like the scan', async () => {
  let open = 0, most = 0;
  const { gmail, sleeps } = setup(async u => {
    const p = path(u);
    if (p === '/profile') return reply(200, { emailAddress: 'me@gmail.com' });
    if (p === '/labels') return reply(200, { labels: Array.from({ length: 5 }, (_, i) => ({ id: `L${i}`, name: `F${i}`, type: 'user' })) });
    open++; most = Math.max(most, open);
    await new Promise(r => setTimeout(r, 1));
    open--;
    return reply(200, { messages: [{ id: 'a' }] });
  });
  await gmail.scanAliases(new Map([['a', { id: 'a', size: 1 }]]));
  assert.equal(most, 1);
  assert.equal(sleeps.length, 5); // the first list goes at once, the five folder lists each wait for a slot
});
