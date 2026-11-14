import test from 'node:test';
import assert from 'node:assert/strict';
import { createGmail, AuthExpired } from '../src/gmail.js';

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
  await gmail.scanAliases(new Map([['4', { id: '4', alias: null, domain: null }]]));
  assert.equal(sleeps.length, 2);              // 3 fetched, the first goes at once, the cached one never waits
  assert.ok(sleeps.every(ms => ms > 0 && ms <= 300));
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
  const { gmail, calls } = setup(() => reply(204));
  await gmail.trash(Array.from({ length: 2500 }, (_, i) => `m${i}`));
  assert.deepEqual(calls.map(c => c.body.ids.length), [1000, 1000, 500]);
  for (const c of calls) {
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
  const { gmail, calls } = setup(() => reply(204));
  await gmail.moveToFolder(['a'], 'L1');
  assert.deepEqual(calls[0].body, { ids: ['a'], addLabelIds: ['L1'], removeLabelIds: ['INBOX'] });
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
  const known = new Map([['1', { id: '1', alias: 'spotify', domain: 'spotify.com' }]]);
  const recs = await gmail.scanAliases(known);
  assert.equal(calls.filter(c => /\/messages\/\w+$/.test(path(c.url))).length, 2);
  assert.deepEqual(recs, [
    { id: '1', alias: 'spotify', domain: 'spotify.com' },
    { id: '2', alias: 'uber', domain: 'uber.com', from: 'no-reply@mail.uber.com' },
    { id: '3', alias: null, domain: 'y.com', from: 'b@y.com' }]);
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
  assert.equal(await gmail.applyRuleToExisting(rule), 3);
  assert.equal(calls[0].url.searchParams.get('q'), 'to:("me+a@gmail.com") -from:("a.com")');
  const posts = calls.filter(c => c.init.method === 'POST');
  assert.equal(posts.length, 1);
  assert.deepEqual(posts[0].body, { ids: ['1', '2', '3'], addLabelIds: ['TRASH'] });
});

test('applyRuleToExisting with nothing matching', async () => {
  const { gmail, calls } = setup(() => reply(200, {}));
  assert.equal(await gmail.applyRuleToExisting(rule), 0);
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
