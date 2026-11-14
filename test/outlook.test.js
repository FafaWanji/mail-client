import test from 'node:test';
import assert from 'node:assert/strict';
import { createOutlook } from '../src/outlook.js';
import { createGmail, AuthExpired } from '../src/gmail.js';
import { toOutlookRule, describeOutlookRule } from '../src/rules.js';

const reply = (status, body, headers = {}) => ({
  status, ok: status >= 200 && status < 300,
  headers: { get: name => headers[name] ?? null },
  text: async () => (body === undefined ? '' : JSON.stringify(body)),
});
function setup(handler, make = createOutlook) {
  const calls = [], sleeps = [];
  const provider = make({
    getToken: () => 'T',
    fetch: async (url, init) => {
      const call = { url: new URL(url), init, body: init.body ? JSON.parse(init.body) : undefined };
      calls.push(call);
      return handler(call.url, call, calls.length - 1);
    },
    sleep: async ms => { sleeps.push(ms); },
  });
  return { provider, calls, sleeps };
}
const path = url => url.pathname.replace('/v1.0/me', '');
const mail = (id, from, to, extra = {}) => ({
  id, subject: `s${id}`, bodyPreview: `p${id}`, receivedDateTime: '2026-11-01T10:00:00Z', isRead: false,
  from: { emailAddress: { name: 'Sender', address: from } },
  toRecipients: to.map(address => ({ emailAddress: { address } })), ...extra,
});

// ---- rule mapping ----
test('outlook rule: addresses become recipients, rest passes through', () => {
  const rule = toOutlookRule({
    conditions: { sentToAddresses: ['me+a@outlook.com'], subjectContains: ['bill'], importance: 'high' },
    exceptions: { senderContains: ['a.com'] },
    actions: { moveToFolder: 'F1', forwardTo: ['x@y.z'], markAsRead: true, delete: false },
  }, new Map([['F1', 'Subs']]));
  assert.deepEqual(rule.conditions, {
    sentToAddresses: [{ emailAddress: { address: 'me+a@outlook.com' } }], subjectContains: ['bill'], importance: 'high' });
  assert.deepEqual(rule.exceptions, { senderContains: ['a.com'] });
  assert.deepEqual(rule.actions, { moveToFolder: 'F1', forwardTo: [{ emailAddress: { address: 'x@y.z' } }], markAsRead: true });
  assert.equal(rule.isEnabled, true);
  assert.equal(rule.sequence, 1);
  assert.match(rule.displayName, /Subs/);
  assert.ok(rule.displayName.length <= 250);
});
test('outlook rule: empty parts are left out', () => {
  const rule = toOutlookRule({ conditions: {}, exceptions: {}, actions: { delete: true } }, new Map());
  assert.equal('conditions' in rule, false);
  assert.equal('exceptions' in rule, false);
});
test('outlook rule: gmail-only action refused', () => assert.throws(
  () => toOutlookRule({ conditions: {}, exceptions: {}, actions: { gmailStar: true } }, new Map()), /Outlook rules cannot do/));
test('outlook rule: needs an action', () => assert.throws(
  () => toOutlookRule({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: {} }, new Map()), /action/));
test('describe outlook rule', () => assert.equal(
  describeOutlookRule({
    conditions: { fromAddresses: [{ emailAddress: { address: 'a@b.c' } }], hasAttachments: true },
    actions: { moveToFolder: 'F1', stopProcessingRules: true },
  }, new Map([['F1', 'Subs']])),
  'If From address: a@b.c; Has attachment: Move to folder (skip inbox): Subs; Stop processing more rules'));

// ---- provider ----
test('sends bearer token to graph', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [] }));
  await provider.listFolders();
  assert.equal(calls[0].init.headers.Authorization, 'Bearer T');
  assert.equal(calls[0].url.origin, 'https://graph.microsoft.com');
});

test('401 throws AuthExpired', async () => {
  const { provider } = setup(() => reply(401, { error: { message: 'x' } }));
  await assert.rejects(provider.listFolders(), AuthExpired);
});

test('429 waits for Retry-After', async () => {
  const { provider, sleeps } = setup((u, c, i) => (i === 0 ? reply(429, undefined, { 'Retry-After': '7' }) : reply(200, { value: [] })));
  await provider.listFolders();
  assert.deepEqual(sleeps, [7000]);
});

test('errors carry status and message', async () => {
  const { provider } = setup(() => reply(400, { error: { message: 'bad rule' } }));
  await assert.rejects(provider.deleteRule('x'), /Outlook error 400: bad rule/);
});

test('listMessages defaults to inbox, newest first, maps fields', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [mail('1', 'a@x.com', ['me+uber@outlook.com', 'b@c.d'])], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/next' }));
  const page = await provider.listMessages({});
  assert.equal(path(calls[0].url), '/mailFolders/inbox/messages');
  assert.equal(calls[0].url.searchParams.get('$orderby'), 'receivedDateTime desc');
  assert.equal(page.nextPageToken, 'https://graph.microsoft.com/v1.0/me/next');
  assert.deepEqual(page.messages[0], {
    id: '1', from: 'Sender <a@x.com>', to: 'me+uber@outlook.com, b@c.d', deliveredTo: '', subject: 's1',
    date: Date.parse('2026-11-01T10:00:00Z'), snippet: 'p1', unread: true });
});

test('listMessages filter searches all folders', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [] }));
  await provider.listMessages({ filter: { to: 'me+a@outlook.com', from: 'x.com' } });
  assert.equal(path(calls[0].url), '/messages');
  assert.equal(calls[0].url.searchParams.get('$search'), '"to:me+a@outlook.com AND from:x.com"');
  assert.equal(calls[0].url.searchParams.get('$orderby'), null);
});

test('listMessages trash and all', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [] }));
  await provider.listMessages({ filter: { trash: true } });
  await provider.listMessages({ folderId: 'ALL' });
  assert.equal(path(calls[0].url), '/mailFolders/deleteditems/messages');
  assert.equal(path(calls[1].url), '/messages');
});

test('page token must point at graph', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [] }));
  await assert.rejects(provider.listMessages({ pageToken: 'https://evil.example/steal' }), /page/i);
  assert.equal(calls.length, 0);
});

test('getMessage splits html and text', async () => {
  const { provider } = setup(() => reply(200, mail('1', 'a@x.com', ['me@outlook.com'], { body: { contentType: 'html', content: '<b>x</b>' } })));
  const m = await provider.getMessage('1');
  assert.equal(m.html, '<b>x</b>');
  assert.equal(m.text, null);
});

test('trash moves each mail to deleted items', async () => {
  const { provider, calls } = setup(() => reply(201, {}));
  await provider.trash(['a', 'b']);
  assert.deepEqual(calls.map(c => path(c.url)).sort(), ['/messages/a/move', '/messages/b/move']);
  assert.deepEqual(calls[0].body, { destinationId: 'deleteditems' });
});

test('markRead patches isRead', async () => {
  const { provider, calls } = setup(() => reply(200, {}));
  await provider.markRead(['a'], false);
  assert.equal(calls[0].init.method, 'PATCH');
  assert.deepEqual(calls[0].body, { isRead: false });
});

test('createFolder returns existing', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [{ id: 'F1', displayName: 'Subscriptions' }] }));
  assert.deepEqual(await provider.createFolder('subscriptions'), { id: 'F1', name: 'Subscriptions' });
  assert.ok(calls.every(c => c.init.method === 'GET'));
});

test('createRule posts to inbox rules', async () => {
  const { provider, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(201, { id: 'R1' }) : reply(200, { value: [] })));
  assert.deepEqual(await provider.createRule({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { delete: true } }), { id: 'R1' });
  const post = calls.find(c => c.init.method === 'POST');
  assert.equal(path(post.url), '/mailFolders/inbox/messageRules');
  assert.deepEqual(post.body.conditions, { subjectContains: ['x'] });
  assert.deepEqual(post.body.actions, { delete: true });
});

test('alias condition matches on recipient text', () => {
  const { provider } = setup(() => reply(200, {}));
  assert.deepEqual(provider.aliasCondition('me+a@outlook.com'), { recipientContains: ['me+a@outlook.com'] });
});

test('scanAliases reads pages, not single mails', async () => {
  const { provider, calls } = setup(u => {
    if (path(u) === '') return reply(200, { mail: 'me@outlook.com' });
    return u.searchParams.get('page') === '2'
      ? reply(200, { value: [mail('2', 'b@y.com', ['me@outlook.com'])] })
      : reply(200, { value: [mail('1', 'no-reply@mail.uber.com', ['me+uber@outlook.com'])], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?page=2' });
  });
  const known = new Map();
  assert.deepEqual(await provider.scanAliases(known), [
    { id: '1', alias: 'uber', domain: 'uber.com' },
    { id: '2', alias: null, domain: 'y.com' }]);
  assert.equal(calls.length, 3);
});

const simple = { conditions: { recipientContains: ['me+a@outlook.com'] }, exceptions: {}, actions: { delete: true } };
test('applyRuleToExisting for a simple rule', async () => {
  const { provider, calls } = setup((u, c) => (c.init.method === 'POST' ? reply(201, {}) : reply(200, { value: [{ id: '1' }, { id: '2' }] })));
  assert.equal(await provider.applyRuleToExisting(simple), 2);
  assert.equal(calls[0].url.searchParams.get('$search'), '"(to:me+a@outlook.com)"');
  assert.equal(calls.filter(c => path(c.url).endsWith('/move')).length, 2);
});
test('rules Outlook cannot replay on old mail report zero', async () => {
  const { provider, calls } = setup(() => reply(200, { value: [] }));
  const hard = { conditions: { importance: 'high' }, exceptions: {}, actions: { delete: true } };
  assert.deepEqual(await provider.countMatching(hard), { count: 0, more: false });
  assert.equal(await provider.applyRuleToExisting(hard), 0);
  assert.equal(calls.length, 0);
});

// ---- gmail side of the shared filter ----
test('gmail turns a filter into a search', async () => {
  const { provider, calls } = setup(() => reply(200, {}), createGmail);
  await provider.listMessages({ filter: { to: 'me+a@gmail.com', from: 'x.com' } });
  await provider.listMessages({ filter: { trash: true } });
  assert.equal(calls[0].url.searchParams.get('q'), 'to:("me+a@gmail.com") from:("x.com")');
  assert.equal(calls[0].url.searchParams.get('labelIds'), null);
  assert.equal(calls[1].url.searchParams.get('q'), 'in:trash');
  assert.deepEqual(provider.aliasCondition('me+a@gmail.com'), { sentToAddresses: ['me+a@gmail.com'] });
});
