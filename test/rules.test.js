import test from 'node:test';
import assert from 'node:assert/strict';
import { predicatesToQuery, toGmailFilter, describeGmailFilter, UnsupportedOnGmail, RULE_CATALOG, GMAIL_SUPPORT } from '../src/rules.js';

test('to address', () => assert.equal(predicatesToQuery({ sentToAddresses: ['me+spotify@gmail.com'] }), 'to:("me+spotify@gmail.com")'));
test('or within, and across', () => assert.equal(
  predicatesToQuery({ fromAddresses: ['a@x.com', 'b@y.com'], subjectContains: ['bill'] }),
  'from:("a@x.com" OR "b@y.com") subject:("bill")'));
test('body words', () => assert.equal(predicatesToQuery({ bodyContains: ['x y'] }), '("x y")'));
test('attachment', () => {
  assert.equal(predicatesToQuery({ hasAttachments: true }), 'has:attachment');
  assert.equal(predicatesToQuery({ hasAttachments: false }), '-has:attachment');
});
test('size is KB in, bytes out', () => assert.equal(
  predicatesToQuery({ withinSizeRange: { minimumSize: 10, maximumSize: 20 } }), 'larger:10240 smaller:20480'));
test('quotes cannot widen a rule', () => assert.equal(
  predicatesToQuery({ subjectContains: ['a" OR from:x'] }), 'subject:("a OR from:x")'));
test('empty values are skipped', () => assert.equal(predicatesToQuery({ subjectContains: ['', '  '] }), ''));
test('unsupported predicate', () => assert.throws(
  () => predicatesToQuery({ importance: 'high' }), e => e instanceof UnsupportedOnGmail && e.key === 'importance'));

test('leak rule', () => assert.deepEqual(
  toGmailFilter({ conditions: { sentToAddresses: ['me+spotify@gmail.com'] }, exceptions: { senderContains: ['spotify.com'] }, actions: { delete: true } }),
  { criteria: { query: 'to:("me+spotify@gmail.com")', negatedQuery: 'from:("spotify.com")' }, action: { addLabelIds: ['TRASH'] } }));
test('move and mark read', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { moveToFolder: 'Label_1', markAsRead: true } }).action,
  { addLabelIds: ['Label_1'], removeLabelIds: ['INBOX', 'UNREAD'] }));
test('copy keeps inbox', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { copyToFolder: 'Label_1' } }).action,
  { addLabelIds: ['Label_1'] }));
test('importance', () => {
  const f = v => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { markImportance: v } }).action;
  assert.deepEqual(f('high'), { addLabelIds: ['IMPORTANT'] });
  assert.deepEqual(f('low'), { removeLabelIds: ['IMPORTANT'] });
});
test('gmail extras', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { gmailStar: true, gmailNeverSpam: true } }).action,
  { addLabelIds: ['STARRED'], removeLabelIds: ['SPAM'] }));
test('forward one address only', () => {
  const f = a => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { forwardTo: a } });
  assert.equal(f(['a@b.c']).action.forward, 'a@b.c');
  assert.throws(() => f(['a@b.c', 'd@e.f']), UnsupportedOnGmail);
});
test('permanent delete refused', () => assert.throws(
  () => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { permanentDelete: true } }), UnsupportedOnGmail));
test('rule without conditions refused', () => assert.throws(
  () => toGmailFilter({ conditions: {}, exceptions: {}, actions: { delete: true } }), /condition/));
test('rule without actions refused', () => assert.throws(
  () => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: {} }), /action/));
test('false action is no action', () => assert.throws(
  () => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { delete: false } }), /action/));
test('catalog covers every supported key', () => {
  const p = new Set(RULE_CATALOG.predicates.map(x => x.key)), a = new Set(RULE_CATALOG.actions.map(x => x.key));
  for (const k of GMAIL_SUPPORT.predicates) assert.ok(p.has(k), k);
  for (const k of GMAIL_SUPPORT.actions) assert.ok(a.has(k), k);
});
test('describe', () => assert.equal(
  describeGmailFilter(
    { criteria: { query: 'to:("a")', negatedQuery: 'from:("b")' }, action: { addLabelIds: ['L1'], removeLabelIds: ['INBOX'] } },
    new Map([['L1', 'Subscriptions']])),
  'If to:("a") except from:("b"): add Subscriptions, skip inbox'));
test('describe native gmail fields', () => assert.equal(
  describeGmailFilter({ criteria: { from: 'x@y.z', hasAttachment: true }, action: { addLabelIds: ['TRASH'] } }, new Map()),
  'If from:x@y.z has:attachment: trash'));
import { ruleToSearch } from '../src/rules.js';
test('search for existing mail negates each exception', () => assert.equal(
  ruleToSearch({ conditions: { sentToAddresses: ['a@b.c'] },
    exceptions: { senderContains: ['x.com'], hasAttachments: false, withinSizeRange: { minimumSize: 1, maximumSize: 2 } } }),
  'to:("a@b.c") -from:("x.com") has:attachment -(larger:1024 smaller:2048)'));
test('several exceptions are OR in the filter', () => assert.equal(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: { senderContains: ['a'], subjectContains: ['b'] }, actions: { delete: true } }).criteria.negatedQuery,
  'from:("a") OR subject:("b")'));
