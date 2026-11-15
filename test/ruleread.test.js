import test from 'node:test';
import assert from 'node:assert/strict';
import { toGmailFilter, fromGmailFilter, toOutlookRule, fromOutlookRule } from '../src/rules.js';

const rule = {
  conditions: { fromAddresses: ['a@b.com', 'c@d.com'], recipientContains: ['shop'], subjectContains: ['sale', 'deal'], bodyContains: ['x y'], hasAttachments: true, withinSizeRange: { minimumSize: 10, maximumSize: 20 } },
  exceptions: { senderContains: ['boss.com'], withinSizeRange: { minimumSize: 0, maximumSize: 5 }, hasAttachments: false, sentToAddresses: ['me+a@gmail.com'] },
  actions: { moveToFolder: 'L1', markAsRead: true, gmailStar: true, forwardTo: ['f@g.com'] },
};

test('a filter this app made reads back to the same filter', () => {
  const back = fromGmailFilter({ id: 'F', ...toGmailFilter(rule) });
  assert.ok(back);
  assert.deepEqual(toGmailFilter(back), toGmailFilter(rule));
  assert.deepEqual(back.actions, rule.actions);
});

test('copy to folder, trash and importance read back', () => {
  const r = { conditions: { subjectContains: ['a'] }, exceptions: {}, actions: { copyToFolder: 'L2', delete: true, markImportance: 'low' } };
  assert.deepEqual(fromGmailFilter(toGmailFilter(r)), r);
});

test('a filter made in Gmail itself is not editable', () => {
  assert.equal(fromGmailFilter({ criteria: { from: 'a@b.com' }, action: { addLabelIds: ['TRASH'] } }), null);
  assert.equal(fromGmailFilter({ criteria: { query: 'from:a@b.com' }, action: { addLabelIds: ['TRASH'] } }), null);
  assert.equal(fromGmailFilter({ criteria: { query: 'larger:1000' }, action: { addLabelIds: ['TRASH'] } }), null);
});

test('outlook rule reads back', () => {
  const r = { conditions: { fromAddresses: ['a@b.com'], importance: 'high' }, exceptions: {}, actions: { moveToFolder: 'F1', stopProcessingRules: true } };
  assert.deepEqual(fromOutlookRule(toOutlookRule(r, new Map())), r);
  assert.equal(fromOutlookRule({ conditions: { somethingNew: true }, actions: { delete: true } }), null);
});
