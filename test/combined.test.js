import test from 'node:test';
import assert from 'node:assert/strict';
import { createCombined, AccountFailed } from '../src/combined.js';
import { createDemo } from '../src/demo.js';

function setup() {
  const a = { key: 'demo:a@gmail.com', address: 'a@gmail.com', provider: createDemo('a@gmail.com') };
  const b = { key: 'demo:b@work.com', address: 'b@work.com', provider: createDemo('b@work.com') };
  return { a, b, all: createCombined([a, b]) };
}

test('merged inbox, newest first, ids lead back to their account', async () => {
  const { all } = setup();
  const { messages } = await all.listMessages({});
  assert.equal(messages.length, 14);
  assert.ok(messages.every((m, i) => i === 0 || messages[i - 1].date >= m.date));
  const mail = await all.getMessage(messages.find(m => m.account === 'b@work.com').id);
  assert.equal(mail.account, 'b@work.com');
  assert.match(mail.to, /work\.com/);
});

test('trash across both accounts, then undo', async () => {
  const { all, a, b } = setup();
  const ids = (await all.listMessages({})).messages.slice(0, 4).map(m => m.id);
  const receipt = await all.trash(ids);
  assert.equal(receipt.count, 4);
  const left = async p => (await p.listMessages({})).messages.length;
  assert.equal(await left(a.provider) + await left(b.provider), 10);
  await all.undo(receipt);
  assert.equal(await left(a.provider) + await left(b.provider), 14);
});

test('a failing account keeps what the others did, for undo', async () => {
  const { all, a, b } = setup();
  b.provider.applyRuleToExisting = async () => { throw new Error('quota'); };
  const rule = { conditions: { senderContains: ['spotify.com'] }, exceptions: {}, actions: { delete: true } };
  const error = await all.applyRuleToExisting(rule).catch(e => e);
  assert.ok(error instanceof AccountFailed);
  assert.match(error.message, /^b@work\.com: quota/);
  assert.equal(error.partial.count, 2);
  assert.equal(error.partial.failed, true);
  await all.undo(error.partial);
  assert.equal((await a.provider.listMessages({})).messages.length, 7);
});

test('a rule is made in every account and deleted from every account', async () => {
  const { all, a, b } = setup();
  const made = await all.createRule({ conditions: { senderContains: ['x.com'] }, exceptions: {}, actions: { delete: true } });
  assert.equal((await a.provider.listRules()).length + (await b.provider.listRules()).length, 2);
  await all.deleteRule(made.id);
  assert.equal((await a.provider.listRules()).length + (await b.provider.listRules()).length, 0);
});

test('a reply goes out from the account that got the mail', async () => {
  const { all, b } = setup();
  const id = (await all.listMessages({})).messages.find(m => m.account === 'b@work.com').id;
  await all.sendMail({ to: 'x@y.com', subject: 'Re', text: 'hi', replyTo: await all.getMessage(id) });
  assert.equal(b.provider.sent.length, 1);
  assert.doesNotMatch(b.provider.sent[0].replyTo.id, /\u0001/);
});

test('accounts without rules for apps are skipped when a rule is made', async () => {
  const { all, a, b } = setup();
  b.provider.rules = false;
  b.provider.createRule = async () => { throw new Error('no rules'); };
  b.provider.validateRule = () => { throw new Error('no rules'); };
  const rule = { conditions: { senderContains: ['x.com'] }, exceptions: {}, actions: { delete: true } };
  all.validateRule(rule);
  await all.createRule(rule);
  assert.equal((await a.provider.listRules()).length, 1);
});
