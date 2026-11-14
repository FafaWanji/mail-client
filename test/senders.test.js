import test from 'node:test';
import assert from 'node:assert/strict';
import { senderReport, senderKey } from '../src/senders.js';
import { senderAddress } from '../src/alias.js';

const rec = (from, n = 1) => Array.from({ length: n }, (_, i) => ({ id: `${from}${i}`, alias: null, domain: from.split('@')[1].split('.').slice(-2).join('.'), from }));

test('companies group by domain, most mail first', () => assert.deepEqual(
  senderReport([...rec('a@news.shop.com', 2), ...rec('b@shop.com', 3), ...rec('x@other.org', 1)], 'me@gmail.com'),
  [{ key: 'shop.com', count: 5 }, { key: 'other.org', count: 1 }]));

test('freemail senders stay separate people', () => assert.deepEqual(
  senderReport([...rec('anna@gmail.com', 2), ...rec('ben@gmail.com', 1)], 'me@gmail.com'),
  [{ key: 'anna@gmail.com', count: 2 }, { key: 'ben@gmail.com', count: 1 }]));

test('own sent mail is not a sender', () => assert.deepEqual(senderReport(rec('me@gmail.com', 4), 'me@gmail.com'), []));

test('records without sender are skipped', () => assert.deepEqual(
  senderReport([{ id: '1', alias: null, domain: null, from: '' }], 'me@gmail.com'), []));

test('senderKey', () => {
  assert.equal(senderKey({ domain: 'shop.com', from: 'a@shop.com' }), 'shop.com');
  assert.equal(senderKey({ domain: 'outlook.com', from: 'a@outlook.com' }), 'a@outlook.com');
});

test('senderAddress', () => {
  assert.equal(senderAddress('Shop <News@Shop.com>'), 'news@shop.com');
  assert.equal(senderAddress('garbage'), '');
});
