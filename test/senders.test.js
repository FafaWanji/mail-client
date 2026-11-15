import test from 'node:test';
import assert from 'node:assert/strict';
import { senderReport, senderKey, chunks, folderReport, sizeText } from '../src/senders.js';
import { senderAddress } from '../src/alias.js';

const rec = (from, n = 1) => Array.from({ length: n }, (_, i) => ({ id: `${from}${i}`, alias: null, domain: from.split('@')[1].split('.').slice(-2).join('.'), from }));

test('companies group by domain, most mail first', () => assert.deepEqual(
  senderReport([...rec('a@news.shop.com', 2), ...rec('b@shop.com', 3), ...rec('x@other.org', 1)], 'me@gmail.com'),
  [{ key: 'shop.com', count: 5, size: 0 }, { key: 'other.org', count: 1, size: 0 }]));

test('freemail senders stay separate people', () => assert.deepEqual(
  senderReport([...rec('anna@gmail.com', 2), ...rec('ben@gmail.com', 1)], 'me@gmail.com'),
  [{ key: 'anna@gmail.com', count: 2, size: 0 }, { key: 'ben@gmail.com', count: 1, size: 0 }]));

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

test('chunks', () => {
  assert.deepEqual(chunks([1, 2, 3, 4, 5], 2), [[1, 2], [3, 4], [5]]);
  assert.deepEqual(chunks([], 2), []);
});

test('sizes add up per sender and per folder', () => {
  const records = [
    { from: 'a@shop.com', domain: 'shop.com', size: 1000, folders: ['INBOX', 'L1'] },
    { from: 'b@shop.com', domain: 'shop.com', size: 500, folders: ['L1'] },
    { from: 'c@x.org', domain: 'x.org' }];
  assert.deepEqual(senderReport(records, 'me@gmail.com')[0], { key: 'shop.com', count: 2, size: 1500 });
  assert.deepEqual(Object.fromEntries(folderReport(records)), { INBOX: { count: 1, size: 1000 }, L1: { count: 2, size: 1500 } });
});

test('sizeText', () => {
  assert.equal(sizeText(512), '512 B');
  assert.equal(sizeText(3500), '3.4 KB');
  assert.equal(sizeText(12 * 1048576), '12 MB');
  assert.equal(sizeText(1.25 * 1073741824), '1.3 GB');
});
