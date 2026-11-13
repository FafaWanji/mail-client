import test from 'node:test';
import assert from 'node:assert/strict';
import { leakReport } from '../src/leaks.js';

const r = (alias, domain, n) => Array.from({ length: n }, (_, i) => ({ id: `${alias}${domain}${i}`, alias, domain }));
const recs = [...r('spotify', 'spotify.com', 5), ...r('spotify', 'randomshop.com', 2), ...r('amazon', 'amazon.de', 9)];

test('flags the foreign sender', () => {
  const [first, second] = leakReport(recs, {});
  assert.equal(first.alias, 'spotify');
  assert.equal(first.total, 7);
  assert.equal(first.home, 'spotify.com');
  assert.equal(first.flagged, true);
  assert.deepEqual(first.domains, [
    { domain: 'spotify.com', count: 5, trusted: true },
    { domain: 'randomshop.com', count: 2, trusted: false }]);
  assert.equal(second.alias, 'amazon');
  assert.equal(second.flagged, false);
});
test('trusting clears the flag', () =>
  assert.equal(leakReport(recs, { spotify: ['randomshop.com'] }).find(x => x.alias === 'spotify').flagged, false));
test('tie goes to alphabetical first', () =>
  assert.equal(leakReport([...r('a', 'zed.com', 1), ...r('a', 'abc.com', 1)], {})[0].home, 'abc.com'));
test('unknown domain', () =>
  assert.equal(leakReport([{ id: '1', alias: 'a', domain: null }], {})[0].home, '(unknown)'));
test('empty', () => assert.deepEqual(leakReport([], {}), []));
