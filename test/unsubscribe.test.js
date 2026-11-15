import test from 'node:test';
import assert from 'node:assert/strict';
import { parseUnsubscribe, mailtoParts } from '../src/unsubscribe.js';

test('one-click needs the post header and an https link', () => {
  assert.deepEqual(parseUnsubscribe('<mailto:u@list.com?subject=stop>, <https://list.com/u?id=1>', 'List-Unsubscribe=One-Click'),
    { oneClick: 'https://list.com/u?id=1', link: 'https://list.com/u?id=1', mailto: 'mailto:u@list.com?subject=stop' });
  assert.equal(parseUnsubscribe('<https://list.com/u>', '').oneClick, null);
});

test('only https and mailto links are kept', () => {
  assert.deepEqual(parseUnsubscribe('<http://list.com/u>, <javascript:alert(1)>, <not a url>', 'List-Unsubscribe=One-Click'),
    { oneClick: null, link: null, mailto: null });
  assert.deepEqual(parseUnsubscribe(undefined), { oneClick: null, link: null, mailto: null });
});

test('mailto parts', () => {
  assert.deepEqual(mailtoParts('mailto:u@list.com?subject=stop%20it'), { to: 'u@list.com', subject: 'stop it', body: 'unsubscribe' });
  assert.deepEqual(mailtoParts('mailto:u@list.com'), { to: 'u@list.com', subject: 'unsubscribe', body: 'unsubscribe' });
});
