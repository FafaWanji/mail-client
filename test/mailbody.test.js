import test from 'node:test';
import assert from 'node:assert/strict';
import { pickBody, buildFrameDoc, FRAME_SANDBOX } from '../src/mailbody.js';

const b64 = s => Buffer.from(s, 'utf8').toString('base64url');
const part = (mimeType, s) => ({ mimeType, body: { data: b64(s) } });

test('nested multipart prefers html', () => {
  const payload = { mimeType: 'multipart/mixed', parts: [
    { mimeType: 'multipart/alternative', parts: [part('text/plain', 'hi'), part('text/html', '<b>hi</b>')] },
    { mimeType: 'application/pdf', filename: 'a.pdf', body: { attachmentId: 'x' } }] };
  assert.deepEqual(pickBody(payload), { html: '<b>hi</b>', text: 'hi' });
});
test('single part plain, utf-8', () => assert.deepEqual(pickBody(part('text/plain', 'grüße')), { html: null, text: 'grüße' }));
test('text attachment is not the body', () => assert.deepEqual(
  pickBody({ mimeType: 'multipart/mixed', parts: [
    part('text/plain', 'body'),
    { mimeType: 'text/plain', filename: 'notes.txt', body: { data: b64('attached') } }] }),
  { html: null, text: 'body' }));
test('no body', () => assert.deepEqual(pickBody({ mimeType: 'multipart/mixed', parts: [] }), { html: null, text: null }));
test('sandbox never allows scripts or same-origin', () => {
  assert.doesNotMatch(FRAME_SANDBOX, /allow-scripts|allow-same-origin|allow-forms|allow-top-navigation/);
});
test('images blocked by default', () => {
  const doc = buildFrameDoc({ html: '<img src="https://t.example/p.gif">', text: null }, false);
  assert.match(doc, /Content-Security-Policy/);
  assert.match(doc, /default-src 'none'/);
  assert.doesNotMatch(doc, /img-src[^;"]*https:/);
});
test('policy comes before the mail', () => {
  const doc = buildFrameDoc({ html: '<p>MAILBODY</p>', text: null }, false);
  assert.ok(doc.indexOf('Content-Security-Policy') < doc.indexOf('MAILBODY'));
});
test('images on request', () => assert.match(buildFrameDoc({ html: '<p>x</p>', text: null }, true), /img-src[^;"]*https:/));
test('plain text is escaped', () => {
  const doc = buildFrameDoc({ html: null, text: '<script>alert(1)</script>' }, false);
  assert.doesNotMatch(doc, /<script>alert/);
  assert.match(doc, /&lt;script&gt;/);
});
test('empty body', () => assert.match(buildFrameDoc({ html: null, text: null }, false), /no content/i));
