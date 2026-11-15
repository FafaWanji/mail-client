import test from 'node:test';
import assert from 'node:assert/strict';
import { buildMime, parseAddresses, replySubject } from '../src/compose.js';

const decode = mime => {
  const [head, body] = mime.split('\r\n\r\n');
  return { head: head.split('\r\n'), body: Buffer.from(body.replace(/\r\n/g, ''), 'base64').toString('utf8') };
};

test('plain mail', () => {
  const { head, body } = decode(buildMime({ to: 'a@b.com', subject: 'Hi', text: 'line 1\nline 2' }));
  assert.deepEqual(head.slice(0, 2), ['To: a@b.com', 'Subject: Hi']);
  assert.equal(body, 'line 1\r\nline 2');
  assert.ok(!head.some(h => h.startsWith('In-Reply-To')));
});

test('a line break in a header cannot add a header', () => {
  const { head } = decode(buildMime({ to: 'a@b.com', subject: 'Hi\r\nBcc: evil@x.com', text: '' }));
  assert.equal(head[1], 'Subject: Hi Bcc: evil@x.com');
  assert.ok(!head.some(h => h.startsWith('Bcc')));
  assert.throws(() => buildMime({ to: 'a@b.com\nBcc: evil@x.com', subject: '', text: '' }), /not a mail address/);
});

test('non-ASCII subject and body', () => {
  const { head, body } = decode(buildMime({ to: 'a@b.com', subject: 'Grüße', text: 'Grüße 👋' }));
  assert.equal(head[1], `Subject: =?UTF-8?B?${Buffer.from('Grüße').toString('base64')}?=`);
  assert.equal(body, 'Grüße 👋');
});

test('reply threads on the parent', () => {
  const { head } = decode(buildMime({ to: 'a@b.com', subject: 'Re: x', text: '', replyTo: { messageId: '<2@b>', references: '<1@b>' } }));
  assert.ok(head.includes('In-Reply-To: <2@b>'));
  assert.ok(head.includes('References: <1@b> <2@b>'));
});

test('addresses', () => {
  assert.deepEqual(parseAddresses('a@b.com; c@d.org ,'), ['a@b.com', 'c@d.org']);
  assert.throws(() => parseAddresses(''), /at least one/);
  assert.throws(() => parseAddresses('Bob <b@c.com>'), /not a mail address/);
});

test('reply subject', () => {
  assert.equal(replySubject('Hello'), 'Re: Hello');
  assert.equal(replySubject('RE: Hello'), 'RE: Hello');
});
