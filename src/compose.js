// Outgoing mail. Every header value passes oneLine(), so a pasted line break cannot add a header.

const oneLine = value => String(value ?? '').replace(/[\r\n]+/g, ' ').trim();
const ADDRESS = /^[^\s@<>(),;:"]+@[^\s@<>(),;:"]+\.[^\s@<>(),;:"]+$/;

/** Splits "a@b.com, c@d.com" into addresses. Throws with a readable message on anything else. */
export function parseAddresses(text) {
  const list = oneLine(text).split(/[,;]/).map(s => s.trim()).filter(Boolean);
  if (!list.length) throw new Error('Add at least one address.');
  const bad = list.find(a => !ADDRESS.test(a));
  if (bad) throw new Error(`"${bad}" is not a mail address.`);
  return list;
}

const utf8Base64 = text => {
  const bytes = new TextEncoder().encode(text);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
};
// RFC 2047, only when the subject is not plain ASCII.
const encodeWord = text => (/^[\x20-\x7e]*$/.test(text) ? text : `=?UTF-8?B?${utf8Base64(text)}?=`);

/** A plain-text RFC 5322 message. replyTo: { messageId, references } of the mail answered. */
export function buildMime({ to, subject, text, replyTo }) {
  const headers = [
    `To: ${parseAddresses([].concat(to).join(', ')).join(', ')}`,
    `Subject: ${encodeWord(oneLine(subject))}`,
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
  ];
  const parent = oneLine(replyTo?.messageId);
  if (parent) headers.push(`In-Reply-To: ${parent}`, `References: ${[oneLine(replyTo.references), parent].filter(Boolean).join(' ')}`);
  const body = utf8Base64(String(text ?? '').replace(/\r?\n/g, '\r\n')).replace(/.{76}/g, '$&\r\n');
  return `${headers.join('\r\n')}\r\n\r\n${body}`;
}

export const base64url = text => utf8Base64(text).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '');

/** "Re: " once, not "Re: Re: ". */
export const replySubject = subject => (/^re:/i.test(subject || '') ? subject : `Re: ${subject || ''}`.trim());
