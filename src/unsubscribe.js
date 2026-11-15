/**
 * Reads the List-Unsubscribe header (RFC 2369) and List-Unsubscribe-Post (RFC 8058).
 * Returns { oneClick, link, mailto }: oneClick is an https URL that takes a POST, link an https URL to open,
 * mailto an address URL. Missing ones are null. Only https links are kept: the header comes from the sender.
 */
export function parseUnsubscribe(header, postHeader) {
  const out = { oneClick: null, link: null, mailto: null };
  for (const [, raw] of String(header || '').matchAll(/<([^>]+)>/g)) {
    let url;
    try { url = new URL(raw.trim()); } catch { continue; }
    if (url.protocol === 'https:' && !out.link) out.link = url.href;
    if (url.protocol === 'mailto:' && !out.mailto) out.mailto = url.href;
  }
  if (out.link && /^\s*List-Unsubscribe=One-Click\s*$/i.test(postHeader || '')) out.oneClick = out.link;
  return out;
}

/** Address, subject and body of a mailto: unsubscribe URL. */
export function mailtoParts(mailto) {
  const url = new URL(mailto);
  return {
    to: decodeURIComponent(url.pathname),
    subject: url.searchParams.get('subject') || 'unsubscribe',
    body: url.searchParams.get('body') || 'unsubscribe',
  };
}
