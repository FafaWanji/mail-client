// No scripts, no same-origin. Popups only so that links in a mail open in a normal new tab.
export const FRAME_SANDBOX = 'allow-popups allow-popups-to-escape-sandbox';

function decode(base64url) {
  const binary = atob(base64url.replaceAll('-', '+').replaceAll('_', '/'));
  return new TextDecoder().decode(Uint8Array.from(binary, c => c.charCodeAt(0)));
}

/** First text/html and first text/plain part of a Gmail payload tree, attachments skipped. */
export function pickBody(payload) {
  const out = { html: null, text: null };
  const walk = part => {
    if (!part) return;
    if (!part.filename && part.body?.data) {
      if (part.mimeType === 'text/html' && out.html === null) out.html = decode(part.body.data);
      if (part.mimeType === 'text/plain' && out.text === null) out.text = decode(part.body.data);
    }
    (part.parts || []).forEach(walk);
  };
  walk(payload);
  return out;
}

const escapeHtml = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Document for the mail iframe. The policy is the first element, so it applies to everything in the mail. */
export function buildFrameDoc(body, allowImages) {
  const policy = `default-src 'none'; style-src 'unsafe-inline'; img-src data:${allowImages ? ' https:' : ''}`;
  const content =
    body.html !== null ? body.html
    : body.text !== null ? `<pre style="white-space:pre-wrap;font:inherit">${escapeHtml(body.text)}</pre>`
    : '<p>(no content)</p>';
  return `<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="${policy}">` +
    '<meta charset="utf-8"><base target="_blank">' +
    '<style>body{font:14px system-ui,sans-serif;margin:12px;background:#fff;color:#111;overflow-wrap:anywhere}img{max-width:100%}</style>' +
    `</head><body>${content}</body></html>`;
}
