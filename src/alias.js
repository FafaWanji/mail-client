// Gmail ignores dots in the local part; other providers do not.
const DOTLESS = new Set(['gmail.com', 'googlemail.com']);
const ADDRESS = /[^\s<>,;"']+@[^\s<>,;"']+/g;

function split(address) {
  const a = address.toLowerCase();
  const at = a.lastIndexOf('@');
  return [a.slice(0, at), a.slice(at + 1)];
}

/** The +tag of the owner's own address found in an address header, or null. */
export function extractAlias(headerValue, ownAddress) {
  const [ownLocal, ownDomain] = split(ownAddress);
  const norm = local => (DOTLESS.has(ownDomain) ? local.replaceAll('.', '') : local);
  for (const found of (headerValue || '').match(ADDRESS) || []) {
    const [local, domain] = split(found);
    const plus = local.indexOf('+');
    if (plus < 0 || domain !== ownDomain) continue;
    const tag = local.slice(plus + 1);
    if (tag && norm(local.slice(0, plus)) === norm(ownLocal)) return tag;
  }
  return null;
}

export function aliasAddress(ownAddress, alias) {
  const [local, domain] = split(ownAddress);
  return `${local}+${alias}@${domain}`;
}

/** The sender's address in lower case, or '' when the header has none. */
export function senderAddress(fromHeader) {
  const found = (fromHeader || '').match(ADDRESS);
  return found ? found[found.length - 1].toLowerCase() : '';
}

export function senderDomain(fromHeader) {
  const found = (fromHeader || '').match(ADDRESS);
  if (!found) return null;
  // shortcut: last two labels, wrong for co.uk style domains; switch to a public-suffix list if that bites
  return split(found[found.length - 1])[1].split('.').slice(-2).join('.');
}
