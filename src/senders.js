// On these domains every address is a different person, so they are never lumped together.
const FREEMAIL = new Set(['gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'live.com', 'msn.com',
  'yahoo.com', 'yahoo.de', 'icloud.com', 'me.com', 'aol.com', 'gmx.de', 'gmx.net', 'web.de', 't-online.de', 'proton.me', 'protonmail.com']);

/** What a scan record is grouped and searched by: the domain for companies, the full address for freemail. */
export function senderKey(record) {
  return record.domain && !FREEMAIL.has(record.domain) ? record.domain : record.from || null;
}

/** Senders with their mail count, most mail first. Mail the owner sent is left out. */
export function senderReport(records, ownAddress) {
  const counts = new Map();
  for (const record of records) {
    const key = senderKey(record);
    if (key && record.from !== ownAddress) counts.set(key, (counts.get(key) || 0) + 1);
  }
  return [...counts].map(([key, count]) => ({ key, count }))
    .sort((a, b) => b.count - a.count || (a.key < b.key ? -1 : 1));
}
