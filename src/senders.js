// On these domains every address is a different person, so they are never lumped together.
const FREEMAIL = new Set(['gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'live.com', 'msn.com',
  'yahoo.com', 'yahoo.de', 'icloud.com', 'me.com', 'aol.com', 'gmx.de', 'gmx.net', 'web.de', 't-online.de', 'proton.me', 'protonmail.com']);

/** What a scan record is grouped and searched by: the domain for companies, the full address for freemail. */
export function senderKey(record) {
  return record.domain && !FREEMAIL.has(record.domain) ? record.domain : record.from || null;
}

/** Senders with their mail count and bytes, most mail first. Mail the owner sent is left out. */
export function senderReport(records, ownAddress) {
  const totals = new Map();
  for (const record of records) {
    const key = senderKey(record);
    if (!key || record.from === ownAddress) continue;
    const t = totals.get(key) || totals.set(key, { key, count: 0, size: 0 }).get(key);
    t.count++;
    t.size += record.size || 0;
  }
  return [...totals.values()].sort((a, b) => b.count - a.count || (a.key < b.key ? -1 : 1));
}

/** Mail count and bytes per folder id, from the scan records' `folders`. */
export function folderReport(records) {
  const totals = new Map();
  for (const record of records) {
    for (const id of record.folders || []) {
      const t = totals.get(id) || totals.set(id, { count: 0, size: 0 }).get(id);
      t.count++;
      t.size += record.size || 0;
    }
  }
  return totals;
}

/** Bytes as a short text: 512 B, 3.4 KB, 12 MB, 1.2 GB. */
export function sizeText(bytes) {
  const units = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  while (bytes >= 1024 && i < units.length - 1) { bytes /= 1024; i++; }
  return `${i && bytes < 10 ? bytes.toFixed(1) : Math.round(bytes)} ${units[i]}`;
}

/** Splits a list into pieces of at most `size`. */
export function chunks(list, size) {
  const out = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}
