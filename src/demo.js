// In-memory stand-in for a mail account, opened with ?demo=1. No network, nothing is saved.
import { extractAlias, senderDomain } from './alias.js';
import { toGmailFilter, describeGmailFilter, GMAIL_SUPPORT } from './rules.js';

const ME = 'demo@gmail.com';
const SAMPLE = [
  ['Spotify <no-reply@spotify.com>', 'demo+spotify@gmail.com', 'Your receipt', 'Thanks for your payment.'],
  ['Spotify <no-reply@spotify.com>', 'demo+spotify@gmail.com', 'New releases for you', 'Fresh music this week.'],
  ['Random Shop <deals@randomshop.com>', 'demo+spotify@gmail.com', '90% OFF everything', '<p>Buy <b>now</b></p><img src="https://example.com/track.gif"><script>alert(1)</script>'],
  ['Uber <no-reply@uber.com>', 'demo+uber@gmail.com', 'Your Tuesday trip', 'Total: 12.40'],
  ['Netflix <info@netflix.com>', 'demo+netflix@gmail.com', 'New sign-in', 'A new device signed in.'],
  ['Netflix <info@netflix.com>', 'demo+netflix@gmail.com', 'Coming soon', 'New this month.'],
  ['Alex <alex@example.org>', 'demo@gmail.com', 'Lunch?', 'Friday at 12?'],
];

export function createDemo() {
  let nextId = 1;
  const mails = SAMPLE.map(([from, to, subject, body], i) => ({
    id: `m${i}`, from, to, deliveredTo: to, subject, date: Date.now() - i * 3600e3, snippet: body.replace(/<[^>]+>/g, ' '),
    unread: i % 2 === 0, labels: new Set(['INBOX']),
    html: body.startsWith('<') ? body : null, text: body.startsWith('<') ? null : body,
  }));
  const folders = [], rules = [];
  const each = (ids, fn) => { for (const m of mails) if (ids.includes(m.id)) fn(m); };
  const names = () => new Map(folders.map(f => [f.id, f.name]));

  return {
    name: 'Gmail',
    support: GMAIL_SUPPORT,
    validateRule: toGmailFilter,
    ownAddress: async () => ME,
    async listMessages({ folderId, query } = {}) {
      const trash = query === 'in:trash';
      const words = trash || !query ? [] : [...query.matchAll(/"([^"]+)"/g)].map(m => m[1].toLowerCase());
      const messages = mails.filter(m =>
        m.labels.has('TRASH') === trash &&
        (trash || query || folderId === 'ALL' || m.labels.has(folderId || 'INBOX')) &&
        words.every(w => `${m.from} ${m.to} ${m.subject}`.toLowerCase().includes(w)));
      return { messages, nextPageToken: null };
    },
    getMessage: async id => mails.find(m => m.id === id),
    trash: async ids => each(ids, m => m.labels.add('TRASH')),
    markRead: async (ids, read) => each(ids, m => { m.unread = !read; }),
    moveToFolder: async (ids, folderId) => each(ids, m => { m.labels.add(folderId); m.labels.delete('INBOX'); }),
    listFolders: async () => folders.slice(),
    async createFolder(name) {
      let folder = folders.find(f => f.name.toLowerCase() === name.trim().toLowerCase());
      if (!folder) folders.push(folder = { id: `L${nextId++}`, name: name.trim() });
      return folder;
    },
    listRules: async () => rules.map(r => ({ id: r.id, summary: describeGmailFilter(r.raw, names()), raw: r.raw })),
    async createRule(rule) {
      const id = `F${nextId++}`;
      rules.push({ id, raw: toGmailFilter(rule) });
      return { id };
    },
    deleteRule: async id => { rules.splice(rules.findIndex(r => r.id === id), 1); },
    applyRuleToExisting: async () => 0,
    countMatching: async () => ({ count: 0, more: false }),
    scanAliases: async () => mails.filter(m => !m.labels.has('TRASH'))
      .map(m => ({ id: m.id, alias: extractAlias(m.to, ME), domain: senderDomain(m.from) })),
  };
}
