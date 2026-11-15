// In-memory stand-in for a mail account, opened with ?demo=1. No network, nothing is saved.
import { extractAlias, senderDomain, senderAddress } from './alias.js';
import { toGmailFilter, fromGmailFilter, describeGmailFilter, GMAIL_SUPPORT } from './rules.js';

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

const AGE_DAYS = [0, 1, 800, 2, 3, 1000, 5];
const SIZE = [30e3, 25e3, 6e6, 12e3, 8e3, 7.5e6, 3e3];
const cleanupHit = (m, f) => (!f.before || m.date < f.before) && (!f.larger || m.size > f.larger) && (!f.unread || m.unread);

export function createDemo() {
  let nextId = 1;
  const mails = SAMPLE.map(([from, to, subject, body], i) => ({
    id: `m${i}`, from, to, deliveredTo: to, subject, date: Date.now() - AGE_DAYS[i] * 864e5 - i * 3600e3, size: SIZE[i], snippet: body.replace(/<[^>]+>/g, ' '),
    unread: i % 2 === 0, labels: new Set(['INBOX']),
    html: body.startsWith('<') ? body : null, text: body.startsWith('<') ? null : body,
  }));
  const folders = [], rules = [];
  const each = (ids, fn) => { for (const m of mails) if (ids.includes(m.id)) fn(m); };
  // The demo's receipt is a copy of each mail's state before the change.
  const changed = (hit, fn) => {
    const before = hit.map(m => ({ m, labels: new Set(m.labels), unread: m.unread }));
    hit.forEach(fn);
    return { count: hit.length, before };
  };
  const pick = ids => mails.filter(m => ids.includes(m.id));
  const names = () => new Map(folders.map(f => [f.id, f.name]));
  // Only the two conditions the demo screens use.
  const matching = rule => mails.filter(m => !m.labels.has('TRASH') &&
    (!rule.conditions.senderContains || rule.conditions.senderContains.some(s => m.from.toLowerCase().includes(s))) &&
    (rule.conditions.sentToAddresses || []).every(a => m.to.toLowerCase().includes(a)));

  return {
    name: 'Gmail',
    support: GMAIL_SUPPORT,
    validateRule: toGmailFilter,
    ownAddress: async () => ME,
    aliasCondition: address => ({ sentToAddresses: [address] }),
    async listMessages({ folderId, query, filter = {} } = {}) {
      const trash = Boolean(filter.trash);
      const words = [query, filter.to].filter(Boolean).map(w => w.toLowerCase());
      const senders = [].concat(filter.from || []);
      const messages = mails.filter(m =>
        m.labels.has('TRASH') === trash &&
        (trash || words.length || senders.length || folderId === 'ALL' || m.labels.has(folderId || 'INBOX')) &&
        (!senders.length || senders.some(s => m.from.toLowerCase().includes(s))) &&
        words.every(w => `${m.from} ${m.to} ${m.subject}`.toLowerCase().includes(w)) && cleanupHit(m, filter));
      return { messages, nextPageToken: null };
    },
    countMessages: async filter => ({ count: mails.filter(m => !m.labels.has('TRASH') && cleanupHit(m, filter)).length, more: false }),
    trashMessages: async filter => changed(mails.filter(m => !m.labels.has('TRASH') && cleanupHit(m, filter)), m => m.labels.add('TRASH')),
    getMessage: async id => mails.find(m => m.id === id),
    trash: async ids => changed(pick(ids), m => m.labels.add('TRASH')),
    undo: async receipt => { for (const { m, labels, unread } of receipt.before) Object.assign(m, { labels, unread }); },
    markRead: async (ids, read) => each(ids, m => { m.unread = !read; }),
    moveToFolder: async (ids, folderId) => changed(pick(ids), m => { m.labels.add(folderId); m.labels.delete('INBOX'); }),
    listFolders: async () => folders.slice(),
    async createFolder(name) {
      let folder = folders.find(f => f.name.toLowerCase() === name.trim().toLowerCase());
      if (!folder) folders.push(folder = { id: `L${nextId++}`, name: name.trim() });
      return folder;
    },
    listRules: async () => rules.map(r => ({ id: r.id, summary: describeGmailFilter(r.raw, names()), raw: r.raw, rule: fromGmailFilter(r.raw), enabled: true })),
    async createRule(rule) {
      const id = `F${nextId++}`;
      rules.push({ id, raw: toGmailFilter(rule) });
      return { id };
    },
    deleteRule: async id => { rules.splice(rules.findIndex(r => r.id === id), 1); },
    async updateRule(id, rule) {
      const made = await this.createRule(rule);
      await this.deleteRule(id);
      return made;
    },
    async recreateRule(raw) {
      const id = `F${nextId++}`;
      rules.push({ id, raw: { criteria: raw.criteria, action: raw.action } });
      return { id };
    },
    async applyRuleToExisting(rule) {
      const { actions } = rule;
      return changed(matching(rule), m => {
        if (actions.markAsRead) m.unread = false;
        if (actions.delete) m.labels.add('TRASH');
        else if (actions.moveToFolder) { m.labels.add(actions.moveToFolder); m.labels.delete('INBOX'); }
      });
    },
    countMatching: async rule => ({ count: matching(rule).length, more: false }),
    scanAliases: async () => mails.filter(m => !m.labels.has('TRASH'))
      .map(m => ({ id: m.id, alias: extractAlias(m.to, ME), domain: senderDomain(m.from), from: senderAddress(m.from) })),
  };
}
