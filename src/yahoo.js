import { extractAlias, senderDomain, senderAddress } from './alias.js';
import { AuthExpired } from './gmail.js';
import { buildMime, parseAddresses } from './compose.js';

const PAGE = 50;
const SCAN_LIMIT = 2000;
const FETCH_BATCH = 200; // mails per bridge call: one IMAP connection, so few large calls beat many small ones
const NO_RULES = 'Yahoo gives apps no access to its filters. Make this rule in Yahoo Mail under Settings > Filters.';

// The part of a rule that IMAP search can run: senders, one recipient, one subject. Null for anything else.
function ruleCriteria(rule) {
  const c = rule.conditions || {};
  const set = v => v !== undefined && v !== null && v !== false && !(Array.isArray(v) && !v.length);
  if (Object.values(rule.exceptions || {}).some(set)) return null;
  const criteria = {};
  for (const [key, value] of Object.entries(c)) {
    if (!set(value)) continue;
    const list = [].concat(value);
    if (key === 'senderContains' || key === 'fromAddresses') criteria.from = [...(criteria.from || []), ...list];
    else if ((key === 'sentToAddresses' || key === 'recipientContains') && list.length === 1 && !criteria.to) criteria.to = list[0];
    else if (key === 'subjectContains' && list.length === 1) criteria.subject = list[0];
    else return null;
  }
  return Object.keys(criteria).length ? criteria : null;
}

/**
 * Yahoo Mail through the local IMAP bridge (serve.py, POST /imap). Same methods as the Gmail provider.
 * The app password is the token. Yahoo has no rules for apps: `rules: false` tells the app to hide them.
 */
export function createYahoo({ user, getToken, fetch: doFetch = (...a) => globalThis.fetch(...a) }) {
  let folderList;

  async function bridge(op, args = {}) {
    let res;
    try {
      res = await doFetch('/imap', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ user, password: getToken(), op, ...args }) });
    } catch {
      throw new Error('The local server does not answer. Is Start.bat running?');
    }
    const data = await res.json().catch(() => ({}));
    if (res.status === 401) throw new AuthExpired(data.error || 'Yahoo refused the login.');
    if (res.status === 404 || res.status === 501) throw new Error('The local server is too old for Yahoo. Close Start.bat and start it again.');
    if (!res.ok) throw new Error(data.error || `Yahoo error ${res.status}`);
    return data.result;
  }

  const folders = async () => (folderList ??= await bridge('folders'));
  const role = async name => (await folders()).find(f => f.role === name)?.name;
  // Everything except Trash, Spam and Drafts, like Gmail's All Mail.
  const everywhere = async () => (await folders()).filter(f => !['Trash', 'Junk', 'Drafts'].includes(f.role)).map(f => f.name);
  const folderOf = id => id.split(':').slice(2).join(':');

  // Moves and returns a receipt. Mail the server gave no new id for cannot be moved back, so it is left out.
  async function move(ids, to, unread = new Set()) {
    if (!ids.length) return { count: 0, moves: [] };
    const moved = await bridge('move', { ids, to });
    return { count: ids.length, moves: moved.filter(m => m.new).map(m => ({ id: m.new, back: folderOf(m.old), unread: unread.has(m.old) })) };
  }

  async function where({ folderId, query, filter = {} }) {
    if (filter.trash) return [await role('Trash')];
    if (filter.spam) return [await role('Junk')];
    if (folderId && folderId !== 'ALL') return [folderId];
    return folderId === 'ALL' || query || filter.to || filter.from || filter.before || filter.larger || filter.unread ? everywhere() : ['INBOX'];
  }
  const criteriaOf = (query, filter = {}) => ({
    text: query || undefined, to: filter.to, from: filter.from && [].concat(filter.from),
    before: filter.before, larger: filter.larger, unseen: filter.unread || undefined,
  });

  function cleanupCriteria({ before, larger, unread }) {
    if (!before && !larger && !unread) throw new Error('Pick at least one cleanup condition.');
    return { before, larger, unseen: unread || undefined };
  }

  return {
    name: 'Yahoo',
    rules: false,
    support: { predicates: new Set(), actions: new Set() },
    validateRule: () => { throw new Error(NO_RULES); },
    aliasCondition: address => ({ sentToAddresses: [address] }),
    async ownAddress() { await folders(); return user; }, // the first call also checks the login

    /** pageToken is the offset into the newest-first hit list. */
    async listMessages({ folderId, query, filter = {}, pageToken } = {}) {
      const offset = Number(pageToken) || 0;
      const { hits, more } = await bridge('search', { folders: await where({ folderId, query, filter }), criteria: criteriaOf(query, filter), limit: offset + PAGE });
      const page = hits.slice(offset, offset + PAGE);
      const messages = page.length ? await bridge('fetch', { ids: page.map(h => h.id) }) : [];
      return { messages, nextPageToken: more ? String(offset + PAGE) : null };
    },

    async getMessage(id) {
      const [mail] = await bridge('fetch', { ids: [id], body: true });
      if (!mail) throw new Error('This mail is gone from that folder. Reload the list.');
      return mail;
    },

    trash: async ids => move(ids, await role('Trash')),
    moveToFolder: (ids, folderId) => move(ids, folderId),
    markRead: (ids, read) => bridge('flag', { ids, seen: read }),
    async undo(receipt) {
      const unread = receipt.moves.filter(m => m.unread).map(m => m.id);
      if (unread.length) await bridge('flag', { ids: unread, seen: false });
      const byFolder = new Map();
      for (const m of receipt.moves) byFolder.set(m.back, [...(byFolder.get(m.back) || []), m.id]);
      for (const [folder, ids] of byFolder) await bridge('move', { ids, to: folder });
    },

    async listFolders() {
      return (await folders()).filter(f => !f.role).map(f => ({ id: f.name, name: f.name })).sort((a, b) => a.name.localeCompare(b.name));
    },
    async createFolder(name) {
      const wanted = name.trim();
      const existing = (await folders()).find(f => f.name.toLowerCase() === wanted.toLowerCase());
      if (existing) return { id: existing.name, name: existing.name };
      const made = await bridge('create', { name: wanted });
      folderList = null;
      return made;
    },

    listRules: async () => [],
    createRule: async () => { throw new Error(NO_RULES); },
    updateRule: async () => { throw new Error(NO_RULES); },
    deleteRule: async () => { throw new Error(NO_RULES); },

    // "Trash all" and "Move" for senders use these: the rule only describes which mail, nothing is saved in Yahoo.
    async countMatching(rule) {
      const criteria = ruleCriteria(rule);
      if (!criteria) return { count: 0, more: false };
      const { hits } = await bridge('search', { folders: await everywhere(), criteria });
      return { count: hits.length, more: false };
    },
    async applyRuleToExisting(rule) {
      const criteria = ruleCriteria(rule);
      const { delete: toTrash, moveToFolder, markAsRead } = rule.actions || {};
      if (!criteria) return { count: 0, moves: [] };
      const ids = (await bridge('search', { folders: await everywhere(), criteria })).hits.map(h => h.id);
      let unread = new Set();
      if (markAsRead) {
        unread = new Set((await bridge('search', { folders: await everywhere(), criteria: { ...criteria, unseen: true } })).hits.map(h => h.id));
        if (unread.size) await bridge('flag', { ids: [...unread], seen: true });
      }
      const to = toTrash ? await role('Trash') : moveToFolder;
      if (to) return move(ids, to, unread);
      return { count: ids.length, moves: [] }; // shortcut: mark-read alone has no undo here, like the Mark read button
    },

    async countMessages(filter) {
      const { hits } = await bridge('search', { folders: await everywhere(), criteria: cleanupCriteria(filter) });
      return { count: hits.length, more: false };
    },
    async trashMessages(filter) {
      const { hits } = await bridge('search', { folders: await everywhere(), criteria: cleanupCriteria(filter) });
      return move(hits.map(h => h.id), await role('Trash'));
    },

    folderCount: async which => bridge('count', { folder: await role(which === 'trash' ? 'Trash' : 'Junk') }),
    emptyFolder: async which => bridge('empty', { folder: await role(which === 'trash' ? 'Trash' : 'Junk') }),

    /** Newest `limit` mails over all folders except Trash, Spam and Drafts. Ids hold the folder, so `folders` is always current. */
    async scanAliases(known, onProgress = () => {}, limit = SCAN_LIMIT) {
      const { hits } = await bridge('search', { folders: await everywhere(), criteria: {}, limit: limit === Infinity ? undefined : limit });
      const missing = hits.map(h => h.id).filter(id => !known.has(id));
      let done = hits.length - missing.length;
      for (let i = 0; i < missing.length; i += FETCH_BATCH) {
        for (const m of await bridge('fetch', { ids: missing.slice(i, i + FETCH_BATCH) })) {
          known.set(m.id, { id: m.id, alias: extractAlias(m.deliveredTo, user) ?? extractAlias(m.to, user), domain: senderDomain(m.from),
            from: senderAddress(m.from), size: m.size, folders: [folderOf(m.id)] });
        }
        done += Math.min(FETCH_BATCH, missing.length - i);
        onProgress(done, hits.length);
      }
      return hits.map(h => known.get(h.id)).filter(Boolean);
    },

    async sendMail({ to, subject, text, replyTo }) {
      const list = parseAddresses([].concat(to).join(', '));
      await bridge('send', { to: list, raw: buildMime({ to: list, subject, text, replyTo }) });
    },
  };
}
