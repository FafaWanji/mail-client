import { extractAlias, senderDomain, senderAddress } from './alias.js';
import { toGmailFilter, fromGmailFilter, ruleToSearch, describeGmailFilter, GMAIL_SUPPORT } from './rules.js';
import { pickBody } from './mailbody.js';

const BASE = 'https://gmail.googleapis.com/gmail/v1/users/me';
const PAGE = 50;          // mails per list page
const SCAN_LIMIT = 2000;  // newest mails read by the alias scan
const BATCH = 1000;       // Gmail's limit for batchModify
const PARALLEL = 10;
const RETRIES = 9;          // waits 0.5 s doubling to 32 s: about two minutes, Gmail's quota refills per minute
const SCAN_GAP = 100;       // ms between scan requests: 10 per second is 3,000 of the 15,000 quota units per minute
const HEADERS = ['From', 'To', 'Subject', 'Delivered-To'];

export class AuthExpired extends Error {}

export async function pool(items, worker, parallel = PARALLEL) {
  const out = new Array(items.length);
  let next = 0;
  const run = async () => { while (next < items.length) { const i = next++; out[i] = await worker(items[i], i); } };
  await Promise.all(Array.from({ length: Math.min(parallel, items.length) }, run));
  return out;
}

function toMsg(m) {
  const h = {};
  for (const { name, value } of m.payload?.headers || []) h[name.toLowerCase()] = value;
  return {
    id: m.id,
    from: h.from || '',
    to: h.to || '',
    deliveredTo: h['delivered-to'] || '',
    subject: h.subject || '',
    date: Number(m.internalDate) || 0,
    snippet: m.snippet || '',
    unread: (m.labelIds || []).includes('UNREAD'),
  };
}

/**
 * Gmail provider. The Outlook provider implements the same methods.
 * folderId 'ALL' means no folder filter (All Mail).
 */
export function createGmail({ getToken, fetch: doFetch = (...a) => globalThis.fetch(...a), sleep = ms => new Promise(r => setTimeout(r, ms)) }) {
  let own;

  async function api(path, { method = 'GET', params = {}, body } = {}) {
    const url = new URL(BASE + path);
    for (const [key, value] of Object.entries(params)) {
      for (const v of [].concat(value)) if (v !== undefined && v !== null && v !== '') url.searchParams.append(key, v);
    }
    for (let attempt = 0; ; attempt++) {
      const headers = { Authorization: `Bearer ${getToken()}` };
      if (body) headers['Content-Type'] = 'application/json';
      const res = await doFetch(url.toString(), { method, headers, body: body && JSON.stringify(body) });
      if (res.status === 401) throw new AuthExpired('Signed out. Sign in again.');
      const text = await res.text();
      let data = null;
      try { data = text ? JSON.parse(text) : null; } catch { /* non-JSON error page */ }
      // Gmail reports a used-up quota as 403, not only as 429.
      const busy = res.status === 429 || res.status === 503 ||
        (res.status === 403 && /quota|rate limit/i.test(data?.error?.message || ''));
      if (busy && attempt < RETRIES) { await sleep(Math.min(500 * 2 ** attempt, 32000)); continue; }
      if (!res.ok) throw new Error(`Gmail error ${res.status}${data?.error?.message ? `: ${data.error.message}` : ''}`);
      return data;
    }
  }

  // The scan is the one bulk reader. Bursting it empties the per-minute quota, so its requests take turns.
  let nextSlot = 0;
  async function scanSlot() {
    const now = Date.now();
    const at = Math.max(now, nextSlot);
    nextSlot = at + SCAN_GAP;
    if (at > now) await sleep(at - now);
  }

  async function modify(ids, add, remove) {
    for (let i = 0; i < ids.length; i += BATCH) {
      const body = { ids: ids.slice(i, i + BATCH) };
      if (add?.length) body.addLabelIds = add;
      if (remove?.length) body.removeLabelIds = remove;
      await api('/messages/batchModify', { method: 'POST', body });
    }
  }

  // Undo must reverse only what really changed: `had` maps each label to the ids that carried it before.
  async function change(ids, add = [], remove = [], had) {
    const ops = [];
    for (const label of add) { const x = ids.filter(id => !had.get(label)?.has(id)); if (x.length) ops.push({ ids: x, remove: [label] }); }
    for (const label of remove) { const x = ids.filter(id => had.get(label)?.has(id)); if (x.length) ops.push({ ids: x, add: [label] }); }
    await modify(ids, add, remove);
    return { count: ids.length, ops };
  }

  // Labels of a few chosen mails, one request each.
  async function labelsOf(ids, labels) {
    const had = new Map(labels.map(l => [l, new Set()]));
    await pool(ids, async id => {
      const m = await api(`/messages/${id}`, { params: { format: 'minimal' } });
      for (const l of m?.labelIds || []) had.get(l)?.add(id);
    });
    return had;
  }

  // Labels of every mail a search finds: one list per label, not one request per mail.
  async function labelsOfSearch(q, labels) {
    return new Map(await Promise.all(labels.map(async l => [l, new Set(await listIds({ q, labelIds: l }, Infinity))])));
  }

  const changeChosen = async (ids, add, remove) => (ids.length ? change(ids, add, remove, await labelsOf(ids, [...add, ...remove])) : { count: 0, ops: [] });

  const metadata = id => api(`/messages/${id}`, { params: { format: 'metadata', metadataHeaders: HEADERS } });

  async function listIds(params, limit) {
    const ids = [];
    let pageToken;
    do {
      const page = await api('/messages', { params: { ...params, maxResults: 500, pageToken } });
      for (const m of page?.messages || []) ids.push(m.id);
      pageToken = page?.nextPageToken;
    } while (pageToken && ids.length < limit);
    return ids.slice(0, limit);
  }

  async function ownAddress() {
    return own ??= (await api('/profile')).emailAddress;
  }

  async function listFolders() {
    const data = await api('/labels');
    return (data?.labels || []).filter(l => l.type === 'user').map(({ id, name }) => ({ id, name }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  return {
    ownAddress,

    /** query: the user's own search text. filter: { to?, from? (one or a list), trash? } built by the app, same for every provider. */
    async listMessages({ folderId, query, filter = {}, pageToken } = {}) {
      const q = [query, filter.to && `to:("${filter.to}")`, filter.from && `from:(${[].concat(filter.from).map(f => `"${f}"`).join(' OR ')})`, filter.trash && 'in:trash']
        .filter(Boolean).join(' ');
      const labelIds = folderId === 'ALL' ? undefined : folderId || (q ? undefined : 'INBOX');
      const page = await api('/messages', { params: { maxResults: PAGE, q, labelIds, pageToken } });
      const messages = await pool(page?.messages || [], async m => toMsg(await metadata(m.id)));
      return { messages, nextPageToken: page?.nextPageToken || null };
    },

    async getMessage(id) {
      const m = await api(`/messages/${id}`, { params: { format: 'full' } });
      return { ...toMsg(m), ...pickBody(m.payload) };
    },

    name: 'Gmail',
    aliasCondition: address => ({ sentToAddresses: [address] }),
    support: GMAIL_SUPPORT,
    validateRule: toGmailFilter, // throws with a readable message when Gmail cannot run the rule

    // Each returns a receipt for undo().
    trash: ids => changeChosen(ids, ['TRASH'], []),
    markRead: (ids, read) => (read ? modify(ids, [], ['UNREAD']) : modify(ids, ['UNREAD'], [])), // no undo: the opposite button is undo
    moveToFolder: (ids, folderId) => changeChosen(ids, [folderId], ['INBOX']),
    async undo(receipt) {
      for (const op of receipt.ops) await modify(op.ids, op.add, op.remove);
    },

    listFolders,

    async createFolder(name) {
      const wanted = name.trim();
      const existing = (await listFolders()).find(f => f.name.toLowerCase() === wanted.toLowerCase());
      if (existing) return existing;
      const made = await api('/labels', { method: 'POST', body: { name: wanted, labelListVisibility: 'labelShow', messageListVisibility: 'show' } });
      return { id: made.id, name: made.name };
    },

    async listRules() {
      const [filters, labels] = await Promise.all([api('/settings/filters'), api('/labels')]);
      const names = new Map((labels?.labels || []).map(l => [l.id, l.name]));
      return (filters?.filter || []).map(raw => ({ id: raw.id, summary: describeGmailFilter(raw, names), raw, rule: fromGmailFilter(raw), enabled: true }));
    },

    async createRule(rule) {
      const made = await api('/settings/filters', { method: 'POST', body: toGmailFilter(rule) });
      return { id: made.id };
    },

    deleteRule: id => api(`/settings/filters/${id}`, { method: 'DELETE' }),

    // Gmail filters cannot be changed or switched off. Edit makes the new one before deleting the old; pause is delete plus a copy kept by the app.
    async updateRule(id, rule) {
      const made = await api('/settings/filters', { method: 'POST', body: toGmailFilter(rule) });
      await api(`/settings/filters/${id}`, { method: 'DELETE' });
      return { id: made.id };
    },
    async recreateRule(raw) {
      const made = await api('/settings/filters', { method: 'POST', body: { criteria: raw.criteria, action: raw.action } });
      return { id: made.id };
    },

    /** Runs the rule's label changes on mail already in the mailbox. Never forwards. Returns a receipt with the count. */
    async applyRuleToExisting(rule) {
      const { action } = toGmailFilter(rule);
      const add = action.addLabelIds || [], remove = action.removeLabelIds || [];
      if (!add.length && !remove.length) return { count: 0, ops: [] };
      const q = ruleToSearch(rule);
      const ids = await listIds({ q }, Infinity);
      if (!ids.length) return { count: 0, ops: [] };
      // A search never finds mail in Trash or Spam, so those labels need no lookup.
      return change(ids, add, remove, await labelsOfSearch(q, [...add, ...remove].filter(l => l !== 'TRASH' && l !== 'SPAM')));
    },

    async countMatching(rule) {
      const page = await api('/messages', { params: { q: ruleToSearch(rule), maxResults: 500 } });
      return { count: (page?.messages || []).length, more: Boolean(page?.nextPageToken) };
    },

    /**
     * Alias and sender domain of the newest mails. `known` (id -> record) is the cache: only new ids are fetched, and new records are added to it.
     * Records without an alias are returned too, so the cache stops them being fetched again.
     */
    async scanAliases(known, onProgress = () => {}) {
      const me = await ownAddress();
      const ids = await listIds({}, SCAN_LIMIT);
      let done = 0;
      return pool(ids, async id => {
        let rec = known.get(id);
        if (!rec) {
          await scanSlot();
          const m = toMsg(await metadata(id));
          rec = { id, alias: extractAlias(m.deliveredTo, me) ?? extractAlias(m.to, me), domain: senderDomain(m.from), from: senderAddress(m.from) };
          known.set(id, rec); // progress survives a failed scan: the caller can save `known` and resume
        }
        onProgress(++done, ids.length);
        return rec;
      });
    },
  };
}
