import { extractAlias, senderDomain } from './alias.js';
import { toOutlookRule, describeOutlookRule, OUTLOOK_SUPPORT } from './rules.js';
import { AuthExpired, pool } from './gmail.js';

const GRAPH = 'https://graph.microsoft.com/';
const BASE = `${GRAPH}v1.0/me`;
const PAGE = 50;
const SCAN_LIMIT = 2000;
const PARALLEL = 4; // Outlook allows 4 requests at once per mailbox
const RETRIES = 9;
const SELECT = 'id,from,toRecipients,subject,receivedDateTime,bodyPreview,isRead';

const address = recipient => recipient?.emailAddress?.address || '';

function toMsg(m) {
  const from = m.from?.emailAddress;
  return {
    id: m.id,
    from: !from ? '' : from.name && from.name !== from.address ? `${from.name} <${from.address}>` : from.address,
    to: (m.toRecipients || []).map(address).join(', '),
    deliveredTo: '',
    subject: m.subject || '',
    date: Date.parse(m.receivedDateTime) || 0,
    snippet: m.bodyPreview || '',
    unread: m.isRead === false,
  };
}

// Outlook cannot run a saved rule on old mail. Simple rules are replayed through search; others are not offered.
const SEARCH_FIELD = { fromAddresses: 'from', senderContains: 'from', sentToAddresses: 'to', recipientContains: 'to', subjectContains: 'subject' };
function replaySearch(rule) {
  const set = part => Object.entries(part || {}).filter(([, v]) => v !== undefined && v !== null && v !== false && !(Array.isArray(v) && !v.length));
  const actions = rule.actions || {};
  if (set(rule.exceptions).length || !(actions.moveToFolder || actions.delete || actions.markAsRead)) return null;
  const parts = [];
  for (const [key, value] of set(rule.conditions)) {
    if (!SEARCH_FIELD[key]) return null;
    parts.push(`(${[].concat(value).map(v => `${SEARCH_FIELD[key]}:${String(v).replaceAll('"', '')}`).join(' OR ')})`);
  }
  return parts.length ? parts.join(' AND ') : null;
}

/** Outlook provider (Microsoft Graph). Same methods as the Gmail provider. */
export function createOutlook({ getToken, fetch: doFetch = (...a) => globalThis.fetch(...a), sleep = ms => new Promise(r => setTimeout(r, ms)) }) {
  let own;

  async function api(pathOrUrl, { method = 'GET', params = {}, body } = {}) {
    const url = new URL(pathOrUrl.startsWith('https://') ? pathOrUrl : BASE + pathOrUrl);
    // A paging link comes from a server response. The token must never follow it to another host.
    if (!url.href.startsWith(GRAPH)) throw new Error('Refused a page link that does not point at Microsoft Graph.');
    for (const [key, value] of Object.entries(params)) if (value !== undefined && value !== null && value !== '') url.searchParams.set(key, value);
    for (let attempt = 0; ; attempt++) {
      const headers = { Authorization: `Bearer ${getToken()}` };
      if (body) headers['Content-Type'] = 'application/json';
      const res = await doFetch(url.toString(), { method, headers, body: body && JSON.stringify(body) });
      if (res.status === 401) throw new AuthExpired('Signed out. Sign in again.');
      if ([429, 503, 504].includes(res.status) && attempt < RETRIES) {
        const after = Number(res.headers?.get?.('Retry-After'));
        await sleep(after > 0 ? after * 1000 : Math.min(500 * 2 ** attempt, 32000));
        continue;
      }
      const text = await res.text();
      let data = null;
      try { data = text ? JSON.parse(text) : null; } catch { /* non-JSON error page */ }
      if (!res.ok) throw new Error(`Outlook error ${res.status}${data?.error?.message ? `: ${data.error.message}` : ''}`);
      return data;
    }
  }

  const each = (ids, fn) => pool(ids, fn, PARALLEL);
  const move = (id, destinationId) => api(`/messages/${id}/move`, { method: 'POST', body: { destinationId } });
  const setRead = (id, isRead) => api(`/messages/${id}`, { method: 'PATCH', body: { isRead } });

  async function pages(first, params, limit) {
    const out = [];
    let page = await api(first, { params });
    for (;;) {
      out.push(...(page?.value || []));
      if (!page?.['@odata.nextLink'] || out.length >= limit) return out.slice(0, limit);
      page = await api(page['@odata.nextLink']);
    }
  }

  async function ownAddress() {
    if (!own) {
      const user = await api('', { params: { $select: 'mail,userPrincipalName' } });
      own = (user.mail || user.userPrincipalName).toLowerCase();
    }
    return own;
  }

  // shortcut: top-level folders only, built-in ones (Drafts, Sent Items) included; filter them if they get in the way
  async function listFolders() {
    const data = await api('/mailFolders', { params: { $top: 200, $select: 'id,displayName' } });
    return (data?.value || []).map(f => ({ id: f.id, name: f.displayName })).sort((a, b) => a.name.localeCompare(b.name));
  }
  const folderNames = async () => new Map((await listFolders()).map(f => [f.id, f.name]));

  return {
    name: 'Outlook',
    support: OUTLOOK_SUPPORT,
    validateRule: rule => toOutlookRule(rule, new Map()),
    // Outlook matches a +alias reliably as text in the recipient line.
    aliasCondition: alias => ({ recipientContains: [alias] }),
    ownAddress,

    async listMessages({ folderId, query, filter = {}, pageToken } = {}) {
      let page;
      if (pageToken) {
        page = await api(pageToken);
      } else {
        const search = [query, filter.to && `to:${filter.to}`, filter.from && `from:${filter.from}`].filter(Boolean).join(' AND ');
        const folder = filter.trash ? 'deleteditems' : folderId === 'ALL' ? null : folderId || (search ? null : 'inbox');
        const params = { $top: PAGE, $select: SELECT };
        if (search) params.$search = `"${search.replaceAll('"', '')}"`; // Graph does not allow sorting together with search
        else params.$orderby = 'receivedDateTime desc';
        page = await api(folder ? `/mailFolders/${folder}/messages` : '/messages', { params });
      }
      return { messages: (page?.value || []).map(toMsg), nextPageToken: page?.['@odata.nextLink'] || null };
    },

    async getMessage(id) {
      const m = await api(`/messages/${id}`, { params: { $select: `${SELECT},body` } });
      const html = m.body?.contentType === 'html';
      return { ...toMsg(m), html: html ? m.body.content : null, text: !html && m.body ? m.body.content : null };
    },

    trash: async ids => { await each(ids, id => move(id, 'deleteditems')); },
    markRead: async (ids, read) => { await each(ids, id => setRead(id, read)); },
    moveToFolder: async (ids, folderId) => { await each(ids, id => move(id, folderId)); },

    listFolders,

    async createFolder(name) {
      const wanted = name.trim();
      const existing = (await listFolders()).find(f => f.name.toLowerCase() === wanted.toLowerCase());
      if (existing) return existing;
      const made = await api('/mailFolders', { method: 'POST', body: { displayName: wanted } });
      return { id: made.id, name: made.displayName };
    },

    async listRules() {
      const [rules, names] = await Promise.all([api('/mailFolders/inbox/messageRules'), folderNames()]);
      return (rules?.value || []).map(raw => ({ id: raw.id, summary: describeOutlookRule(raw, names), raw }));
    },

    async createRule(rule) {
      const made = await api('/mailFolders/inbox/messageRules', { method: 'POST', body: toOutlookRule(rule, await folderNames()) });
      return { id: made.id };
    },

    deleteRule: id => api(`/mailFolders/inbox/messageRules/${id}`, { method: 'DELETE' }),

    async countMatching(rule) {
      const search = replaySearch(rule);
      if (!search) return { count: 0, more: false };
      const page = await api('/messages', { params: { $search: `"${search}"`, $select: 'id', $top: 500 } });
      return { count: (page?.value || []).length, more: Boolean(page?.['@odata.nextLink']) };
    },

    async applyRuleToExisting(rule) {
      const search = replaySearch(rule);
      if (!search) return 0;
      const found = await pages('/messages', { $search: `"${search}"`, $select: 'id', $top: 500 }, SCAN_LIMIT);
      const { markAsRead, moveToFolder, delete: toTrash } = rule.actions;
      await each(found, async ({ id }) => {
        if (markAsRead) await setRead(id, true);
        if (toTrash) await move(id, 'deleteditems');
        else if (moveToFolder) await move(id, moveToFolder);
      });
      return found.length;
    },

    /** Alias and sender domain of the newest mails. Outlook returns recipients in the list, so this takes a few requests. */
    async scanAliases(known, onProgress = () => {}) {
      const me = await ownAddress();
      const found = await pages('/messages', { $select: 'id,from,toRecipients', $top: 500, $orderby: 'receivedDateTime desc' }, SCAN_LIMIT);
      const records = found.map(m => {
        let rec = known.get(m.id);
        if (!rec) {
          rec = { id: m.id, alias: extractAlias((m.toRecipients || []).map(address).join(', '), me), domain: senderDomain(address(m.from)), from: address(m.from).toLowerCase() };
          known.set(m.id, rec);
        }
        return rec;
      });
      onProgress(records.length, records.length);
      return records;
    },
  };
}
