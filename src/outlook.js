import { extractAlias, senderDomain } from './alias.js';
import { toOutlookRule, fromOutlookRule, describeOutlookRule, OUTLOOK_SUPPORT } from './rules.js';
import { AuthExpired, pool } from './gmail.js';

const GRAPH = 'https://graph.microsoft.com/';
const BASE = `${GRAPH}v1.0/me`;
const PAGE = 50;
const SCAN_LIMIT = 2000;
const PARALLEL = 4; // Outlook allows 4 requests at once per mailbox
const RETRIES = 9;
const SELECT = 'id,from,toRecipients,subject,receivedDateTime,bodyPreview,isRead';
// Graph has no size field on a message; PR_MESSAGE_SIZE (0x0E08) holds it.
const SIZE_PROP = "singleValueExtendedProperties($filter=id eq 'Integer 0x0E08')";

const WELL_KNOWN = { trash: 'deleteditems', spam: 'junkemail' };

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

// Cleanup parts of a filter in Outlook's search language. Unread cannot be searched: callers drop read mail themselves.
const day = ms => new Date(ms).toISOString().slice(0, 10);
const cleanupSearch = filter => [filter.before && `received<${day(filter.before)}`, filter.larger && `size>${Math.round(filter.larger)}`].filter(Boolean);

// One sender, or a list of senders of which any may match.
function fromSearch(from) {
  const list = [].concat(from || []).map(f => `from:${f}`);
  return list.length > 1 ? `(${list.join(' OR ')})` : list[0];
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
  // A move gives the mail a new id. Receipts keep that id and the folder it came from.
  const move = async (id, destinationId) => (await api(`/messages/${id}/move`, { method: 'POST', body: { destinationId } }))?.id || id;
  async function moveChosen(ids, destinationId) {
    const moves = await each(ids, async id => {
      const { parentFolderId } = await api(`/messages/${id}`, { params: { $select: 'parentFolderId' } });
      return { id: await move(id, destinationId), folderId: parentFolderId };
    });
    return { count: ids.length, moves };
  }
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

  // shortcut: at most `limit` mails per click, like the other bulk actions on Outlook
  async function cleanupFind(filter, limit) {
    const search = cleanupSearch(filter);
    if (!search.length && !filter.unread) throw new Error('Pick at least one cleanup condition.');
    const params = { $select: 'id,parentFolderId,isRead', $top: 500 };
    if (search.length) params.$search = `"${search.join(' AND ')}"`; else params.$filter = 'isRead eq false';
    const found = await pages('/messages', params, limit);
    return found.filter(m => !filter.unread || m.isRead === false);
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
        const search = [query, filter.to && `to:${filter.to}`, fromSearch(filter.from), ...cleanupSearch(filter)].filter(Boolean).join(' AND ');
        const folder = filter.trash ? 'deleteditems' : filter.spam ? 'junkemail' : folderId === 'ALL' ? null : folderId || (search ? null : 'inbox');
        const params = { $top: PAGE, $select: SELECT };
        if (search) params.$search = `"${search.replaceAll('"', '')}"`; // Graph does not allow sorting together with search
        else params.$orderby = 'receivedDateTime desc';
        page = await api(folder ? `/mailFolders/${folder}/messages` : '/messages', { params });
      }
      // shortcut: unread is dropped per page, so an unread-only page can show fewer than 50; use $filter if that bothers
      const messages = (page?.value || []).map(toMsg).filter(m => !filter.unread || m.unread);
      return { messages, nextPageToken: page?.['@odata.nextLink'] || null };
    },

    async getMessage(id) {
      const m = await api(`/messages/${id}`, { params: { $select: `${SELECT},body,internetMessageHeaders` } });
      const html = m.body?.contentType === 'html';
      const header = name => m.internetMessageHeaders?.find(h => h.name.toLowerCase() === name)?.value || '';
      return { ...toMsg(m), html: html ? m.body.content : null, text: !html && m.body ? m.body.content : null,
        listUnsubscribe: header('list-unsubscribe'), listUnsubscribePost: header('list-unsubscribe-post') };
    },

    // trash, moveToFolder and applyRuleToExisting return a receipt for undo().
    trash: ids => moveChosen(ids, 'deleteditems'),
    markRead: async (ids, read) => { await each(ids, id => setRead(id, read)); },
    moveToFolder: moveChosen,
    async undo(receipt) {
      await each(receipt.moves, async ({ id, folderId, unread }) => {
        if (unread) await setRead(id, false);
        if (folderId) await move(id, folderId);
      });
    },

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
      return (rules?.value || []).map(raw => ({ id: raw.id, summary: describeOutlookRule(raw, names), raw, rule: fromOutlookRule(raw), enabled: raw.isEnabled !== false }));
    },

    async createRule(rule) {
      const made = await api('/mailFolders/inbox/messageRules', { method: 'POST', body: toOutlookRule(rule, await folderNames()) });
      return { id: made.id };
    },

    deleteRule: id => api(`/mailFolders/inbox/messageRules/${id}`, { method: 'DELETE' }),

    async updateRule(id, rule) {
      const { sequence, isEnabled, ...body } = toOutlookRule(rule, await folderNames()); // keep its place in the order and its on/off state
      await api(`/mailFolders/inbox/messageRules/${id}`, { method: 'PATCH', body });
      return { id };
    },
    setRuleEnabled: (id, isEnabled) => api(`/mailFolders/inbox/messageRules/${id}`, { method: 'PATCH', body: { isEnabled } }),

    async folderCount(which) {
      return (await api(`/mailFolders/${WELL_KNOWN[which]}`, { params: { $select: 'totalItemCount' } }))?.totalItemCount || 0;
    },
    /** Deletes every mail in Deleted Items or Junk for good (not to Recoverable Items' undo list). Returns the count. */
    async emptyFolder(which) {
      const found = await pages(`/mailFolders/${WELL_KNOWN[which]}/messages`, { $select: 'id', $top: 500 }, Infinity);
      await each(found, ({ id }) => api(`/messages/${id}/permanentDelete`, { method: 'POST' }));
      return found.length;
    },

    async countMessages(filter) {
      const found = await cleanupFind(filter, 500);
      return { count: found.length, more: found.length >= 500 };
    },
    async trashMessages(filter) {
      const found = await cleanupFind(filter, SCAN_LIMIT);
      const moves = await each(found, async ({ id, parentFolderId }) => ({ id: await move(id, 'deleteditems'), folderId: parentFolderId }));
      return { count: found.length, moves };
    },

    async countMatching(rule) {
      const search = replaySearch(rule);
      if (!search) return { count: 0, more: false };
      const page = await api('/messages', { params: { $search: `"${search}"`, $select: 'id', $top: 500 } });
      return { count: (page?.value || []).length, more: Boolean(page?.['@odata.nextLink']) };
    },

    async applyRuleToExisting(rule) {
      const search = replaySearch(rule);
      if (!search) return { count: 0, moves: [] };
      const found = await pages('/messages', { $search: `"${search}"`, $select: 'id,parentFolderId,isRead', $top: 500 }, SCAN_LIMIT);
      const { markAsRead, moveToFolder, delete: toTrash } = rule.actions;
      const to = toTrash ? 'deleteditems' : moveToFolder;
      const moves = await each(found, async ({ id, parentFolderId, isRead }) => {
        const unread = Boolean(markAsRead && isRead === false);
        if (unread) await setRead(id, true);
        return to ? { id: await move(id, to), folderId: parentFolderId, unread } : { id, unread };
      });
      return { count: found.length, moves };
    },

    /** Alias and sender domain of the newest mails. Outlook returns recipients in the list, so this takes a few requests. */
    async scanAliases(known, onProgress = () => {}, limit = SCAN_LIMIT) {
      const me = await ownAddress();
      const found = await pages('/messages', { $select: 'id,from,toRecipients,parentFolderId', $top: 500, $orderby: 'receivedDateTime desc', $expand: SIZE_PROP }, limit);
      const records = found.map(m => {
        let rec = known.get(m.id);
        if (rec?.size === undefined || rec.folders?.[0] !== m.parentFolderId) {
          rec = { id: m.id, alias: extractAlias((m.toRecipients || []).map(address).join(', '), me), domain: senderDomain(address(m.from)), from: address(m.from).toLowerCase(),
            size: Number(m.singleValueExtendedProperties?.[0]?.value) || 0, folders: [m.parentFolderId] };
          known.set(m.id, rec);
        }
        return rec;
      });
      onProgress(records.length, records.length);
      return records;
    },
  };
}
