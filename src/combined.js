// Every connected account as one provider: the merged inbox, search, senders and cleanup.
// Folders, rules, aliases and leaks belong to one account and are not offered here.
// Ids carry their account as `<key>\u0001<id>`; the key is kind and address, so it stays right when the account list changes.

const SEP = '\u0001';
export const combinedId = (key, id) => `${key}${SEP}${id}`;

/** One account failed. `partial` holds what the accounts before it already did, so undo still covers that. */
export class AccountFailed extends Error {
  constructor(account, cause, partial) {
    super(`${account.address}: ${cause.message}`);
    Object.assign(this, { account, cause, partial });
  }
}

/** accounts: [{ key, address, provider }] */
export function createCombined(accounts) {
  const byKey = new Map(accounts.map(a => [a.key, a]));
  function split(id) {
    const at = id.indexOf(SEP);
    const account = byKey.get(id.slice(0, at));
    if (!account) throw new Error('That mail belongs to an account that is signed out.');
    return [account, id.slice(at + 1)];
  }
  const tag = (account, m) => ({ ...m, id: combinedId(account.key, m.id), account: account.address });

  // Runs fn for each account in turn. Results come back per account; a failure carries the results so far.
  async function each(list, fn, summarize = () => undefined) {
    const done = [];
    for (const [account, arg] of list) {
      try {
        done.push({ key: account.key, result: await fn(account, arg) });
      } catch (error) {
        const partial = summarize(done);
        throw new AccountFailed(account, error, partial && { ...partial, failed: true });
      }
    }
    return summarize(done) ?? done;
  }
  const all = () => accounts.map(a => [a, undefined]);
  function grouped(ids) {
    const out = new Map();
    for (const id of ids) {
      const [account, raw] = split(id);
      if (!out.has(account)) out.set(account, []);
      out.get(account).push(raw);
    }
    return [...out];
  }
  // A combined receipt: each account's own receipt, and the total for the app.
  const receipt = done => ({ count: done.reduce((n, d) => n + (d.result?.count || 0), 0), parts: done });
  const counted = done => ({ count: done.reduce((n, d) => n + d.result.count, 0), more: done.some(d => d.result.more) });

  return {
    name: 'All accounts',
    combined: true,
    support: { predicates: new Set(), actions: new Set() },
    validateRule: rule => accounts.forEach(a => a.provider.validateRule(rule)),
    ownAddress: async () => 'All accounts',

    /** Pages from every account, newest first. The page token holds each account's own token. */
    async listMessages({ folderId, query, filter, pageToken } = {}) {
      const tokens = pageToken ? JSON.parse(pageToken) : null;
      const list = accounts.filter(a => !tokens || tokens[a.key]).map(a => [a, tokens?.[a.key]]);
      const pages = await each(list, (a, token) => a.provider.listMessages({ folderId, query, filter, pageToken: token }));
      const next = {};
      const messages = [];
      for (const { key, result } of pages) {
        if (result.nextPageToken) next[key] = result.nextPageToken;
        messages.push(...result.messages.map(m => tag(byKey.get(key), m)));
      }
      return { messages: messages.sort((a, b) => b.date - a.date), nextPageToken: Object.keys(next).length ? JSON.stringify(next) : null };
    },

    async getMessage(id) {
      const [account, raw] = split(id);
      return tag(account, await account.provider.getMessage(raw));
    },

    trash: ids => each(grouped(ids), (a, raw) => a.provider.trash(raw), receipt),
    markRead: (ids, read) => each(grouped(ids), (a, raw) => a.provider.markRead(raw, read)),
    async undo(combinedReceipt) {
      for (const { key, result } of combinedReceipt.parts) {
        if (!byKey.has(key)) throw new Error('An account of this action is signed out.');
        await byKey.get(key).provider.undo(result);
      }
    },

    /** Sends from the account that holds `replyTo`, or `viaId` (a mail of that account) when there is no reply. */
    async sendMail(mail) {
      const [account, raw] = split(mail.replyTo?.id || mail.viaId || '');
      await account.provider.sendMail({ ...mail, replyTo: mail.replyTo && { ...mail.replyTo, id: raw } });
    },

    countMatching: rule => each(all(), a => a.provider.countMatching(rule), counted),
    applyRuleToExisting: rule => each(all(), a => a.provider.applyRuleToExisting(rule), receipt),
    // The rule is made in every account. Its id lists each part, so deleteRule (and undo) remove all of them.
    createRule: rule => each(all(), a => a.provider.createRule(rule),
      done => ({ id: JSON.stringify(done.map(d => [d.key, d.result.id])) })),
    async deleteRule(id) {
      for (const [key, raw] of JSON.parse(id)) await byKey.get(key)?.provider.deleteRule(raw);
    },

    countMessages: filter => each(all(), a => a.provider.countMessages(filter), counted),
    trashMessages: filter => each(all(), a => a.provider.trashMessages(filter), receipt),
  };
}
