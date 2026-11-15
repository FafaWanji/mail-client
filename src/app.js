
import { googleSignIn, microsoftSignIn, GOOGLE_ID, MICROSOFT_ID, GOOGLE_FULL } from './auth.js';
import { createGmail, AuthExpired } from './gmail.js';
import { createOutlook } from './outlook.js';
import { createDemo } from './demo.js';
import { buildFrameDoc, FRAME_SANDBOX } from './mailbody.js';
import { aliasAddress, extractAlias, senderAddress } from './alias.js';
import { senderReport, senderKey, chunks, folderReport, sizeText } from './senders.js';
import { RULE_CATALOG } from './rules.js';
import { leakReport } from './leaks.js';
import { parseUnsubscribe, mailtoParts } from './unsubscribe.js';

const BULK_CONFIRM = 20; // trashing more than this many mails asks first

const $ = id => document.getElementById(id);
// Mail data only ever enters the page through textContent, never innerHTML.
function el(tag, props = {}, ...kids) {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...kids.filter(k => k !== null && k !== undefined));
  return node;
}
const decodeEntities = s => new DOMParser().parseFromString(s, 'text/html').documentElement.textContent;
const senderName = from => from.replace(/<[^>]*>/, '').replace(/"/g, '').trim() || from;

const accounts = [];            // connected this session: { kind, address, token, provider }
let current = null;             // the account on screen
let provider, me;               // shortcuts for current.provider and current.address
let box = { title: 'Inbox' };  // { title, folderId?, query?, filter? }
let messages = [], nextPageToken = null;
const selected = new Set();
let openMail = null;
let folders = [];
let scanRecords = [];           // every scanned mail, also those without alias (cache)
let aliasRecords = [];          // only those with an alias

// ---------- storage (this browser only) ----------
function load(key, fallback) {
  try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; }
}
function save(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* storage full or blocked: cache is optional */ }
}

// The scan cache outgrows localStorage (about 5 million characters) on a whole-mailbox scan, so it lives in IndexedDB.
const scanDb = new Promise((resolve, reject) => {
  const open = indexedDB.open('mailclient', 1);
  open.onupgradeneeded = () => open.result.createObjectStore('kv');
  open.onsuccess = () => resolve(open.result);
  open.onerror = () => reject(open.error);
});
async function idb(mode, fn) {
  const request = fn((await scanDb).transaction('kv', mode).objectStore('kv'));
  return new Promise((resolve, reject) => { request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); });
}
const saveScan = (address, records) => idb('readwrite', store => store.put(records, `scan:${address}`)).catch(() => { /* cache is optional */ });
async function loadScan(address) {
  try {
    const kept = await idb('readonly', store => store.get(`scan:${address}`));
    if (kept) return kept;
    const old = load(`scan2:${address}`, []); // the cache before IndexedDB, moved over once
    if (old.length) await saveScan(address, old);
    localStorage.removeItem(`scan2:${address}`);
    return old;
  } catch {
    return [];
  }
}

// ---------- status and errors ----------
let statusTimer;
function status(text, isError = false) {
  clearTimeout(statusTimer);
  $('status').textContent = text;
  $('status').className = isError ? 'error' : '';
  if (text && !isError) statusTimer = setTimeout(() => { $('status').textContent = ''; }, 4000);
}

/** Every provider call goes through here: busy cursor, errors shown, expired session back to sign-in. */
async function run(fn) {
  document.body.classList.add('busy');
  try {
    return await fn();
  } catch (error) {
    if (error instanceof AuthExpired) expire(current);
    else status(error.message, true);
    return undefined;
  } finally {
    document.body.classList.remove('busy');
  }
}

// ---------- accounts ----------
const KINDS = {
  gmail: { label: 'Gmail', id: GOOGLE_ID, signIn: googleSignIn, create: createGmail },
  outlook: { label: 'Outlook', id: MICROSOFT_ID, signIn: microsoftSignIn, create: createOutlook },
};
// Only kind and address are remembered, never a token. Each session signs in again.
const remembered = () => load('accounts', []).filter(a => KINDS[a.kind]);
const sameAccount = (a, b) => a.kind === b.kind && a.address === b.address;

function showGate(text) {
  $('app').hidden = true;
  $('gate').hidden = false;
  $('gateText').textContent = text;
  $('gateAccounts').replaceChildren(...remembered().map(known => {
    const live = accounts.find(a => sameAccount(a, known));
    return el('button', {
      textContent: live ? `Open ${known.address}` : `Sign in to ${known.address} (${KINDS[known.kind].label})`,
      onclick: () => (live ? switchTo(live) : addAccount(known.kind, known.address)),
    });
  }));
  $('addGmailGate').hidden = !GOOGLE_ID;
  $('addOutlookGate').hidden = !MICROSOFT_ID;
}

/** Signs in and returns the connected account. `hint` is the address to preselect. */
async function connect(kind, hint) {
  const account = { kind, token: await KINDS[kind].signIn(hint) };
  account.provider = KINDS[kind].create({ getToken: () => account.token });
  account.address = await account.provider.ownAddress();
  const index = accounts.findIndex(a => sameAccount(a, account));
  if (index >= 0) accounts[index] = account; else accounts.push(account);
  save('accounts', [...remembered().filter(a => !sameAccount(a, account)), { kind, address: account.address }]);
  return account;
}

async function addAccount(kind, hint) {
  try {
    await switchTo(await connect(kind, hint));
  } catch (error) {
    if ($('app').hidden) showGate(error.message);
    else { renderAccounts(); status(error.message, true); }
  }
}

function expire(account) {
  const index = accounts.indexOf(account);
  if (index >= 0) accounts.splice(index, 1);
  showGate(`The session for ${account?.address || 'this account'} ended. Sign in again.`);
}

function renderAccounts() {
  const others = remembered().filter(known => !accounts.some(a => sameAccount(a, known)));
  $('accountSelect').replaceChildren(
    ...accounts.map((a, i) => el('option', { value: `live:${i}`, textContent: a.address, selected: a === current })),
    ...others.map((a, i) => el('option', { value: `known:${i}`, textContent: `${a.address} (sign in)` })));
  $('accountSelect').onchange = () => {
    const [type, index] = $('accountSelect').value.split(':');
    if (type === 'live') switchTo(accounts[index]);
    else addAccount(others[index].kind, others[index].address);
  };
  $('addGmail').hidden = !GOOGLE_ID;
  $('addOutlook').hidden = !MICROSOFT_ID;
}

async function switchTo(account) {
  current = account;
  provider = account.provider;
  me = account.address;
  $('gate').hidden = true;
  $('app').hidden = false;
  $('builder').hidden = true;
  pickedSenders.clear();
  offerUndo(null);
  renderAccounts();
  try { localStorage.removeItem(`aliasScan:${me}`); } catch { /* old cache format, best effort */ }
  $('fullScan').checked = load(`fullScan:${me}`, false);
  scanRecords = await loadScan(me); // show the last scan at once, the new one updates it
  setAliasRecords();
  await refreshFolders();
  await openBox(BOXES.inbox);
  scan();
}

for (const [id, kind] of [['addGmail', 'gmail'], ['addGmailGate', 'gmail'], ['addOutlook', 'outlook'], ['addOutlookGate', 'outlook']]) {
  $(id).onclick = () => addAccount(kind);
}
$('signOut').onclick = () => location.reload(); // tokens live in memory only, so a reload signs every account out

// ---------- views ----------
function showView(name) {
  for (const view of ['mail', 'senders', 'cleanup', 'rules', 'leaks']) $(`view-${view}`).hidden = view !== name;
  if (name === 'senders') renderSenders();
  if (name === 'rules') refreshRules();
  if (name === 'leaks') renderLeaks();
}
for (const button of document.querySelectorAll('[data-view]')) button.onclick = () => showView(button.dataset.view);

const BOXES = {
  inbox: { title: 'Inbox' },
  all: { title: 'All Mail', folderId: 'ALL' },
  trash: { title: 'Trash', filter: { trash: true } },
  spam: { title: 'Spam', filter: { spam: true } },
};
for (const button of document.querySelectorAll('[data-box]')) button.onclick = () => openBox(BOXES[button.dataset.box]);

// ---------- mail list ----------
async function openBox(next) {
  box = next;
  showView('mail');
  $('boxTitle').textContent = box.title;
  $('actEmpty').hidden = !(box === BOXES.trash || box === BOXES.spam);
  messages = [];
  selected.clear();
  closeMail();
  await loadPage(undefined);
}

async function loadPage(pageToken) {
  const page = await run(() => provider.listMessages({ folderId: box.folderId, query: box.query, filter: box.filter, pageToken }));
  if (!page) return;
  messages.push(...page.messages);
  nextPageToken = page.nextPageToken;
  renderList();
}

function renderList() {
  $('selectAll').checked = false;
  $('more').hidden = !nextPageToken;
  $('messages').replaceChildren(...messages.map(m => {
    const check = el('input', { type: 'checkbox', checked: selected.has(m.id), ariaLabel: `Select ${m.subject}` });
    check.onclick = event => {
      event.stopPropagation();
      if (check.checked) selected.add(m.id); else selected.delete(m.id);
    };
    const row = el('li', { className: `${m.unread ? 'unread' : ''} ${openMail?.id === m.id ? 'open' : ''}`, tabIndex: 0 },
      check,
      el('span', { className: 'from', textContent: senderName(m.from) }),
      el('span', { className: 'date', textContent: new Date(m.date).toLocaleDateString() }),
      el('span', { className: 'subject', textContent: m.subject || '(no subject)' }),
      el('span', { className: 'snippet', textContent: decodeEntities(m.snippet) }));
    row.onclick = () => openMessage(m.id);
    row.onkeydown = event => { if (event.key === 'Enter') openMessage(m.id); };
    return row;
  }));
  if (!messages.length) $('messages').append(el('li', { className: 'muted', textContent: 'No mail here.' }));
}

$('more').onclick = () => loadPage(nextPageToken);
$('selectAll').onchange = () => {
  selected.clear();
  if ($('selectAll').checked) for (const m of messages) selected.add(m.id);
  renderList();
  $('selectAll').checked = selected.size > 0;
};
$('search').onsubmit = event => {
  event.preventDefault();
  const query = $('query').value.trim();
  openBox(query ? { title: `Search: ${query}`, query } : BOXES.inbox);
};

// ---------- reading ----------
function closeMail() {
  openMail = null;
  $('readHead').replaceChildren('Select a mail.');
  $('readHead').className = 'muted';
  $('readTools').hidden = true;
  $('frame').hidden = true;
  $('frame').srcdoc = '';
}

function renderFrame(allowImages) {
  $('frame').setAttribute('sandbox', FRAME_SANDBOX);
  $('frame').srcdoc = buildFrameDoc(openMail, allowImages);
  $('showImages').hidden = allowImages || openMail.html === null;
}

async function openMessage(id) {
  const mail = await run(() => provider.getMessage(id));
  if (!mail) return;
  openMail = mail;
  $('readHead').className = '';
  $('readHead').replaceChildren(
    el('strong', { textContent: mail.subject || '(no subject)' }),
    el('div', { textContent: `From: ${mail.from}` }),
    el('div', { textContent: `To: ${mail.deliveredTo || mail.to}` }),
    el('div', { className: 'muted', textContent: new Date(mail.date).toLocaleString() }));
  $('readTools').hidden = false;
  const unsub = parseUnsubscribe(mail.listUnsubscribe, mail.listUnsubscribePost);
  $('unsubscribe').hidden = !(unsub.oneClick || unsub.link || (unsub.mailto && provider.sendMail));
  $('frame').hidden = false;
  renderFrame(false);
  const listed = messages.find(m => m.id === id);
  if (listed?.unread) {
    listed.unread = false;
    run(() => provider.markRead([id], true));
  }
  renderList();
}
$('showImages').onclick = () => renderFrame(true);
$('unsubscribe').onclick = async () => {
  const info = parseUnsubscribe(openMail.listUnsubscribe, openMail.listUnsubscribePost);
  const done = await unsubscribe(info, senderName(openMail.from));
  if (done.link) window.open(done.link, '_blank', 'noopener,noreferrer'); // the sender's own page finishes it
  else if (done.text) status(done.text, !done.ok);
};

// ---------- unsubscribe ----------
/**
 * Uses the sender's List-Unsubscribe header. One-click (RFC 8058) is a bare POST; a mailto: is sent as a short mail.
 * Returns { ok, text }, or { link } when only the sender's web page can do it.
 */
async function unsubscribe(info, name) {
  if (info.oneClick) {
    // no-cors and no cookies: the request carries nothing of this app or the browser, and its answer cannot be read.
    const sent = current.kind === 'demo' || await fetch(info.oneClick, { method: 'POST', mode: 'no-cors', credentials: 'omit', referrerPolicy: 'no-referrer',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'List-Unsubscribe=One-Click' }).then(() => true, () => false);
    return sent ? { ok: true, text: `Unsubscribe request sent to ${name}.` } : { ok: false, text: `Could not reach ${name}'s unsubscribe address.` };
  }
  if (info.mailto && provider.sendMail) {
    const { to, subject, body } = mailtoParts(info.mailto);
    if (!confirm(`${name} unsubscribes by mail. Send "${subject}" from ${me} to ${to}?`)) return { ok: false, text: 'Nothing sent.' };
    const sent = await run(() => provider.sendMail({ to, subject, text: body }).then(() => true));
    return { ok: Boolean(sent), text: sent ? `Unsubscribe mail sent to ${to}.` : '' };
  }
  if (info.link) return { link: info.link };
  return { ok: false, text: `${name} gives no way to unsubscribe.` };
}

// ---------- undo ----------
// Only the last action can be undone. Receipts come from the provider; rules and scan records are what the app added or dropped.
let lastUndo = null;
function offerUndo(text, receipts = [], { rules = [], records = [] } = {}) {
  receipts = receipts.filter(r => r?.count);
  lastUndo = text && (receipts.length || rules.length) ? { account: current, receipts, rules, records } : null;
  $('undo').hidden = !lastUndo;
  $('undo').textContent = `Undo: ${text}`;
}
$('undo').onclick = async () => {
  const undo = lastUndo;
  offerUndo(null);
  if (!undo || undo.account !== current) return;
  const done = await run(async () => {
    for (const receipt of undo.receipts) await provider.undo(receipt);
    for (const id of undo.rules) await provider.deleteRule(id);
    return true;
  });
  if (undo.records.length) {
    scanRecords = [...scanRecords, ...undo.records];
    saveScan(me, scanRecords);
    setAliasRecords();
  }
  if (done) status('Undone.');
  await openBox(box);
};
const mailsText = n => `${n} mail${n === 1 ? '' : 's'}`;

// ---------- cleanup actions ----------
function targets() {
  if (selected.size) return [...selected];
  if (openMail) return [openMail.id];
  status('Select a mail first.', true);
  return [];
}
async function act(fn, undoText) {
  const ids = targets();
  if (!ids.length) return;
  const receipt = await run(async () => (await fn(ids)) || true);
  if (!receipt) return;
  if (undoText) offerUndo(`${undoText} ${mailsText(ids.length)}`, [receipt]);
  await openBox(box);
}
$('actTrash').onclick = () => {
  const count = selected.size || (openMail ? 1 : 0);
  if (count > BULK_CONFIRM && !confirm(`Move ${count} mails to Trash?`)) return;
  act(ids => provider.trash(ids), 'trash');
};
$('actRead').onclick = () => act(ids => provider.markRead(ids, true));
$('actUnread').onclick = () => act(ids => provider.markRead(ids, false));
$('actMove').onclick = async () => {
  if (!targets().length) return;
  const folderId = await pickFolder();
  if (folderId) act(ids => provider.moveToFolder(ids, folderId), 'move');
};

// The only permanent delete in the app.
$('actEmpty').onclick = async () => {
  const which = box === BOXES.trash ? 'trash' : 'spam', name = box.title, account = current;
  if (account.kind === 'gmail' && !account.fullAccess) {
    // Asked here, while the click still allows a popup. The token then also holds the normal scopes.
    try { account.token = await googleSignIn(account.address, [GOOGLE_FULL]); } catch (error) { return status(error.message, true); }
    account.fullAccess = true;
  }
  const count = await run(() => provider.folderCount(which));
  if (count === undefined) return;
  if (!count) return status(`${name} is empty.`);
  const typed = prompt(`Delete all ${mailsText(count)} in ${name} of ${me} for good?\nThis cannot be undone, not even in ${account.kind === 'outlook' ? 'Outlook' : 'Gmail'}.\n\nType DELETE to confirm.`);
  if (typed !== 'DELETE') return status('Nothing deleted.');
  const gone = await run(() => provider.emptyFolder(which));
  if (gone !== undefined) status(`${mailsText(gone)} deleted for good.`);
  offerUndo(null); // an older receipt may point at mail that is gone now
  await openBox(box);
};

// ---------- folders ----------
async function refreshFolders() {
  folders = (await run(() => provider.listFolders())) || folders;
  const totals = folderReport(scanRecords);
  $('folders').replaceChildren(...folders.map(f => el('li', {},
    el('button', { className: 'link', textContent: f.name, onclick: () => openBox({ title: f.name, folderId: f.id }) }),
    totals.has(f.id) ? el('span', { className: 'count', title: 'Scanned mails in this folder', textContent: `${totals.get(f.id).count} · ${sizeText(totals.get(f.id).size)}` }) : null)));
  if (!folders.length) $('folders').append(el('li', { className: 'muted', textContent: 'None yet.' }));
}

async function makeFolder(name) {
  const folder = await run(() => provider.createFolder(name));
  if (folder) await refreshFolders();
  return folder;
}
$('newFolder').onclick = async () => {
  const name = (prompt('Folder name') || '').trim();
  if (name) await makeFolder(name);
};

/** Dialog to pick an existing folder or name a new one. Resolves to a folder id, or null. */
function pickFolder() {
  const NEW = '\u0000new';
  const select = $('folderSelect'), input = $('folderNew'), dialog = $('folderDialog');
  select.replaceChildren(
    ...folders.map(f => el('option', { value: f.id, textContent: f.name })),
    el('option', { value: NEW, textContent: 'New folder…' }));
  input.value = '';
  select.onchange = () => { input.hidden = select.value !== NEW; };
  select.onchange();
  return new Promise(resolve => {
    dialog.onclose = async () => {
      if (dialog.returnValue !== 'ok') return resolve(null);
      if (select.value !== NEW) return resolve(select.value);
      const name = input.value.trim();
      if (!name) { status('No folder name given.', true); return resolve(null); }
      resolve((await makeFolder(name))?.id || null);
    };
    dialog.returnValue = '';
    dialog.showModal();
  });
}

// ---------- rules shared ----------
/** Saves a rule, then offers to run it on mail already in the mailbox. */
async function saveRule(rule, replacesId) {
  const made = await run(() => (replacesId ? provider.updateRule(replacesId, rule) : provider.createRule(rule)));
  if (!made) return false;
  status('Rule saved.');
  const match = await run(() => provider.countMatching(rule));
  if (match?.count && confirm(`Rule saved. Also apply it to the ${match.count}${match.more ? '+' : ''} mails already here?`)) {
    const changed = await run(() => provider.applyRuleToExisting(rule));
    if (changed) {
      status(`Applied to ${mailsText(changed.count)}.`);
      offerUndo(`rule on ${mailsText(changed.count)}`, [changed]);
    }
  }
  return true;
}

// ---------- aliases ----------
async function scan() {
  const account = current;
  if (account.scanning) return status('A scan is already running.');
  account.scanning = true;
  try { await scanAccount(account); } finally { account.scanning = false; }
}

async function scanAccount(account) {
  const whole = load(`fullScan:${account.address}`, false);
  const known = new Map((await loadScan(account.address)).map(r => [r.id, r]));
  const onScreen = () => current === account; // the user may switch account while a scan runs
  let records;
  try {
    records = await account.provider.scanAliases(known, (done, total) => {
      if (done % 500 === 0) saveScan(account.address, [...known.values()]); // a long scan keeps its progress if the tab closes
      if (onScreen()) $('scanStatus').textContent = `Scanning ${done} / ${total}`;
    }, whole ? Infinity : undefined);
  } catch (error) {
    await saveScan(account.address, [...known.values()]); // keep what was read, the next scan continues from there
    if (!onScreen()) return;
    $('scanStatus').textContent = `Scan stopped at ${known.size} mails. Press the scan button to continue.`;
    if (error instanceof AuthExpired) expire(account);
    else status(`Alias scan failed: ${error.message}`, true);
    return;
  }
  await saveScan(account.address, records);
  if (!onScreen()) return;
  scanRecords = records;
  $('scanStatus').textContent = `${whole ? 'Whole mailbox' : 'Newest mails'}: ${records.length} scanned, ${sizeText(records.reduce((sum, r) => sum + (r.size || 0), 0))}.`;
  setAliasRecords();
  refreshFolders();
}

function setAliasRecords() {
  aliasRecords = scanRecords.filter(r => r.alias);
  refreshSidebar();
  if (!$('view-leaks').hidden) renderLeaks();
  if (!$('view-senders').hidden) renderSenders();
}

const aliasFilter = alias => ({ to: aliasAddress(me, alias) });

function refreshSidebar() {
  const counts = new Map();
  for (const r of aliasRecords) counts.set(r.alias, (counts.get(r.alias) || 0) + 1);
  const rows = [...counts].sort((a, b) => b[1] - a[1]).map(([alias, count]) => el('li', {},
    el('button', { className: 'link', textContent: `+${alias}`, onclick: () => openBox({ title: `+${alias}`, filter: aliasFilter(alias) }) }),
    el('span', { className: 'count', textContent: count }),
    el('button', { textContent: 'Folder', title: `Sort +${alias} into a folder`, onclick: () => aliasToFolder(alias) })));
  $('aliases').replaceChildren(...rows);
  if (!rows.length) $('aliases').append(el('li', { className: 'muted', textContent: 'No +alias mail found.' }));
  const flagged = leakReport(aliasRecords, trusted()).filter(a => a.flagged).length;
  $('leakCount').textContent = flagged ? `(${flagged})` : '';
  $('leakCount').className = flagged ? 'flag' : '';
}

async function aliasToFolder(alias) {
  const folderId = await pickFolder();
  if (!folderId) return;
  await saveRule({ conditions: provider.aliasCondition(aliasAddress(me, alias)), exceptions: {}, actions: { moveToFolder: folderId } });
}
$('rescan').onclick = scan;
$('fullScan').onchange = () => {
  save(`fullScan:${me}`, $('fullScan').checked);
  if ($('fullScan').checked) scan();
};

// ---------- rules view ----------
// Gmail cannot switch a filter off. A paused Gmail filter is deleted and its definition kept here until resumed.
// shortcut: paused filters live in this browser only, like trust marks
const pausedKey = () => `paused:${me}`;
const presetsOf = rule => Object.fromEntries(Object.keys(KIND).map(section =>
  [section, Object.entries(rule[section] || {}).map(([key, value]) => ({ key, value }))]));

async function refreshRules() {
  const rules = await run(() => provider.listRules());
  if (!rules) return;
  const paused = provider.setRuleEnabled ? [] : load(pausedKey(), []);
  const button = (textContent, fn) => el('button', { textContent, onclick: async () => { if (await fn() !== false) refreshRules(); } });
  const row = (summary, on, ...buttons) => el('li', { className: on ? '' : 'muted' }, el('span', { textContent: `${on ? '' : '(paused) '}${summary}` }), ...buttons);
  const dropPaused = i => save(pausedKey(), load(pausedKey(), []).filter((_, j) => j !== i));
  $('rules').replaceChildren(
    ...rules.map(rule => row(rule.summary, rule.enabled,
      rule.rule
        ? button('Edit', () => { openBuilder(presetsOf(rule.rule), rule.id); return false; })
        : el('button', { textContent: 'Edit', disabled: true, title: 'Made outside this app in a form it cannot read. Delete and rebuild it here.' }),
      button(rule.enabled ? 'Pause' : 'Resume', () => pauseRule(rule, !rule.enabled)),
      button('Delete', () => confirm(`Delete this rule?\n\n${rule.summary}`) && run(() => provider.deleteRule(rule.id))))),
    ...paused.map((kept, i) => row(kept.summary, false,
      button('Resume', async () => { if (await run(() => provider.recreateRule(kept.raw))) dropPaused(i); }),
      button('Delete', () => (confirm(`Delete this paused rule?\n\n${kept.summary}`) ? dropPaused(i) : false)))));
  if (!rules.length && !paused.length) $('rules').append(el('li', { className: 'muted', textContent: 'No rules yet.' }));
}

async function pauseRule(rule, enable) {
  if (provider.setRuleEnabled) return run(() => provider.setRuleEnabled(rule.id, enable));
  const kept = load(pausedKey(), []);
  save(pausedKey(), [...kept, { summary: rule.summary, raw: rule.raw }]); // copy first, so a failed delete cannot lose the rule
  if (!(await run(() => provider.deleteRule(rule.id).then(() => true)))) save(pausedKey(), kept);
}

const KIND = { conditions: 'predicates', exceptions: 'predicates', actions: 'actions' };

/** Input for one catalog entry. The returned element has read() giving the rule value. */
function valueInput(entry, section, value) {
  const box = el('span', { className: 'value' });
  if (entry.type === 'strings' || entry.type === 'addresses') {
    const input = el('input', { value: [].concat(value ?? []).join(', '), placeholder: 'Several: separate with commas', ariaLabel: entry.label });
    box.append(input);
    box.read = () => input.value.split(',').map(s => s.trim()).filter(Boolean);
  } else if (entry.type === 'bool' && section === 'actions') {
    box.read = () => true;
  } else if (entry.type === 'bool') {
    const select = el('select', { ariaLabel: entry.label }, el('option', { value: 'yes', textContent: 'yes' }), el('option', { value: 'no', textContent: 'no' }));
    if (value === false) select.value = 'no';
    box.append(select);
    box.read = () => select.value === 'yes';
  } else if (entry.type === 'enum') {
    const select = el('select', { ariaLabel: entry.label }, ...entry.options.map(o => el('option', { value: o, textContent: o })));
    if (value) select.value = value;
    box.append(select);
    box.read = () => select.value;
  } else if (entry.type === 'sizeRange') {
    const min = el('input', { type: 'number', min: 0, placeholder: 'min KB', ariaLabel: 'Minimum size in KB', value: value?.minimumSize || '' });
    const max = el('input', { type: 'number', min: 0, placeholder: 'max KB', ariaLabel: 'Maximum size in KB', value: value?.maximumSize || '' });
    box.append(min, max);
    box.read = () => ({ minimumSize: Number(min.value) || 0, maximumSize: Number(max.value) || 0 });
  } else if (entry.type === 'folder') {
    const select = el('select', { ariaLabel: entry.label }, ...folders.map(f => el('option', { value: f.id, textContent: f.name })));
    if (value) select.value = value;
    box.append(select);
    box.read = () => select.value;
  }
  return box;
}

function ruleRow(section, preset) {
  const entries = RULE_CATALOG[KIND[section]];
  const supported = provider.support[KIND[section]];
  const select = el('select', { ariaLabel: 'Type' }, ...entries.map(e => el('option', {
    value: e.key,
    textContent: supported.has(e.key) ? e.label : `${e.label} (not on ${provider.name})`,
    disabled: !supported.has(e.key),
    title: supported.has(e.key) ? '' : `${provider.name} cannot do this.`,
  })));
  select.value = preset?.key || entries.find(e => supported.has(e.key)).key;
  const row = el('div', { className: 'rule-row' }, select);
  const draw = value => {
    row.querySelector('.value')?.remove();
    select.after(valueInput(entries.find(e => e.key === select.value), section, value));
  };
  select.onchange = () => draw();
  draw(preset?.value);
  row.append(el('button', { type: 'button', textContent: 'Remove', onclick: () => row.remove() }));
  row.read = () => [select.value, row.querySelector('.value').read()];
  return row;
}

function readSection(section) {
  const out = {};
  for (const row of $(section).children) {
    const [key, value] = row.read();
    if (Array.isArray(value) && !value.length) continue;
    out[key] = Array.isArray(value) && Array.isArray(out[key]) ? out[key].concat(value) : value;
  }
  return out;
}

let editingId = null; // the rule the builder replaces, or null for a new rule
function openBuilder(presets = {}, replacesId = null) {
  editingId = replacesId;
  showView('rules');
  $('builder').hidden = false;
  $('builderError').textContent = '';
  for (const section of Object.keys(KIND)) {
    $(section).replaceChildren(...(presets[section] || []).map(p => ruleRow(section, p)));
  }
  if (!$('conditions').children.length) $('conditions').append(ruleRow('conditions'));
  if (!$('actions').children.length) $('actions').append(ruleRow('actions'));
  $('builder').scrollIntoView();
}
$('newRule').onclick = () => openBuilder();
$('cancelRule').onclick = () => { $('builder').hidden = true; };
for (const button of document.querySelectorAll('[data-add]')) {
  button.onclick = () => $(button.dataset.add).append(ruleRow(button.dataset.add));
}

$('builder').onsubmit = async event => {
  event.preventDefault();
  const rule = { conditions: readSection('conditions'), exceptions: readSection('exceptions'), actions: readSection('actions') };
  try {
    provider.validateRule(rule);
  } catch (error) {
    $('builderError').textContent = error.message;
    return;
  }
  $('builderError').textContent = '';
  const sendsTo = [rule.actions.forwardTo, rule.actions.forwardAsAttachmentTo, rule.actions.redirectTo].flat().filter(Boolean);
  if (sendsTo.length && !confirm(`Every matching mail will be sent on to ${sendsTo.join(', ')}. Save?`)) return;
  if (rule.actions.permanentDelete && !confirm('Matching mail will be deleted for good, with no Trash. Save?')) return;
  if (await saveRule(rule, editingId)) {
    $('builder').hidden = true;
    refreshRules();
  }
};

$('ruleFromMail').onclick = () => {
  if (!openMail) return;
  const conditions = [{ key: 'fromAddresses', value: [senderAddress(openMail.from)] }];
  const alias = extractAlias(openMail.deliveredTo, me) ?? extractAlias(openMail.to, me);
  if (alias) for (const [key, value] of Object.entries(provider.aliasCondition(aliasAddress(me, alias)))) conditions.push({ key, value });
  openBuilder({ conditions });
};

// ---------- senders ----------
const SENDERS_SHOWN = 100;
const SENDERS_PER_RULE = 20;   // keeps one rule's search text well under Gmail's length limit
const pickedSenders = new Set();

function renderSenders() {
  const report = senderReport(scanRecords, me).slice(0, SENDERS_SHOWN);
  const shown = new Set(report.map(entry => entry.key));
  for (const key of pickedSenders) if (!shown.has(key)) pickedSenders.delete(key);
  $('senders').replaceChildren(...report.map(entry => {
    const check = el('input', { type: 'checkbox', checked: pickedSenders.has(entry.key), ariaLabel: `Select ${entry.key}` });
    const toggle = on => {
      if (on) pickedSenders.add(entry.key); else pickedSenders.delete(entry.key);
      check.checked = on;
      showPicked(report.length);
    };
    check.onclick = event => { event.stopPropagation(); toggle(check.checked); };
    return el('li', { onclick: () => toggle(!check.checked) },
      check,
      el('span', { className: 'name', textContent: entry.key }),
      el('span', { className: 'count', textContent: entry.count }),
      el('span', { className: 'count', textContent: sizeText(entry.size) }));
  }));
  if (!report.length) $('senders').append(el('li', { className: 'muted', textContent: 'No scanned mail yet.' }));
  $('sendersAll').onchange = () => {
    pickedSenders.clear();
    if ($('sendersAll').checked) for (const entry of report) pickedSenders.add(entry.key);
    renderSenders();
  };
  showPicked(report.length);
}

function showPicked(total) {
  $('sendersPicked').textContent = pickedSenders.size ? `${pickedSenders.size} selected` : 'Tick one or more senders.';
  $('sendersAll').checked = total > 0 && pickedSenders.size === total;
}

function pickedList() {
  if (!pickedSenders.size) status('Tick at least one sender first.', true);
  return [...pickedSenders];
}
const namesOf = list => (list.length <= 5 ? list.join(', ') : `${list.slice(0, 5).join(', ')} and ${list.length - 5} more`);

$('sendersShow').onclick = () => {
  const senders = pickedList();
  if (senders.length) openBox({ title: `From ${namesOf(senders)}`, filter: { from: senders } });
};
$('sendersUnsub').onclick = async () => {
  const senders = pickedList();
  if (!senders.length || !confirm(`Unsubscribe from ${namesOf(senders)}? One-click requests go out at once; the others show a link or ask first.`)) return;
  const sent = [], none = [], links = [];
  for (const sender of senders) {
    const page = await run(() => provider.listMessages({ folderId: 'ALL', filter: { from: sender } }));
    // Not every mail of a sender carries the header, so the newest three are tried.
    let info = parseUnsubscribe('');
    for (const { id } of (page?.messages || []).slice(0, 3)) {
      const mail = await run(() => provider.getMessage(id));
      info = parseUnsubscribe(mail?.listUnsubscribe, mail?.listUnsubscribePost);
      if (info.oneClick || info.link || info.mailto) break;
    }
    const done = await unsubscribe(info, sender);
    if (done.link) links.push([sender, done.link]);
    else (done.ok ? sent : none).push(sender);
  }
  // Pages cannot all be opened at once (popup blocker), so each gets its own button.
  $('unsubLinks').replaceChildren(...links.map(([sender, link]) => el('li', {},
    el('span', { textContent: `${sender} needs its own page:` }),
    el('button', { textContent: 'Open unsubscribe page', onclick: () => window.open(link, '_blank', 'noopener,noreferrer') }))));
  status([sent.length && `Done: ${namesOf(sent)}.`, links.length && `${links.length} need a page, see the list.`, none.length && `Not done: ${namesOf(none)}.`].filter(Boolean).join(' '));
};
$('sendersTrash').onclick = () => trashSenders(pickedList(), false);
$('sendersBlock').onclick = () => trashSenders(pickedList(), true);
$('sendersMove').onclick = async () => {
  const senders = pickedList();
  if (!senders.length) return;
  const folderId = await pickFolder();
  if (folderId) trashSenders(senders, true, folderId);
};

/**
 * Trashes every mail from the given senders. With `block`, rules also trash future mail from them.
 * With `folderId`, moves into that folder instead of Trash (and the rules do the same).
 */
async function trashSenders(senders, block, folderId) {
  if (!senders.length) return;
  const where = folderId ? `to ${folders.find(f => f.id === folderId)?.name || 'the folder'}` : 'to Trash';
  const rules = chunks(senders, SENDERS_PER_RULE)
    .map(part => ({ conditions: { senderContains: part }, exceptions: {}, actions: folderId ? { moveToFolder: folderId } : { delete: true } }));
  let count = 0, more = false;
  for (const rule of rules) {
    const match = await run(() => provider.countMatching(rule));
    if (!match) return;
    count += match.count;
    more ||= match.more;
  }
  if (!count && !block) return status('No mail from these senders outside Trash.');
  const question = `Move ${count}${more ? '+' : ''} mails from ${namesOf(senders)} ${where}` +
    `${block ? `, and send all future mail from them ${where}` : ''}?`;
  if (!confirm(question)) return;
  let changed = 0;
  const receipts = [], made = [], records = [];
  for (const rule of rules) {
    if (block) {
      const saved = await run(() => provider.createRule(rule));
      if (!saved) break;
      made.push(saved.id);
    }
    const done = await run(() => provider.applyRuleToExisting(rule));
    if (!done) break;
    receipts.push(done);
    changed += done.count;
    const gone = new Set(rule.conditions.senderContains);
    for (const key of gone) pickedSenders.delete(key);
    if (folderId) continue; // moved mail stays in the mailbox, so its scan records stay
    // Drop what is finished at once, so a failure half way leaves the list matching the mailbox.
    records.push(...scanRecords.filter(r => gone.has(senderKey(r))));
    scanRecords = scanRecords.filter(r => !gone.has(senderKey(r)));
  }
  saveScan(me, scanRecords);
  setAliasRecords();
  offerUndo(`${folderId ? 'move + rule' : block ? 'trash + block' : 'trash'} ${mailsText(changed)}`, receipts, { rules: made, records });
  if (changed || block) status(`${changed} mails moved ${where}.${block ? ` Future mail from these senders goes ${where}.` : ''}`);
}

// ---------- cleanup ----------
function cleanupFilter() {
  const years = Number($('cleanYears').value), mb = Number($('cleanMB').value), unread = $('cleanUnread').checked;
  const filter = {};
  if (years > 0) filter.before = Date.now() - years * 365.25 * 864e5;
  if (mb > 0) filter.larger = mb * 1048576;
  if (unread) filter.unread = true;
  const words = [years > 0 && `older than ${years} years`, mb > 0 && `larger than ${mb} MB`, unread && 'never opened'].filter(Boolean);
  if (!words.length) status('Fill in at least one condition.', true);
  return words.length ? { filter, words: words.join(', ') } : null;
}
$('cleanup').onsubmit = event => {
  event.preventDefault();
  const pick = cleanupFilter();
  if (pick) openBox({ title: `Cleanup: ${pick.words}`, folderId: 'ALL', filter: pick.filter });
};
$('cleanTrash').onclick = async () => {
  const pick = cleanupFilter();
  if (!pick) return;
  const match = await run(() => provider.countMessages(pick.filter));
  if (!match) return;
  if (!match.count) return status('No mail outside Trash matches.');
  if (!confirm(`Move ${match.count}${match.more ? '+' : ''} mails (${pick.words}) to Trash?`)) return;
  const receipt = await run(() => provider.trashMessages(pick.filter));
  if (!receipt) return;
  status(`${mailsText(receipt.count)} moved to Trash.`);
  offerUndo(`cleanup ${mailsText(receipt.count)}`, [receipt]);
};

// ---------- leaks ----------
// shortcut: trust marks live in this browser only; move them into the mail account if they must survive a browser reset
const trusted = () => load(`trusted:${me}`, {});

function renderLeaks() {
  const report = leakReport(aliasRecords, trusted());
  $('leaks').replaceChildren(...report.map(entry => el('div', { className: `leak ${entry.flagged ? 'flagged' : ''}` },
    el('h3', { textContent: `+${entry.alias} (${entry.total} mail${entry.total === 1 ? '' : 's'})${entry.flagged ? ' — possible leak' : ''}` }),
    el('ul', {}, ...entry.domains.map(d => {
      const real = d.domain !== '(unknown)';
      const foreign = d.domain !== entry.home;
      return el('li', {},
        el('span', { className: `domain ${d.trusted ? '' : 'flag'}`, textContent: `${d.domain}${foreign ? '' : ' (main sender)'}` }),
        el('span', { className: 'count', textContent: d.count }),
        real ? el('button', { textContent: 'Show mail', onclick: () => openBox({ title: `+${entry.alias} from ${d.domain}`, filter: { ...aliasFilter(entry.alias), from: d.domain } }) }) : null,
        foreign && !d.trusted ? el('button', { textContent: 'Trust', onclick: () => trust(entry.alias, d.domain) }) : null,
        foreign && real ? el('button', { textContent: 'Block', onclick: () => block(entry.alias, d.domain) }) : null);
    })))));
  if (!report.length) $('leaks').append(el('p', { className: 'muted', textContent: 'No +alias mail found yet.' }));
}

function trust(alias, domain) {
  const all = trusted();
  all[alias] = [...new Set([...(all[alias] || []), domain])];
  save(`trusted:${me}`, all);
  renderLeaks();
  refreshSidebar();
}

async function block(alias, domain) {
  if (!confirm(`Trash all mail from ${domain} to +${alias}, now and in future?`)) return;
  const rule = { conditions: { ...provider.aliasCondition(aliasAddress(me, alias)), senderContains: [domain] }, exceptions: {}, actions: { delete: true } };
  const made = await run(() => provider.createRule(rule));
  if (!made) return;
  const changed = await run(() => provider.applyRuleToExisting(rule));
  if (!changed) return offerUndo('block', [], { rules: [made.id] });
  status(`Blocked. ${changed.count} mails moved to Trash.`);
  const hit = r => r.alias === alias && r.domain === domain;
  const records = scanRecords.filter(hit);
  scanRecords = scanRecords.filter(r => !hit(r));
  saveScan(me, scanRecords);
  setAliasRecords();
  offerUndo(`block ${mailsText(changed.count)}`, [changed], { rules: [made.id], records });
}

// ---------- boot ----------
if (new URLSearchParams(location.search).has('demo')) {
  const demo = { kind: 'demo', provider: createDemo() };
  demo.address = await demo.provider.ownAddress();
  accounts.push(demo);
  switchTo(demo);
} else {
  showGate(GOOGLE_ID || MICROSOFT_ID
    ? 'Sign in to an account. Add more accounts later from the sidebar.'
    : 'Not set up yet. Follow SETUP.md (Gmail) or SETUP-OUTLOOK.md, paste the client ID into config.js and reload.');
}
