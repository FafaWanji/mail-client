import { GOOGLE_CLIENT_ID } from '../config.js';
import { signIn, getToken } from './auth.js';
import { createGmail, AuthExpired } from './gmail.js';
import { createDemo } from './demo.js';
import { buildFrameDoc, FRAME_SANDBOX } from './mailbody.js';
import { aliasAddress, extractAlias } from './alias.js';
import { RULE_CATALOG } from './rules.js';
import { leakReport } from './leaks.js';

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
const senderAddress = from => (from.match(/[^\s<>"]+@[^\s<>"]+/) || [''])[0].toLowerCase();

let provider, me;
let box = { title: 'Inbox' };  // { title, folderId?, query? }
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
    if (error instanceof AuthExpired) showGate('Session expired. Sign in again.');
    else status(error.message, true);
    return undefined;
  } finally {
    document.body.classList.remove('busy');
  }
}

// ---------- sign-in ----------
function showGate(text) {
  $('app').hidden = true;
  $('gate').hidden = false;
  $('gateText').textContent = text;
  $('signIn').hidden = !GOOGLE_CLIENT_ID;
}

async function start() {
  me = await provider.ownAddress();
  $('gate').hidden = true;
  $('app').hidden = false;
  $('account').textContent = me;
  await refreshFolders();
  await openBox({ title: 'Inbox' });
  scan();
}

$('signIn').onclick = async () => {
  try {
    await signIn();
    provider = createGmail({ getToken });
    await start();
  } catch (error) {
    showGate(error.message);
  }
};
$('signOut').onclick = () => location.reload(); // the token lives in memory only, so a reload signs out

// ---------- views ----------
function showView(name) {
  for (const view of ['mail', 'rules', 'leaks']) $(`view-${view}`).hidden = view !== name;
  if (name === 'rules') refreshRules();
  if (name === 'leaks') renderLeaks();
}
for (const button of document.querySelectorAll('[data-view]')) button.onclick = () => showView(button.dataset.view);

const BOXES = {
  inbox: { title: 'Inbox' },
  all: { title: 'All Mail', folderId: 'ALL' },
  trash: { title: 'Trash', query: 'in:trash' },
};
for (const button of document.querySelectorAll('[data-box]')) button.onclick = () => openBox(BOXES[button.dataset.box]);

// ---------- mail list ----------
async function openBox(next) {
  box = next;
  showView('mail');
  $('boxTitle').textContent = box.title;
  messages = [];
  selected.clear();
  closeMail();
  await loadPage(undefined);
}

async function loadPage(pageToken) {
  const page = await run(() => provider.listMessages({ folderId: box.folderId, query: box.query, pageToken }));
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

// ---------- cleanup actions ----------
function targets() {
  if (selected.size) return [...selected];
  if (openMail) return [openMail.id];
  status('Select a mail first.', true);
  return [];
}
async function act(fn) {
  const ids = targets();
  if (!ids.length) return;
  const done = await run(async () => { await fn(ids); return true; });
  if (done) await openBox(box);
}
$('actTrash').onclick = () => {
  const count = selected.size || (openMail ? 1 : 0);
  if (count > BULK_CONFIRM && !confirm(`Move ${count} mails to Trash?`)) return;
  act(ids => provider.trash(ids));
};
$('actRead').onclick = () => act(ids => provider.markRead(ids, true));
$('actUnread').onclick = () => act(ids => provider.markRead(ids, false));
$('actMove').onclick = async () => {
  if (!targets().length) return;
  const folderId = await pickFolder();
  if (folderId) act(ids => provider.moveToFolder(ids, folderId));
};

// ---------- folders ----------
async function refreshFolders() {
  folders = (await run(() => provider.listFolders())) || folders;
  $('folders').replaceChildren(...folders.map(f => el('li', {},
    el('button', { className: 'link', textContent: f.name, onclick: () => openBox({ title: f.name, folderId: f.id }) }))));
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
async function saveRule(rule) {
  const made = await run(() => provider.createRule(rule));
  if (!made) return false;
  status('Rule saved.');
  const match = await run(() => provider.countMatching(rule));
  if (match?.count && confirm(`Rule saved. Also apply it to the ${match.count}${match.more ? '+' : ''} mails already here?`)) {
    const changed = await run(() => provider.applyRuleToExisting(rule));
    if (changed !== undefined) status(`Applied to ${changed} mails.`);
  }
  return true;
}

// ---------- aliases ----------
async function scan() {
  const key = `aliasScan:${me}`;
  const known = new Map(load(key, []).map(r => [r.id, r]));
  try {
    scanRecords = await provider.scanAliases(known, (done, total) => { $('scanStatus').textContent = `Scanning ${done} / ${total}`; });
  } catch (error) {
    $('scanStatus').textContent = '';
    if (error instanceof AuthExpired) showGate('Session expired. Sign in again.');
    else status(`Alias scan failed: ${error.message}`, true);
    return;
  }
  save(key, scanRecords);
  $('scanStatus').textContent = `Newest ${scanRecords.length} mails scanned.`;
  setAliasRecords();
}

function setAliasRecords() {
  aliasRecords = scanRecords.filter(r => r.alias);
  refreshSidebar();
  if (!$('view-leaks').hidden) renderLeaks();
}

const aliasQuery = alias => `to:("${aliasAddress(me, alias)}")`;

function refreshSidebar() {
  const counts = new Map();
  for (const r of aliasRecords) counts.set(r.alias, (counts.get(r.alias) || 0) + 1);
  const rows = [...counts].sort((a, b) => b[1] - a[1]).map(([alias, count]) => el('li', {},
    el('button', { className: 'link', textContent: `+${alias}`, onclick: () => openBox({ title: `+${alias}`, query: aliasQuery(alias) }) }),
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
  await saveRule({ conditions: { sentToAddresses: [aliasAddress(me, alias)] }, exceptions: {}, actions: { moveToFolder: folderId } });
}
$('rescan').onclick = scan;

// ---------- rules view ----------
async function refreshRules() {
  const rules = await run(() => provider.listRules());
  if (!rules) return;
  $('rules').replaceChildren(...rules.map(rule => el('li', {},
    el('span', { textContent: rule.summary }),
    el('button', {
      textContent: 'Delete',
      onclick: async () => {
        if (!confirm(`Delete this rule?\n\n${rule.summary}`)) return;
        await run(() => provider.deleteRule(rule.id));
        refreshRules();
      },
    }))));
  if (!rules.length) $('rules').append(el('li', { className: 'muted', textContent: 'No rules yet.' }));
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
    const min = el('input', { type: 'number', min: 0, placeholder: 'min KB', ariaLabel: 'Minimum size in KB' });
    const max = el('input', { type: 'number', min: 0, placeholder: 'max KB', ariaLabel: 'Maximum size in KB' });
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
    title: supported.has(e.key) ? '' : `${provider.name} filters cannot do this. Available on Outlook accounts.`,
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

function openBuilder(presets = {}) {
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
  if (await saveRule(rule)) {
    $('builder').hidden = true;
    refreshRules();
  }
};

$('ruleFromMail').onclick = () => {
  if (!openMail) return;
  const conditions = [{ key: 'fromAddresses', value: [senderAddress(openMail.from)] }];
  const alias = extractAlias(openMail.deliveredTo, me) ?? extractAlias(openMail.to, me);
  if (alias) conditions.push({ key: 'sentToAddresses', value: [aliasAddress(me, alias)] });
  openBuilder({ conditions });
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
        real ? el('button', { textContent: 'Show mail', onclick: () => openBox({ title: `+${entry.alias} from ${d.domain}`, query: `${aliasQuery(entry.alias)} from:("${d.domain}")` }) }) : null,
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
  const rule = { conditions: { sentToAddresses: [aliasAddress(me, alias)], senderContains: [domain] }, exceptions: {}, actions: { delete: true } };
  const made = await run(() => provider.createRule(rule));
  if (!made) return;
  const changed = await run(() => provider.applyRuleToExisting(rule));
  if (changed === undefined) return;
  status(`Blocked. ${changed} mails moved to Trash.`);
  scanRecords = scanRecords.filter(r => !(r.alias === alias && r.domain === domain));
  save(`aliasScan:${me}`, scanRecords);
  setAliasRecords();
}

// ---------- boot ----------
if (new URLSearchParams(location.search).has('demo')) {
  provider = createDemo();
  start();
} else {
  showGate(GOOGLE_CLIENT_ID
    ? 'Sign in to read and sort your Gmail.'
    : 'Not set up yet. Follow SETUP.md, then paste the client ID into config.js and reload.');
}
