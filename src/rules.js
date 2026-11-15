// Rule model: { id?, conditions, exceptions, actions }.
// Keys are Microsoft Graph messageRulePredicates / messageRuleActions names (checked 2026-11-13),
// so an Outlook rule passes through unchanged. Address predicates hold plain strings here.

const IMPORTANCE = ['low', 'normal', 'high'];

export const RULE_CATALOG = {
  predicates: [
    { key: 'fromAddresses', label: 'From address', type: 'addresses' },
    { key: 'senderContains', label: 'Sender contains', type: 'strings' },
    { key: 'sentToAddresses', label: 'Sent to address', type: 'addresses' },
    { key: 'recipientContains', label: 'Recipient contains', type: 'strings' },
    { key: 'subjectContains', label: 'Subject contains', type: 'strings' },
    { key: 'bodyContains', label: 'Body contains', type: 'strings' },
    { key: 'bodyOrSubjectContains', label: 'Subject or body contains', type: 'strings' },
    { key: 'hasAttachments', label: 'Has attachment', type: 'bool' },
    { key: 'withinSizeRange', label: 'Size between (KB)', type: 'sizeRange' },
    { key: 'headerContains', label: 'Header contains', type: 'strings' },
    { key: 'categories', label: 'Has category', type: 'strings' },
    { key: 'importance', label: 'Importance is', type: 'enum', options: IMPORTANCE },
    { key: 'sensitivity', label: 'Sensitivity is', type: 'enum', options: ['normal', 'personal', 'private', 'confidential'] },
    { key: 'messageActionFlag', label: 'Flagged for', type: 'enum', options: ['any', 'call', 'doNotForward', 'followUp', 'fyi', 'forward', 'noResponseNecessary', 'read', 'reply', 'replyToAll', 'review'] },
    { key: 'sentToMe', label: 'I am on the To line', type: 'bool' },
    { key: 'sentOnlyToMe', label: 'Sent only to me', type: 'bool' },
    { key: 'sentCcMe', label: 'I am on the Cc line', type: 'bool' },
    { key: 'sentToOrCcMe', label: 'I am on To or Cc', type: 'bool' },
    { key: 'notSentToMe', label: 'I am not on the To line', type: 'bool' },
    { key: 'isApprovalRequest', label: 'Is approval request', type: 'bool' },
    { key: 'isAutomaticForward', label: 'Is automatic forward', type: 'bool' },
    { key: 'isAutomaticReply', label: 'Is automatic reply', type: 'bool' },
    { key: 'isEncrypted', label: 'Is encrypted', type: 'bool' },
    { key: 'isMeetingRequest', label: 'Is meeting request', type: 'bool' },
    { key: 'isMeetingResponse', label: 'Is meeting response', type: 'bool' },
    { key: 'isNonDeliveryReport', label: 'Is non-delivery report', type: 'bool' },
    { key: 'isPermissionControlled', label: 'Is permission controlled', type: 'bool' },
    { key: 'isReadReceipt', label: 'Is read receipt', type: 'bool' },
    { key: 'isSigned', label: 'Is signed', type: 'bool' },
    { key: 'isVoicemail', label: 'Is voicemail', type: 'bool' },
  ],
  actions: [
    { key: 'moveToFolder', label: 'Move to folder (skip inbox)', type: 'folder' },
    { key: 'copyToFolder', label: 'Copy to folder (keep in inbox)', type: 'folder' },
    { key: 'markAsRead', label: 'Mark as read', type: 'bool' },
    { key: 'delete', label: 'Move to Trash', type: 'bool' },
    { key: 'markImportance', label: 'Mark importance', type: 'enum', options: IMPORTANCE },
    { key: 'forwardTo', label: 'Forward to', type: 'addresses' },
    { key: 'gmailStar', label: 'Star (Gmail only)', type: 'bool' },
    { key: 'gmailNeverSpam', label: 'Never send to Spam (Gmail only)', type: 'bool' },
    { key: 'assignCategories', label: 'Assign categories', type: 'strings' },
    { key: 'forwardAsAttachmentTo', label: 'Forward as attachment to', type: 'addresses' },
    { key: 'redirectTo', label: 'Redirect to', type: 'addresses' },
    { key: 'permanentDelete', label: 'Delete permanently', type: 'bool' },
    { key: 'stopProcessingRules', label: 'Stop processing more rules', type: 'bool' },
  ],
};

// Order here is the order of clauses in the generated query.
export const GMAIL_SUPPORT = {
  predicates: new Set(['fromAddresses', 'senderContains', 'sentToAddresses', 'recipientContains', 'subjectContains',
    'bodyContains', 'bodyOrSubjectContains', 'hasAttachments', 'withinSizeRange']),
  actions: new Set(['moveToFolder', 'copyToFolder', 'markAsRead', 'delete', 'markImportance', 'forwardTo',
    'gmailStar', 'gmailNeverSpam']),
};

export class UnsupportedOnGmail extends Error {
  constructor(key, detail) {
    const entry = [...RULE_CATALOG.predicates, ...RULE_CATALOG.actions].find(e => e.key === key);
    super(`Gmail filters cannot do "${entry ? entry.label : key}"${detail ? `: ${detail}` : ''}`);
    this.key = key;
  }
}

const PREFIX = {
  fromAddresses: 'from:', senderContains: 'from:',
  sentToAddresses: 'to:', recipientContains: 'to:',
  subjectContains: 'subject:', bodyContains: '', bodyOrSubjectContains: '',
};

const isSet = v => v !== undefined && v !== null && v !== '' && !(Array.isArray(v) && v.length === 0);

// A value is always one quoted literal, so a quote or operator inside it cannot widen the rule.
function terms(values) {
  return [].concat(values).map(v => String(v).replaceAll('"', '').trim()).filter(Boolean).map(v => `"${v}"`);
}

function clauses(predicates) {
  for (const key of Object.keys(predicates || {})) {
    if (isSet(predicates[key]) && !GMAIL_SUPPORT.predicates.has(key)) throw new UnsupportedOnGmail(key);
  }
  const out = [];
  for (const key of GMAIL_SUPPORT.predicates) {
    const value = (predicates || {})[key];
    if (!isSet(value)) continue;
    if (key === 'hasAttachments') out.push(value ? 'has:attachment' : '-has:attachment');
    else if (key === 'withinSizeRange') {
      const size = [];
      if (value.minimumSize > 0) size.push(`larger:${value.minimumSize * 1024}`);
      if (value.maximumSize > 0) size.push(`smaller:${value.maximumSize * 1024}`);
      if (size.length) out.push(size.join(' '));
    } else {
      const t = terms(value);
      if (t.length) out.push(`${PREFIX[key]}(${t.join(' OR ')})`);
    }
  }
  return out;
}

/** Gmail search string that matches when every predicate matches. */
export function predicatesToQuery(predicates) {
  return clauses(predicates).join(' ');
}

const negate = clause =>
  clause.startsWith('-') ? clause.slice(1) : `-${/^(larger|smaller):/.test(clause) ? `(${clause})` : clause}`;

/** Search string for mail already in the mailbox: all conditions, none of the exceptions. */
export function ruleToSearch(rule) {
  return [...clauses(rule.conditions), ...clauses(rule.exceptions).map(negate)].join(' ');
}

export function toGmailFilter(rule) {
  const query = predicatesToQuery(rule.conditions);
  if (!query) throw new Error('A rule needs at least one condition.');
  // Outlook exceptions mean "unless any of these matches", hence OR.
  const negatedQuery = clauses(rule.exceptions).map(c => (c.includes(' ') ? `(${c})` : c)).join(' OR ');

  const actions = rule.actions || {};
  for (const key of Object.keys(actions)) {
    if (isSet(actions[key]) && actions[key] !== false && !GMAIL_SUPPORT.actions.has(key)) throw new UnsupportedOnGmail(key);
  }
  const add = [], remove = [];
  let forward;
  if (isSet(actions.moveToFolder)) { add.push(actions.moveToFolder); remove.push('INBOX'); }
  if (isSet(actions.copyToFolder)) add.push(actions.copyToFolder);
  if (actions.markAsRead) remove.push('UNREAD');
  if (actions.delete) add.push('TRASH');
  if (actions.markImportance === 'high') add.push('IMPORTANT');
  if (actions.markImportance === 'low') remove.push('IMPORTANT');
  if (isSet(actions.forwardTo)) {
    const to = [].concat(actions.forwardTo);
    if (to.length !== 1) throw new UnsupportedOnGmail('forwardTo', 'only one forwarding address per filter');
    forward = to[0];
  }
  if (actions.gmailStar) add.push('STARRED');
  if (actions.gmailNeverSpam) remove.push('SPAM');
  if (!add.length && !remove.length && !forward) throw new Error('A rule needs at least one action.');

  const filter = { criteria: { query }, action: {} };
  if (negatedQuery) filter.criteria.negatedQuery = negatedQuery;
  if (add.length) filter.action.addLabelIds = add;
  if (remove.length) filter.action.removeLabelIds = remove;
  if (forward) filter.action.forward = forward;
  return filter;
}

const ADDED = { TRASH: 'trash', STARRED: 'star', IMPORTANT: 'mark important' };
const REMOVED = { INBOX: 'skip inbox', UNREAD: 'mark read', SPAM: 'never spam', IMPORTANT: 'mark not important' };

/** One readable line for any Gmail filter, also ones made in Gmail itself. */
export function describeGmailFilter(filter, labelNames) {
  const c = filter.criteria || {}, a = filter.action || {};
  const when = [];
  if (c.from) when.push(`from:${c.from}`);
  if (c.to) when.push(`to:${c.to}`);
  if (c.subject) when.push(`subject:${c.subject}`);
  if (c.query) when.push(c.query);
  if (c.hasAttachment) when.push('has:attachment');
  if (c.size) when.push(`${c.sizeComparison === 'smaller' ? 'smaller' : 'larger'}:${c.size}`);
  const name = id => labelNames.get(id) || id;
  const then = [
    ...(a.addLabelIds || []).map(id => ADDED[id] || `add ${name(id)}`),
    ...(a.removeLabelIds || []).map(id => REMOVED[id] || `remove ${name(id)}`),
    ...(a.forward ? [`forward to ${a.forward}`] : []),
  ];
  return `If ${when.join(' ') || '(any mail)'}${c.negatedQuery ? ` except ${c.negatedQuery}` : ''}: ${then.join(', ') || '(nothing)'}`;
}

// ---------- Outlook ----------

const GMAIL_ONLY = new Set(['gmailStar', 'gmailNeverSpam']);
export const OUTLOOK_SUPPORT = {
  predicates: new Set(RULE_CATALOG.predicates.map(e => e.key)),
  actions: new Set(RULE_CATALOG.actions.map(e => e.key).filter(key => !GMAIL_ONLY.has(key))),
};

// Graph wants these as recipient objects, the rule model holds plain addresses.
const RECIPIENT_KEYS = new Set(['fromAddresses', 'sentToAddresses', 'forwardTo', 'forwardAsAttachmentTo', 'redirectTo']);
const FOLDER_KEYS = new Set(['moveToFolder', 'copyToFolder']);
const labelOf = key => [...RULE_CATALOG.predicates, ...RULE_CATALOG.actions].find(e => e.key === key)?.label || key;

function graphPart(part, kind) {
  const out = {};
  for (const [key, value] of Object.entries(part || {})) {
    if (!isSet(value) || (kind === 'actions' && value === false)) continue;
    if (!OUTLOOK_SUPPORT[kind].has(key)) throw new Error(`Outlook rules cannot do "${labelOf(key)}".`);
    out[key] = RECIPIENT_KEYS.has(key) ? [].concat(value).map(address => ({ emailAddress: { address } })) : value;
  }
  return out;
}

function describePart(part, folderNames) {
  return Object.entries(part || {}).map(([key, value]) => {
    if (value === true) return labelOf(key);
    let text;
    if (RECIPIENT_KEYS.has(key)) text = [].concat(value).map(r => r?.emailAddress?.address ?? r).join(', ');
    else if (FOLDER_KEYS.has(key)) text = folderNames.get(value) || value;
    else if (key === 'withinSizeRange') text = `${value.minimumSize}-${value.maximumSize} KB`;
    else text = [].concat(value).join(', ');
    return `${labelOf(key)}: ${text}`;
  }).join('; ');
}

/** One readable line for an Outlook inbox rule (Graph messageRule). */
export function describeOutlookRule(raw, folderNames) {
  const when = describePart(raw.conditions, folderNames) || '(any mail)';
  const unless = describePart(raw.exceptions, folderNames);
  return `If ${when}${unless ? ` except ${unless}` : ''}: ${describePart(raw.actions, folderNames) || '(nothing)'}`;
}

/** Graph messageRule body for a rule. */
export function toOutlookRule(rule, folderNames) {
  const actions = graphPart(rule.actions, 'actions');
  if (!Object.keys(actions).length) throw new Error('A rule needs at least one action.');
  const body = { sequence: 1, isEnabled: true, actions };
  const conditions = graphPart(rule.conditions, 'predicates');
  const exceptions = graphPart(rule.exceptions, 'predicates');
  if (Object.keys(conditions).length) body.conditions = conditions;
  if (Object.keys(exceptions).length) body.exceptions = exceptions;
  body.displayName = describeOutlookRule(body, folderNames).slice(0, 250);
  return body;
}

// ---------- reading rules back, for editing ----------

const sameSet = (a = [], b = []) => a.length === b.length && [...a].sort().join('\n') === [...b].sort().join('\n');
const sameFilter = (a, b) => (a.criteria.query || '') === (b.criteria.query || '') &&
  (a.criteria.negatedQuery || '') === (b.criteria.negatedQuery || '') &&
  sameSet(a.action.addLabelIds, b.action.addLabelIds) && sameSet(a.action.removeLabelIds, b.action.removeLabelIds) &&
  (a.action.forward || '') === (b.action.forward || '');

// One clause as predicatesToQuery writes it. Sticky, so it only matches at the current position.
const CLAUSE = /(from:|to:|subject:|)\(("[^"]*"(?: OR "[^"]*")*)\)|(-?)has:attachment|larger:(\d+)(?: smaller:(\d+))?|smaller:(\d+)/y;
const KEY_OF = { 'from:': 'senderContains', 'to:': 'recipientContains', 'subject:': 'subjectContains', '': 'bodyContains' };

// Reads clauses joined by `glue` (' ' for conditions, ' OR ' for exceptions). Null if any part is not ours.
function parseClauses(text, glue) {
  const out = {};
  let at = 0;
  while (at < text.length) {
    if (at > 0) { if (!text.startsWith(glue, at)) return null; at += glue.length; }
    const wrapped = text[at] === '(' && text[at + 1] !== '"';
    CLAUSE.lastIndex = at + (wrapped ? 1 : 0);
    const m = CLAUSE.exec(text);
    if (!m) return null;
    at = CLAUSE.lastIndex;
    if (wrapped) { if (text[at] !== ')') return null; at++; }
    let key, value;
    if (m[2] !== undefined) {
      value = m[2].slice(1, -1).split('" OR "');
      key = KEY_OF[m[1]];
      if (key !== 'subjectContains' && key !== 'bodyContains' && value.every(v => v.includes('@'))) key = key === 'senderContains' ? 'fromAddresses' : 'sentToAddresses';
    } else if (m[0].endsWith('has:attachment')) { key = 'hasAttachments'; value = m[3] !== '-'; }
    else {
      const [min, max] = [m[4], m[5] || m[6]].map(n => Number(n || 0));
      if (min % 1024 || max % 1024) return null; // the builder works in whole KB
      key = 'withinSizeRange'; value = { minimumSize: min / 1024, maximumSize: max / 1024 };
    }
    if (key in out) return null;
    out[key] = value;
  }
  return out;
}

/** The rule model behind a Gmail filter, or null when the filter was not made the way this app makes them. */
export function fromGmailFilter(raw) {
  const c = raw.criteria || {}, a = raw.action || {};
  if (Object.keys(c).some(k => k !== 'query' && k !== 'negatedQuery') || !c.query) return null;
  const conditions = parseClauses(c.query, ' ');
  const exceptions = c.negatedQuery ? parseClauses(c.negatedQuery, ' OR ') : {};
  if (!conditions || !exceptions) return null;
  const add = a.addLabelIds || [], remove = a.removeLabelIds || [];
  const actions = {};
  for (const id of add) {
    if (id === 'TRASH') actions.delete = true;
    else if (id === 'STARRED') actions.gmailStar = true;
    else if (id === 'IMPORTANT') actions.markImportance = 'high';
    else if (remove.includes('INBOX') && !actions.moveToFolder) actions.moveToFolder = id;
    else if (!actions.copyToFolder) actions.copyToFolder = id;
    else return null;
  }
  if (remove.includes('UNREAD')) actions.markAsRead = true;
  if (remove.includes('SPAM')) actions.gmailNeverSpam = true;
  if (remove.includes('IMPORTANT')) actions.markImportance = 'low';
  if (a.forward) actions.forwardTo = [a.forward];
  const rule = { conditions, exceptions, actions };
  try {
    return sameFilter(toGmailFilter(rule), { criteria: c, action: a }) ? rule : null;
  } catch {
    return null;
  }
}

/** The rule model behind an Outlook rule, or null when it uses something this app does not know. */
export function fromOutlookRule(raw) {
  const part = (graph, kind) => {
    const out = {};
    for (const [key, value] of Object.entries(graph || {})) {
      if (!OUTLOOK_SUPPORT[kind].has(key)) return null;
      out[key] = RECIPIENT_KEYS.has(key) ? value.map(r => r?.emailAddress?.address).filter(Boolean) : value;
    }
    return out;
  };
  const rule = { conditions: part(raw.conditions, 'predicates'), exceptions: part(raw.exceptions, 'predicates'), actions: part(raw.actions, 'actions') };
  return rule.conditions && rule.exceptions && rule.actions ? rule : null;
}
