# Mail Client (Gmail part) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A local web page that reads and cleans a Gmail mailbox by `+alias`, sorts aliases into folders, flags leaked aliases, and saves rules as real Gmail filters.

**Architecture:** Static page, plain ES modules, no build step and no npm dependencies. Pure logic (alias parsing, rule mapping, leak report, mail body) lives in small modules tested with `node --test`. One provider module (`gmail.js`) wraps the Gmail REST API behind an interface the later Outlook provider will also implement. `app.js` is the only file that touches the DOM.

**Tech Stack:** HTML/CSS/vanilla JS (ES modules), Google Identity Services token client, Gmail REST API v1, Node 24 built-in test runner, `python -m http.server` to serve.

**Spec:** `docs/design.md`. Outlook (design order step 5) gets its own plan after this one works.

## Global Constraints

- Root: `C:\Users\FAyaz\Desktop\MailClient`. Own git repo (`git init` in Task 1).
- No npm dependencies. No build step. `package.json` holds only `"type": "module"` and the test script.
- Served at exactly `http://localhost:8080` (the OAuth origin is registered for it).
- Access tokens live in a module variable only. Never `localStorage`, never disk.
- OAuth scopes: `https://www.googleapis.com/auth/gmail.modify` and `https://www.googleapis.com/auth/gmail.settings.basic`. Nothing wider.
- "Delete" always means Trash. Permanent delete does not exist in the Gmail part.
- Rule model uses Microsoft Graph `messageRule` property names, so Outlook rules pass through unchanged later.
- Every rule is a real Gmail filter. No rule may depend on the app being open.
- Rules with a forward action ask for confirmation before saving.
- No compose/send. No combined inbox. No mobile layout.
- Verify every external API name against current docs before using it (Gmail API reference, Google Identity Services, Graph `messageRulePredicates` / `messageRuleActions`). This plan's lists were written from memory.

## Review Focus

1. **Token expires after about 1 hour mid-session.** Any call returning 401 must surface a "sign in again" prompt, not a blank list. Test in Task 5.
2. **Address header shapes.** Display names, several recipients, upper case, dots in the Gmail local part, `+` in a stranger's address, empty tag. Only the owner's own address counts as an alias. Tests in Task 1.
3. **Rule values containing quotes or Gmail search operators** (`a" OR from:x`). Must stay one literal term, never widen the rule. Test in Task 2.
4. **Large mailboxes and rate limits.** More than 1,000 ids in one bulk action, and HTTP 429. Must chunk and retry, not drop mail silently. Tests in Task 5.
5. **Hostile HTML mail.** Scripts must not run and remote images must not load until asked. Test in Task 4.

---

## File Structure

| File | Responsibility |
|---|---|
| `Start.bat` | Serve the folder on `localhost:8080`, open the browser |
| `SETUP.md` | One-time Google Cloud steps for the owner |
| `config.js` | `export const GOOGLE_CLIENT_ID = ''` (owner pastes it; not a secret) |
| `index.html`, `style.css` | Three-pane shell |
| `src/alias.js` | Alias extraction from address headers |
| `src/rules.js` | Rule catalog, Gmail support table, rule to Gmail filter mapping |
| `src/leaks.js` | Leak report from scan records |
| `src/mailbody.js` | Pick body part, build sandboxed frame document |
| `src/gmail.js` | Gmail provider (all network calls) |
| `src/auth.js` | Google sign-in, in-memory token |
| `src/app.js` | DOM, views, event wiring |
| `test/*.test.js` | One test file per `src` logic module |

---

### Task 1: Scaffold and alias parsing

**Files:**
- Create: `package.json`, `.gitignore`, `Start.bat`, `src/alias.js`
- Test: `test/alias.test.js`

**Interfaces:**
- Produces: `extractAlias(headerValue: string, ownAddress: string): string | null`, `aliasAddress(ownAddress: string, alias: string): string`, `senderDomain(fromHeader: string): string | null`

- [ ] **Step 1: `git init`; write `package.json`** with `"type": "module"`, `"private": true`, `"scripts": {"test": "node --test"}`. `.gitignore`: `node_modules/`.

- [ ] **Step 2: Write `Start.bat`**

```bat
@echo off
cd /d "%~dp0"
start "" http://localhost:8080
python -m http.server 8080 --bind 127.0.0.1
```

- [ ] **Step 3: Write the failing tests** in `test/alias.test.js`

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { extractAlias, aliasAddress, senderDomain } from '../src/alias.js';

const me = 'me@gmail.com';
test('plain', () => assert.equal(extractAlias('me+spotify@gmail.com', me), 'spotify'));
test('display name and case', () => assert.equal(extractAlias('"X" <Me+Spotify@Gmail.com>', me), 'spotify'));
test('several recipients', () => assert.equal(extractAlias('a@x.com, me+uber@gmail.com', me), 'uber'));
test('gmail dots ignored, tag kept', () => assert.equal(extractAlias('m.e+net-flix.1@gmail.com', me), 'net-flix.1'));
test('stranger plus address', () => assert.equal(extractAlias('other+tag@gmail.com', me), null));
test('no tag', () => assert.equal(extractAlias('me@gmail.com', me), null));
test('empty tag', () => assert.equal(extractAlias('me+@gmail.com', me), null));
test('empty header', () => assert.equal(extractAlias('', me), null));
test('aliasAddress', () => assert.equal(aliasAddress(me, 'spotify'), 'me+spotify@gmail.com'));
test('senderDomain', () => assert.equal(senderDomain('Spotify <no-reply@email.spotify.com>'), 'spotify.com'));
test('senderDomain none', () => assert.equal(senderDomain('garbage'), null));
```

- [ ] **Step 4: Run** `npm test`. Expected: FAIL, cannot find `../src/alias.js`.

- [ ] **Step 5: Implement `src/alias.js`.** Local-part dots are ignored only for `gmail.com` and `googlemail.com`. `senderDomain` returns the last two labels and carries this comment: `// shortcut: last two labels, wrong for co.uk style domains; switch to a public-suffix list if that bites`.

- [ ] **Step 6: Run** `npm test`. Expected: 11 pass.

- [ ] **Step 7: Commit** `feat: scaffold and alias parsing`.

---

### Task 2: Rule model and Gmail mapping

**Files:**
- Create: `src/rules.js`
- Test: `test/rules.test.js`

**Interfaces:**
- Produces:
  - `Rule = { id?: string, conditions: object, exceptions: object, actions: object }`, keys are Graph `messageRulePredicates` / `messageRuleActions` names; address predicates hold plain strings.
  - `RULE_CATALOG = { predicates: [{key, label, type}], actions: [{key, label, type}] }`, `type` one of `'strings' | 'bool' | 'enum' | 'sizeRange' | 'folder' | 'addresses'` (enums add `options`).
  - `GMAIL_SUPPORT = { predicates: Set<string>, actions: Set<string> }`
  - `class UnsupportedOnGmail extends Error` with `.key`
  - `predicatesToQuery(predicates: object): string`
  - `toGmailFilter(rule: Rule): { criteria: {query?, negatedQuery?}, action: {addLabelIds?, removeLabelIds?, forward?} }`
  - `describeGmailFilter(filter: object, labelNames: Map<string,string>): string`

- [ ] **Step 1: Write the failing tests**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { predicatesToQuery, toGmailFilter, UnsupportedOnGmail, RULE_CATALOG, GMAIL_SUPPORT } from '../src/rules.js';

test('to address', () => assert.equal(predicatesToQuery({ sentToAddresses: ['me+spotify@gmail.com'] }), 'to:("me+spotify@gmail.com")'));
test('or within, and across', () => assert.equal(
  predicatesToQuery({ fromAddresses: ['a@x.com', 'b@y.com'], subjectContains: ['bill'] }),
  'from:("a@x.com" OR "b@y.com") subject:("bill")'));
test('body words', () => assert.equal(predicatesToQuery({ bodyContains: ['x y'] }), '("x y")'));
test('attachment', () => {
  assert.equal(predicatesToQuery({ hasAttachments: true }), 'has:attachment');
  assert.equal(predicatesToQuery({ hasAttachments: false }), '-has:attachment');
});
test('size is KB in, bytes out', () => assert.equal(
  predicatesToQuery({ withinSizeRange: { minimumSize: 10, maximumSize: 20 } }), 'larger:10240 smaller:20480'));
test('quotes cannot widen a rule', () => assert.equal(
  predicatesToQuery({ subjectContains: ['a" OR from:x'] }), 'subject:("a OR from:x")'));
test('empty values are skipped', () => assert.equal(predicatesToQuery({ subjectContains: ['', '  '] }), ''));
test('unsupported predicate', () => assert.throws(
  () => predicatesToQuery({ importance: 'high' }), e => e instanceof UnsupportedOnGmail && e.key === 'importance'));

test('leak rule', () => assert.deepEqual(
  toGmailFilter({ conditions: { sentToAddresses: ['me+spotify@gmail.com'] }, exceptions: { senderContains: ['spotify.com'] }, actions: { delete: true } }),
  { criteria: { query: 'to:("me+spotify@gmail.com")', negatedQuery: 'from:("spotify.com")' }, action: { addLabelIds: ['TRASH'] } }));
test('move and mark read', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { moveToFolder: 'Label_1', markAsRead: true } }).action,
  { addLabelIds: ['Label_1'], removeLabelIds: ['INBOX', 'UNREAD'] }));
test('copy keeps inbox', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { copyToFolder: 'Label_1' } }).action,
  { addLabelIds: ['Label_1'] }));
test('importance', () => {
  const f = v => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { markImportance: v } }).action;
  assert.deepEqual(f('high'), { addLabelIds: ['IMPORTANT'] });
  assert.deepEqual(f('low'), { removeLabelIds: ['IMPORTANT'] });
});
test('gmail extras', () => assert.deepEqual(
  toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { gmailStar: true, gmailNeverSpam: true } }).action,
  { addLabelIds: ['STARRED'], removeLabelIds: ['SPAM'] }));
test('forward one address only', () => {
  const f = a => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { forwardTo: a } });
  assert.equal(f(['a@b.c']).action.forward, 'a@b.c');
  assert.throws(() => f(['a@b.c', 'd@e.f']), UnsupportedOnGmail);
});
test('permanent delete refused', () => assert.throws(
  () => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: { permanentDelete: true } }), UnsupportedOnGmail));
test('rule without conditions refused', () => assert.throws(
  () => toGmailFilter({ conditions: {}, exceptions: {}, actions: { delete: true } }), /condition/));
test('rule without actions refused', () => assert.throws(
  () => toGmailFilter({ conditions: { subjectContains: ['x'] }, exceptions: {}, actions: {} }), /action/));
test('catalog covers every supported key', () => {
  const p = new Set(RULE_CATALOG.predicates.map(x => x.key)), a = new Set(RULE_CATALOG.actions.map(x => x.key));
  for (const k of GMAIL_SUPPORT.predicates) assert.ok(p.has(k), k);
  for (const k of GMAIL_SUPPORT.actions) assert.ok(a.has(k), k);
});
```

- [ ] **Step 2: Run** `npm test`. Expected: FAIL, module not found.

- [ ] **Step 3: Implement `src/rules.js`.**
  - `RULE_CATALOG`: every property of Graph `messageRulePredicates` and `messageRuleActions`, read from the current Microsoft docs, plus the two Gmail-only actions `gmailStar` and `gmailNeverSpam`.
  - `GMAIL_SUPPORT.predicates`: `fromAddresses`, `senderContains` (both `from:`), `sentToAddresses`, `recipientContains` (both `to:`), `subjectContains` (`subject:`), `bodyContains`, `bodyOrSubjectContains` (bare terms), `hasAttachments`, `withinSizeRange`.
  - `GMAIL_SUPPORT.actions`: `moveToFolder`, `copyToFolder`, `markAsRead`, `delete`, `markImportance`, `forwardTo`, `gmailStar`, `gmailNeverSpam`.
  - Quoting: strip `"` from the value, wrap in `"`. Predicate order in the output follows `GMAIL_SUPPORT.predicates` order above.
  - Everything goes into `criteria.query` / `criteria.negatedQuery`, never the dedicated `from`/`to`/`subject` fields: one string then serves both the filter and "apply to existing".
  - Omit empty `addLabelIds` / `removeLabelIds` / `negatedQuery` keys.

- [ ] **Step 4: Run** `npm test`. Expected: all pass.

- [ ] **Step 5: Commit** `feat: rule model and gmail filter mapping`.

---

### Task 3: Leak report

**Files:**
- Create: `src/leaks.js`
- Test: `test/leaks.test.js`

**Interfaces:**
- Consumes: scan records `{ id: string, alias: string, domain: string | null }`
- Produces: `leakReport(records, trusted: {[alias]: string[]}): [{ alias, total, home, domains: [{domain, count, trusted}], flagged }]`

Rules the tests pin: `home` is the domain with the most mail for that alias (ties: alphabetical first). `home` is always trusted. `flagged` is true when any other domain is not in `trusted[alias]`. Output sorted flagged first, then by `total` descending. Records with `domain: null` count under `'(unknown)'`.

- [ ] **Step 1: Write the failing tests**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { leakReport } from '../src/leaks.js';

const r = (alias, domain, n) => Array.from({ length: n }, (_, i) => ({ id: `${alias}${domain}${i}`, alias, domain }));
const recs = [...r('spotify', 'spotify.com', 5), ...r('spotify', 'randomshop.com', 2), ...r('amazon', 'amazon.de', 9)];

test('flags the foreign sender', () => {
  const [first, second] = leakReport(recs, {});
  assert.equal(first.alias, 'spotify');
  assert.equal(first.home, 'spotify.com');
  assert.equal(first.flagged, true);
  assert.deepEqual(first.domains, [
    { domain: 'spotify.com', count: 5, trusted: true },
    { domain: 'randomshop.com', count: 2, trusted: false }]);
  assert.equal(second.alias, 'amazon');
  assert.equal(second.flagged, false);
});
test('trusting clears the flag', () =>
  assert.equal(leakReport(recs, { spotify: ['randomshop.com'] }).find(x => x.alias === 'spotify').flagged, false));
test('unknown domain', () =>
  assert.equal(leakReport([{ id: '1', alias: 'a', domain: null }], {})[0].home, '(unknown)'));
test('empty', () => assert.deepEqual(leakReport([], {}), []));
```

- [ ] **Step 2: Run**, expect FAIL. **Step 3: Implement.** **Step 4: Run**, expect pass. **Step 5: Commit** `feat: leak report`.

---

### Task 4: Mail body and safe frame

**Files:**
- Create: `src/mailbody.js`
- Test: `test/mailbody.test.js`

**Interfaces:**
- Produces: `pickBody(payload): { html: string | null, text: string | null }` (Gmail `payload` tree, base64url data, nested multipart), `buildFrameDoc(body: {html, text}, allowImages: boolean): string`, `FRAME_SANDBOX = ''`

- [ ] **Step 1: Write the failing tests**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { pickBody, buildFrameDoc, FRAME_SANDBOX } from '../src/mailbody.js';

const b64 = s => Buffer.from(s, 'utf8').toString('base64url');
const part = (mimeType, s) => ({ mimeType, body: { data: b64(s) } });

test('nested multipart prefers html', () => {
  const payload = { mimeType: 'multipart/mixed', parts: [
    { mimeType: 'multipart/alternative', parts: [part('text/plain', 'hi'), part('text/html', '<b>hi</b>')] },
    { mimeType: 'application/pdf', filename: 'a.pdf', body: { attachmentId: 'x' } }] };
  assert.deepEqual(pickBody(payload), { html: '<b>hi</b>', text: 'hi' });
});
test('single part plain, utf-8', () => assert.deepEqual(pickBody(part('text/plain', 'grüße')), { html: null, text: 'grüße' }));
test('no body', () => assert.deepEqual(pickBody({ mimeType: 'multipart/mixed', parts: [] }), { html: null, text: null }));
test('sandbox allows nothing', () => assert.equal(FRAME_SANDBOX, ''));
test('images blocked by default', () => {
  const doc = buildFrameDoc({ html: '<img src="https://t.example/p.gif">', text: null }, false);
  assert.match(doc, /Content-Security-Policy/);
  assert.match(doc, /default-src 'none'/);
  assert.doesNotMatch(doc, /img-src[^;"]*https:/);
});
test('images on request', () => assert.match(buildFrameDoc({ html: '<p>x</p>', text: null }, true), /img-src[^;"]*https:/));
test('plain text is escaped', () => {
  const doc = buildFrameDoc({ html: null, text: '<script>alert(1)</script>' }, false);
  assert.doesNotMatch(doc, /<script>alert/);
  assert.match(doc, /&lt;script&gt;/);
});
test('empty body', () => assert.match(buildFrameDoc({ html: null, text: null }, false), /no content/i));
```

- [ ] **Step 2: Run**, expect FAIL.

- [ ] **Step 3: Implement.** CSP meta tag is the first element in the frame document: `default-src 'none'; style-src 'unsafe-inline'; img-src data:` plus ` https:` on `img-src` when `allowImages`. Links get `<base target="_blank">`. Decode with `atob` + `TextDecoder` (works in browser and Node 24; no `Buffer` in `src`).

- [ ] **Step 4: Run**, expect pass. **Step 5: Commit** `feat: mail body and sandboxed frame`.

---

### Task 5: Gmail provider

**Files:**
- Create: `src/gmail.js`
- Test: `test/gmail.test.js`

**Interfaces:**
- Consumes: `extractAlias`, `aliasAddress`, `senderDomain` (Task 1); `toGmailFilter`, `predicatesToQuery` (Task 2); `pickBody` (Task 4)
- Produces: `class AuthExpired extends Error`, and `createGmail({ getToken: () => string, fetch?, sleep? })` returning the provider. This is the interface the Outlook provider will copy:

```
ownAddress(): Promise<string>
listMessages({ folderId?, query?, pageToken? }): Promise<{ messages: Msg[], nextPageToken: string | null }>
   Msg = { id, from, to, subject, date: number, snippet, unread: boolean }
getMessage(id): Promise<Msg & { html, text }>
trash(ids: string[]): Promise<void>
markRead(ids: string[], read: boolean): Promise<void>
moveToFolder(ids: string[], folderId: string): Promise<void>      // add label, remove INBOX
listFolders(): Promise<[{ id, name }]>                             // user labels only
createFolder(name: string): Promise<{ id, name }>                  // returns the existing one on a name match
listRules(): Promise<[{ id, summary: string, raw: object }]>
createRule(rule: Rule): Promise<{ id }>
deleteRule(id: string): Promise<void>
applyRuleToExisting(rule: Rule, onProgress?): Promise<number>      // count changed; never forwards
scanAliases(known: Map<string, Rec>, onProgress?): Promise<Rec[]>  // Rec = { id, alias, domain }; fetches only ids not in `known`
```

Pinned values: page size 50 for `listMessages`; `scanAliases` reads the newest 2,000 message ids with at most 10 requests in flight and keeps only records with an alias; `batchModify` chunks of 1,000 ids; on HTTP 429 or 503 retry up to 4 times with `sleep(500 * 2 ** attempt)`; on 401 throw `AuthExpired`; any other non-2xx throws an `Error` carrying the status and Gmail's error message.

- [ ] **Step 1: Write the failing tests** with a fake `fetch` that records calls and replays queued responses, and a `sleep` that records delays:
  - `sends bearer token`: every call carries `Authorization: Bearer T`.
  - `401 throws AuthExpired`.
  - `429 retries then succeeds`: two 429s then 200, result returned, `sleep` called with `500` then `1000`.
  - `429 five times gives up`: throws, exactly 5 fetches.
  - `trash chunks by 1000`: 2,500 ids, exactly 3 `batchModify` posts with 1000/1000/500 ids, each body `addLabelIds: ['TRASH']`.
  - `trash with no ids makes no request`.
  - `moveToFolder body`: `addLabelIds: ['L1'], removeLabelIds: ['INBOX']`.
  - `createFolder returns existing`: label list contains `Subscriptions`, call with `subscriptions` returns it, no POST.
  - `listMessages maps headers`: metadata response maps to `Msg`, `unread` from the `UNREAD` label, `date` from `internalDate` as number.
  - `scanAliases skips known ids`: 3 listed ids, 1 in `known`, exactly 2 `messages.get` calls.
  - `scanAliases reads Delivered-To before To`: message with `Delivered-To: me+uber@gmail.com` and `To: list@x.com` gives alias `uber`.
  - `applyRuleToExisting pages and does not forward`: two list pages, one `batchModify`, returned count equals ids, request body has no `forward`.
  - `createRule posts toGmailFilter output` to `/settings/filters`.

- [ ] **Step 2: Run**, expect FAIL.

- [ ] **Step 3: Implement `src/gmail.js`.** Base URL `https://gmail.googleapis.com/gmail/v1/users/me`. `listMessages` with `folderId` uses `labelIds`; without both `folderId` and `query` it lists `INBOX`. Message metadata via `format=metadata` with `metadataHeaders` `From`, `To`, `Subject`, `Delivered-To`. `applyRuleToExisting` searches with `predicatesToQuery(conditions)` plus each exception clause negated, and applies the filter's `addLabelIds` / `removeLabelIds`. `listRules` builds `summary` with `describeGmailFilter`.

- [ ] **Step 4: Run** `npm test`. Expected: all pass.

- [ ] **Step 5: Commit** `feat: gmail provider`.

---

### Task 6: Sign-in, shell, read and clean up

**Files:**
- Create: `config.js`, `SETUP.md`, `index.html`, `style.css`, `src/auth.js`, `src/app.js`

**Interfaces:**
- Consumes: `createGmail`, `AuthExpired` (Task 5); `buildFrameDoc`, `FRAME_SANDBOX` (Task 4)
- Produces: `signIn(): Promise<void>`, `getToken(): string`, `signOut(): void` in `src/auth.js`; in `src/app.js` a `showView(name: 'mail' | 'rules' | 'leaks')` switch and a `run(fn)` wrapper that later tasks use for every provider call (shows a busy state, shows errors in a status bar, turns `AuthExpired` into the sign-in prompt).

- [ ] **Step 1: Write `SETUP.md`**: numbered click steps, checked against Google's current console: create project, enable Gmail API, configure the consent screen as External with the owner's address as test user, add the two scopes from Global Constraints, create an OAuth client of type "Web application" with authorized JavaScript origin `http://localhost:8080`, paste the client ID into `config.js`.

- [ ] **Step 2: `src/auth.js`** using `google.accounts.oauth2.initTokenClient` (script `https://accounts.google.com/gsi/client`). If `GOOGLE_CLIENT_ID` is empty, the page shows "Follow SETUP.md first" instead of a sign-in button.

- [ ] **Step 3: Shell and mail view.** Left: account address, Inbox, All Mail, Trash, then placeholders for Folders and Aliases, then Rules and Leaks links. Middle: search box (Gmail query syntax), message list with checkbox per row, select-all, toolbar (Trash, Mark read, Mark unread, Move to folder), "Load more". Right: message header, "Show images" button, `<iframe sandbox>` filled through `srcdoc` from `buildFrameDoc`.

- [ ] **Step 4: Cleanup behaviour.** Trashing more than 20 messages at once asks "Move N messages to Trash?" first. After any bulk action the list reloads. Opening an unread message marks it read.

- [ ] **Step 5: Verify by hand** (owner follows `SETUP.md`, runs `Start.bat`): sign in; inbox lists; open a mail with remote images, they stay blocked until "Show images"; trash one mail and find it in Gmail's Trash; search `from:` something. `npm test` still passes.

- [ ] **Step 6: Commit** `feat: sign-in, mail list, read, trash`.

---

### Task 7: Alias sidebar and folders

**Files:**
- Modify: `src/app.js`, `index.html`, `style.css`

**Interfaces:**
- Consumes: `scanAliases`, `listFolders`, `createFolder`, `createRule`, `applyRuleToExisting`, `listMessages` (Task 5); `aliasAddress` (Task 1)
- Produces: `aliasRecords` (module state, `Rec[]`) that Task 9 reads; `refreshSidebar()`.

- [ ] **Step 1: Scan and cache.** After sign-in run `scanAliases` with a progress line. Cache records in `localStorage` under `aliasScan:<ownAddress>` (ids, aliases and sender domains only, no mail content). "Rescan" button.

- [ ] **Step 2: Alias list.** Sidebar shows each alias with its count, most mail first. Click lists `to:("<aliasAddress>")`.

- [ ] **Step 3: Folders.** Sidebar lists folders from `listFolders`. "New folder" asks for a name. Click lists that folder.

- [ ] **Step 4: Put alias in folder.** Each alias row has "Move to folder…". Choosing one creates the rule `{ conditions: { sentToAddresses: [aliasAddress] }, exceptions: {}, actions: { moveToFolder: folderId } }` and then asks "Also move the N existing mails?" which runs `applyRuleToExisting`.

- [ ] **Step 5: Verify by hand:** create folder "Subscriptions", move two aliases in, confirm in Gmail's own settings that two filters exist with "Skip Inbox" and the label, confirm the existing mail left the inbox, send a test mail to one alias and see it arrive only in the folder.

- [ ] **Step 6: Commit** `feat: alias sidebar and folders`.

---

### Task 8: Rule builder

**Files:**
- Modify: `src/app.js`, `index.html`, `style.css`

**Interfaces:**
- Consumes: `RULE_CATALOG`, `GMAIL_SUPPORT`, `toGmailFilter`, `UnsupportedOnGmail` (Task 2); `listRules`, `createRule`, `deleteRule`, `applyRuleToExisting`, `listFolders` (Task 5)

- [ ] **Step 1: Rules view.** List from `listRules` (summary text, Delete button with confirm). "New rule" opens the builder.

- [ ] **Step 2: Builder.** Three sections: Conditions, Exceptions, Actions. Each is a list of rows: a dropdown fed from `RULE_CATALOG`, then an input matching the entry's `type` (`strings`/`addresses`: comma-separated text; `bool`: yes/no; `enum`: select; `sizeRange`: two KB numbers; `folder`: select from `listFolders`). Entries not in `GMAIL_SUPPORT` are shown disabled with the title text "Gmail filters cannot do this. Available on Outlook accounts."

- [ ] **Step 3: Save.** Run `toGmailFilter` first and show its error text in the builder if it throws. A rule with `forwardTo` asks "Every matching mail will be forwarded to <address>. Save?". After saving ask "Apply to N existing mails?".

- [ ] **Step 4: Shortcuts into the builder.** Message toolbar gets "Rule from this mail" (prefills `fromAddresses`, and `sentToAddresses` when the mail came through an alias).

- [ ] **Step 5: Verify by hand:** build the leak rule (to `+alias`, exception sender contains the home domain, action delete), confirm the filter in Gmail's settings reads the same, apply to existing; confirm a disabled entry cannot be picked; try a forward rule to an unverified address and confirm Gmail's error is shown in the builder, not swallowed.

- [ ] **Step 6: Commit** `feat: rule builder`.

---

### Task 9: Leak view

**Files:**
- Modify: `src/app.js`, `index.html`, `style.css`

**Interfaces:**
- Consumes: `leakReport` (Task 3); `aliasRecords` (Task 7); `createRule`, `applyRuleToExisting`, `trash`, `listMessages` (Task 5)

- [ ] **Step 1: View.** One block per alias from `leakReport`, flagged ones first with a marker. Rows: domain, count, and for non-home domains three buttons: "Trust", "Show mail", "Block".

- [ ] **Step 2: Trust.** Stored in `localStorage` under `trusted:<ownAddress>` as `{[alias]: string[]}`; re-renders. Code comment: `// shortcut: trust marks live in this browser only; move to a Gmail label or draft if they must survive a browser reset`.

- [ ] **Step 3: Block.** Creates `{ conditions: { sentToAddresses: [aliasAddress], senderContains: [domain] }, exceptions: {}, actions: { delete: true } }` after the confirm "Trash all mail from <domain> to +<alias>, now and in future?", then runs `applyRuleToExisting`.

- [ ] **Step 4: Verify by hand:** an alias with two sender domains shows flagged; Trust clears it and survives a page reload; Block creates the filter and the existing mail lands in Trash.

- [ ] **Step 5: Commit** `feat: leak view`.
