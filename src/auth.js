import * as config from '../config.js';

// A pasted ID often carries a stray space.
export const GOOGLE_ID = (config.GOOGLE_CLIENT_ID || '').trim();
export const MICROSOFT_ID = (config.MICROSOFT_CLIENT_ID || '').trim();

// Tokens are returned to the caller and kept in memory only. Nothing here writes to storage.

// ---------- Google ----------
const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',         // read, label, trash, send
  'https://www.googleapis.com/auth/gmail.settings.basic', // filters
];

// Full access is the only Gmail scope that can delete for good. It is asked for on the Empty click, never at sign-in.
export const GOOGLE_FULL = 'https://mail.google.com/';

/** Opens Google's sign-in and resolves to an access token. `hint` preselects an address; `extra` adds scopes. */
export function googleSignIn(hint, extra = []) {
  const scopes = [...GOOGLE_SCOPES, ...extra];
  return new Promise((resolve, reject) => {
    const oauth = window.google?.accounts?.oauth2;
    if (!oauth) return reject(new Error('Google sign-in did not load. Check the internet connection and reload.'));
    oauth.initTokenClient({
      client_id: GOOGLE_ID,
      scope: scopes.join(' '),
      prompt: hint ? '' : 'select_account',
      login_hint: hint,
      callback: response => {
        if (response.error) return reject(new Error(response.error_description || response.error));
        if (!oauth.hasGrantedAllScopes(response, ...scopes)) {
          return reject(new Error(extra.length ? 'The extra permission was not given. Nothing was deleted.' : 'Both permissions are needed (mail and filters). Sign in again and tick both.'));
        }
        resolve(response.access_token);
      },
      error_callback: error => reject(new Error(error?.message || 'Sign-in was closed.')),
    }).requestAccessToken();
  });
}

// ---------- Microsoft (authorization code flow with PKCE, in a popup) ----------
// consumers: personal accounts only. It works for apps set to "Personal accounts only", which /common/ refuses.
const MS_LOGIN = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const MS_SCOPES = [
  'https://graph.microsoft.com/Mail.ReadWrite',            // read, move, mark
  'https://graph.microsoft.com/MailboxSettings.ReadWrite', // inbox rules
  'https://graph.microsoft.com/User.Read',                 // own address
  'https://graph.microsoft.com/Mail.Send',                 // reply, compose, unsubscribe by mail
].join(' ');

const base64url = bytes => btoa(String.fromCharCode(...bytes)).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '');
const randomString = () => base64url(crypto.getRandomValues(new Uint8Array(32)));

/** Waits for auth.html in the popup to hand back the redirect's query string. */
function popupResult(popup) {
  return new Promise((resolve, reject) => {
    const finish = () => { removeEventListener('message', onMessage); clearInterval(watch); };
    const onMessage = event => {
      if (event.origin !== location.origin || event.source !== popup || event.data?.type !== 'ms-auth') return;
      finish();
      resolve(new URLSearchParams(event.data.search));
    };
    // The grace period lets a message that was sent just before the popup closed arrive first.
    const watch = setInterval(() => {
      if (popup.closed) setTimeout(() => { finish(); reject(new Error('Sign-in was closed.')); }, 500);
    }, 500);
    addEventListener('message', onMessage);
  });
}

/** Opens Microsoft's sign-in and resolves to an access token for Microsoft Graph. */
export async function microsoftSignIn(hint) {
  const redirectUri = `${location.origin}/auth.html`;
  const verifier = randomString(), state = randomString();
  const popup = window.open('about:blank', 'ms-signin', 'width=520,height=680'); // opened first, while the click still counts
  if (!popup) throw new Error('The browser blocked the sign-in window. Allow popups for this page.');

  const challenge = base64url(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))));
  const query = new URLSearchParams({
    client_id: MICROSOFT_ID, response_type: 'code', response_mode: 'query', redirect_uri: redirectUri, scope: MS_SCOPES,
    state, code_challenge: challenge, code_challenge_method: 'S256', prompt: 'select_account',
  });
  if (hint) query.set('login_hint', hint);
  popup.location = `${MS_LOGIN}/authorize?${query}`;

  const result = await popupResult(popup);
  if (result.get('error')) throw new Error(result.get('error_description') || result.get('error'));
  if (result.get('state') !== state || !result.get('code')) throw new Error('Sign-in answer did not match the request. Try again.');

  const response = await fetch(`${MS_LOGIN}/token`, {
    method: 'POST',
    body: new URLSearchParams({
      client_id: MICROSOFT_ID, grant_type: 'authorization_code', code: result.get('code'),
      redirect_uri: redirectUri, code_verifier: verifier, scope: MS_SCOPES,
    }),
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error_description || data.error || `Microsoft sign-in failed (${response.status}).`);
  return data.access_token; // the refresh token in `data` is dropped on purpose: nothing long-lived is kept
}
