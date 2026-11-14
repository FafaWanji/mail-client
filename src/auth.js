import { GOOGLE_CLIENT_ID } from '../config.js';

export const CLIENT_ID = GOOGLE_CLIENT_ID.trim(); // a pasted ID often carries a stray space

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',         // read, label, trash
  'https://www.googleapis.com/auth/gmail.settings.basic', // filters
];

// The token lives here only. It is never written to storage, so closing the tab signs out.
let token = '';

export const getToken = () => token;

export function signIn() {
  return new Promise((resolve, reject) => {
    const oauth = window.google?.accounts?.oauth2;
    if (!oauth) return reject(new Error('Google sign-in did not load. Check the internet connection and reload.'));
    oauth.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES.join(' '),
      callback: response => {
        if (response.error) return reject(new Error(response.error_description || response.error));
        if (!oauth.hasGrantedAllScopes(response, ...SCOPES)) {
          return reject(new Error('Both permissions are needed (mail and filters). Sign in again and tick both.'));
        }
        token = response.access_token;
        resolve();
      },
      error_callback: error => reject(new Error(error?.message || 'Sign-in was closed.')),
    }).requestAccessToken();
  });
}
