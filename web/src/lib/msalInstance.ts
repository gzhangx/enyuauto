import { PublicClientApplication } from '@azure/msal-browser';

// Replace with your Azure AD app's client ID from https://portal.azure.com
const MSAL_CLIENT_ID = '72f543e0-817c-4939-8925-898b1048762c';

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: 'localStorage' },
});

// Single shared ready-promise; handleRedirectPromise clears leftover redirect state
export const msalReady = msalInstance
  .initialize()
  .then(() => msalInstance.handleRedirectPromise().catch(() => null));

/** Clear any stale interaction-in-progress lock before opening a popup. */
export function clearMsalInteractionLock() {
  [sessionStorage, localStorage].forEach(store => {
    Object.keys(store)
      .filter(k => k.includes('interaction.status'))
      .forEach(k => store.removeItem(k));
  });
}
