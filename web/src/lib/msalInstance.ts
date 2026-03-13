import { PublicClientApplication, type AuthenticationResult } from '@azure/msal-browser';

// Replace with your Azure AD app's client ID from https://portal.azure.com
const MSAL_CLIENT_ID = '72f543e0-817c-4939-8925-898b1048762c';

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    //redirectUri: window.location.origin + import.meta.env.BASE_URL,
    redirectUri: window.location.origin, // + import.meta.env.BASE_URL,
  },
  cache: { cacheLocation: 'localStorage' },
});

// Resolves to the AuthenticationResult if we just returned from a redirect, otherwise null
export const msalReady: Promise<AuthenticationResult | null> = msalInstance
  .initialize()
  .then(() => msalInstance.handleRedirectPromise())
  .catch(() => null);
