import {
  BrowserCacheLocation,
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo
} from '@azure/msal-browser';
import { config, sharePointScope } from './config';

export interface IAuthState {
  account: AccountInfo;
  /** The `state` string passed to loginRedirect, returned after the redirect round-trip. */
  state?: string;
}

const scopes: string[] = [sharePointScope];

const msal: PublicClientApplication = new PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    redirectUri: new URL(import.meta.env.BASE_URL, window.location.origin).href
  },
  cache: { cacheLocation: BrowserCacheLocation.LocalStorage }
});

/**
 * Completes a pending redirect, or returns the cached account.
 * If nobody is signed in, starts a redirect login (page navigates away) and resolves to undefined.
 */
export async function ensureSignedIn(state: string, loginHint?: string): Promise<IAuthState | undefined> {
  await msal.initialize();
  const result = await msal.handleRedirectPromise();
  if (result?.account) {
    msal.setActiveAccount(result.account);
    return { account: result.account, state: result.state };
  }

  const accounts: AccountInfo[] = msal.getAllAccounts();
  const account: AccountInfo | undefined =
    (loginHint && accounts.find(a => a.username.toLowerCase() === loginHint.toLowerCase())) ||
    (!loginHint ? accounts[0] : undefined);
  if (account) {
    msal.setActiveAccount(account);
    return { account };
  }

  await msal.loginRedirect({ scopes, state, loginHint, prompt: loginHint ? undefined : 'select_account' });
  return undefined;
}

export async function getSharePointToken(allowRedirect = false): Promise<string> {
  const account: AccountInfo | null = msal.getActiveAccount();
  if (!account) { throw new Error('Not signed in.'); }
  try {
    return (await msal.acquireTokenSilent({ scopes, account })).accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      if (allowRedirect) { await msal.acquireTokenRedirect({ scopes, account }); }
      throw new Error('Your sign-in needs to be refreshed. Uploaded photos stay in SharePoint; switch accounts to sign in again. This discards unsaved page selections.');
    }
    throw e;
  }
}

export async function signOut(): Promise<void> {
  await msal.logoutRedirect({ account: msal.getActiveAccount() ?? undefined });
}
