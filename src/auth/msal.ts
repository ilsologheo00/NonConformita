import { PublicClientApplication, InteractionRequiredAuthError, AuthError, type AccountInfo } from "@azure/msal-browser";
import { config } from "../config";

// Delegated permissions: the item is created as the signed-in user, so
// "Creato da" in SharePoint shows the real author.
export const graphScopes = ["Sites.ReadWrite.All", "User.ReadBasic.All"];

export const msal = new PublicClientApplication({
  auth: {
    clientId: config.clientId || "00000000-0000-0000-0000-000000000000",
    authority: `https://login.microsoftonline.com/${config.tenantId || "organizations"}`,
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "localStorage" },
});

/** Set when the return from the Microsoft sign-in page failed; shown on the sign-in screen. */
export let signInError = "";

export function describeAuthError(e: unknown): string {
  if (e instanceof AuthError) return `${e.errorCode}: ${e.errorMessage.split("\r\n")[0] || e.message}`;
  return (e as Error)?.message || String(e);
}

/** Processes the response in the URL after coming back from login.microsoftonline.com. */
export async function completeRedirect(): Promise<void> {
  try {
    const result = await msal.handleRedirectPromise();
    if (result?.account) msal.setActiveAccount(result.account);
  } catch (e) {
    signInError = describeAuthError(e);
  }
  if (!msal.getActiveAccount()) {
    const accounts = msal.getAllAccounts();
    if (accounts.length) msal.setActiveAccount(accounts[0]);
  }
}

// Guards against bouncing to Microsoft forever when a token still cannot be
// obtained right after coming back from an interactive request.
const REDIRECT_KEY = "nc-token-redirect";

export async function getToken(account: AccountInfo): Promise<string> {
  try {
    const r = await msal.acquireTokenSilent({ scopes: graphScopes, account });
    sessionStorage.removeItem(REDIRECT_KEY);
    return r.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const last = Number(sessionStorage.getItem(REDIRECT_KEY) || 0);
      if (Date.now() - last > 60_000) {
        sessionStorage.setItem(REDIRECT_KEY, String(Date.now()));
        await msal.acquireTokenRedirect({ scopes: graphScopes, account });
      }
      throw new Error(`Microsoft non ha concesso l'accesso a SharePoint (${describeAuthError(e)}).`);
    }
    throw e;
  }
}
