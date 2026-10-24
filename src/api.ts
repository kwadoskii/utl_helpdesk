import {
  AccountInfo,
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const API_URL = (process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api").replace(/\/$/, "");
const SCOPES = [process.env.NEXT_PUBLIC_ENTRA_API_SCOPE || "User.Read"];

export async function apiRequest<T = any>(
  msal: PublicClientApplication,
  account: AccountInfo,
  path: string,
  init: RequestInit = {},
): Promise<T> {
  let token;
  try {
    token = await msal.acquireTokenSilent({ account, scopes: SCOPES });
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      await msal.acquireTokenRedirect({ account, scopes: SCOPES });
      throw new Error("Redirecting to Microsoft to renew your session.");
    }
    throw error;
  }
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token.accessToken}`);
  if (!(init.body instanceof FormData) && init.body && !headers.has("Content-Type"))
    headers.set("Content-Type", "application/json");
  const response = await fetch(`${API_URL}${path.startsWith("/") ? path : `/${path}`}`, {
    ...init,
    headers,
    cache: "no-store",
  });
  if (!response.ok) {
    const body = await response
      .json()
      .catch(() => ({ message: `Request failed (${response.status})` }));
    throw new Error(body.message || `Request failed (${response.status})`);
  }
  if (response.status === 204) return undefined as T;
  return response.json() as Promise<T>;
}
