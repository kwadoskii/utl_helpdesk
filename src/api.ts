import { PublicClientApplication, AccountInfo } from "@azure/msal-browser";

export async function apiRequest(msal: PublicClientApplication, account: AccountInfo, path: string, init: RequestInit = {}) {
  const token = await msal.acquireTokenSilent({
    account,
    scopes: [process.env.NEXT_PUBLIC_ENTRA_API_SCOPE || "User.Read"],
  });
  const response = await fetch(`${process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api"}${path}`, {
    ...init,
    headers: { ...(init.headers || {}), Authorization: `Bearer ${token.accessToken}` },
  });
  if (!response.ok) throw new Error((await response.json().catch(() => ({ message: "Request failed" }))).message || "Request failed");
  return response.json();
}
