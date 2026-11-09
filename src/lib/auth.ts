import type { AccountInfo } from "@azure/msal-browser";

export function hasSupportAccess(account: AccountInfo | null): boolean {
  if (!account?.idTokenClaims) return false;

  const claims = account.idTokenClaims as {
    roles?: string[];
    groups?: string[];
  };

  const supportGroupId = process.env.NEXT_PUBLIC_IT_SUPPORT_GROUP_ID;

  return (
    (claims.roles ?? []).some((role) => ["Support.Agent", "Support.Access"].includes(role)) ||
    (!!supportGroupId && (claims.groups ?? []).includes(supportGroupId))
  );
}
