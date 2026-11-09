import { AccountInfo } from "@azure/msal-browser";

export const isSupport = (account: AccountInfo | null) =>
  Boolean(
    account?.idTokenClaims &&
    (((account.idTokenClaims as any).roles || []).some((r: string) =>
      ["Support.Agent", "Support.Access"].includes(r),
    ) ||
      ((account.idTokenClaims as any).groups || []).includes(
        process.env.NEXT_PUBLIC_IT_SUPPORT_GROUP_ID || "",
      )),
  );
