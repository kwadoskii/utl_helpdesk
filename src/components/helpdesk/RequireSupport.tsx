"use client";

import { useEffect, useState, type ReactNode } from "react";
import { useMsal } from "@azure/msal-react";
import { useRouter } from "next/navigation";
import { InteractionStatus } from "@azure/msal-browser";
import { hasSupportAccess } from "@/lib/auth";

export default function RequireSupport({ children }: { children: ReactNode }) {
  const { instance, accounts, inProgress } = useMsal();
  const router = useRouter();
  const account = accounts[0] ?? instance.getActiveAccount();
  const isSupport = hasSupportAccess(account ?? null);

  const [authorized, setAuthorized] = useState(false);

  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;

    if (!account) {
      router.replace("/my-tickets");
      return;
    }

    if (!isSupport) {
      router.replace("/my-tickets");
      return;
    }

    setAuthorized(true);
  }, [account, inProgress, isSupport, router]);

  if (!authorized) {
    return <div className="boot-screen">Checking access…</div>;
  }

  return <>{children}</>;
}
