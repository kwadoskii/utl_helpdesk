"use client";

import { ReactNode, useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: process.env.NEXT_PUBLIC_ENTRA_CLIENT_ID || "configure-client-id",
    authority: `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_ENTRA_TENANT_ID || "common"}`,
    redirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI || (typeof window !== "undefined" ? window.location.origin : "http://localhost:3000"),
  },
  cache: { cacheLocation: "sessionStorage" },
});

export default function Providers({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false);
  useEffect(() => {
    msalInstance.initialize().then(() => setReady(true)).catch(console.error);
  }, []);
  if (!ready) return <div className="boot-screen">Loading Helpdesk…</div>;
  return <MsalProvider instance={msalInstance}>{children}</MsalProvider>;
}
