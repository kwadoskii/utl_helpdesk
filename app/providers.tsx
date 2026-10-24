"use client";

import { ReactNode, useEffect, useState } from "react";
import { PublicClientApplication, EventType, AuthenticationResult } from "@azure/msal-browser";
import { MsalProvider, useMsal } from "@azure/msal-react";

const clientId = process.env.NEXT_PUBLIC_ENTRA_CLIENT_ID || "";
const tenantId = process.env.NEXT_PUBLIC_ENTRA_TENANT_ID || "";
const msalInstance = new PublicClientApplication({
  auth: {
    clientId: clientId || "00000000-0000-0000-0000-000000000000",
    authority: `https://login.microsoftonline.com/${tenantId || "common"}`,
    redirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000",
    navigateToLoginRequestUrl: true,
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
});

function AuthenticationGate({ children }: { children: ReactNode }) {
  const { instance, accounts, inProgress } = useMsal();
  const [redirectStarted, setRedirectStarted] = useState(false);
  const configured = Boolean(
    clientId && tenantId && !clientId.includes("your-") && !tenantId.includes("your-"),
  );

  useEffect(() => {
    if (inProgress !== "none" || accounts.length || redirectStarted || !configured) return;
    setRedirectStarted(true);
    instance.loginRedirect({ scopes: ["User.Read"] }).catch((error) => {
      console.error("Microsoft sign-in failed", error);
      setRedirectStarted(false);
    });
  }, [instance, accounts.length, inProgress, redirectStarted, configured]);

  if (!configured)
    return (
      <div className="auth-screen">
        <div className="auth-card">
          <div className="brand-mark">MIS</div>
          <h1>Configure Microsoft sign-in</h1>
          <p>
            The dashboard is protected. Add your Microsoft Entra application settings to{" "}
            <code>client/.env.local</code>, then restart Next.js.
          </p>
          <pre>
            NEXT_PUBLIC_ENTRA_CLIENT_ID=your-client-id{"\n"}
            NEXT_PUBLIC_ENTRA_TENANT_ID=your-tenant-id{"\n"}
            NEXT_PUBLIC_ENTRA_API_SCOPE=api://your-api-app-id/access_as_user
          </pre>
        </div>
      </div>
    );
  if (!accounts.length || inProgress !== "none")
    return (
      <div className="boot-screen">
        <div className="brand-mark">MIS</div>
        <h2>Connecting to Microsoft</h2>
        <p>Redirecting you to your organisation’s sign-in page…</p>
        <button
          className="primary-btn"
          onClick={() => instance.loginRedirect({ scopes: ["User.Read"] })}
        >
          Continue to sign in
        </button>
      </div>
    );
  return <>{children}</>;
}

export default function Providers({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false);
  useEffect(() => {
    msalInstance
      .initialize()
      .then(() => {
        const result = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
        if (result) msalInstance.setActiveAccount(result);
        msalInstance.addEventCallback((event) => {
          if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
            const account = (event.payload as AuthenticationResult).account;
            if (account) msalInstance.setActiveAccount(account);
          }
        });
        setReady(true);
      })
      .catch((error) => console.error("MSAL initialization failed", error));
  }, []);
  if (!ready) return <div className="boot-screen">Loading secure support portal…</div>;
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticationGate>{children}</AuthenticationGate>
    </MsalProvider>
  );
}
