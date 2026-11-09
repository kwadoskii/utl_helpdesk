"use client";

import { Moon, RefreshCw, Sun } from "lucide-react";
import { useEffect, useState } from "react";
import { toggleTheme } from "@/lib/toggleTheme";

export default function SettingsPage() {
  const [dark, setDark] = useState(false);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("helpdesk-theme");
      const enabled = saved === "dark";
      setDark(enabled);
      document.documentElement.classList.toggle("dark", enabled);
    } catch {}
    // setThemeReady(true);
  }, []);

  return (
    <div className="page">
      <div className="page-heading">
        <div>
          <div className="eyebrow">PREFERENCES</div>
          <h1>Settings</h1>
          <p>Manage your local display preferences and connection details.</p>
        </div>
      </div>
      <section className="panel settings-panel">
        <h2>Appearance</h2>
        <p>Dark mode preference is saved in this browser and restored on reload.</p>
        <button className="secondary-btn" onClick={() => void toggleTheme(dark, setDark)}>
          {dark ? <Sun size={16} /> : <Moon size={16} />} Switch to {dark ? "light" : "dark"} mode
        </button>
        <h2>API connection</h2>
        <p>
          API endpoint:{" "}
          <code>{process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api"}</code>
        </p>
        {/* <button className="secondary-btn" onClick={() => void loadTickets()}>
          <RefreshCw size={16} /> Test connection / refresh tickets
        </button> */}
      </section>
    </div>
  );
}
