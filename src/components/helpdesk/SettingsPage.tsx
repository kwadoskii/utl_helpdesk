"use client";

import { apiRequest } from "@/api";
import { useMsal } from "@azure/msal-react";
import type { AccountInfo, PublicClientApplication } from "@azure/msal-browser";
import { toast } from "sonner";

import {
  Mail,
  Moon,
  Sun,
  Monitor,
  Bell,
  Settings2,
  Save,
  Send,
  Loader2,
  CheckCircle2,
  SunMoon,
} from "lucide-react";

import { SyntheticEvent, useCallback, useEffect, useState } from "react";

type Tab = "smtp" | "appearance" | "notifications" | "general";
type Theme = "light" | "dark" | "system";
type Encryption = "none" | "ssl" | "starttls";
type Priority = "Low" | "Medium" | "High" | "Urgent";

interface SmtpSettings {
  host: string;
  port: number;
  encryption: Encryption;
  username: string;
  password: string;
  senderName: string;
  senderEmail: string;
  configured: boolean;
  hasPassword: boolean;
}

interface NotificationSettings {
  ticketCreated: boolean;
  newMessage: boolean;
  internalNote: boolean;
  ticketClosed: boolean;
}

interface GeneralSettings {
  helpdeskName: string;
  defaultPriority: Priority;
  maxFileSizeMB: number;
  allowedFileTypes: string[];
}

const defaultSmtp: SmtpSettings = {
  host: "",
  port: 587,
  encryption: "starttls",
  username: "",
  password: "",
  senderName: "IT Helpdesk",
  senderEmail: "",
  configured: false,
  hasPassword: false,
};

const defaultNotifications: NotificationSettings = {
  ticketCreated: true,
  newMessage: true,
  internalNote: true,
  ticketClosed: true,
};

const defaultGeneral: GeneralSettings = {
  helpdeskName: "",
  defaultPriority: "Medium",
  maxFileSizeMB: 10,
  allowedFileTypes: [],
};

const tabs = [
  { id: "smtp", label: "SMTP", icon: Mail },
  { id: "appearance", label: "Appearance", icon: Sun },
  { id: "notifications", label: "Notifications", icon: Bell },
  { id: "general", label: "General", icon: Settings2 },
] as const;

export default function SettingsPage() {
  const { instance: msalInstance, accounts } = useMsal();

  const instance = msalInstance as PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo | null;

  const [activeTab, setActiveTab] = useState<Tab>("smtp");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [testing, setTesting] = useState(false);

  const [theme, setTheme] = useState<Theme>("system");

  const [smtp, setSmtp] = useState<SmtpSettings>(defaultSmtp);
  const [testRecipient, setTestRecipient] = useState("");

  const [notificationsSettings, setNotificationsSettings] =
    useState<NotificationSettings>(defaultNotifications);

  const [general, setGeneral] = useState<GeneralSettings>(defaultGeneral);
  const [allowedFileTypesInput, setAllowedFileTypesInput] = useState("");

  // Load all settings from the backend
  const loadSettings = useCallback(async () => {
    if (!account) {
      setLoading(false);
      return;
    }

    setLoading(true);

    try {
      const response = await apiRequest<any>(instance, account, "/settings");

      // Supports either direct response or { settings: ... }
      const data = response.settings || response;

      setSmtp({
        ...defaultSmtp,
        ...(data.smtp || {}),
        password: "",
      });

      setNotificationsSettings({
        ...defaultNotifications,
        ...(data.notifications || {}),
      });

      setGeneral({
        ...defaultGeneral,
        ...(data.general || {}),
      });
      setAllowedFileTypesInput(data.general.allowedFileTypes.join(", "));
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Could not load settings.");
    } finally {
      setLoading(false);
    }
  }, [account, instance]);

  useEffect(() => {
    void loadSettings();
  }, [loadSettings]);

  // Load saved theme
  useEffect(() => {
    try {
      const saved = localStorage.getItem("helpdesk-theme") as Theme | null;

      if (saved === "light" || saved === "dark" || saved === "system") {
        setTheme(saved);
      }
    } catch {
      // Ignore localStorage errors
    }
  }, []);

  // Apply theme and follow OS preference in system mode
  useEffect(() => {
    const media = window.matchMedia("(prefers-color-scheme: dark)");

    const applyTheme = () => {
      const dark = theme === "dark" || (theme === "system" && media.matches);

      document.documentElement.classList.toggle("dark", dark);
    };

    applyTheme();

    if (theme === "system") {
      media.addEventListener("change", applyTheme);
    }

    return () => {
      media.removeEventListener("change", applyTheme);
    };
  }, [theme]);

  const changeTheme = (value: Theme) => {
    setTheme(value);
    try {
      localStorage.setItem("helpdesk-theme", value);
    } catch {
      toast.error("Could not save appearance preference.");
    }
  };

  // Save SMTP configuration
  async function saveSmtp(e: SyntheticEvent<HTMLFormElement>) {
    e.preventDefault();

    if (!account) {
      toast.error("Your Microsoft session is not available.");
      return;
    }

    setSaving(true);

    try {
      const payload: Record<string, unknown> = {
        host: smtp.host.trim(),
        port: Number(smtp.port),
        encryption: smtp.encryption,
        username: smtp.username.trim(),
        senderName: smtp.senderName.trim(),
        senderEmail: smtp.senderEmail.trim(),
      };

      // Only send password when a new one has been entered
      if (smtp.password.trim()) {
        payload.password = smtp.password;
      }

      await apiRequest(instance, account, "/settings/smtp", {
        method: "PUT",
        body: JSON.stringify(payload),
      });

      setSmtp((old) => ({
        ...old,
        password: "",
        configured: true,
        hasPassword: old.hasPassword || Boolean(payload.password),
      }));

      toast.success("SMTP settings saved successfully.");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Could not save SMTP settings.");
    } finally {
      setSaving(false);
    }
  }

  // Send SMTP test email
  async function testSmtp() {
    if (!account) {
      toast.error("Your Microsoft session is not available.");
      return;
    }

    if (!testRecipient.trim()) {
      toast.error("Enter a recipient email address.");
      return;
    }

    setTesting(true);

    try {
      await apiRequest(instance, account, "/settings/smtp/test", {
        method: "POST",
        body: JSON.stringify({
          recipient: testRecipient.trim(),
        }),
      });

      toast.success("Test email sent successfully.");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Test email could not be sent.");
    } finally {
      setTesting(false);
    }
  }

  // Save notification preferences
  async function saveNotifications() {
    if (!account) return;

    setSaving(true);

    try {
      await apiRequest(instance, account, "/settings/notifications", {
        method: "PUT",
        body: JSON.stringify(notificationsSettings),
      });

      toast.success("Notification preferences saved.");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Could not save notification settings.");
    } finally {
      setSaving(false);
    }
  }

  // Save general settings
  async function saveGeneral(e: SyntheticEvent<HTMLFormElement>) {
    e.preventDefault();

    if (!account) return;

    setSaving(true);

    try {
      const payload = {
        ...general,
        helpdeskName: general.helpdeskName.trim(),
        allowedFileTypes: general.allowedFileTypes
          .map((type) => type.trim().toLowerCase().replace(/^\./, ""))
          .filter(Boolean),
      };

      await apiRequest(instance, account, "/settings/general", {
        method: "PUT",
        body: JSON.stringify(payload),
      });

      setGeneral(payload);

      toast.success("General settings saved.");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Could not save general settings.");
    } finally {
      setSaving(false);
    }
  }

  const apiUrl = process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api";

  if (loading) {
    return (
      <div className="page">
        <div className="page-heading">
          <div>
            {/* <div className="eyebrow">PREFERENCES</div>
            <h1>Settings</h1> */}
            {/* <p>Loading your helpdesk configuration...</p> */}
          </div>
        </div>

        <section className="panel settings-loading">
          <div className="flex gap-0.5 flex-col justify-center text-center">
            <Loader2 className="animate-spin self-center" size={24} />
            <p>Loading your helpdesk configuration...</p>
            <span>Loading settings...</span>
          </div>
        </section>
      </div>
    );
  }

  if (!account) {
    return (
      <div className="page">
        <section className="panel">
          <h2>Authentication required</h2>
          <p>Please sign in with your Microsoft account to access settings.</p>
        </section>
      </div>
    );
  }

  return (
    <div className="page">
      {/* <div className="page-heading">
        <div>
          <h1>Settings</h1>
          <p>Manage your helpdesk configuration and preferences.</p>
        </div>
      </div> */}

      <section className="panel md:p-5 md:px-10">
        {/* Tab Navigation */}
        <div className="settings-tabs" role="tablist">
          {tabs.map((tab) => {
            const Icon = tab.icon;
            const selected = activeTab === tab.id;

            return (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={selected}
                className={`settings-tab ${selected ? "active" : ""}`}
                onClick={() => setActiveTab(tab.id)}
              >
                <Icon size={17} />
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>

        <div className="settings-tab-content">
          {/* SMTP SETTINGS */}
          {activeTab === "smtp" && (
            <div>
              <div className="settings-section-heading">
                <div>
                  <h2>SMTP Configuration</h2>
                  <p>Configure the outgoing email server used for ticket notifications.</p>
                </div>

                <span className={`settings-status ${smtp.configured ? "success" : "warning"}`}>
                  {smtp.configured ? "Configured" : "Not configured"}
                </span>
              </div>

              <form onSubmit={saveSmtp} className="settings-form">
                <div className="settings-form-grid">
                  <label>
                    SMTP Host
                    <input
                      required
                      type="text"
                      placeholder="smtp.company.com"
                      value={smtp.host}
                      onChange={(e) => setSmtp({ ...smtp, host: e.target.value })}
                    />
                  </label>

                  <label>
                    SMTP Port
                    <input
                      required
                      type="number"
                      min={1}
                      max={65535}
                      value={smtp.port}
                      onChange={(e) => setSmtp({ ...smtp, port: Number(e.target.value) })}
                    />
                  </label>

                  <label>
                    Encryption
                    <select
                      value={smtp.encryption}
                      onChange={(e) =>
                        setSmtp({
                          ...smtp,
                          encryption: e.target.value as Encryption,
                        })
                      }
                    >
                      <option value="starttls">STARTTLS</option>
                      <option value="ssl">SSL / TLS</option>
                      <option value="none">None</option>
                    </select>
                  </label>

                  <label>
                    SMTP Username
                    <input
                      required
                      type="text"
                      autoComplete="username"
                      placeholder="notifications@company.com"
                      value={smtp.username}
                      onChange={(e) => setSmtp({ ...smtp, username: e.target.value })}
                    />
                  </label>

                  <label>
                    SMTP Password
                    <input
                      type="password"
                      autoComplete="new-password"
                      placeholder={
                        smtp.hasPassword
                          ? "Leave blank to keep existing password"
                          : "Enter SMTP password"
                      }
                      value={smtp.password}
                      onChange={(e) => setSmtp({ ...smtp, password: e.target.value })}
                    />
                    <small>The saved password is encrypted and never displayed.</small>
                  </label>

                  <label>
                    Sender Name
                    <input
                      required
                      type="text"
                      placeholder="IT Helpdesk"
                      value={smtp.senderName}
                      onChange={(e) => setSmtp({ ...smtp, senderName: e.target.value })}
                    />
                  </label>

                  <label className="full-width">
                    Sender Email
                    <input
                      required
                      type="email"
                      placeholder="helpdesk@company.com"
                      value={smtp.senderEmail}
                      onChange={(e) => setSmtp({ ...smtp, senderEmail: e.target.value })}
                    />
                  </label>
                </div>

                <div className="settings-actions">
                  <button type="submit" className="primary-btn" disabled={saving}>
                    {saving ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                    {saving ? "Saving..." : "Save SMTP Settings"}
                  </button>
                </div>
              </form>

              <div className="settings-divider" />

              <div>
                <h3>Test Email Configuration</h3>
                <p>
                  Send a test email to verify that your SMTP configuration is working correctly.
                </p>

                <div className="settings-test-email">
                  <input
                    type="email"
                    placeholder="Enter recipient email address"
                    value={testRecipient}
                    onChange={(e) => setTestRecipient(e.target.value)}
                  />

                  <button
                    type="button"
                    className="secondary-btn"
                    onClick={testSmtp}
                    disabled={testing}
                  >
                    {testing ? <Loader2 size={16} className="animate-spin" /> : <Send size={16} />}
                    {testing ? "Sending..." : "Send Test Email"}
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Appearance */}
          {activeTab === "appearance" && (
            <div>
              <h2>Appearance</h2>
              <p>
                Choose how the helpdesk looks. System mode follows your computer's appearance
                preference.
              </p>

              <div className="theme-options">
                <button
                  type="button"
                  className={`theme-option ${theme === "light" ? "selected" : ""}`}
                  onClick={() => changeTheme("light")}
                >
                  <Sun size={23} />
                  <strong>Light</strong>
                  <span>Light appearance</span>
                  {theme === "light" && <CheckCircle2 size={16} />}
                </button>

                <button
                  type="button"
                  className={`theme-option ${theme === "dark" ? "selected" : ""}`}
                  onClick={() => changeTheme("dark")}
                >
                  <Moon size={23} />
                  <strong>Dark</strong>
                  <span>Dark appearance</span>
                  {theme === "dark" && <CheckCircle2 size={16} />}
                </button>

                <button
                  type="button"
                  className={`theme-option ${theme === "system" ? "selected" : ""}`}
                  onClick={() => changeTheme("system")}
                >
                  <SunMoon size={23} />
                  <strong>System</strong>
                  <span>Follow device settings</span>
                  {theme === "system" && <CheckCircle2 size={16} />}
                </button>
              </div>

              <div className="settings-info">
                Your appearance preference is saved locally in this browser.
              </div>
            </div>
          )}

          {/* NOTIFICATIONS */}
          {activeTab === "notifications" && (
            <div>
              <h2>Notification Settings</h2>
              <p>
                Choose which ticket activities should trigger email notifications. These preferences
                do not disable in-app notifications.
              </p>

              <div className="settings-notification-list">
                {[
                  {
                    key: "ticketCreated",
                    title: "New Ticket Created",
                    description: "Notify MIS Support when a user creates a ticket.",
                  },
                  {
                    key: "newMessage",
                    title: "New Ticket Message",
                    description: "Notify ticket participants when a new reply is posted.",
                  },
                  {
                    key: "internalNote",
                    title: "Internal Note",
                    description: "Notify MIS Support when an internal note is added.",
                  },
                  {
                    key: "ticketClosed",
                    title: "Ticket Closed",
                    description: "Notify ticket participants when a ticket is closed.",
                  },
                ].map((item) => {
                  const key = item.key as keyof NotificationSettings;

                  return (
                    <label key={item.key} className="settings-notification-item">
                      <div>
                        <strong>{item.title}</strong>
                        <span>{item.description}</span>
                      </div>

                      <input
                        type="checkbox"
                        checked={notificationsSettings[key]}
                        onChange={(e) =>
                          setNotificationsSettings((old) => ({
                            ...old,
                            [key]: e.target.checked,
                          }))
                        }
                      />
                    </label>
                  );
                })}
              </div>

              <div className="settings-actions">
                <button
                  type="button"
                  className="primary-btn"
                  onClick={saveNotifications}
                  disabled={saving}
                >
                  {saving ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                  {saving ? "Saving..." : "Save Notification Settings"}
                </button>
              </div>
            </div>
          )}

          {/* GENERAL */}
          {activeTab === "general" && (
            <div>
              <h2>General Settings</h2>
              <p>Configure general helpdesk preferences and attachment restrictions.</p>

              <form onSubmit={saveGeneral} className="settings-form">
                <div className="settings-form-grid">
                  <label className="full-width">
                    Helpdesk Name
                    <input
                      required
                      maxLength={100}
                      value={general.helpdeskName}
                      onChange={(e) =>
                        setGeneral({
                          ...general,
                          helpdeskName: e.target.value,
                        })
                      }
                    />
                  </label>

                  <label>
                    Default Ticket Priority
                    <select
                      value={general.defaultPriority}
                      onChange={(e) =>
                        setGeneral({
                          ...general,
                          defaultPriority: e.target.value as Priority,
                        })
                      }
                    >
                      <option value="Low">Low</option>
                      <option value="Medium">Medium</option>
                      <option value="High">High</option>
                      <option value="Urgent">Urgent</option>
                    </select>
                  </label>

                  <label>
                    Maximum Attachment Size (MB)
                    <input
                      type="number"
                      min={1}
                      max={10}
                      required
                      value={general.maxFileSizeMB}
                      onChange={(e) =>
                        setGeneral({
                          ...general,
                          maxFileSizeMB: Number(e.target.value),
                        })
                      }
                    />
                  </label>

                  <label className="full-width">
                    Allowed File Types
                    <input
                      type="text"
                      value={allowedFileTypesInput}
                      onChange={(e) => setAllowedFileTypesInput(e.target.value)}
                      onBlur={() => {
                        setGeneral({
                          ...general,
                          allowedFileTypes: allowedFileTypesInput
                            .split(",")
                            .map((type) => type.trim().toLowerCase())
                            .filter(Boolean),
                        });
                        // setAllowedFileTypesInput(general.allowedFileTypes.join(", "));
                      }}
                      placeholder="jpg, jpeg, png, pdf, doc, docx, txt"
                    />
                    <small>
                      Separate file extensions with commas. Do not include wildcard characters.
                    </small>
                  </label>
                </div>

                <div className="settings-actions">
                  <button type="submit" className="primary-btn" disabled={saving}>
                    {saving ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                    {saving ? "Saving..." : "Save General Settings"}
                  </button>
                </div>
              </form>
            </div>
          )}
        </div>

        {/* API Connection */}
        <div className="settings-api">
          <h2>API Connection</h2>
          <p>
            API endpoint: <code>{apiUrl}</code>
          </p>
        </div>
      </section>
    </div>
  );
}
