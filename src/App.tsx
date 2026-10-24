"use client";

import {
  ChangeEvent,
  FormEvent,
  SyntheticEvent,
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";
import { usePathname, useRouter } from "next/navigation";
import { useMsal } from "@azure/msal-react";
import { AccountInfo } from "@azure/msal-browser";
import { apiRequest } from "./api";
import {
  Bell,
  BookOpen,
  CheckCircle2,
  Clock3,
  FileText,
  Filter,
  Headphones,
  LayoutDashboard,
  LogOut,
  Menu,
  Moon,
  Paperclip,
  Plus,
  Search,
  Send,
  Settings,
  ShieldCheck,
  Sun,
  Ticket,
  Users,
  X,
  CircleHelp,
  RefreshCw,
} from "lucide-react";

type Status = "Open" | "In progress" | "Pending" | "Resolved" | "Closed";
type TicketItem = {
  _id: string;
  ticketNumber: string;
  subject: string;
  requesterName?: string;
  requesterEmail?: string;
  department?: string;
  category?: string;
  priority: "Low" | "Medium" | "High" | "Urgent";
  status: Status;
  updatedAt: string;
  assigneeName?: string;
  messages?: Message[];
  createdAt?: string;
  resolvedAt?: string;
};
type Message = {
  _id?: string;
  authorId: string;
  authorName: string;
  body: string;
  createdAt?: string;
  attachments?: { originalName: string; size: number }[];
  isInternalNote?: boolean;
};
type Article = {
  _id: string;
  title: string;
  slug: string;
  category: string;
  summary: string;
  content?: string;
  tags?: string[];
};
const nav = [
  {
    label: "Workspace",
    items: [
      { name: "Dashboard", path: "/", icon: LayoutDashboard },
      { name: "All tickets", path: "/all-tickets", icon: Ticket },
      { name: "My tickets", path: "/my-tickets", icon: Headphones },
      { name: "Knowledge base", path: "/knowledge-base", icon: BookOpen },
    ],
  },
  {
    label: "Manage",
    items: [
      { name: "Team queue", path: "/team-queue", icon: Users },
      { name: "Reports", path: "/reports", icon: FileText },
      { name: "Settings", path: "/settings", icon: Settings },
    ],
  },
];
const routeNames: Record<string, string> = {
  "/": "Dashboard",
  "/all-tickets": "All tickets",
  "/my-tickets": "My tickets",
  "/knowledge-base": "Knowledge base",
  "/team-queue": "Team queue",
  "/reports": "Reports",
  "/settings": "Settings",
};
const API_TICKET_FIELDS = "";

export default function App() {
  const { instance, accounts } = useMsal();
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo | null;
  const router = useRouter();
  const pathname = usePathname();
  const active = routeNames[pathname] || "Dashboard";
  const [dark, setDark] = useState(false);
  const [themeReady, setThemeReady] = useState(false);
  const [tickets, setTickets] = useState<TicketItem[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [selectedTicket, setSelectedTicket] = useState<TicketItem | null>(null);
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState("All");
  const [showNew, setShowNew] = useState(false);
  const [newTitle, setNewTitle] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [newCategory, setNewCategory] = useState("General");
  const [newPriority, setNewPriority] = useState("Medium");
  const [newFiles, setNewFiles] = useState<FileList | null>(null);
  const [reply, setReply] = useState("");
  const [replyFiles, setReplyFiles] = useState<FileList | null>(null);
  const [mobileNav, setMobileNav] = useState(false);
  const [articles, setArticles] = useState<Article[]>([]);
  const [articleQuery, setArticleQuery] = useState("");
  const [openedArticle, setOpenedArticle] = useState<Article | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [metrics, setMetrics] = useState<any>(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("helpdesk-theme");
      const enabled = saved === "dark";
      setDark(enabled);
      document.documentElement.classList.toggle("dark", enabled);
    } catch {}
    setThemeReady(true);
  }, []);
  const toggleTheme = () => {
    const next = !dark;
    setDark(next);
    document.documentElement.classList.toggle("dark", next);
    try {
      localStorage.setItem("helpdesk-theme", next ? "dark" : "light");
    } catch {}
  };

  const loadTickets = useCallback(async () => {
    if (!account) return;
    setLoading(true);
    setError("");
    try {
      const path = active === "My tickets" ? "/tickets?mine=true" : "/tickets";
      const data = await apiRequest<TicketItem[]>(instance, account, path);
      setTickets(data);
      setSelectedId((current) =>
        data.some((t) => t._id === current) ? current : data[0]?._id || "",
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
    } finally {
      setLoading(false);
    }
  }, [account, active, instance]);
  useEffect(() => {
    if (active !== "Knowledge base" && active !== "Settings") void loadTickets();
  }, [loadTickets, active]);
  useEffect(() => {
    if (!account || !selectedId) {
      setSelectedTicket(null);
      return;
    }
    let cancelled = false;
    apiRequest<TicketItem>(instance, account, `/tickets/${selectedId}`)
      .then((t) => {
        if (!cancelled) setSelectedTicket(t);
      })
      .catch((e) => {
        if (!cancelled) setError(e instanceof Error ? e.message : "Unable to load ticket details.");
      });
    return () => {
      cancelled = true;
    };
  }, [account, instance, selectedId]);
  useEffect(() => {
    if (!account || active !== "Knowledge base") return;
    apiRequest<Article[]>(
      instance,
      account,
      `/knowledge${articleQuery ? `?q=${encodeURIComponent(articleQuery)}` : ""}`,
    )
      .then(setArticles)
      .catch((e) =>
        setError(e instanceof Error ? e.message : "Could not load knowledge articles."),
      );
  }, [account, active, articleQuery, instance]);
  useEffect(() => {
    if (!account || (active !== "Dashboard" && active !== "Reports")) return;
    apiRequest(instance, account, "/tickets/metrics")
      .then(setMetrics)
      .catch(() => setMetrics(null));
  }, [account, active, instance, tickets.length]);

  const filtered = useMemo(
    () =>
      tickets.filter(
        (t) =>
          (filter === "All" || t.status === filter) &&
          `${t.ticketNumber} ${t.subject} ${t.requesterName || ""}`
            .toLowerCase()
            .includes(query.toLowerCase()),
      ),
    [tickets, filter, query],
  );
  const isSupport = Boolean(
    account?.idTokenClaims &&
    ((account.idTokenClaims as any).roles || []).some((r: string) =>
      ["Support.Agent", "Support.Access"].includes(r),
    ),
  );
  const navigate = (path: string) => {
    setMobileNav(false);
    setError("");
    router.push(path);
  };
  const signOut = () => instance.logoutRedirect({ account: account || undefined });

  async function createTicket(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || !newTitle.trim() || !newDesc.trim()) return;
    setBusy(true);
    setError("");
    try {
      const form = new FormData();
      form.append("subject", newTitle.trim());
      form.append("description", newDesc.trim());
      form.append("category", newCategory);
      form.append("priority", newPriority);
      Array.from(newFiles || []).forEach((file) => form.append("attachments", file));
      const created = await apiRequest<TicketItem>(instance, account, "/tickets", {
        method: "POST",
        body: form,
      });
      setShowNew(false);
      setNewTitle("");
      setNewDesc("");
      setNewFiles(null);
      await loadTickets();
      setSelectedId(created._id);
      navigate("/all-tickets");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Ticket could not be created.");
    } finally {
      setBusy(false);
    }
  }
  async function sendReply(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || !selectedTicket || (!reply.trim() && !replyFiles?.length)) return;
    setBusy(true);
    setError("");
    try {
      const form = new FormData();
      form.append("body", reply.trim());
      Array.from(replyFiles || []).forEach((file) => form.append("attachments", file));
      const updated = await apiRequest<TicketItem>(
        instance,
        account,
        `/tickets/${selectedTicket._id}/messages`,
        { method: "POST", body: form },
      );
      setSelectedTicket(updated);
      setTickets((old) => old.map((t) => (t._id === updated._id ? { ...t, ...updated } : t)));
      setReply("");
      setReplyFiles(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Reply could not be sent.");
    } finally {
      setBusy(false);
    }
  }
  async function claimTicket() {
    if (!account || !selectedTicket) return;
    setBusy(true);
    try {
      const t = await apiRequest<TicketItem>(
        instance,
        account,
        `/tickets/${selectedTicket._id}/claim`,
        { method: "PATCH", body: JSON.stringify({}) },
      );
      setSelectedTicket(t);
      await loadTickets();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not claim ticket.");
    } finally {
      setBusy(false);
    }
  }
  async function changeStatus(status: string) {
    if (!account || !selectedTicket) return;
    setBusy(true);
    try {
      const t = await apiRequest<TicketItem>(
        instance,
        account,
        `/tickets/${selectedTicket._id}/status`,
        { method: "PATCH", body: JSON.stringify({ status }) },
      );
      setSelectedTicket(t);
      setTickets((old) => old.map((x) => (x._id === t._id ? { ...x, ...t } : x)));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not update ticket status.");
    } finally {
      setBusy(false);
    }
  }

  const title = selectedTicket?.subject || "Select a ticket";
  const resolved =
    metrics?.resolved ??
    tickets.filter((t) => t.status === "Resolved" || t.status === "Closed").length;
  const openCount =
    (metrics?.open ?? tickets.filter((t) => t.status === "Open").length) +
    (metrics?.inProgress ?? tickets.filter((t) => t.status === "In progress").length);
  if (!themeReady) return <div className="boot-screen">Loading your workspace…</div>;

  return (
    <div className={dark ? "app dark" : "app"}>
      <aside className={`sidebar ${mobileNav ? "mobile-open" : ""}`}>
        <div className="brand">
          <div className="brand-mark">
            <Headphones size={20} />
          </div>
          <div>
            <strong>Helpdesk</strong>
            <small>MIS SUPPORT PORTAL</small>
          </div>
          <button className="icon-btn mobile-close" onClick={() => setMobileNav(false)}>
            <X size={18} />
          </button>
        </div>
        <button className="new-ticket" onClick={() => setShowNew(true)}>
          <Plus size={17} /> Create a ticket
        </button>
        {nav.map((group) => (
          <div className="nav-group" key={group.label}>
            <div className="nav-label">{group.label}</div>
            {group.items.map((item) => (
              <button
                key={item.name}
                className={`nav-item ${active === item.name ? "active" : ""}`}
                onClick={() => navigate(item.path)}
              >
                <item.icon size={18} />
                <span>{item.name}</span>
                {item.name === "All tickets" && <span className="nav-count">{tickets.length}</span>}
              </button>
            ))}
          </div>
        ))}
        <div className="sidebar-bottom">
          <div className="help-card">
            <div className="help-icon">
              <CircleHelp size={18} />
            </div>
            <strong>Need a hand?</strong>
            <p>Check the knowledge base for quick answers.</p>
            <button onClick={() => navigate("/knowledge-base")}>
              Browse articles <span>→</span>
            </button>
          </div>
          <div className="user-row">
            <div className="avatar">
              {(account?.name || "MIS")
                .split(" ")
                .map((x) => x[0])
                .slice(0, 2)
                .join("")}
            </div>
            <div className="user-info">
              <strong>{account?.name || "Signed-in user"}</strong>
              <span>Microsoft Entra ID</span>
            </div>
            <button className="icon-btn" title="Sign out" onClick={signOut}>
              <LogOut size={16} />
            </button>
          </div>
        </div>
      </aside>
      <main className="main">
        <header className="topbar">
          <div className="top-left">
            <button className="icon-btn mobile-menu" onClick={() => setMobileNav(true)}>
              <Menu size={20} />
            </button>
            <div className="breadcrumb">
              Workspace <span>/</span> <strong>{active}</strong>
            </div>
          </div>
          <div className="top-actions">
            <div className="search-global">
              <Search size={16} />
              <input
                placeholder="Search tickets..."
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
              <kbd>⌘ K</kbd>
            </div>
            <button className="icon-btn" onClick={toggleTheme} title="Toggle dark mode">
              {dark ? <Sun size={18} /> : <Moon size={18} />}
            </button>
            <button className="notification icon-btn" title="Notifications">
              <Bell size={18} />
            </button>
            <div className="avatar profile">
              {(account?.name || "MIS")
                .split(" ")
                .map((x) => x[0])
                .slice(0, 2)
                .join("")}
            </div>
          </div>
        </header>
        {error && (
          <div className="error-banner" role="alert">
            <span>{error}</span>
            <button className="icon-btn" onClick={() => setError("")}>
              <X size={16} />
            </button>
          </div>
        )}
        {active === "Dashboard" || active === "Reports" ? (
          <div className="page">
            <div className="page-heading">
              <div>
                <div className="eyebrow">MIS SUPPORT WORKSPACE</div>
                <h1>
                  {active === "Reports"
                    ? "Support reports"
                    : `Good day${account?.name ? `, ${account.name.split(" ")[0]}` : ""}`}{" "}
                  <span className="wave">✦</span>
                </h1>
                <p>
                  {active === "Reports"
                    ? "Review ticket volumes and resolution performance."
                    : "Here’s what’s happening across your MIS support desk."}
                </p>
              </div>
              <button className="primary-btn" onClick={() => setShowNew(true)}>
                <Plus size={17} /> New ticket
              </button>
            </div>
            <div className="metric-grid">
              <Metric
                icon={Ticket}
                label="Total tickets"
                value={String(metrics?.total ?? tickets.length)}
                delta="Live API"
                tone="purple"
                foot="from the ticket database"
              />
              <Metric
                icon={Clock3}
                label="Open tickets"
                value={String(openCount)}
                delta="Needs attention"
                tone="orange"
                foot="open and in progress"
              />
              <Metric
                icon={CheckCircle2}
                label="Resolved"
                value={String(resolved)}
                delta="Live API"
                tone="green"
                foot="resolved or closed"
              />
              <Metric
                icon={Clock3}
                label="Avg. close time"
                value={
                  metrics?.averageCloseTimeMs
                    ? `${(metrics.averageCloseTimeMs / 3600000).toFixed(1)}h`
                    : "—"
                }
                delta="Calculated"
                tone="blue"
                foot="based on resolved tickets"
              />
            </div>
            <div className="dashboard-grid">
              <section className="panel tickets-panel">
                <div className="panel-head">
                  <div>
                    <h2>{active === "Reports" ? "Ticket activity" : "Recent tickets"}</h2>
                    <p>Loaded from the support API</p>
                  </div>
                  <button className="text-btn" onClick={() => navigate("/all-tickets")}>
                    View all <span>→</span>
                  </button>
                </div>
                <div className="filter-row">
                  <div className="filter-pills">
                    {["All", "Open", "In progress", "Pending", "Resolved", "Closed"].map((f) => (
                      <button
                        key={f}
                        onClick={() => setFilter(f)}
                        className={filter === f ? "selected" : ""}
                      >
                        {f}
                      </button>
                    ))}
                  </div>
                  <button
                    className="icon-btn filter-icon"
                    onClick={() => {
                      setFilter("All");
                      void loadTickets();
                    }}
                    title="Refresh tickets"
                  >
                    <RefreshCw size={16} />
                  </button>
                </div>
                <TicketTable
                  tickets={filtered.slice(0, 8)}
                  selectedId={selectedId}
                  onSelect={(id) => {
                    setSelectedId(id);
                    navigate("/all-tickets");
                  }}
                  loading={loading}
                />
                <div className="panel-footer">
                  <span>
                    Showing {Math.min(filtered.length, 8)} of {tickets.length} tickets
                  </span>
                  <button onClick={() => navigate("/all-tickets")}>Open ticket queue →</button>
                </div>
              </section>
              <TicketConversation
                ticket={selectedTicket}
                account={account}
                reply={reply}
                setReply={setReply}
                files={replyFiles}
                setFiles={setReplyFiles}
                onReply={sendReply}
                onClaim={claimTicket}
                onStatus={changeStatus}
                busy={busy}
                isSupport={isSupport}
              />
            </div>
          </div>
        ) : active === "Knowledge base" ? (
          <div className="page">
            <div className="page-heading">
              <div>
                <div className="eyebrow">SELF-SERVICE SUPPORT</div>
                <h1>Knowledge base</h1>
                <p>Search published MIS help articles from the backend.</p>
              </div>
            </div>
            <div className="kb-search">
              <Search size={20} />
              <input
                placeholder="Search articles, topics and how-tos..."
                value={articleQuery}
                onChange={(e) => setArticleQuery(e.target.value)}
              />
            </div>
            <div className="kb-grid">
              {articles.map((a) => (
                <article className="kb-card" key={a._id || a.slug}>
                  <div className="kb-icon">
                    <BookOpen size={20} />
                  </div>
                  <span className="kb-category">{a.category}</span>
                  <h3>{a.title}</h3>
                  <p>{a.summary}</p>
                  <button onClick={() => setOpenedArticle(a)}>
                    Read article <span>→</span>
                  </button>
                </article>
              ))}
            </div>
            {!articles.length && !loading && (
              <div className="empty">
                No published articles found. Add articles to the knowledge base in MongoDB or
                through the API.
              </div>
            )}
          </div>
        ) : active === "Settings" ? (
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
              <button className="secondary-btn" onClick={toggleTheme}>
                {dark ? <Sun size={16} /> : <Moon size={16} />} Switch to {dark ? "light" : "dark"}{" "}
                mode
              </button>
              <h2>API connection</h2>
              <p>
                API endpoint:{" "}
                <code>{process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api"}</code>
              </p>
              <button className="secondary-btn" onClick={() => void loadTickets()}>
                <RefreshCw size={16} /> Test connection / refresh tickets
              </button>
            </section>
          </div>
        ) : (
          <div className="page">
            <div className="page-heading">
              <div>
                <div className="eyebrow">SUPPORT WORKSPACE</div>
                <h1>{active}</h1>
                <p>
                  {active === "My tickets"
                    ? "Tickets you submitted or are responsible for."
                    : active === "Team queue"
                      ? "Claim and work on unassigned support requests."
                      : "Browse and manage support requests."}
                </p>
              </div>
              <button className="primary-btn" onClick={() => setShowNew(true)}>
                <Plus size={17} /> New ticket
              </button>
            </div>
            <section className="panel">
              <div className="panel-head">
                <div>
                  <h2>{active}</h2>
                  <p>
                    {loading ? "Loading tickets from the backend…" : `${filtered.length} ticket(s)`}
                  </p>
                </div>
                <button className="icon-btn" onClick={() => void loadTickets()} title="Refresh">
                  <RefreshCw size={16} />
                </button>
              </div>
              <TicketTable
                tickets={
                  active === "Team queue"
                    ? filtered.filter((t) => !t.assigneeName || t.assigneeName === "Unassigned")
                    : filtered
                }
                selectedId={selectedId}
                onSelect={setSelectedId}
                loading={loading}
              />
              {selectedTicket && (
                <div className="queue-actions">
                  <button className="secondary-btn" onClick={() => navigate("/")}>
                    Open conversation
                  </button>
                  {active === "Team queue" && (
                    <button className="primary-btn" disabled={busy} onClick={claimTicket}>
                      Claim ticket
                    </button>
                  )}
                </div>
              )}
            </section>
          </div>
        )}
      </main>
      {showNew && (
        <div className="modal-backdrop" onClick={() => setShowNew(false)}>
          <form className="modal" onSubmit={createTicket} onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h2>Create a support ticket</h2>
                <p>Tell the MIS team what you need help with.</p>
              </div>
              <button type="button" className="icon-btn" onClick={() => setShowNew(false)}>
                <X size={18} />
              </button>
            </div>
            <label>
              Subject
              <input
                required
                maxLength={180}
                value={newTitle}
                onChange={(e) => setNewTitle(e.target.value)}
                placeholder="e.g. Unable to access Outlook"
              />
            </label>
            <label>
              Description
              <textarea
                required
                value={newDesc}
                onChange={(e) => setNewDesc(e.target.value)}
                placeholder="Describe the issue and what you have tried..."
                rows={5}
              />
            </label>
            <div className="form-two">
              <label>
                Category
                <select value={newCategory} onChange={(e) => setNewCategory(e.target.value)}>
                  {[
                    "General",
                    "Printer",
                    "Network",
                    "Email",
                    "Access",
                    "Hardware",
                    "Software",
                    "Security",
                    "Other",
                  ].map((x) => (
                    <option key={x}>{x}</option>
                  ))}
                </select>
              </label>
              <label>
                Priority
                <select value={newPriority} onChange={(e) => setNewPriority(e.target.value)}>
                  {["Low", "Medium", "High", "Urgent"].map((x) => (
                    <option key={x}>{x}</option>
                  ))}
                </select>
              </label>
            </div>
            <label className="file-label">
              <Paperclip size={15} /> Attach images, PDF, DOC/DOCX or TXT
              <input
                type="file"
                multiple
                accept="image/*,.pdf,.doc,.docx,.txt"
                onChange={(e: ChangeEvent<HTMLInputElement>) => setNewFiles(e.target.files)}
              />
            </label>
            <div className="modal-actions">
              <button type="button" className="secondary-btn" onClick={() => setShowNew(false)}>
                Cancel
              </button>
              <button
                className="primary-btn"
                disabled={busy || !newTitle.trim() || !newDesc.trim()}
              >
                {busy ? (
                  "Submitting…"
                ) : (
                  <>
                    <Plus size={16} /> Submit ticket
                  </>
                )}
              </button>
            </div>
          </form>
        </div>
      )}
      {openedArticle && (
        <div className="modal-backdrop" onClick={() => setOpenedArticle(null)}>
          <div className="modal article-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <div>
                <span className="kb-category">{openedArticle.category}</span>
                <h2>{openedArticle.title}</h2>
              </div>
              <button className="icon-btn" onClick={() => setOpenedArticle(null)}>
                <X size={18} />
              </button>
            </div>
            <p>{openedArticle.summary}</p>
            <div className="article-content">
              {openedArticle.content ||
                "This article does not yet have full content. Please contact MIS Support for assistance."}
            </div>
            <div className="modal-actions">
              <button className="primary-btn" onClick={() => setOpenedArticle(null)}>
                Close article
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function TicketTable({
  tickets,
  selectedId,
  onSelect,
  loading,
}: {
  tickets: TicketItem[];
  selectedId: string;
  onSelect: (id: string) => void;
  loading: boolean;
}) {
  return (
    <div className="ticket-table">
      <div className="table-head">
        <span>Ticket</span>
        <span>Priority</span>
        <span>Status</span>
        <span>Updated</span>
      </div>
      {tickets.map((t) => (
        <button
          className={`ticket-row ${selectedId === t._id ? "row-selected" : ""}`}
          key={t._id}
          onClick={() => onSelect(t._id)}
        >
          <div className="ticket-title">
            <span className="ticket-id">{t.ticketNumber}</span>
            <strong>{t.subject}</strong>
            <small>
              {t.requesterName || t.requesterEmail || "Requester"} · {t.department || "General"}
            </small>
          </div>
          <div>
            <StatusPill value={t.priority} />
          </div>
          <div>
            <StatusPill value={t.status} />
          </div>
          <div className="updated">
            {t.updatedAt ? new Date(t.updatedAt).toLocaleString() : "—"}
          </div>
        </button>
      ))}
      {loading && !tickets.length && <div className="empty">Loading tickets…</div>}
      {!loading && !tickets.length && <div className="empty">No tickets found.</div>}
    </div>
  );
}
function TicketConversation({
  ticket,
  account,
  reply,
  setReply,
  files,
  setFiles,
  onReply,
  onClaim,
  onStatus,
  busy,
  isSupport,
}: any) {
  return (
    <section className="panel detail-panel">
      <div className="panel-head detail-head">
        <div>
          <span className="ticket-id">{ticket?.ticketNumber || "TICKET"}</span>
          <h2>{ticket?.subject || "Select a ticket"}</h2>
        </div>
      </div>
      {ticket ? (
        <>
          <div className="detail-meta">
            <StatusPill value={ticket.status} />
            <StatusPill value={ticket.priority} />
            <span className="meta-text">
              Updated {ticket.updatedAt ? new Date(ticket.updatedAt).toLocaleString() : "recently"}
            </span>
          </div>
          <div className="detail-info">
            <div>
              <span>Requester</span>
              <strong>{ticket.requesterName || ticket.requesterEmail}</strong>
            </div>
            <div>
              <span>Assigned to</span>
              <strong>{ticket.assigneeName || "Unassigned"}</strong>
            </div>
            <div>
              <span>Category</span>
              <strong>{ticket.category || "General"}</strong>
            </div>
          </div>
          <div className="conversation-heading">
            <h3>Conversation</h3>
            <span>{ticket.messages?.length || 0} messages</span>
          </div>
          <div className="messages">
            {(ticket.messages || []).map((m: Message, i: number) => (
              <div
                className={`message ${m.authorId === account?.localAccountId || m.authorEmail === account?.username ? "staff" : ""}`}
                key={m._id || i}
              >
                <div className="message-avatar">
                  {(m.authorName || "U")
                    .split(" ")
                    .map((x: string) => x[0])
                    .slice(0, 2)
                    .join("")}
                </div>
                <div className="message-body">
                  <div className="message-by">
                    <strong>{m.authorName}</strong>
                    <span>{m.createdAt ? new Date(m.createdAt).toLocaleString() : ""}</span>
                  </div>
                  <p>{m.body}</p>
                  {m.attachments?.map((a, i) => (
                    <div className="attachment-chip" key={i}>
                      <Paperclip size={13} />
                      {a.originalName}
                    </div>
                  ))}
                </div>
              </div>
            ))}
            {!ticket.messages?.length && <div className="empty">No messages yet.</div>}
          </div>
          <form className="reply-box" onSubmit={onReply}>
            <textarea
              value={reply}
              onChange={(e) => setReply(e.target.value)}
              placeholder="Write a reply to the requester…"
            />
            <div className="reply-tools">
              <label className="attach-control" title="Attach files">
                <Paperclip size={17} />
                <input
                  type="file"
                  multiple
                  accept="image/*,.pdf,.doc,.docx,.txt"
                  onChange={(e: ChangeEvent<HTMLInputElement>) => setFiles(e.target.files)}
                />
              </label>
              <span>
                {files?.length
                  ? `${files.length} attachment(s) selected`
                  : "Attach image, PDF, DOC or TXT"}
              </span>
              <button className="send-btn" disabled={busy || (!reply.trim() && !files?.length)}>
                <Send size={15} />
                {busy ? "Sending…" : "Reply"}
              </button>
            </div>
          </form>
          {isSupport && (
            <div className="conversation-actions">
              <button className="secondary-btn" disabled={busy} onClick={onClaim}>
                Claim ticket
              </button>
              <label>
                Status{" "}
                <select
                  value={ticket.status}
                  onChange={(e) => onStatus(e.target.value)}
                  disabled={busy}
                >
                  {["Open", "In progress", "Pending", "Resolved", "Closed"].map((s: string) => (
                    <option key={s}>{s}</option>
                  ))}
                </select>
              </label>
            </div>
          )}
        </>
      ) : (
        <div className="empty">Select a ticket from the list to view its conversation.</div>
      )}
    </section>
  );
}
function Metric({
  icon: Icon,
  label,
  value,
  delta,
  tone,
  foot,
}: {
  icon: any;
  label: string;
  value: string;
  delta: string;
  tone: string;
  foot: string;
}) {
  return (
    <div className="metric-card">
      <div className="metric-top">
        <div className={`metric-icon ${tone}`}>
          <Icon size={18} />
        </div>
        <span className={`metric-delta ${tone}`}>{delta}</span>
      </div>
      <div className="metric-label">{label}</div>
      <div className="metric-value">{value}</div>
      <div className="metric-foot">{foot}</div>
    </div>
  );
}
function StatusPill({ value }: { value: string }) {
  const c = value.toLowerCase().replaceAll(" ", "-");
  return <span className={`pill pill-${c}`}>{value}</span>;
}
