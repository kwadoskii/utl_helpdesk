"use client";
import { apiRequest } from "@/api";
import { relativeDate } from "@/lib/relativeDate";
import { Message, TicketItem } from "@/lib/types";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { CheckCircle2, Clock3, Paperclip, Plus, RefreshCw, Send, Ticket, X } from "lucide-react";
import { useRouter } from "next/navigation";
import {
  ChangeEvent,
  SyntheticEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import NewButton from "../NewTicketButton";
import { greetUser } from "@/lib/greetUser";
import StatusPill from "../StatusPill";
import TicketTable from "../TicketTable";
import { toast } from "sonner";

export default function DashboardPage() {
  const [showNew, setShowNew] = useState(false);
  const [loading, setLoading] = useState(true);
  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;
  // const [error, setError] = useState("");
  const [tickets, setTickets] = useState<TicketItem[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [filter, setFilter] = useState("All");
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState(false);
  const [selectedTicket, setSelectedTicket] = useState<TicketItem | null>(null);
  const [mobileNav, setMobileNav] = useState(false);
  const [newTitle, setNewTitle] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [newCategory, setNewCategory] = useState("General");
  const [newPriority, setNewPriority] = useState("Medium");
  const [newFiles, setNewFiles] = useState<FileList | null>(null);
  const [metrics, setMetrics] = useState<any>(null);
  const [reply, setReply] = useState("");
  const [replyFiles, setReplyFiles] = useState<FileList | null>(null);
  const [showCloseModal, setShowCloseModal] = useState(false);
  const [close, setClose] = useState(false);

  const router = useRouter();

  const navigate = (path: string) => {
    setMobileNav(false);
    router.push(path);
  };

  async function createTicket(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || !newTitle.trim() || !newDesc.trim()) return;
    setBusy(true);
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
      // navigate("/all-tickets");
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Ticket could not be created.");
      toast.error(e instanceof Error ? e.message : "Ticket could not be created.");
    } finally {
      setBusy(false);
    }
  }

  const isSupport = Boolean(
    account?.idTokenClaims &&
    (((account.idTokenClaims as any).roles || []).some((r: string) =>
      ["Support.Agent", "Support.Access"].includes(r),
    ) ||
      ((account.idTokenClaims as any).groups || []).includes(
        process.env.NEXT_PUBLIC_IT_SUPPORT_GROUP_ID || "",
      )),
  );

  const loadTickets = useCallback(async () => {
    if (!account) return;
    setLoading(true);
    try {
      const path = "/tickets";

      const data = await apiRequest<TicketItem[]>(instance, account, path);
      setTickets(data);
      setSelectedId((current) =>
        data.some((t) => t._id === current) ? current : data[0]?._id || "",
      );
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
      toast.error(e instanceof Error ? e.message : "Could not load tickets from the API.");
    } finally {
      setLoading(false);
    }
  }, [account, instance]);

  async function sendReply(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || !selectedTicket || (!reply.trim() && !replyFiles?.length)) return;
    setBusy(true);
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
      toast.success("Message sent.");
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Reply could not be sent.");
      toast.error(e instanceof Error ? e.message : "Reply could not be sent.");
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    void loadTickets();
  }, [loadTickets]);

  useEffect(() => {
    apiRequest(instance, account, "/tickets/metrics")
      .then(setMetrics)
      .catch(() => setMetrics(null));
  }, [account, instance, tickets.length]);

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
      // setSelectedTicket(t);
      await loadTickets();
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Could not claim ticket.");
      toast.error(e instanceof Error ? e.message : "Could not claim ticket.");
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
      // setError(e instanceof Error ? e.message : "Could not update ticket status.");
      toast.error(e instanceof Error ? e.message : "Could not update ticket status.");
    } finally {
      setBusy(false);
    }
  }

  async function closeTicket() {
    if (!account || !selectedTicket) return;
    setClose(true);
    setBusy(true);

    try {
      const t = await apiRequest<TicketItem>(
        instance,
        account,
        `/tickets/${selectedTicket._id}/close`,
        { method: "PATCH", body: JSON.stringify({}) },
      );
      setSelectedTicket(t);
      setTickets((old) => old.map((x) => (x._id === t._id ? { ...x, ...t } : x)));

      toast.success("Ticket closed successfully.");
      setShowCloseModal(false);
      await loadTickets();
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Could not close ticket.");
      toast.error(e instanceof Error ? e.message : "Could not close ticket.");
    } finally {
      setSelectedId("");
      setSelectedTicket(null);
      setBusy(false);
      setClose(false);
    }
  }

  const r = metrics?.resolved + metrics?.closed;
  const resolved =
    r ?? tickets.filter((t) => t.status === "Resolved" || t.status === "Closed").length;
  const openCount =
    (metrics?.open ?? tickets.filter((t) => t.status === "Open").length) +
    (metrics?.inProgress ?? tickets.filter((t) => t.status === "In progress").length);

  const NUMBER_OF_RECENT_TICKETS = 5;
  const topTickets = ((numberOfTickets: number | undefined = NUMBER_OF_RECENT_TICKETS) => {
    const priorityOrder = {
      Urgent: 1,
      High: 2,
      Medium: 3,
      Low: 4,
    };

    return [...filtered]
      .filter((t) => t.status !== "Closed")
      .sort((a, b) => new Date(b.createdAt ?? 0).getTime() - new Date(a.createdAt ?? 0).getTime())
      .sort((a, b) => priorityOrder[a.priority] - priorityOrder[b.priority])
      .slice(0, numberOfTickets);
  })();

  return (
    <>
      <div className="page">
        <div className="page-heading">
          <div>
            {/* <div className="eyebrow">SUPPORT WORKSPACE</div> */}
            <h1>
              {`${account?.name ? `${greetUser(account.name.split(" ")[0])}` : ""}`}{" "}
              <span className="wave">✦</span>
            </h1>
            <p>Here’s what’s happening across your MIS support desk.</p>
          </div>

          <NewButton setShowNew={() => setShowNew(true)} />
        </div>
        <div className="metric-grid">
          <Metric
            icon={Ticket}
            label="Total tickets"
            value={String(metrics?.total ?? tickets.length)}
            delta=""
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
            delta=""
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
                <h2>Recent tickets</h2>
                <p>
                  {loading
                    ? "Loading tickets from the backend…"
                    : topTickets.length > 0
                      ? `Showing top ${topTickets.length} recent open tickets of ${filtered.length} ticket(s)`
                      : "No recent open tickets"}
                </p>
              </div>

              <div className="flex">
                <button className="icon-btn" onClick={() => void loadTickets()} title="Refresh">
                  <RefreshCw size={16} />
                </button>
                <button className="text-btn" onClick={() => router.push("/all-tickets")}>
                  View all <span>→</span>
                </button>
              </div>
            </div>
            <TicketTable
              tickets={topTickets}
              selectedTicket={(ticket) => setSelectedTicket(ticket ?? null)}
              selectedId={selectedId}
              onSelect={setSelectedId}
              loading={loading}
            />
          </section>

          <>
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
              onClose={() => setShowCloseModal(true)}
              busy={busy}
              isSupport={isSupport}
            />

            {/* {selectedTicket && (
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
            )} */}
          </>
          {showCloseModal && (
            <div className="modal-backdrop">
              <div
                className="modal"
                role="dialog"
                aria-modal="true"
                aria-labelledby="close-ticket-title"
                style={{ maxWidth: "420px" }}
              >
                <div className="modal-head">
                  <div>
                    <h2 id="close-ticket-title">Close Ticket?</h2>
                    <p className="!text-sm text-black">
                      Are you sure you want to close this ticket? You will no longer be able to
                      reply once it is closed.
                    </p>
                  </div>
                </div>

                <div className="modal-actions">
                  <button
                    type="button"
                    className="secondary-btn"
                    onClick={() => setShowCloseModal(false)}
                    disabled={close}
                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    className="primary-btn"
                    onClick={closeTicket}
                    disabled={close}
                    style={{
                      background: "#dc2626",
                      boxShadow: "none",
                    }}
                  >
                    {close ? "Closing..." : "Yes, Close Ticket"}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

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
    </>
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
  onClose,
  busy,
  isSupport,
}: any) {
  const messagesRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const messages = messagesRef.current;
    if (messages) {
      messages.scrollTop = messages.scrollHeight;
    }
  }, [ticket]);

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
              Last updated {ticket.updatedAt ? relativeDate(ticket.updatedAt) : "recently"}
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
          <div className="messages overflow-hidden" ref={messagesRef}>
            {(ticket.messages || []).map((m: Message, i: number) => (
              <div
                className={`message ${m.authorId === account?.localAccountId || m.authorEmail === account?.username ? "staff sent" : ""}`}
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
                  <div
                    className={`message-by ${m.authorId === account?.localAccountId ? "sent" : ""}`}
                  >
                    <strong>{m.authorName}</strong>
                    <span>{m.createdAt ? new Date(m.createdAt).toLocaleString() : ""}</span>
                  </div>
                  <p className="">{m.body}</p>
                  {m.attachments?.map((a, i) => (
                    <div className="attachment-chip" key={i}>
                      <Paperclip size={13} />
                      {a.originalName}
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          {!ticket.messages?.length && <div className="empty">No messages yet.</div>}

          <form className="reply-box" onSubmit={onReply}>
            <textarea
              value={reply}
              onChange={(e) => setReply(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                  e.preventDefault();
                  e.currentTarget.form?.requestSubmit();
                }
              }}
              placeholder="Type a message"
              className="overflow-hidden py-1.5"
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
          {(isSupport ||
            (ticket.requesterId &&
              ticket.requesterId ===
                ((account?.idTokenClaims as any)?.oid || (account?.idTokenClaims as any)?.sub))) &&
            ticket.status !== "Closed" && (
              <div className="conversation-actions">
                {isSupport && (
                  <>
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
                        {["Open", "In progress", "Pending", "Resolved", "Closed"].map(
                          (s: string) => (
                            <option key={s}>{s}</option>
                          ),
                        )}
                      </select>
                    </label>
                  </>
                )}
                <button className="secondary-btn" disabled={busy} onClick={onClose}>
                  Close ticket
                </button>
              </div>
            )}
        </>
      ) : (
        <div className="empty">Select a ticket from the list to view its conversation.</div>
      )}
    </section>
  );
}
