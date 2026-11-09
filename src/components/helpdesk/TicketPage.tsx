"use client";

import { apiRequest } from "@/api";
import { formatDate, relativeDate } from "@/lib/relativeDate";
import { TicketItem } from "@/lib/types";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { Paperclip, Send } from "lucide-react";
import { notFound } from "next/navigation";
import { ChangeEvent, SyntheticEvent, useCallback, useEffect, useState } from "react";

export default function TicketPage({ id }: { id: string }) {
  const [replyFiles, setReplyFiles] = useState<FileList | null>(null);
  const [busy, setBusy] = useState(false);
  const [close, setClose] = useState(false);
  const [loading, setLoading] = useState(false);
  const [reply, setReply] = useState("");
  const [error, setError] = useState("");
  const [ticket, setTicket] = useState<TicketItem | null>(null);

  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;

  async function sendReply(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || (!reply.trim() && !replyFiles?.length)) return;
    setBusy(true);
    setError("");
    try {
      const form = new FormData();
      form.append("body", reply.trim());
      Array.from(replyFiles || []).forEach((file) => form.append("attachments", file));
      const updated = await apiRequest<TicketItem>(instance, account, `/tickets/${id}/messages`, {
        method: "POST",
        body: form,
      });
      setTicket(updated);
      // setTickets((old) => old.map((t) => (t._id === updated._id ? { ...t, ...updated } : t)));
      setReply("");
      setReplyFiles(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Reply could not be sent.");
    } finally {
      setBusy(false);
    }
  }

  async function closeTicket() {
    if (!account) return;
    if (!window.confirm("Close this ticket? You will not be able to reply after it is closed."))
      return;
    setClose(true);
    setError("");
    try {
      const t = await apiRequest<TicketItem>(instance, account, `/tickets/${id}/close`, {
        method: "PATCH",
        body: JSON.stringify({}),
      });
      // setSelectedTicket(t);
      // setTickets((old) => old.map((x) => (x._id === t._id ? { ...x, ...t } : x)));
      await loadTicket();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not close ticket.");
    } finally {
      // setSelectedId("");
      // setSelectedTicket(null);
      setClose(false);
    }
  }

  const loadTicket = useCallback(async () => {
    if (!account) return;
    setLoading(true);
    setError("");
    try {
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];
      const ticket = await apiRequest<TicketItem>(msalClient, account, `/tickets/${id}`);

      setTicket(ticket);
      // setAllTickets(allTickets);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
    } finally {
      setLoading(false);
    }
  }, [account, instance, id]);

  useEffect(() => {
    void loadTicket();

    // if (!ticket) {
    //   return notFound();
    // }
  }, [loadTicket]);

  return (
    <div className="page">
      <div className="page-heading">
        <div>
          {/* <div className="eyebrow">SUPPORT WORKSPACE</div> */}
          <h2 className="text-xl font-bold">
            <span className="text-sm">Ticket {ticket?.ticketNumber}:</span> {ticket?.subject}
          </h2>
          <p>Raised by {ticket?.requesterName}.</p>
        </div>
      </div>

      <div className="dashboard-grid-ticket relative">
        <div className="space-y-4">
          <section className="panel">
            <div className="bg-gray-100 dark:bg-[#151a2b] px-5 py-3 flex items-center text-sm border-b dark:border-b-[--border]">
              <p>{ticket?.status === "Closed" ? "Closed" : "Post a reply"}</p>
            </div>
            <div className="panel-head flex !p-4">
              <div className="flex-1">
                <form className="reply-box !m-0" onSubmit={sendReply}>
                  <textarea
                    value={reply}
                    onChange={(e) => setReply(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                        e.preventDefault();
                        e.currentTarget.form?.requestSubmit();
                      }
                    }}
                    disabled={loading || busy || ticket?.status === "Closed"}
                    placeholder="Type a message"
                    className="overflow-hidden py-1.5 !text-sm"
                  />
                  <div className="reply-tools">
                    <label className="attach-control" title="Attach files">
                      <Paperclip size={17} />
                      <input
                        type="file"
                        multiple
                        accept="image/*,.pdf,.doc,.docx,.txt"
                        onChange={(e: ChangeEvent<HTMLInputElement>) =>
                          setReplyFiles(e.target.files)
                        }
                        disabled={loading || busy || ticket?.status === "Closed"}
                      />
                    </label>
                    <span>
                      {replyFiles?.length
                        ? `${replyFiles.length} attachment(s) selected`
                        : "Attach image, PDF, DOC or TXT"}
                    </span>
                    <button
                      className="send-btn !px-10"
                      disabled={busy || (!reply.trim() && !replyFiles?.length)}
                    >
                      <Send size={15} />
                      {busy ? "Sending…" : "Reply"}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          </section>

          <section className="mgs space-y-4">
            {ticket?.messages
              ?.sort(
                (a, b) =>
                  Date.parse(String(b.createdAt ?? "")) - Date.parse(String(a.createdAt ?? "")),
              )
              .map((m) => (
                <div className="panel" key={m._id}>
                  <div className="bg-gray-100 dark:bg-[#151a2b] px-5 flex flex-col gap-0.5 py-2 text-xs border-b dark:border-b-[--border]">
                    <p>{m.authorName}</p>
                    <p className="text-xs text-gray-500">
                      {m.createdAt ? relativeDate(m.createdAt) : "—"}
                    </p>
                  </div>
                  <div className="panel-head flex !p-4 !pt-1">
                    <p className="!text-sm">{m.body}</p>
                  </div>
                </div>
              ))}
          </section>
        </div>

        <section className="space-y-4 sticky top-6">
          <div className="panel shadow-md p-5 text-sm">
            <h2 className="mb-5 font-medium text-base">Ticket Details</h2>

            <div className="flex flex-col gap-y-4">
              <div className="flex flex-col">
                <p className="text-gray-500 font-light">Status</p>
                <p>{ticket?.status}</p>
              </div>
              <div className="flex flex-col">
                <p className="text-gray-500 font-light">Category</p>
                <p>{ticket?.category}</p>
              </div>
              <div className="flex flex-col">
                <p className="text-gray-500 font-light">Priority</p>
                <p>{ticket?.priority}</p>
              </div>
              <div className="flex flex-col">
                <p className="text-gray-500 font-light">Created</p>
                <p>{ticket?.createdAt ? formatDate(ticket.createdAt) : "—"}</p>
              </div>
              <div className="flex flex-col">
                <p className="text-gray-500 font-light">Last Updated</p>
                <p>{ticket?.createdAt ? formatDate(ticket.updatedAt) : "—"}</p>
              </div>
            </div>
          </div>

          <div className="flex flex-col">
            {ticket?.status !== "Closed" && (
              <button
                className="btn secondary-btn !bg-red-600/90 !text-white !text-sm"
                onClick={closeTicket}
                disabled={close || loading}
              >
                {close ? "Closing..." : "Close Ticket"}
              </button>
            )}
          </div>
        </section>
      </div>
    </div>
  );
}
