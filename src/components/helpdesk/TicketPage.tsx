"use client";

import { apiRequest } from "@/api";
import { formatDate, relativeDate } from "@/lib/relativeDate";
import { MessageAttachment, TicketItem } from "@/lib/types";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { ExternalLink, File, FileDown, FileImage, FileText, Paperclip, Send } from "lucide-react";
import { notFound } from "next/navigation";
import { ChangeEvent, SyntheticEvent, useCallback, useEffect, useState } from "react";
import { toast } from "sonner";

const API_ORIGIN = (process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000/api").replace(
  /\/api\/?$/,
  "",
);

function getAttachmentUrl(storageName: string) {
  if (storageName?.startsWith("http://") || storageName?.startsWith("https://")) {
    return storageName;
  }

  return `${API_ORIGIN}${storageName?.startsWith("/") ? storageName : `/uploads/${storageName}`}`;
}

function formatFileSize(size: number) {
  if (!size) return "";

  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;

  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

export default function TicketPage({ id }: { id: string }) {
  const [replyFiles, setReplyFiles] = useState<FileList | null>(null);
  const [busy, setBusy] = useState(false);
  const [close, setClose] = useState(false);
  const [loading, setLoading] = useState(false);
  const [reply, setReply] = useState("");
  // const [error, setError] = useState("");
  const [ticket, setTicket] = useState<TicketItem | null>(null);
  const [showCloseModal, setShowCloseModal] = useState(false);

  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;

  async function sendReply(e: SyntheticEvent) {
    e.preventDefault();
    if (!account || (!reply.trim() && !replyFiles?.length)) return;
    setBusy(true);
    // setError("");
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
      toast.success("Message sent.");
      setReply("");
      setReplyFiles(null);
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Reply could not be sent.");
      // setError(e instanceof Error ? e.message : "Reply could not be sent.");
    } finally {
      setBusy(false);
    }
  }

  async function closeTicket() {
    if (!account) return;

    setClose(true);

    try {
      await apiRequest<TicketItem>(instance, account, `/tickets/${id}/close`, {
        method: "PATCH",
        body: JSON.stringify({}),
      });

      toast.success("Ticket closed successfully.");
      setShowCloseModal(false);
      await loadTicket();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Could not close ticket.");
    } finally {
      setClose(false);
    }
  }

  const loadTicket = useCallback(async () => {
    if (!account) return;
    setLoading(true);
    // setError("");
    try {
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];
      const ticket = await apiRequest<TicketItem>(msalClient, account, `/tickets/${id}`);

      setTicket(ticket);
      // setAllTickets(allTickets);
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
      toast.error(e instanceof Error ? e.message : "Could not load tickets from the API.");
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
          <h2 className="text-xl font-bold">{ticket?.subject}</h2>
          <h3 className="text-sm">{ticket?.ticketNumber}</h3>
          <p>Raised by {ticket?.requesterName}</p>
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
                  <div className="panel-head flex flex-col !p-4 !pt-1 !self !items-start">
                    {m.body && (
                      <p className="!text-sm !text-[--text] whitespace-pre-wrap">{m.body}</p>
                    )}

                    <MessageAttachments attachments={m.attachments || []} />
                  </div>
                </div>
              ))}
          </section>
        </div>

        <section className="space-y-4 sticky top-20">
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
                className="primary-btn"
                onClick={() => setShowCloseModal(true)}
                disabled={close}
              >
                Close Ticket
              </button>
            )}
          </div>
        </section>
      </div>

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
                  Are you sure you want to close this ticket? You will no longer be able to reply
                  once it is closed.
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
  );
}

function MessageAttachments({ attachments }: { attachments: MessageAttachment[] }) {
  if (!attachments?.length) return null;

  return (
    <div className="space-y-3">
      <p className="text-xs font-medium text-gray-500">Attachments ({attachments.length})</p>

      {attachments.map((file, index) => {
        const fileUrl = getAttachmentUrl(file.storageName);

        const mime = file.mimeType?.toLowerCase() || "";
        const name = file.originalName || `Attachment ${index + 1}`;

        const isImage = mime.startsWith("image/") || /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(name);

        const isPdf = mime === "application/pdf" || /\.pdf$/i.test(name);

        const isText = mime.startsWith("text/") || /\.txt$/i.test(name);

        const isWord =
          /\.(doc|docx)$/i.test(name) ||
          mime.includes("wordprocessingml") ||
          mime.includes("msword");

        // Image attachment
        if (isImage) {
          return (
            <div key={file._id || index} className="space-y-2">
              <a href={fileUrl} target="_blank" rel="noopener noreferrer">
                <img
                  src={fileUrl}
                  alt={name}
                  className="max-w-full max-h-[320px] rounded-lg border dark:border-[--border] object-contain cursor-pointer"
                />
              </a>

              <p className="text-xs text-gray-500">{name}</p>
            </div>
          );
        }

        // PDF attachment
        if (isPdf) {
          return (
            <div
              key={file._id || index}
              className="rounded-lg border dark:border-[--border] overflow-hidden"
            >
              <div className="flex items-center justify-between gap-3 p-3 bg-gray-50 dark:bg-[#151a2b]">
                <div className="flex items-center gap-2 min-w-0">
                  <FileText size={20} className="text-red-500 shrink-0" />

                  <div className="min-w-0">
                    <p className="text-sm font-medium truncate">{name}</p>
                    <p className="text-xs text-gray-500">
                      PDF {file.size ? `· ${formatFileSize(file.size)}` : ""}
                    </p>
                  </div>
                </div>

                <a
                  href={fileUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-sm text-indigo-500 flex items-center gap-1 shrink-0"
                >
                  Open <ExternalLink size={14} />
                </a>
              </div>
              {/* <iframe src={fileUrl} title={name} className="w-full h-[400px] bg-white" /> */}
              {/* <object type="application/pdf" data={fileUrl} width="600" height="400">
                <p>
                  Your browser does not support PDFs. <a href={fileUrl}>Download the PDF</a>.
                </p>
              </object> */}
            </div>
          );
        }

        // Word and text attachments
        return (
          <a
            key={file._id || index}
            href={fileUrl}
            target="_blank"
            rel="noopener noreferrer"
            download={!isText}
            className="flex items-center justify-between gap-3 p-3 rounded-lg border dark:border-[--border] hover:bg-gray-50 dark:hover:bg-[#151a2b] transition"
          >
            <div className="flex items-center gap-3 min-w-0">
              {isWord ? (
                <FileText size={22} className="text-blue-500 shrink-0" />
              ) : isText ? (
                <File size={22} className="text-gray-500 shrink-0" />
              ) : (
                <FileImage size={22} className="text-indigo-500 shrink-0" />
              )}

              <div className="min-w-0">
                <p className="text-sm font-medium truncate">{name}</p>
                <p className="text-xs text-gray-500">
                  {isWord ? "Word document" : isText ? "Text file" : "Attachment"}
                  {file.size ? ` · ${formatFileSize(file.size)}` : ""}
                </p>
              </div>
            </div>

            <FileDown size={18} className="text-gray-500 shrink-0" />
          </a>
        );
      })}
    </div>
  );
}
