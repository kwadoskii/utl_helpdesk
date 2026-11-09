import { apiRequest } from "@/api";
import { TicketItem } from "@/lib/types";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { Paperclip, Plus, X } from "lucide-react";
import { useRouter } from "next/navigation";
import { ChangeEvent, SyntheticEvent, useState } from "react";

export default function NewTicketModal({ setShowNew }: { setShowNew: (show: boolean) => void }) {
  const [loading, setLoading] = useState(true);
  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;
  const [error, setError] = useState("");
  // const [tickets, setTickets] = useState<TicketItem[]>([]);

  const [busy, setBusy] = useState(false);
  // const [selectedTicket, setSelectedTicket] = useState<TicketItem | null>(null);

  const [newTitle, setNewTitle] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [newCategory, setNewCategory] = useState("General");
  const [newPriority, setNewPriority] = useState("Medium");
  const [newFiles, setNewFiles] = useState<FileList | null>(null);

  const router = useRouter();

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
      // await loadTickets();
      // setSelectedId(created._id);
      // navigate("/all-tickets");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Ticket could not be created.");
    } finally {
      setBusy(false);
    }
    router.refresh();
  }

  return (
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
            className="!text-sm"
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
          <button className="primary-btn" disabled={busy || !newTitle.trim() || !newDesc.trim()}>
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
  );
}
