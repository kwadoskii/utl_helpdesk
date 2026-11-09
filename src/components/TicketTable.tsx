import { formatDate, relativeDate } from "@/lib/relativeDate";
import { TicketItem } from "@/lib/types";
import StatusPill from "./StatusPill";

export default function TicketTable({
  tickets,
  selectedId,
  selectedTicket,
  onSelect,
  loading,
}: {
  tickets: TicketItem[];
  selectedId: string;
  selectedTicket: (t?: TicketItem | null | undefined) => void;
  onSelect: (id: string) => void;
  loading: boolean;
}) {
  return (
    <div className="ticket-table">
      <div className="table-head">
        <span>Ticket</span>
        <span>Priority</span>
        <span>Status</span>
        <span>Created</span>
        <span>Updated</span>
      </div>
      {tickets.map((t) => (
        <button
          className={`ticket-row ${selectedId === t._id ? "row-selected" : ""}`}
          key={t._id}
          onClick={() => {
            if (selectedId !== t._id) {
              onSelect(t._id);
              selectedTicket(t);
            } else {
              onSelect("");
              selectedTicket(null);
            }
          }}
        >
          <div className="ticket-title">
            <span className="ticket-id">{t.ticketNumber}</span>
            <strong>{t.subject}</strong>
            <small>
              {t.requesterName || t.requesterEmail || "Requester"} · {t.category || "General"}
            </small>
          </div>
          <div>
            <StatusPill value={t.priority} />
          </div>
          <div>
            <StatusPill value={t.status} />
          </div>
          <div className="updated flex flex-col gap-1">
            <span>{t.createdAt ? formatDate(t.createdAt).split("  ")[1] : "—"}</span>
            <span>{t.createdAt ? formatDate(t.createdAt).split("  ")[0] : "—"}</span>
          </div>
          <div className="updated">{t.updatedAt ? relativeDate(t.updatedAt) : "—"}</div>
        </button>
      ))}
      {loading && !tickets.length && <div className="empty">Loading tickets…</div>}
      {!loading && !tickets.length && <div className="empty">No tickets found.</div>}
    </div>
  );
}
