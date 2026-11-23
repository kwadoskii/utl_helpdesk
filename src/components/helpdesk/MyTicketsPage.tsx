"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import NewButton from "../NewTicketButton";
import { RefreshCw, Search } from "lucide-react";
import TicketTable from "../TicketTable";
import { useMsal } from "@azure/msal-react";
import { AccountInfo } from "@azure/msal-browser";
import { TicketItem } from "@/lib/types";
import { apiRequest } from "@/api";
import NewTicketModal from "../NewTicketModal";
import { useRouter } from "next/navigation";
import { toast } from "sonner";

export default function MyTicketsPage() {
  const [showNew, setShowNew] = useState(false);
  const [loading, setLoading] = useState(true);
  const [selectedId, setSelectedId] = useState("");
  const [selectedTicket, setSelectedTicket] = useState<TicketItem | null>(null);
  const [query, setQuery] = useState("");
  const searchRef = useRef<HTMLInputElement | null>(null);

  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;
  // const [error, setError] = useState("");
  const [tickets, setTickets] = useState<TicketItem[]>([]);
  const router = useRouter();

  const loadTickets = useCallback(async () => {
    if (!account) return;
    setLoading(true);
    // setError("");

    try {
      const path = "/tickets?mine=true";

      const data = await apiRequest<TicketItem[]>(instance, account, path);
      setTickets(data);
      // setSelectedId((current) =>
      //   data.some((t) => t._id === current) ? current : data[0]?._id || "",
      // );
    } catch (e) {
      // setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
      toast.error(e instanceof Error ? e.message : "Could not load tickets from the API.");
    } finally {
      setLoading(false);
    }
  }, [account, instance]);

  const onSelect = (id: string) => {
    router.push(`/tickets/${id}`);
  };

  const filtered = [
    ...tickets.filter((t) => t.status !== "Closed"),
    ...tickets.filter((t) => t.status === "Closed"),
  ].filter((ticket) => {
    if (!query) return true;

    const search = query.toLowerCase();

    return (
      ticket.requesterName?.toLowerCase().includes(search) ||
      ticket.subject?.toLowerCase().includes(search) ||
      ticket.status?.toLowerCase().includes(search) ||
      ticket.messages?.some((message) => message.body?.toLowerCase().includes(search))
    );
  });

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        searchRef.current?.focus();
      }

      if (e.key === "Escape") {
        searchRef.current?.blur();
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, []);

  useEffect(() => {
    void loadTickets();
  }, [loadTickets, showNew]);

  return (
    <>
      <div className="page">
        <div className="page-heading">
          <div>
            {/* <div className="eyebrow">SUPPORT WORKSPACE</div> */}
            <h1>My Tickets</h1>
            <p>Tickets you submitted or are responsible for.</p>
          </div>

          <NewButton setShowNew={() => setShowNew(true)} />
        </div>

        <div className="search-global my-3 mt-0">
          <Search size={16} />
          <input
            placeholder="Search tickets..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            ref={searchRef}
          />
          <kbd>⌘ K</kbd>
        </div>

        <section className="panel">
          <div className="panel-head">
            <div>
              {/* <h2>My Tickets</h2> */}
              <p>
                {loading ? "Loading tickets from the backend…" : `${filtered.length} ticket(s)`}
              </p>
            </div>
            <button className="icon-btn" onClick={() => void loadTickets()} title="Refresh">
              <RefreshCw size={16} />
            </button>
          </div>

          <TicketTable
            tickets={filtered}
            selectedId={selectedId}
            onSelect={(id) => onSelect(id)}
            loading={loading}
            selectedTicket={(ticket) => setSelectedTicket(ticket ?? null)}
          />
        </section>
      </div>
      {showNew && <NewTicketModal setShowNew={setShowNew} />}
    </>
  );
}
