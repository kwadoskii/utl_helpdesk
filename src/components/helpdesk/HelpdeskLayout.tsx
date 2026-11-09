"use client";

import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import {
  Bell,
  BookOpen,
  FileText,
  Headphones,
  LayoutDashboard,
  LogOut,
  Menu,
  Moon,
  Search,
  Settings2,
  Sun,
  Ticket,
  Users,
  X,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";
import { apiRequest } from "@/api";
import { TicketItem } from "@/lib/types";
import { isSupport } from "@/lib/isSupport";

export default function HelpdeskLayout({ children }: { title?: string; children: ReactNode }) {
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
        { name: "Settings", path: "/settings", icon: Settings2 },
      ],
    },
  ];
  const [mobileNav, setMobileNav] = useState(false);
  const routeNames: Record<string, string> = {
    "/": "Dashboard",
    "/all-tickets": "All tickets",
    "/my-tickets": "My tickets",
    "/knowledge-base": "Knowledge base",
    "/knowledge-base/:slug": "Knowledge base",
    "/team-queue": "Team queue",
    "/reports": "Reports",
    "/settings": "Settings",
    "/tickets/:id": "Ticket",
  };
  const router = useRouter();
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [query, setQuery] = useState("");
  const [dark, setDark] = useState(false);
  const [themeReady, setThemeReady] = useState(false);
  const [allTickets, setAllTickets] = useState<TicketItem[]>([]);
  const [myTickets, setMyTickets] = useState<TicketItem[]>([]);

  const pathname = usePathname();
  console.log(pathname);
  const active =
    routeNames[pathname] ||
    Object.entries(routeNames).find(([route]) => {
      const pattern = new RegExp(`^${route.replace(/:[^/]+/g, "[^/]+")}$`);

      return pattern.test(pathname);
    })?.[1] ||
    "Dashboard";

  const navigate = (path: string) => {
    setMobileNav(false);
    setError("");
    router.push(path);
  };

  const { instance, accounts } = useMsal();
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo | null;

  const visibleNav = isSupport(account)
    ? nav
    : [
        {
          label: "Workspace",
          items: nav
            .flatMap((group) => group.items)
            .filter((item) => ["/my-tickets", "/knowledge-base"].includes(item.path)),
        },
      ];

  // const signOut = () => instance.logoutRedirect({ account: account || undefined });

  async function signOut() {
    try {
      const account = instance.getActiveAccount();

      await instance.logoutRedirect({
        account: account || undefined,
        postLogoutRedirectUri: window.location.origin,
      });
    } catch (error) {
      console.error("[MSAL LOGOUT ERROR]", error);

      alert("An error occurred while logging out. Please try again.");
    }
  }

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
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];
      const myTickets = await apiRequest<TicketItem[]>(msalClient, account, "/tickets?mine=true");
      const allTickets = await apiRequest<TicketItem[]>(msalClient, account, "/tickets");

      setMyTickets(myTickets);
      setAllTickets(allTickets);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not load tickets from the API.");
    } finally {
      setLoading(false);
    }
  }, [account, instance]);

  const searchRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    void loadTickets();
  }, [loadTickets]);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("helpdesk-theme");
      const enabled = saved === "dark";
      setDark(enabled);
      document.documentElement.classList.toggle("dark", enabled);
    } catch {}
    setThemeReady(true);
  }, []);

  // activate central search
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

  if (!themeReady)
    return <div className={"boot-screen " + dark ? "app dark" : ""}>Loading your workspace…</div>;

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

        {visibleNav.map((group) => (
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
                {(item.name === "All tickets" && (
                  <span className="nav-count">
                    {allTickets.filter((f) => f.status !== "Closed").length}
                  </span>
                )) ||
                  (item.name === "My tickets" &&
                    myTickets.filter((f) => f.status !== "Closed").length > 0 && (
                      <span className="nav-count">
                        {myTickets.filter((f) => f.status !== "Closed").length}
                      </span>
                    ))}
              </button>
            ))}
          </div>
        ))}
        <div className="sidebar-bottom">
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
              <span>{account?.username}</span>
              <span>{account?.username.includes("utltrustees") ? "UTL" : "UTLAM"}</span>
            </div>
            <button className="icon-btn" title="Sign out" onClick={signOut}>
              <LogOut size={16} />
            </button>
          </div>
        </div>
      </aside>

      {/* Move the existing header JSX here. */}
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

          {/* Preserve your existing header actions here. */}
          <div className="top-actions">
            <div className="search-global">
              <Search size={16} />
              <input
                placeholder="Search tickets..."
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                ref={searchRef}
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
        {children}
      </main>
    </div>
  );
}
