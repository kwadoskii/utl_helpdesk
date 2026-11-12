"use client";

import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import {
  Bell,
  CheckCheck,
  MessageSquare,
  CircleCheck,
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
  Monitor,
  MonitorCog,
  SunMoon,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";
import { apiRequest } from "@/api";
import { NotificationItem, NotificationResponse, TicketItem } from "@/lib/types";
import { isSupport } from "@/lib/isSupport";
import { relativeDate } from "@/lib/relativeDate";

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
  const [themeMode, setThemeMode] = useState<"system" | "light" | "dark">("system");
  const [dark, setDark] = useState(false);
  const [themeReady, setThemeReady] = useState(false);
  const [allTickets, setAllTickets] = useState<TicketItem[]>([]);
  const [myTickets, setMyTickets] = useState<TicketItem[]>([]);

  const [notifications, setNotifications] = useState<NotificationItem[]>([]);
  const [unreadNotifsCount, setUnreadNotifsCount] = useState(0);
  const [notificationOpen, setNotificationOpen] = useState(false);
  const [notificationLoading, setNotificationLoading] = useState(false);

  const pathname = usePathname();
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
    const nextMode = (
      {
        system: "light",
        light: "dark",
        dark: "system",
      } as const
    )[themeMode];

    setThemeMode(nextMode);
    localStorage.setItem("helpdesk-theme", nextMode);

    const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");

    const isDark = nextMode === "dark" || (nextMode === "system" && mediaQuery.matches);

    setDark(isDark);
    document.documentElement.classList.toggle("dark", isDark);
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

  const markAsRead = async (notification: NotificationItem) => {
    if (notification.isRead) return;

    try {
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];

      await apiRequest(msalClient, account!, `/notifications/${notification._id}/read`, {
        method: "PATCH",
      });

      setNotifications((previous) =>
        previous.map((item) => (item._id === notification._id ? { ...item, isRead: true } : item)),
      );

      setUnreadNotifsCount((previous) => Math.max(0, previous - 1));
    } catch (error) {
      console.error("Failed to mark notification as read:", error);
    }
  };

  const markAllAsRead = async () => {
    if (!account || unreadNotifsCount === 0) return;

    setNotificationLoading(true);

    try {
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];

      await apiRequest(msalClient, account, "/notifications/read-all", { method: "PATCH" });

      setNotifications((previous) => previous.map((item) => ({ ...item, isRead: true })));

      setUnreadNotifsCount(0);
    } catch (error) {
      console.error("Failed to mark all notifications as read:", error);
    } finally {
      setNotificationLoading(false);
    }
  };

  const openNotification = async (notification: NotificationItem) => {
    await markAsRead(notification);

    const ticketId =
      typeof notification.ticketId === "string"
        ? notification.ticketId
        : notification.ticketId?._id;

    setNotificationOpen(false);

    if (ticketId) {
      navigate(`/tickets/${ticketId}`);
    }
  };

  const getNotificationIcon = (type: NotificationItem["type"]) => {
    switch (type) {
      case "ticket_created":
        return <Ticket size={17} />;

      case "new_message":
        return <MessageSquare size={17} />;

      case "ticket_closed":
        return <CircleCheck size={17} />;

      default:
        return <Bell size={17} />;
    }
  };

  const getNotifications = useCallback(async () => {
    if (!account) return;

    try {
      const msalClient = instance as unknown as Parameters<typeof apiRequest>[0];

      const response = await apiRequest<NotificationResponse>(
        msalClient,
        account,
        "/notifications",
      );

      setNotifications(response.notifications || []);
      setUnreadNotifsCount(response.unreadCount || 0);
    } catch (error) {
      console.error("Failed to load notifications:", error);
    }
  }, [account, instance]);

  useEffect(() => {
    if (!account) return;

    void getNotifications();

    const interval = window.setInterval(() => {
      void getNotifications();
    }, 5_000);

    return () => window.clearInterval(interval);
  }, [account, getNotifications]);
  const searchRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    void loadTickets();
  }, [loadTickets]);

  useEffect(() => {
    const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");

    const applyTheme = (mode: string) => {
      const isDark = mode === "dark" || (mode === "system" && mediaQuery.matches);

      setDark(isDark);
      document.documentElement.classList.toggle("dark", isDark);
    };

    const saved = localStorage.getItem("helpdesk-theme");
    const initialMode: "light" | "dark" | "system" =
      saved === "light" || saved === "dark" || saved === "system" ? saved : "system";

    setThemeMode(initialMode);
    applyTheme(initialMode);
    setThemeReady(true);

    const handleSystemChange = () => {
      if (initialMode === "system" || localStorage.getItem("helpdesk-theme") === "system") {
        applyTheme("system");
      }
    };

    mediaQuery.addEventListener("change", handleSystemChange);

    return () => {
      mediaQuery.removeEventListener("change", handleSystemChange);
    };
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
                {(item.name === "All tickets" &&
                  allTickets.filter((f) => f.status !== "Closed").length > 0 && (
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
        <header className="topbar sticky top-0 shadow-sm z-30">
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

            <button className="icon-btn" onClick={toggleTheme} title="Toggle theme">
              {themeMode === "system" ? (
                <SunMoon size={18} />
              ) : dark ? (
                <Sun size={18} />
              ) : (
                <Moon size={18} />
              )}
            </button>

            <div className="notification-wrapper">
              <button
                className={`notification icon-btn ${notificationOpen ? "notification-active" : ""}`}
                title="Notifications"
                aria-label={`Notifications, ${unreadNotifsCount} unread`}
                aria-expanded={notificationOpen}
                onClick={() => setNotificationOpen((previous) => !previous)}
              >
                <Bell size={19} />

                {unreadNotifsCount > 0 && (
                  <span className="notification-count">
                    {unreadNotifsCount > 99 ? "99+" : unreadNotifsCount}
                  </span>
                )}
              </button>

              {notificationOpen && (
                <div className="notification-dropdown">
                  <div className="notification-dropdown-header">
                    <div>
                      <h3>Notifications</h3>
                      <span>
                        {unreadNotifsCount > 0
                          ? `${unreadNotifsCount} unread notifications`
                          : "You're all caught up"}
                      </span>
                    </div>

                    <button
                      className="notification-mark-all"
                      onClick={markAllAsRead}
                      disabled={unreadNotifsCount === 0 || notificationLoading}
                    >
                      <CheckCheck size={15} />
                      Mark all read
                    </button>
                  </div>

                  <div className="notification-list">
                    {notifications.length === 0 ? (
                      <div className="notification-empty">
                        <Bell size={28} />
                        <strong>No notifications yet</strong>
                        <span>Updates about your tickets will appear here.</span>
                      </div>
                    ) : (
                      notifications.map((notification) => (
                        <button
                          key={notification._id}
                          className={`notification-item ${!notification.isRead ? "unread" : ""}`}
                          onClick={() => void openNotification(notification)}
                        >
                          <div className={`notification-item-icon ${notification.type}`}>
                            {getNotificationIcon(notification.type)}
                          </div>

                          <div className="notification-item-content">
                            <div className="notification-item-heading">
                              <strong>{notification.title}</strong>

                              {!notification.isRead && <span className="notification-unread-dot" />}
                            </div>

                            <p>{notification.message}</p>

                            <span className="notification-time">
                              {relativeDate(notification.createdAt)}
                            </span>
                          </div>
                        </button>
                      ))
                    )}
                  </div>

                  <div className="notification-dropdown-footer">
                    <button
                      onClick={() => {
                        setNotificationOpen(false);
                        navigate("/my-tickets");
                      }}
                    >
                      View my tickets
                    </button>
                  </div>
                </div>
              )}
            </div>

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
