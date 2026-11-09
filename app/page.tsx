import DashboardPage from "@/components/helpdesk/DashboardPage";
import RequireSupport from "@/components/helpdesk/RequireSupport";

export default function Page() {
  return (
    <RequireSupport>
      <DashboardPage />
    </RequireSupport>
  );
}
