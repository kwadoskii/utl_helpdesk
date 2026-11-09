import AllTicketsPage from "@/components/helpdesk/AllTicketsPage";
import RequireSupport from "@/components/helpdesk/RequireSupport";

export default function AllTicketsRoute() {
  return (
    <RequireSupport>
      <AllTicketsPage />
    </RequireSupport>
  );
}
