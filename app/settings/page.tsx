import RequireSupport from "@/components/helpdesk/RequireSupport";
import SettingsPage from "@/components/helpdesk/SettingsPage";

export default function SettingsRoute() {
  return (
    <RequireSupport>
      <SettingsPage />
    </RequireSupport>
  );
}
