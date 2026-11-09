import TicketPage from "@/components/helpdesk/TicketPage";

export default async function PageRoute({ params }: PageProps) {
  const { id } = await params;

  return <TicketPage id={id} />;
}

interface PageProps {
  params: Promise<{ id: string }>;
}
