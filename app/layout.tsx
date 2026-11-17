import type { Metadata } from "next";
import { Toaster } from "sonner";

import Providers from "./providers";
import "./globals.css";
import HelpdeskLayout from "@/components/helpdesk/HelpdeskLayout";

export const metadata: Metadata = {
  title: "Helpdesk | MIS Support Portal",
  description: "UTL MIS support ticketing and knowledge base portal.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <head>
        <link rel="icon" href="./favicon.ico" sizes="any" />
      </head>

      <body>
        <Providers>
          <HelpdeskLayout>
            {children}
            <Toaster position="top-center" richColors />
          </HelpdeskLayout>
        </Providers>
      </body>
    </html>
  );
}
