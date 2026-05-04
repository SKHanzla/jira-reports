import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "CDEF Monthly Report",
  description: "Active CDEF Monthly Report Generator",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
