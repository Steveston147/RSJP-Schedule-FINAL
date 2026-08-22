import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "RSJP Schedule",
  description: "RSJP and custom programme schedule planning tool",
  robots: { index: false, follow: false },
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
