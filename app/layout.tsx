import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "에이비랩 코리아 Sales Report",
  description: "ABLAB Korea sales dashboard",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
