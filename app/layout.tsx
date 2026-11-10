import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Tách bộ biểu VTVL: HĐND – UBND",
  description: "Tách file bộ biểu vị trí việc làm HĐND-UBND cấp xã thành hai file HĐND và UBND.",
  robots: { index: false, follow: false },
};

export const viewport: Viewport = { width: "device-width", initialScale: 1 };

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
