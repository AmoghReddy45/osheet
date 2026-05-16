import type { Metadata } from "next";
import { GeistSans } from "geist/font/sans";
import { GeistMono } from "geist/font/mono";
import "./globals.css";

export const metadata: Metadata = {
  title: "osheet — AI-native spreadsheet compiler",
  description: "Upload any .xlsx. Get back an AI-native workbook.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className={`${GeistSans.variable} ${GeistMono.variable}`}>
      <body className="bg-[#171717] text-[#ededed] font-sans antialiased min-h-screen">
        {children}
      </body>
    </html>
  );
}
