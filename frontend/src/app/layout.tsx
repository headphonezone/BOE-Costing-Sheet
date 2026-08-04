import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "BOE Costing Dashboard",
  description: "Ferrari Video / Headphone Zone BOE costing & duty dashboard",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className="h-full">
      <body className="min-h-full bg-slate-50 text-slate-900 antialiased dark:bg-slate-950 dark:text-slate-100">
        {children}
      </body>
    </html>
  );
}
