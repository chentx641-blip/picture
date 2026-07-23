import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "图链工坊｜链接转图片",
  description: "按 Excel 行列批量将图片链接转换为图片预览。",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-CN"><body>{children}</body></html>;
}
