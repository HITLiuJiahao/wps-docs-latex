import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: '公式转换器｜LaTeX 转 WPS/Word 原生公式',
  description: '在本地浏览器中将 DOCX 中的 LaTeX 代码转换为可编辑的 WPS/Word 原生公式。',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
