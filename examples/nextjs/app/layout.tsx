import type { Metadata } from 'next';
import '@pptx-editor.dev/core/styles/editor.css';
import './globals.css';

// The icons match the other demos (examples/vite/public, copied by `bun run assets`).
export const metadata: Metadata = {
  title: 'PPTX Editor — Next.js',
  description: 'PPTX editor in the Next.js App Router',
  icons: {
    icon: [
      { url: '/favicon.ico', sizes: 'any' },
      { url: '/icon.svg', type: 'image/svg+xml' },
      { url: '/web-app-manifest-192x192.png', type: 'image/png', sizes: '192x192' },
    ],
    apple: '/icon.png',
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
