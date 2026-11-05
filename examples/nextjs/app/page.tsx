'use client';

import dynamic from 'next/dynamic';

// The editor measures text with HarfBuzz and paints into the DOM, so it renders on the client only.
const Editor = dynamic(() => import('./components/Editor').then((m) => m.Editor), {
  ssr: false,
  loading: () => <div className="demo-loading" />,
});

export default function Page() {
  return <Editor />;
}
