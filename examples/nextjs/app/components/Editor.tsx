'use client';

// PPTX Editor in the Next.js App Router. The `<PptxEditor>` sugar renders the complete default
// editor: toolbar, slide rail, canvas, speaker notes, and the slide show (F5). Compose
// `PptxEditor.Root` / `.Viewport` / `.Toolbar` / `.Content` instead when you want to own the
// layout (see examples/vite).

import { useEffect, useState } from 'react';
import { PptxEditor } from '@pptx-editor.dev/react';
import { drawingModule } from '@pptx-editor.dev/pro/drawing';

// Glow, inner shadow, reflection, and soft edge (EigenPal Pro License).
const modules = [drawingModule({ licenseKey: process.env.NEXT_PUBLIC_PPTX_LICENSE })];

export function Editor() {
  const [bytes, setBytes] = useState<Uint8Array | undefined>();
  useEffect(() => {
    void fetch('/showcase.pptx')
      .then((res) => res.arrayBuffer())
      .then((b) => setBytes(new Uint8Array(b)));
  }, []);
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <PptxEditor
        {...(bytes ? { document: bytes } : {})}
        modules={modules}
        onReady={(editor) => {
          // The e2e suite and the verification scripts drive the editor through this handle.
          (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = editor;
        }}
      />
    </div>
  );
}
