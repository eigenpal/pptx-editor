// The agent replay page: a scripted agent edits the presentation on the server in steps. Each
// step returns the complete PPTX file, and `createDocumentRefresh(editor)` applies it in place.
// `reviewModule()` compares the files and the editor flashes what changed. Step through the
// edits one at a time, or play them all. The controls are in `agent-panel.tsx`.

import { StrictMode, useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { packagedFonts } from '@docx-editor.dev/fonts';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { reviewModule } from '@pptx-editor.dev/pro';
import { PptxEditor, useFonts } from '@pptx-editor.dev/react';
import { AgentPanel, fetchStep } from './agent-panel.tsx';
import { useNarrow } from './narrow.ts';
import './styles.css';

// The review module computes the changes of each step (EigenPal Pro License). Evaluation use
// needs no key.
const MODULES: readonly EditorModule[] = [
  reviewModule({ licenseKey: import.meta.env.VITE_PPTX_LICENSE }),
];

function App() {
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [error, setError] = useState('');
  const fonts = useFonts(packagedFonts());
  const narrow = useNarrow();
  useEffect(() => {
    fetchStep(0).then(setBytes, () => setError('The presentation did not load.'));
  }, []);
  if (!bytes) return <p role="status">{error || 'Loading the presentation'}</p>;
  return (
    <PptxEditor.Root
      document={bytes}
      colorMode="system"
      fonts={fonts}
      modules={MODULES}
      onReady={(instance: PptxEditorInstance) => {
        (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = instance;
      }}
    >
      <PptxEditor.Viewport className="refresh-app">
        <PptxEditor.Toolbar />
        <div className="refresh-workspace">
          <PptxEditor.SlideRail orientation={narrow ? 'horizontal' : 'vertical'} />
          <main className="refresh-docpane">
            <div className="refresh-canvas">
              <PptxEditor.Content />
              <PptxEditor.Loading overlay />
            </div>
            <PptxEditor.NotesPane defaultOpen={!narrow} />
          </main>
          <AgentPanel />
        </div>
      </PptxEditor.Viewport>
    </PptxEditor.Root>
  );
}

const container = document.getElementById('app');
if (container)
  createRoot(container).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
