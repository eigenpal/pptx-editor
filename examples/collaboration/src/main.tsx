import '@pptx-editor.dev/core/styles/editor.css';
import '../../shared/collaboration/room.css';
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App.tsx';

const container = document.getElementById('app');
if (container)
  createRoot(container).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
