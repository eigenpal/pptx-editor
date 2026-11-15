// Geist (OFL-1.1) for the shell, self-hosted: the demo makes no font request to a third party.
import '@fontsource-variable/geist';
import './styles.css';
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App.tsx';
import { PreviewBanner } from '../../shared/PreviewBanner.tsx';

const container = document.getElementById('app');
if (container)
  createRoot(container).render(
    <StrictMode>
      <div className="demo-frame">
        <PreviewBanner />
        <App />
      </div>
    </StrictMode>
  );
