// The React demo's shell in happy-dom: the shared cases of `scripts/test/demo-shell-cases.ts`.

import { afterEach } from 'bun:test';
import { demoShellCases } from '../../../../scripts/test/demo-shell-cases.ts';
import { closeDemo, openDemo } from './demo-harness.tsx';

afterEach(closeDemo);

demoShellCases('React', openDemo);
