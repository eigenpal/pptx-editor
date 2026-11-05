// The editor's format, history, and nudge keys as commands: a key press maps to one
// `EditorCommand`, with no editor state. Split from `pptx-editor.ts`, which runs them, so the
// facade stays under the line cap.

import type { EditorCommand } from '../contracts/index.ts';

/** The modifier state of a key press. */
interface KeyMods {
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
}

/**
 * The format or history command of a key, or null: undo and redo, bold, italic, underline,
 * Clear Formatting (Ctrl/Cmd+\\), copy and paste formatting (Ctrl/Cmd+Shift+C and +V), New Slide
 * (Ctrl+M), and, in text, subscript (Ctrl/Cmd+=) and superscript (Ctrl/Cmd+Shift+=).
 */
export function formatKeyCommand(key: string, e: KeyMods, inText: boolean): EditorCommand | null {
  const mod = e.metaKey || e.ctrlKey;
  const lower = key.toLowerCase();
  if (mod && inText && (key === '=' || key === '+'))
    return { type: e.shiftKey ? 'toggleSuperscript' : 'toggleSubscript' };
  if (mod && e.shiftKey && (lower === 'c' || lower === 'v'))
    return { type: lower === 'c' ? 'copyFormat' : 'pasteFormat' };
  if (mod && lower === 'z') return { type: e.shiftKey ? 'redo' : 'undo' };
  if (mod && lower === 'y') return { type: 'redo' };
  if (mod && lower === 'b') return { type: 'toggleBold' };
  if (mod && key === '\\') return { type: 'clearFormatting' };
  if (mod && lower === 'i') return { type: 'toggleItalic' };
  if (mod && lower === 'u') return { type: 'toggleUnderline' };
  // New Slide: Ctrl+M (a browser keeps ⇧⌘N and ⌘M for itself), also while text is edited.
  if (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && lower === 'm')
    return { type: 'addSlide' };
  return null;
}

/** The nudge of an arrow key on selected shapes: 8 px, 1 px with Ctrl/Cmd, ten times with Shift. */
export function nudgeCommand(key: string, shift: boolean, mod: boolean): EditorCommand | null {
  const step = (shift ? 10 : 1) * 9_525 * (mod ? 1 : 8);
  const nudges: Record<string, [number, number]> = {
    ArrowLeft: [-step, 0],
    ArrowRight: [step, 0],
    ArrowUp: [0, -step],
    ArrowDown: [0, step],
  };
  const d = Object.hasOwn(nudges, key) ? nudges[key] : undefined;
  return d ? { type: 'nudge', dx: d[0], dy: d[1] } : null;
}
