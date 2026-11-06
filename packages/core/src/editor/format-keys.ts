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

/** Where a format key acts: the text being edited, selected shapes, and the selected text. */
export interface FormatKeyTarget {
  /** Text is being edited. */
  readonly inText: boolean;
  /** Shapes are selected, so paragraph and font keys format all their text. */
  readonly shapes?: boolean;
  /** The selected text, read only when a key needs it (Shift+F3). */
  readonly selectedText?: () => string;
}

const ALIGN_KEYS: Readonly<Record<string, 'l' | 'ctr' | 'r' | 'just'>> = {
  l: 'l',
  e: 'ctr',
  r: 'r',
  j: 'just',
};

/**
 * Shift+F3 steps the case of the selected text: lower case to Capitalize Each Word, that to
 * UPPER CASE, and UPPER CASE back to lower case.
 */
export function nextCase(text: string): 'lower' | 'upper' | 'capitalize' {
  const letters = text.replace(/[^\p{L}]/gu, '');
  if (letters && letters === letters.toUpperCase() && letters !== letters.toLowerCase())
    return 'lower';
  if (letters && letters === letters.toLowerCase()) return 'capitalize';
  return 'upper';
}

/**
 * The format or history command of a key, or null: undo and redo, bold, italic, underline,
 * Clear Formatting (Ctrl/Cmd+\\), copy and paste formatting (Ctrl/Cmd+Shift+C and +V), New Slide
 * (Ctrl+M), and, in text, subscript (Ctrl/Cmd+=) and superscript (Ctrl/Cmd+Shift+=).
 *
 * In text and on selected shapes: Left, Center, Right, and Justify (Ctrl/Cmd+L, E, R, J), and
 * the next or the previous font size (Ctrl/Cmd+Shift+> and <). In text: the list level down or
 * up (Ctrl/Cmd+] and [), and Change Case (Shift+F3). With nothing selected, these keys keep
 * their browser meaning.
 */
export function formatKeyCommand(
  key: string,
  e: KeyMods,
  target: boolean | FormatKeyTarget
): EditorCommand | null {
  const where = typeof target === 'boolean' ? { inText: target } : target;
  const inText = where.inText;
  const formats = inText || !!where.shapes;
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
  if (!formats || e.altKey) return null;
  const align = mod && !e.shiftKey && Object.hasOwn(ALIGN_KEYS, lower) ? ALIGN_KEYS[lower] : null;
  if (align) return { type: 'setAlign', align };
  // Shift turns `.` and `,` into `>` and `<` on most layouts; both spellings are accepted.
  if (mod && e.shiftKey && (key === '>' || key === '.'))
    return { type: 'stepFontSize', direction: 1 };
  if (mod && e.shiftKey && (key === '<' || key === ','))
    return { type: 'stepFontSize', direction: -1 };
  if (!inText) return null;
  if (mod && !e.shiftKey && (key === ']' || key === '['))
    return { type: 'indent', direction: key === ']' ? 1 : -1 };
  if (key === 'F3' && e.shiftKey && !mod)
    return { type: 'changeCase', case: nextCase(where.selectedText?.() ?? '') };
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
