/** @jsxImportSource react */
// The find bar: find and replace over the engine's `editor.find`. Ctrl/Cmd+F opens it to find
// and Ctrl/Cmd+H (or Ctrl/Cmd+Shift+H, which macOS leaves to the page) to replace. Enter and
// Shift+Enter step through the matches; Escape closes it and returns the focus.
//
// The toolbar's `find` group and a host's menu open it with `requestFindBar`: an event that
// names the editor, so a bar answers only its own editor (as the link popover does). The
// requests, keys, counter, actions, and focus rules are core's (`overlays.ts`), shared with the
// Vue adapter's `PptxEditorFindBar`.

import {
  dispatchFindRequest,
  findAsYouType,
  findBarOpener,
  findBarView,
  findFieldTakesFocus,
  focusAfterFindBar,
  listenFindRequests,
  onFindShown,
  onNewDocument,
  reserveFindBarSpace,
  runFindBarAction,
  type FindBarAction,
} from '@pptx-editor.dev/core/editor';
import { useCallback, useEffect, useRef, useState, type KeyboardEvent } from 'react';
import type { FindOptions, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { usePopupEntry } from './popups.tsx';
import { useRegisterFindBar } from './chrome-actions.tsx';

/** Open the find bar of `editor`: to find, or with the replace row when `replace`. @public */
export function requestFindBar(editor: PptxEditorInstance | null, replace = false): void {
  dispatchFindRequest(editor, replace);
}

/** The props of `FindBar`. @public */
export interface FindBarProps {
  /** Appended after the `pptx-find` class. */
  className?: string;
}

/** The find and replace bar. Render one inside the editor's viewport, over the canvas. @public */
export function FindBar({ className }: FindBarProps) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  // Find rows and controls are on only while a find bar can answer them.
  useRegisterFindBar();
  // `popups.find`: a renderer replaces the packaged bar here; `false` turns find off.
  const entry = usePopupEntry('find');
  const off = entry === false;
  const root = useRef<HTMLDivElement>(null);
  // Finds the editor scope when a host renderer replaces the packaged bar.
  const anchor = useRef<HTMLSpanElement>(null);
  const findInput = useRef<HTMLInputElement>(null);
  const returnTo = useRef<HTMLElement | null>(null);
  const [open, setOpen] = useState(false);
  const [replaceShown, setReplaceShown] = useState(false);
  const [focusCount, setFocusCount] = useState(0);
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  // The counter reads the deck and the selection: render again when either moves.
  useEditorState((s) => s.revision);
  useEditorState((s) => s.selection);
  useEditorState((s) => s.currentSlide);
  // A step inside the speaker notes moves neither: find tells when it shows another match.
  const [, setShownCount] = useState(0);
  useEffect(
    () => (editor ? onFindShown(editor.find, () => setShownCount((n) => n + 1)) : undefined),
    [editor]
  );
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const options: FindOptions = { matchCase, wholeWord };

  const show = useCallback(
    (replace: boolean) => {
      if (!open) {
        returnTo.current = findBarOpener();
        // A fresh Find shows only the find row; Replace opens both.
        setReplaceShown(replace);
      } else if (replace) setReplaceShown(true);
      setOpen(true);
      setFocusCount((n) => n + 1);
    },
    [open]
  );

  useEffect(() => {
    if (focusCount === 0) return;
    findInput.current?.focus();
    findInput.current?.select();
  }, [focusCount]);

  useEffect(() => {
    if (off || !editor) return;
    // The request for this editor, and Ctrl/Cmd+F and Ctrl/Cmd+H while the focus is in this
    // editor (or nowhere).
    return listenFindRequests(editor, () => root.current ?? anchor.current, show);
  }, [editor, show, off]);

  // The open bar keeps its height free above the slide (core's `reserveFindBarSpace`).
  const packaged = !off && typeof entry !== 'function';
  useEffect(() => (packaged ? reserveFindBarSpace(root.current) : undefined), [packaged]);

  // Another document: the bar closes and forgets the search, which names the old deck's text.
  useEffect(() => {
    if (!editor) return;
    return onNewDocument(editor, () => {
      setOpen(false);
      setQuery('');
      setReplacement('');
      setStatus(null);
    });
  }, [editor]);

  const run = (action: FindBarAction): void => {
    if (!editor || !query) return;
    setStatus(runFindBarAction(t, editor, action, query, replacement, options));
    // The buttons turn off with no match left: focus stays in the bar, in the find field.
    if (findFieldTakesFocus(editor, action, query, options)) findInput.current?.focus();
  };
  const step = (forward: boolean): void => run(forward ? 'next' : 'previous');
  const replaceOne = (): void => run('replace');
  const replaceAll = (): void => run('replaceAll');

  const close = (): void => {
    setOpen(false);
    setStatus(null);
    const back = returnTo.current;
    returnTo.current = null;
    focusAfterFindBar(editor, root.current, root.current ?? anchor.current, back);
  };

  const onBarKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    e.stopPropagation();
    close();
  };

  const { matches, counter, inNotes } = findBarView(t, editor, open, query, options);

  if (off) return null;
  if (typeof entry === 'function')
    return (
      <>
        <span ref={anchor} hidden />
        {open ? entry({ replace: replaceShown, onClose: close }) : null}
      </>
    );
  return (
    <div
      ref={root}
      className={`pptx-find${className ? ` ${className}` : ''}`}
      role="search"
      aria-label={t('find.label')}
      hidden={!open}
      onKeyDown={onBarKey}
    >
      {open ? (
        <>
          <div className="pptx-find__row">
            <ToolButton
              label={t('find.showReplace')}
              icon={replaceShown ? 'keyboard_arrow_up' : 'keyboard_arrow_down'}
              pressed={replaceShown}
              onClick={() => setReplaceShown((v) => !v)}
            />
            <input
              ref={findInput}
              className="pptx-find__input"
              type="text"
              aria-label={t('find.find')}
              placeholder={t('find.findPlaceholder')}
              value={query}
              onChange={(e) => {
                setQuery(e.target.value);
                setStatus(null);
                findAsYouType(editor, e.target.value, options);
              }}
              onKeyDown={(e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                step(!e.shiftKey);
              }}
            />
            <span className="pptx-find__count" aria-live="polite">
              {counter}
            </span>
            <ToolButton
              label={t('find.previous')}
              icon="keyboard_arrow_up"
              disabled={!matches.length}
              onClick={() => step(false)}
            />
            <ToolButton
              label={t('find.next')}
              icon="keyboard_arrow_down"
              disabled={!matches.length}
              onClick={() => step(true)}
            />
            <ToolButton label={t('find.close')} icon="close" onClick={close} />
          </div>
          {replaceShown ? (
            <div className="pptx-find__row pptx-find__row--replace">
              <input
                className="pptx-find__input"
                type="text"
                aria-label={t('find.replace')}
                placeholder={t('find.replacePlaceholder')}
                value={replacement}
                disabled={!editable}
                onChange={(e) => setReplacement(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key !== 'Enter') return;
                  e.preventDefault();
                  replaceOne();
                }}
              />
              <button
                type="button"
                className="pptx-button"
                disabled={!editable || !matches.length}
                onClick={replaceOne}
              >
                {t('find.replaceOne')}
              </button>
              <button
                type="button"
                className="pptx-button"
                disabled={!editable || !matches.length}
                onClick={replaceAll}
              >
                {t('find.replaceAll')}
              </button>
            </div>
          ) : null}
          <div className="pptx-find__options">
            <label className="pptx-find__option">
              <input
                type="checkbox"
                checked={matchCase}
                onChange={(e) => setMatchCase(e.target.checked)}
              />
              {t('find.matchCase')}
            </label>
            <label className="pptx-find__option">
              <input
                type="checkbox"
                checked={wholeWord}
                onChange={(e) => setWholeWord(e.target.checked)}
              />
              {t('find.wholeWord')}
            </label>
          </div>
          {inNotes ? <p className="pptx-find__hint">{t('find.inNotes')}</p> : null}
          {status ? (
            <p className="pptx-find__status" role="status">
              {status}
            </p>
          ) : null}
        </>
      ) : null}
    </div>
  );
}
