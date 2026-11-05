// A chat app with a deck panel, as a host would build it: the chat on the left, and the
// editor in a panel that slides in from the right when you open the deck card. One
// `PptxEditor.Root` wraps both, so the chat reads the deck (the live card thumbnail) and acts
// on it (go to a slide, present, download) through the public API.

import { useCallback, useEffect, useRef, useState, type CSSProperties } from 'react';
import { flushSync } from 'react-dom';
import { packagedFonts } from '@docx-editor.dev/fonts';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { Icon, PptxEditor, useFonts, usePptxEditor, useSlideshow } from '@pptx-editor.dev/react';
import { exportDeck, fetchDeck } from '../../shared/demo-files.ts';
import { FIRST_DECK_TEXT, reply, type AssistantAction } from './assistant.ts';
import { Composer, EmptyState, Thread, type Message } from './Chat.tsx';
import { DeckPanel } from './DeckPanel.tsx';
import { usePanelLayout } from './panel-layout.ts';
import { EDIT_SQUARE } from './icons.ts';
import { TEXT } from './text.ts';

const base = import.meta.env.BASE_URL;
const DECK_FILE = 'showcase.pptx';

/** The conversation that is there when the page opens, so the deck card shows at once. */
const SEED: readonly Message[] = [
  { id: 1, role: 'user', text: TEXT.seedPrompt },
  { id: 2, role: 'assistant', text: FIRST_DECK_TEXT, deck: true },
];

function ChatApp({ loadDeck }: { loadDeck: () => void }) {
  const editor = usePptxEditor();
  const show = useSlideshow();
  const [messages, setMessages] = useState<readonly Message[]>(SEED);
  const [chatTitle, setChatTitle] = useState<string>(TEXT.seedTitle);
  const [thinking, setThinking] = useState<string | null>(null);
  const [streaming, setStreaming] = useState(false);
  const layout = usePanelLayout();
  const { open, setOpen } = layout;
  const card = useRef<HTMLButtonElement>(null);
  const composer = useRef<HTMLTextAreaElement>(null);
  const panel = useRef<HTMLElement>(null);
  const hasDeck = messages.some((m) => m.deck);
  const nextId = useRef(SEED.length + 1);
  /** The reply that is on its way: New chat cancels it, so it never lands in the new chat. */
  const pending = useRef<number | null>(null);
  const cancelReply = useCallback(() => {
    if (pending.current !== null) window.clearTimeout(pending.current);
    pending.current = null;
  }, []);
  useEffect(() => cancelReply, [cancelReply]);

  const closePanel = useCallback(() => {
    // Focus inside the panel goes back to the card that opened it.
    const active = document.activeElement;
    const inside = !active || active === document.body || active.closest('.shell__panel');
    // Render the close first: an expanded panel or a sheet keeps the chat inert until then,
    // and an inert element does not take focus.
    flushSync(() => setOpen(false));
    if (inside) (card.current ?? composer.current)?.focus({ preventScroll: true });
  }, [setOpen]);

  const download = useCallback(() => exportDeck(editor, TEXT.deckTitle), [editor]);

  const run = useCallback(
    (action: AssistantAction | undefined) => {
      if (!action) return;
      if (action.kind === 'goTo') {
        setOpen(true);
        editor?.goToSlide(action.slide);
      } else if (action.kind === 'present') show.start('current');
      else if (action.kind === 'download') download();
      else setOpen(true);
    },
    [editor, show, download, setOpen]
  );

  const send = useCallback(
    (text: string) => {
      const snapshot = editor?.snapshot();
      const answer = reply(text, {
        hasDeck,
        slideCount: snapshot?.slides.length ?? 0,
        currentSlide: snapshot?.currentSlide ?? 0,
      });
      if (!messages.length) setChatTitle(text.length > 34 ? `${text.slice(0, 32)}…` : text);
      setMessages((m) => [...m, { id: nextId.current++, role: 'user', text }]);
      setThinking(answer.deck ? TEXT.thinking : TEXT.working);
      setStreaming(true);
      // A new deck starts from fresh bytes while the assistant "thinks".
      if (answer.deck) loadDeck();
      cancelReply();
      pending.current = window.setTimeout(
        () => {
          pending.current = null;
          setThinking(null);
          setMessages((m) => [
            ...m,
            {
              id: nextId.current++,
              role: 'assistant',
              text: answer.text,
              stream: true,
              ...(answer.deck ? { deck: true } : {}),
            },
          ]);
          run(answer.action);
        },
        answer.deck ? 1100 : 450
      );
    },
    [editor, hasDeck, messages.length, loadDeck, run, cancelReply]
  );

  const onStreamed = useCallback(() => setStreaming(false), []);

  const newChat = useCallback(() => {
    cancelReply();
    setOpen(false);
    setMessages([]);
    setChatTitle(TEXT.newChat);
    setThinking(null);
    setStreaming(false);
    composer.current?.focus();
  }, [setOpen, cancelReply]);

  // A sheet covers the chat, so the focus goes into it: to Back, the way out.
  useEffect(() => {
    if (!open || !layout.sheet) return;
    const frame = requestAnimationFrame(() =>
      panel.current?.querySelector<HTMLElement>('.deck__head button')?.focus()
    );
    return () => cancelAnimationFrame(frame);
  }, [open, layout.sheet]);

  // Expand makes the chat inert: focus that was in it moves to the panel's Show chat button.
  useEffect(() => {
    if (!open || !layout.expanded) return;
    const active = document.activeElement;
    if (active && active !== document.body && !active.closest('.shell__chat')) return;
    panel.current?.querySelector<HTMLElement>('.deck__head [aria-pressed]')?.focus();
  }, [open, layout.expanded]);

  // Escape closes the panel from the chat, the panel header, or the divider. Escape on the
  // slide, in the notes, or in a menu belongs to the editor: it leaves a text box or a group.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape' || e.defaultPrevented) return;
      const target = e.target as Element | null;
      const chrome = target?.closest?.('.deck__head, .shell__handle');
      if (!chrome && target?.closest?.('.shell__panel, .pptx-slideshow')) return;
      closePanel();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, closePanel]);

  const shellStyle = { '--panel-w': `${layout.panelWidth}px` } as CSSProperties;
  return (
    <div
      ref={layout.shell}
      className="shell"
      style={shellStyle}
      data-open={open || undefined}
      data-expanded={layout.expanded || undefined}
      data-sheet={layout.sheet || undefined}
      data-resizing={layout.resizing || undefined}
    >
      <main
        className="shell__chat"
        inert={(open && (layout.expanded || layout.sheet)) || undefined}
      >
        <header className="chat-head">
          <span className="chat-head__brand">{TEXT.appName}</span>
          {messages.length ? (
            <>
              <span className="chat-head__sep" aria-hidden="true">
                /
              </span>
              <span className="chat-head__title">{chatTitle}</span>
            </>
          ) : null}
          <button type="button" className="chat-head__new" onClick={newChat}>
            <Icon paths={EDIT_SQUARE} className="" />
            <span className="chat-head__new-label">{TEXT.newChat}</span>
          </button>
        </header>
        {messages.length ? (
          <Thread
            messages={messages}
            thinking={thinking}
            deckTitle={TEXT.deckTitle}
            deckOpen={open}
            // The card of an open deck moves the focus to the slide, ready to type.
            onOpenDeck={() => (open ? editor?.focus() : setOpen(true))}
            onStreamed={onStreamed}
            cardRef={card}
          />
        ) : (
          <EmptyState onPick={send} />
        )}
        <Composer ref={composer} busy={!!thinking || streaming} onSend={send} />
      </main>
      <div className="shell__spacer" aria-hidden="true" />
      <section
        ref={panel}
        className="shell__panel"
        aria-label={TEXT.deckTitle}
        inert={!open || undefined}
      >
        <div
          className="shell__handle"
          role="separator"
          aria-orientation="vertical"
          aria-label={TEXT.resize}
          aria-valuemin={Math.round(layout.minFraction * 100)}
          aria-valuemax={100}
          aria-valuenow={Math.round(layout.fraction * 100)}
          tabIndex={0}
          {...layout.handleProps}
        />
        {/* Mounted while closed, so the deck is ready the moment the panel opens. */}
        <DeckPanel
          title={TEXT.deckTitle}
          expanded={layout.expanded}
          sheet={layout.sheet}
          targetWidth={layout.panelWidth}
          onClose={closePanel}
          onExpand={layout.toggleExpanded}
          onDownload={download}
        />
      </section>
    </div>
  );
}

export function App() {
  const [bytes, setBytes] = useState<Uint8Array | undefined>();
  const loadDeck = useCallback(() => void fetchDeck(base, DECK_FILE).then(setBytes), []);
  useEffect(loadDeck, [loadDeck]);
  const fonts = useFonts(packagedFonts());
  return (
    <PptxEditor.Root
      {...(bytes ? { document: bytes } : {})}
      colorMode="system"
      fonts={fonts}
      onReady={(instance: PptxEditorInstance) => {
        (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = instance;
      }}
    >
      <ChatApp loadDeck={loadDeck} />
    </PptxEditor.Root>
  );
}
