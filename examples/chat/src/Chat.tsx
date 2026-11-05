// The chat column: the thread, the deck card that an assistant message attaches, and the
// composer. The card shows the deck's first slide live from the editor, so an edit in the
// panel shows in the chat too.

import {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type FormEvent,
  type KeyboardEvent,
} from 'react';
import { Icon, useEditorState, usePptxEditor } from '@pptx-editor.dev/react';
import { ARROW_UPWARD, SPARK } from './icons.ts';
import { TEXT } from './text.ts';

export interface Message {
  readonly id: number;
  readonly role: 'user' | 'assistant';
  readonly text: string;
  /** The assistant attached the deck under its text. */
  readonly deck?: boolean;
  /** The text streams in; false for messages that were there before the page loaded. */
  readonly stream?: boolean;
}

const reducedMotion = (): boolean => matchMedia('(prefers-reduced-motion: reduce)').matches;

/** Reveal `text` a few words at a time, as a model streams it; `done` runs at the end. */
function useStreamed(text: string, stream: boolean, done: () => void): string {
  const [shown, setShown] = useState(stream && !reducedMotion() ? 0 : text.length);
  const finished = useRef(done);
  finished.current = done;
  useEffect(() => {
    if (shown >= text.length) {
      finished.current();
      return;
    }
    // Break at spaces, so a word never shows half-typed.
    const timer = setTimeout(() => {
      const next = text.indexOf(' ', shown + 6);
      setShown(next === -1 ? text.length : next);
    }, 28);
    return () => clearTimeout(timer);
  }, [shown, text]);
  return text.slice(0, shown);
}

/** The deck's first slide, painted by the editor and repainted when it changes. */
function DeckThumb() {
  const editor = usePptxEditor();
  // The key changes whenever the slide may render differently.
  const first = useEditorState((s) => s.slides[0]?.key);
  const ratio = useEditorState((s) => (s.slideWidth ? s.slideHeight / s.slideWidth : 9 / 16));
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (editor && first && box.current) editor.renderThumbnail(0, box.current);
  }, [editor, first]);
  return (
    <div
      ref={box}
      className="card__thumb"
      data-ready={first ? true : undefined}
      style={{ aspectRatio: String(1 / ratio) }}
    />
  );
}

export const DeckCard = forwardRef<
  HTMLButtonElement,
  { title: string; open: boolean; onOpen: () => void }
>(function DeckCard({ title, open, onOpen }, ref) {
  const count = useEditorState((s) => s.slides.length);
  const edited = useEditorState((s) => s.canUndo);
  return (
    <button
      ref={ref}
      type="button"
      className="card"
      data-open={open || undefined}
      aria-label={TEXT.openDeckLabel(title)}
      aria-expanded={open}
      onClick={onOpen}
    >
      <DeckThumb />
      <span className="card__foot">
        <span className="card__body">
          <span className="card__title">{title}</span>
          <span className="card__meta">
            {count ? TEXT.slides(count) : TEXT.presentation}
            {edited ? ` · ${TEXT.edited}` : ''}
          </span>
        </span>
        <span className="card__cta">{open ? TEXT.deckOpen : TEXT.openDeck}</span>
      </span>
    </button>
  );
});

const MessageView = memo(function MessageView({
  message,
  deckTitle,
  deckOpen,
  onOpenDeck,
  onStreamed,
  cardRef,
}: {
  message: Message;
  deckTitle: string;
  deckOpen: boolean;
  onOpenDeck: () => void;
  onStreamed: (id: number) => void;
  cardRef: React.Ref<HTMLButtonElement>;
}) {
  const [streamed, setStreamed] = useState(!message.stream);
  const text = useStreamed(message.text, !!message.stream, () => {
    if (streamed) return;
    setStreamed(true);
    onStreamed(message.id);
  });
  if (message.role === 'user')
    return (
      <div className="msg msg--user">
        <span className="msg__avatar msg__avatar--user" aria-hidden="true">
          {TEXT.youInitial}
        </span>
        <div className="msg__content">
          <span className="msg__name">{TEXT.you}</span>
          <p className="msg__text">{message.text}</p>
        </div>
      </div>
    );
  return (
    <div className="msg msg--assistant">
      <BrandMark />
      <div className="msg__content">
        <span className="msg__name">{TEXT.appName}</span>
        <p className="msg__text">{text}</p>
        {message.deck && streamed ? (
          <DeckCard ref={cardRef} title={deckTitle} open={deckOpen} onOpen={onOpenDeck} />
        ) : null}
      </div>
    </div>
  );
});

/** The assistant's mark: a square with the spark, in the accent color. */
export function BrandMark({ busy = false }: { busy?: boolean }) {
  return (
    <span className="brand-mark" data-busy={busy || undefined} aria-hidden="true">
      <Icon paths={SPARK} className="" />
    </span>
  );
}

export function Thinking({ label }: { label: string }) {
  return (
    <div className="msg msg--assistant" role="status">
      <BrandMark busy />
      <div className="msg__content">
        <span className="msg__name">{TEXT.appName}</span>
        <p className="msg__text msg__thinking">{label}</p>
      </div>
    </div>
  );
}

export function Thread({
  messages,
  thinking,
  deckTitle,
  deckOpen,
  onOpenDeck,
  onStreamed,
  cardRef,
}: {
  messages: readonly Message[];
  thinking: string | null;
  deckTitle: string;
  deckOpen: boolean;
  onOpenDeck: () => void;
  onStreamed: (id: number) => void;
  cardRef: React.Ref<HTMLButtonElement>;
}) {
  const scroller = useRef<HTMLDivElement>(null);
  const pinned = useRef(true);
  // A screen reader hears each reply once, whole, when it has streamed in.
  const [spoken, setSpoken] = useState('');
  const streamed = useCallback(
    (id: number) => {
      setSpoken(messages.find((m) => m.id === id)?.text ?? '');
      onStreamed(id);
    },
    [messages, onStreamed]
  );
  // The thread follows new text while the reader is at the bottom, and stays put once they
  // scroll up to read.
  useLayoutEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const follow = () => {
      if (pinned.current) el.scrollTop = el.scrollHeight;
    };
    follow();
    const observer = new ResizeObserver(follow);
    for (const child of Array.from(el.children)) observer.observe(child);
    return () => observer.disconnect();
  }, [messages.length, thinking]);
  // The newest deck card is the one that opens the panel.
  const lastDeck = messages.reduce((at, m, i) => (m.deck ? i : at), -1);
  return (
    <div
      ref={scroller}
      className="thread"
      onScroll={(e) => {
        const el = e.currentTarget;
        pinned.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48;
      }}
    >
      <div className="thread__inner">
        {messages.map((m, i) => (
          <MessageView
            key={m.id}
            message={m}
            deckTitle={deckTitle}
            deckOpen={deckOpen && i === lastDeck}
            onOpenDeck={onOpenDeck}
            onStreamed={streamed}
            cardRef={i === lastDeck ? cardRef : null}
          />
        ))}
        {thinking ? <Thinking label={thinking} /> : null}
        <p className="sr-only" role="status">
          {spoken}
        </p>
      </div>
    </div>
  );
}

export function EmptyState({ onPick }: { onPick: (text: string) => void }) {
  return (
    <div className="empty">
      <BrandMark />
      <h1 className="empty__title">{TEXT.greeting}</h1>
      <p className="empty__lead">{TEXT.greetingLead}</p>
      <div className="empty__ideas">
        {TEXT.suggestions.map((s, i) => (
          <button
            key={s}
            type="button"
            className="idea"
            style={{ animationDelay: `${i * 50}ms` }}
            onClick={() => onPick(s)}
          >
            <span className="idea__text">{s}</span>
            <Icon name="keyboard_arrow_right" className="" />
          </button>
        ))}
      </div>
    </div>
  );
}

export const Composer = forwardRef<
  HTMLTextAreaElement,
  { busy: boolean; onSend: (text: string) => void }
>(function Composer({ busy, onSend }, ref) {
  const [value, setValue] = useState('');
  const ready = value.trim().length > 0 && !busy;
  const submit = (e?: FormEvent) => {
    e?.preventDefault();
    if (!ready) return;
    onSend(value.trim());
    setValue('');
  };
  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    // Enter sends; Shift+Enter breaks the line; an IME composition keeps its Enter.
    if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) submit(e);
  };
  return (
    <form className="composer" onSubmit={submit}>
      <div className="composer__box">
        <textarea
          ref={ref}
          className="composer__input"
          rows={1}
          value={value}
          placeholder={TEXT.composerPlaceholder}
          aria-label={TEXT.composerPlaceholder}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={onKeyDown}
        />
        <div className="composer__bar">
          <span className="composer__hint">{TEXT.composerHint}</span>
          <button type="submit" className="composer__send" disabled={!ready}>
            {TEXT.send}
            <Icon paths={ARROW_UPWARD} className="" />
          </button>
        </div>
      </div>
      <p className="composer__note">{TEXT.disclaimer}</p>
    </form>
  );
});
