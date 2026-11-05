// The scripted assistant. A real host calls its model here; the shape of the answer stays the
// same: text to stream, an optional deck to attach, and an optional action on the open editor.
// The actions use only the public editor API (`goToSlide`, the slide show, `save`).

export type AssistantAction =
  | { readonly kind: 'goTo'; readonly slide: number }
  | { readonly kind: 'present' }
  | { readonly kind: 'download' }
  | { readonly kind: 'open' };

export interface AssistantReply {
  readonly text: string;
  /** Attach a new deck card under the text. */
  readonly deck?: boolean;
  readonly action?: AssistantAction;
}

/** What the assistant knows about the editor when it answers. */
export interface DeckContext {
  readonly hasDeck: boolean;
  readonly slideCount: number;
  readonly currentSlide: number;
}

export const FIRST_DECK_TEXT =
  'Here is a first draft. It opens in the editor next to this chat, so you can change any text, shape, or picture in place. Ask me to jump to a slide, present, or download it.';

export function reply(input: string, deck: DeckContext): AssistantReply {
  const text = input.toLowerCase();
  if (!deck.hasDeck) return { text: FIRST_DECK_TEXT, deck: true };
  const slide = /slide\s*(\d+)/.exec(text);
  if (slide) {
    const n = Number(slide[1]);
    if (n < 1 || n > deck.slideCount)
      return {
        text: `The deck has ${deck.slideCount} slides. Pick a number from 1 to ${deck.slideCount}.`,
      };
    return {
      text: `Slide ${n} is open in the editor. Click any text on it to edit.`,
      action: { kind: 'goTo', slide: n - 1 },
    };
  }
  if (/present|slide ?show|full ?screen/.test(text))
    return {
      text: 'Starting the slide show. Press Escape to come back.',
      action: { kind: 'present' },
    };
  if (/download|export|save|pptx/.test(text))
    return {
      text: 'Here is the deck as a .pptx file, with every edit you made.',
      action: { kind: 'download' },
    };
  if (/open|edit|show/.test(text))
    return { text: 'The deck is open next to the chat.', action: { kind: 'open' } };
  return {
    text: `You are on slide ${deck.currentSlide + 1} of ${deck.slideCount}. This mock does not rewrite slides, but the editor does: click a text box and type. You can also say “show slide 2”, “present”, or “download”.`,
  };
}
