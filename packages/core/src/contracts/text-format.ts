// The text format types of the public contract: the Paragraph, Font, and Bullets and Numbering
// dialogs' fields and reads, Change Case, Text Direction, and Columns.
// Split from `index.ts`, which re-exports them, for the line cap.

import type { ColorChoice } from './neutral.ts';
import type { ClipboardNode } from './clipboard.ts';

/**
 * Numbering schemes the editor sets: the `ST_TextAutonumberScheme` values that layout numbers
 * exactly. Arabic, Roman, and Latin letters with a period, a right parenthesis, or both
 * parentheses, and Arabic numbers with nothing (`arabicPlain`; the schema has no other plain
 * scheme). @public
 */
export type AutonumberScheme =
  | `${'arabic' | 'romanUc' | 'romanLc' | 'alphaUc' | 'alphaLc'}${'Period' | 'ParenR' | 'ParenBoth'}`
  | 'arabicPlain';

/**
 * A bullet format: a bullet character in a font (`a:buChar`, `a:buFont`), or a numbering scheme
 * (`a:buAutoNum`), with its size and color. An omitted field keeps each paragraph's value, so a
 * command can change only the size, the color, or the start number. With no `kind`, only the
 * size and color change, on the paragraphs that show a list, and each keeps its own bullet or
 * number. A paragraph with no list takes the bullet its placeholder inherits, or the
 * first gallery item; one with the other kind takes the first gallery item for an omitted
 * character or scheme. `null` makes the bullet follow the text (`a:buSzTx`, `a:buClrTx`), also
 * over a list style that sets a value. @public
 */
export type BulletFormat =
  | {
      readonly kind: 'bullet';
      /**
       * One character. In a symbol font it is the font's own code point (`ü` in Wingdings is a
       * check mark), as `a:buChar/@char` stores it.
       */
      readonly char?: string;
      /**
       * The bullet font: a family or a theme font (`+mn-lt`). With a character and no font, the
       * bullet follows the text font (`a:buFontTx`).
       */
      readonly font?: string;
      /** Percent of the text size, 25 to 400 (`a:buSzPct`). */
      readonly sizePercent?: number | null;
      readonly color?: ColorChoice | null;
    }
  | {
      readonly kind: 'number';
      readonly scheme?: AutonumberScheme;
      /**
       * The first number, 1 to 32767. Absent keeps each paragraph's start; a paragraph with no
       * numbering starts at 1.
       */
      readonly startAt?: number;
      /** Percent of the text size, 25 to 400 (`a:buSzPct`). */
      readonly sizePercent?: number | null;
      readonly color?: ColorChoice | null;
    }
  | {
      /**
       * A picture bullet (`a:buBlip`, ECMA-376 Part 1, 21.1.2.4.2): the picture stands on the
       * baseline, 0.7 times the bullet size high, and keeps its aspect.
       */
      readonly kind: 'picture';
      /** PNG, JPEG, or GIF bytes, at most 50 MB. The command adds them as one media part. */
      readonly picture: Uint8Array;
      /** Percent of the text size, 25 to 400 (`a:buSzPct`). */
      readonly sizePercent?: number | null;
    }
  | {
      readonly kind?: undefined;
      /** Percent of the text size, 25 to 400 (`a:buSzPct`). */
      readonly sizePercent?: number | null;
      readonly color?: ColorChoice | null;
    };

/**
 * The bullet format of the selected paragraphs, as the gallery and dialog show it: resolved
 * values. @public
 */
export type BulletFormatRead =
  | {
      readonly kind: 'bullet';
      readonly char: string;
      readonly font?: string;
      /** Percent of the text size. */
      readonly sizePercent: number;
      /** `#RRGGBB`, or absent when the bullet follows the text color. */
      readonly color?: string;
    }
  | {
      readonly kind: 'number';
      /** The scheme as the file names it; it can be one the editor does not set. */
      readonly scheme: string;
      readonly startAt: number;
      /** Percent of the text size. */
      readonly sizePercent: number;
      readonly color?: string;
    }
  | {
      /** A picture bullet (`a:buBlip`). */
      readonly kind: 'picture';
      /** Percent of the text size. */
      readonly sizePercent: number;
    };

/**
 * What the selected paragraphs share when their bullet formats differ, so the Bullets and
 * Numbering dialog opens on their list's tab and shows the values they have in common. @public
 */
export interface BulletFormatMixed {
  /** The list kind of each selected paragraph, each kind once, in selection order. */
  readonly kinds: readonly ('bullet' | 'number' | 'picture' | 'none')[];
  /** Percent of the text size, when every bullet and number has the same; otherwise null. */
  readonly sizePercent: number | null;
  /**
   * The first number, when every numbered paragraph has the same; otherwise null. Absent when no
   * paragraph is numbered.
   */
  readonly startAt?: number | null;
}

/**
 * The orientation of a shape's text (Home > Text Direction), in the words of Office.js's
 * the Office.js `ShapeTextOrientation` enum:`horizontal`, rotated 90° (`vertical`, file
 * `vert`) or 270° (`vertical270`), stacked letters (`wordArtVertical`), upright East Asian text
 * (`eastAsianVertical`, file `eaVert`), and the two the editor reads but does not set
 * (`mongolianVertical`, `wordArtVerticalRtl`). Not DOCX Editor's `TextDirection`, which is
 * bidi. @public
 */
export type TextOrientation =
  | 'horizontal'
  | 'vertical'
  | 'vertical270'
  | 'wordArtVertical'
  | 'eastAsianVertical'
  | 'mongolianVertical'
  | 'wordArtVerticalRtl';

/** Text columns of a shape. @public */
export interface TextColumns {
  /** 1 to 16. */
  readonly count: number;
  /** The gap between columns, EMU (DOCX Editor: `gapTwips`). */
  readonly gapEmu: number;
}

/**
 * Change Case choices: "Sentence case.", "lowercase", "UPPERCASE", "Capitalize Each Word", and
 * "tOGGLE cASE". @public
 */
export type TextCase = 'sentence' | 'lower' | 'upper' | 'capitalize' | 'toggle';

/**
 * Underline styles: `ST_TextUnderlineType` in words, as Office.js names them (`single` is `sng`,
 * `double` is `dbl`, `wavyDouble` is `wavyDbl`). @public
 */
export type UnderlineStyle =
  | 'none'
  | 'words'
  | 'single'
  | 'double'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDouble';

/** Strikethrough: none, single, or double (`ST_TextStrikeType`). @public */
export type StrikeStyle = 'none' | 'single' | 'double';

/** Capitals: none, small caps, or all caps (`ST_TextCapsType`). @public */
export type CapsStyle = 'none' | 'small' | 'all';

/**
 * The fields of the Font dialog (Font and Character Spacing tabs), as `setFontFormat`
 * takes them. An omitted field keeps each run's value. The same shape as
 * `ParagraphFormatCommand`. @public
 */
export interface FontFormatCommand {
  /** The Latin font (`a:latin`): a family name of 1 to 64 letters, digits, spaces, `-._+`. */
  readonly fontFamily?: string;
  /** Size in points, 1 to 4000 (`sz`). */
  readonly fontSizePt?: number;
  /** Bold (`b`). */
  readonly bold?: boolean;
  /** Italic (`i`). */
  readonly italic?: boolean;
  /** A color choice, or null for the inherited color. */
  readonly color?: ColorChoice | null;
  /** The underline style (`u`); `none` removes it. */
  readonly underline?: UnderlineStyle;
  /** Strikethrough (`strike`): none, single, or double. */
  readonly strike?: StrikeStyle;
  /**
   * Offset in percent of the font size, -100 to 100 (`baseline`): positive raises
   * (superscript), negative lowers, 0 turns an inherited offset off.
   */
  readonly baseline?: number;
  /** Small caps or all caps (`cap`). */
  readonly caps?: CapsStyle;
  /** Points added after each character, -1584 to 1584 (`spc`); negative condenses. */
  readonly characterSpacingPt?: number;
  /**
   * Kerning for fonts of this size in points and above, 0.01 to 1584 (`kern`, in hundredths);
   * null turns kerning off, which the engine writes as `kern="0"`.
   */
  readonly kerningPt?: number | null;
  /**
   * The underline color (`a:uFill`), or null for the underline to take the text color
   * (`a:uFillTx`, also over a list style that sets one).
   */
  readonly underlineColor?: ColorChoice | null;
}

/**
 * The selected text as the Font dialog reads it: resolved values. A field is null when the
 * selected runs disagree (`disagrees` says which), so a control can show a blank. @public
 */
export interface FontFormatRead {
  /** The Latin font (`a:latin`), resolved through the theme. */
  readonly fontFamily: string | null;
  /** Points. */
  readonly fontSizePt: number | null;
  /** Bold. */
  readonly bold: boolean | null;
  /** Italic. */
  readonly italic: boolean | null;
  /** The underline style; `none` without an underline. */
  readonly underline: UnderlineStyle | null;
  /** Strikethrough: none, single, or double. */
  readonly strike: StrikeStyle | null;
  /** Percent of the font size; positive is superscript. */
  readonly baseline: number | null;
  /** Capitals: none, small, or all. */
  readonly caps: CapsStyle | null;
  /** Points. */
  readonly characterSpacingPt: number | null;
  /** Points and above; null when kerning is off or the runs disagree (see `disagrees`). */
  readonly kerningPt: number | null;
  /**
   * The underline's own solid color, `#RRGGBB`; null when it takes the text color, when its fill
   * is not solid (a gradient or a pattern), or when the runs disagree.
   */
  readonly underlineColor: string | null;
  /** Which fields differ between the selected runs. */
  readonly disagrees: {
    readonly fontFamily: boolean;
    readonly fontSizePt: boolean;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly underline: boolean;
    readonly strike: boolean;
    readonly baseline: boolean;
    readonly caps: boolean;
    readonly characterSpacingPt: boolean;
    readonly kerningPt: boolean;
    readonly underlineColor: boolean;
  };
}

/** Paragraph alignment as the Paragraph dialog names it (`algn`: `l`, `ctr`, `r`, `just`, `dist`). @public */
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify' | 'distributed';

/**
 * Line spacing as the Paragraph dialog sets it: a `multiple` of single spacing
 * (`a:spcPct`), or `exact` points (`a:spcPts`). DOCX Editor's rule names; `a:lnSpc` has no
 * "at least". @public
 */
export interface ParagraphLineSpacing {
  /** `multiple` of single spacing (`a:spcPct`), or `exact` points (`a:spcPts`). */
  readonly rule: 'multiple' | 'exact';
  /** Lines for `multiple` (1 is single), points for `exact`. */
  readonly value: number;
}

/**
 * The fields of the Paragraph dialog, as `setParagraphFormat` takes them. An omitted
 * field keeps the paragraph's value. `null` removes the paragraph's own value, so it takes the
 * inherited value again (from the layout, the master, or the shape's list style); a zero is a
 * value and blocks the inherited one. @public
 */
export interface ParagraphFormatCommand {
  /** Alignment (`algn`). */
  readonly alignment?: ParagraphAlignment;
  /** Space before the paragraph in points (`a:spcBef/a:spcPts`), 0 to 1584. */
  readonly spaceBeforePt?: number | null;
  /**
   * Space before the paragraph in lines of its text (`a:spcBef/a:spcPct`), 0 to 132. Give
   * this or `spaceBeforePt`, not both.
   */
  readonly spaceBeforeLines?: number | null;
  /** Space after the paragraph in points (`a:spcAft/a:spcPts`), 0 to 1584. */
  readonly spaceAfterPt?: number | null;
  /** Space after the paragraph in lines (`a:spcAft/a:spcPct`), 0 to 132. */
  readonly spaceAfterLines?: number | null;
  /** Line spacing (`a:lnSpc`): a multiple from 0.01 to 9.99, or 0.01 to 1584 points. */
  readonly lineSpacing?: ParagraphLineSpacing | null;
  /** "Before text": the left indent of the text in EMU (`marL`), 0 to 51206400. */
  readonly indentLeftEmu?: number | null;
  /**
   * The first line's offset from the left indent in EMU (`indent`): negative is a hanging
   * indent, positive a first-line indent.
   */
  readonly indentFirstLineEmu?: number | null;
}

/**
 * The selected paragraphs as the Paragraph dialog reads them. A field is null when the
 * paragraphs disagree (`disagrees` says which), so a control can show a blank. Values are
 * resolved: inherited values count. @public
 */
export interface ParagraphFormatRead {
  /** Alignment; null when the paragraphs disagree. */
  readonly alignment: ParagraphAlignment | null;
  /** Points, or null when the space is in lines or the paragraphs disagree. */
  readonly spaceBeforePt: number | null;
  /** Lines, or null when the space is in points or the paragraphs disagree. */
  readonly spaceBeforeLines: number | null;
  /** Points, or null when the space is in lines or the paragraphs disagree. */
  readonly spaceAfterPt: number | null;
  /** Lines, or null when the space is in points or the paragraphs disagree. */
  readonly spaceAfterLines: number | null;
  /** Line spacing; null when the paragraphs disagree. */
  readonly lineSpacing: ParagraphLineSpacing | null;
  /** "Before text", EMU; null when the paragraphs disagree. */
  readonly indentLeftEmu: number | null;
  /** The first line's offset from the left indent, EMU (negative hangs). */
  readonly indentFirstLineEmu: number | null;
  /** Which fields differ between the selected paragraphs. */
  readonly disagrees: {
    readonly alignment: boolean;
    readonly spaceBefore: boolean;
    readonly spaceAfter: boolean;
    readonly lineSpacing: boolean;
    readonly indentLeft: boolean;
    readonly indentFirstLine: boolean;
  };
}

/**
 * The text format commands of `EditorCommand`: the Paragraph, Font, and Bullets and Numbering
 * dialogs, superscript and subscript, Change Case, character spacing, Text Direction, and
 * Columns. Each acts on the selected text, or on all text of the selected shapes. @public
 */
export type TextFormatCommand =
  /**
   * Format > Paragraph: alignment, indents, spacing before and after, and line spacing on every
   * selected paragraph, as one undo step. DOCX Editor's command name; the fields follow its
   * `ParagraphFormatCommand`, in PresentationML units (EMU indents, spacing in points or lines).
   */
  | ({ readonly type: 'setParagraphFormat' } & ParagraphFormatCommand)
  /** Superscript on or off: the button writes `baseline` 30 % up. */
  | { readonly type: 'toggleSuperscript' }
  /** Subscript on or off: the button writes `baseline` 25 % down. */
  | { readonly type: 'toggleSubscript' }
  /** Home > Change Case: rewrite the letters of the selected text. Formatting stays. */
  | { readonly type: 'changeCase'; readonly case: TextCase }
  /** Home > Character Spacing: points added after each character (`spc`); 0 is Normal. */
  | { readonly type: 'setCharacterSpacing'; readonly points: number }
  /** Format > Font: the fields of the Font dialog, as one undo step. */
  | ({ readonly type: 'setFontFormat' } & FontFormatCommand)
  /**
   * Home > Text Direction (`a:bodyPr/@vert`) of the selected shapes. The type leaves out
   * `mongolianVertical` and `wordArtVerticalRtl`, which layout does not draw correctly;
   * the engine refuses them too.
   */
  | {
      readonly type: 'setTextDirection';
      readonly direction: Exclude<TextOrientation, 'mongolianVertical' | 'wordArtVerticalRtl'>;
    }
  /**
   * Home > Columns: the number of text columns (1 to 16, `numCol`) of the selected shapes, and
   * the gap between them in EMU (`spcCol`, 0 to 51206400). An absent `gapEmu` keeps the gap.
   */
  | { readonly type: 'setColumns'; readonly count: number; readonly gapEmu?: number }
  /**
   * Home > Bullets or Numbering (a gallery item), or Format > Bullets and Numbering: the bullet
   * format of the selected paragraphs (Office.js `ParagraphFormat.bulletFormat`).
   * `null` is the gallery's None: the list turns off.
   */
  | { readonly type: 'setBulletFormat'; readonly format: BulletFormat | null }
  /**
   * Paste formatted text: DrawingML paragraphs (`a:p`, as `ClipboardNode`s) replace the selected
   * text, or, with no text selected, go in a new text box in the middle of the slide. The
   * paragraphs are checked as untrusted input: runs, breaks, and their formatting are kept, and
   * links, pictures, and relationship references are left out. A field becomes plain text.
   */
  | { readonly type: 'pasteText'; readonly paragraphs: readonly ClipboardNode[] };
