<p align="center">
  <a href="https://pptx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" alt="PPTX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/react"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/react.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/react"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/react.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=C43E1C" alt="license" /></a>
</p>

# @pptx-editor.dev/react

The React adapter for PPTX Editor. `PptxEditor.Root` owns the engine instance, `.Viewport` scopes the styles, and `.Content` attaches the slide surface. The toolbar, slide rail, notes, side panel, and dialogs read the engine through hooks and hold no editing state.

## Install

```bash
npm install @pptx-editor.dev/core @pptx-editor.dev/react
```

## Use

```tsx
import { PptxEditor } from '@pptx-editor.dev/react';
import '@pptx-editor.dev/core/styles/editor.css';

export function App({ bytes }: { bytes: Uint8Array }) {
  return <PptxEditor document={bytes} mode="edit" />;
}
```

To build your own chrome, use `usePptxEditor()`, `useEditorState(selector)`, and `useEditorCommand(slotId)`. As in DOCX Editor, `useEditorCommand(slotId).execute()` returns `true` when the engine ran the command and `false` on a refusal. `disabledReason` is the engine's catalog key, or its own text for a developer reason: show it with `translateKeyOrText(t, reason)` from `@pptx-editor.dev/core/editor`, with `t` from `useTranslation()`.

## Slides and layouts

The `slide` group of the chrome registry holds the slide commands: New Slide, Layout, Reset, and Hide Slide. It is named after the Office.js `Slide` object. It is a contextual group: the preset formatting bar does not draw it, so you place its slots in your own toolbar, ribbon, or menu.

| Slot identifier | Control | Command |
| --- | --- | --- |
| `slide.add` | New slide: a split button; its arrow opens the layout gallery | `addSlide`, with `layoutId` for a picked layout |
| `slide.layout` | Layout: the layout gallery, with the current slide's layout marked | `applyLayout` |
| `slide.reset` | Reset slide | `resetSlides` |
| `slide.hide` | Hide slide: pressed when every selected slide is hidden | `setSlidesHidden` |

The commands act on the slides selected in the rail. A refused command returns the engine's reason, for example `disabled.viewing` in viewing mode or `errors.layoutNotFound` for a layout id that the deck does not have.

```tsx
import { Toolbar } from '@pptx-editor.dev/react';

export function SlidesRibbon() {
  return (
    <Toolbar preset={false}>
      <Toolbar.Slot slotId="slide.add" />
      <Toolbar.Slot slotId="slide.layout" />
      <Toolbar.Slot slotId="slide.reset" />
      <Toolbar.Slot slotId="slide.hide" />
    </Toolbar>
  );
}
```

Each slot takes the customization ladder: `className`, `icon`, `asChild`, `hidden`, and children to replace the control. Hiding a slot changes the interface only: the command still works from the rail and the keyboard.

### Build your own layout picker

`LayoutGallery` draws the deck's layouts in a gallery: one group per slide master, each tile the engine's own render of a new slide on that layout, with its placeholder prompts. Put it in any menu or popup. The tiles are menu rows, and the arrow keys, Home, and End move between them.

| Prop | Type | Default | What it does |
| --- | --- | --- | --- |
| `value` | `string` | none | The layout marked as current; the tiles are then `menuitemradio` rows |
| `onChange` | `(layoutId: string) => void` | required | Called with the picked layout's id |
| `layouts` | `readonly SlideLayoutSummary[]` | every layout of the deck | The layouts to show |
| `disabled` | `boolean` | `true` in viewing mode | Shows the tiles but takes no pick |
| `disabledReason` | `string \| null` | the viewing reason | Why the gallery takes no pick, in the tiles' tooltips: a catalog key, such as `useEditorCommand(slotId).disabledReason`, or text |
| `className` | `string` | none | A class name added to the gallery |

```tsx
import { LayoutGallery, useEditorState, usePptxEditor } from '@pptx-editor.dev/react';

export function LayoutMenu({ onDone }: { onDone: () => void }) {
  const editor = usePptxEditor();
  const current = useEditorState((s) => s.slides[s.currentSlide]?.layoutId);
  return (
    <div role="menu" aria-label="Layout">
      <LayoutGallery
        value={current ?? ''}
        onChange={(layoutId) => {
          editor?.exec({ type: 'applyLayout', layoutId });
          onDone();
        }}
      />
    </div>
  );
}
```

For a picker of your own, read the layouts with `useSlideLayouts()` (the same hook exists in `@pptx-editor.dev/vue`) and paint a tile with `editor.renderLayoutThumbnail(layoutId, element)`.

## Speaker notes

`PptxEditor.NotesPane` shows the current slide's speaker notes under the slide, with their bold, italic, and underline. Type in it to edit the notes: the text is written to the slide when the field loses focus, when the current slide changes, and before a format or an undo. Ctrl/Cmd+B, I, and U format the selected notes text, or the word that has the caret inside it. While the notes field has the cursor, `useEditorCommand('text.bold')` and the other text formats (italic, underline, strikethrough) format the notes, so the toolbar and your own buttons act on the notes too. Ctrl/Cmd+Z and Shift+Ctrl/Cmd+Z undo and redo the deck's edits, notes included.

| Prop | Type | Default | What it does |
| --- | --- | --- | --- |
| `className` | `string` | none | A class name added to the pane |
| `showFormatButtons` | `boolean` | `true` | Shows the pane's bold, italic, and underline buttons; the keys work either way |

```tsx
import { PptxEditor } from '@pptx-editor.dev/react';

export function Notes() {
  return <PptxEditor.NotesPane className="acme-notes" showFormatButtons={false} />;
}
```

In your own notes chrome, run the engine commands: `setNotes` writes a slide's notes, and `formatNotes` toggles a format on a range of them as one undo step. The offsets count the characters of `snapshot().notes`, with `\n` between paragraphs.

```ts
editor.exec({ type: 'formatNotes', index: 0, start: 0, end: 4, format: 'bold' });
```

## Sections and the slide sorter

The rail and the slide sorter group slides under their sections (`p14:sectionLst`) when the deck has them, as a tree: a section header collapses its slides (the arrow, or Left and Right), selects them (a click on the name, or Enter), renames in place (double-click or F2), and opens the section menu (Shift+F10). Which sections are collapsed is a view setting, `collapsedSections` in the snapshot: the rail and the slide sorter share it, and it is never saved.

The `section` group holds the section commands (`p14:sectionLst`), one slot per row, so a ribbon can show it as one menu and a host can place one row. The group is named after the section object, and the rows act on the current slide's section.

| Slot identifier | Control | Command |
| --- | --- | --- |
| `section.add` | Add section: adds a section at the current slide and starts renaming it | `addSection` |
| `section.rename` | Rename section: opens the name field in the rail or the slide sorter, whose Enter runs `renameSection` | none: `requestSectionRename` |
| `section.remove` | Remove section: its slides join the section before it | `removeSection` |
| `section.removeAll` | Remove all sections: the slides stay | `removeAllSections` |
| `section.collapseAll`, `section.expandAll` | Collapse all, Expand all | `setSectionsCollapsed` |
| `view.presentationView` | Normal or Slide sorter | `setPresentationView` |

Without sections, every row but Add section is off, with the reason `disabled.noSections`. Rename has no engine command of its own: `section.rename` opens the name field in the rail or the slide sorter, and `section.add` opens it for the new section. Both work the same from the toolbar part, `Menu.Item`, `ContextMenu.Slot`, and `useEditorCommand`. In your own chrome, start a rename with `requestSectionRename(editor, sectionId)`, or run `renameSection` with a name. An open rail or a sorter takes the request also in a deck without sections, so Add section asks for the name there too. Without a mounted rail or sorter, `section.rename` is off with the reason `disabled.noSectionPane`, Add section adds the section without the field, `runToolbarCommand(editor, 'section.rename')` returns a reason that names `requestSectionRename` and prints a development warning, and so does a request that no rail or sorter takes. Only the packaged rail and slide sorter take rename requests; a public point where your own pane registers to take them is a later change. The command slots take the whole customization ladder; `view.presentationView` is a menu, so it takes `className`, `hidden`, and children, not `icon` or `asChild`.

The section menu (right-click a section header, or Shift+F10 on it) is the `sectionContextMenu` popup entry, and the slide menu is the `railContextMenu` entry. Both cover the rail and the slide sorter. `false` shows no menu; a function renders your own menu, or the packaged `SectionContextMenu` or `RailContextMenu` with its props. The section and slide commands keep working from the API.

`PptxEditor.SlideSorter` takes one prop:

| Prop        | Type     | Default | What it does                     |
| ----------- | -------- | ------- | -------------------------------- |
| `className` | `string` | none    | A class name added to the sorter |

`setPresentationView` changes a view setting only: it makes no undo step and is never saved. The default `PptxEditor` swaps its rail, slide, and notes for the slide sorter when the view is `slideSorter`. When you compose the editor yourself, make the same switch. If the view becomes `slideSorter` and no `PptxEditor.SlideSorter` of the editor is mounted, a development warning says so:

```tsx
import { PptxEditor, useEditorState } from '@pptx-editor.dev/react';

export function Workspace() {
  const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
  return sorter ? (
    <PptxEditor.SlideSorter />
  ) : (
    <>
      <PptxEditor.SlideRail />
      <PptxEditor.Content />
      <PptxEditor.NotesPane />
    </>
  );
}
```

## Dialogs

`PptxEditor.Root` mounts the editor's dialogs from one registry, as DOCX Editor does. Pass `popups` to change one: a render function replaces the packaged dialog, and `false` turns it off; the menu rows that open it show as disabled with the reason. The command behind a dialog always works.

| Key | Dialog | Opened from |
| --- | --- | --- |
| `pageSetup` | Page setup: slide size, appearance, and fonts | Your own menu or button (`useEditorDialog('pageSetup')`) |
| `background` | Format background | Slide context menu |
| `paragraph` | Paragraph: alignment, indents, spacing, and line spacing | Format > Paragraph…, Line spacing > Line spacing options…, text context menu > Paragraph… |
| `font` | Font: font, style, size, font color, underline and underline color, effects, character spacing, and kerning | Format > Font…, Character spacing > More spacing…, text context menu > Font… |
| `bulletsAndNumbering` | Bullets and numbering: bullet or number style, size, color, start number, and a custom character and font | Format > Bullets and numbering…, Bullets or Numbering > Bullets and numbering… |
| `columns` | Columns: the number of columns and the spacing between them | Columns > More columns… |
| `imageAltText` | Alt Text of the selected element | Picture toolbar, shape and picture context menus |
| `imageProperties` | Format Picture of the selected picture | Picture toolbar, Format > Format picture…, picture context menu |
| `hyperlink` | The link popup | Ctrl/Cmd+K, the link control, Insert > Link… |
| `find` | The find and replace bar | Ctrl/Cmd+F, Ctrl/Cmd+H, Edit > Find |
| `contextMenu` | The canvas context menu | Right-click, Shift+F10, or the Menu key on the slide |
| `railContextMenu` | The context menu of a slide in the rail or the slide sorter | Right-click a slide |
| `sectionContextMenu` | The menu of a section header in the rail or the slide sorter | Right-click a section header |

```tsx
import { PptxEditor, definePopup, type ParagraphDialogProps } from '@pptx-editor.dev/react';

function MyParagraphDialog({ open, onClose }: ParagraphDialogProps) {
  return open ? <MyForm onDone={onClose} /> : null;
}

<PptxEditor.Root popups={{ paragraph: definePopup(MyParagraphDialog), background: false }}>
  {/* … */}
</PptxEditor.Root>;
```

To change the packaged Paragraph dialog instead, pass `className` (it goes after `pptx-dialog` on the dialog element) or its named parts: `Header`, `Title`, `Body`, `Footer`, `Field` (`general`, `indentation`, `spacing`), `Error`, `Apply`, and `Cancel`. A host field inside the dialog reads the draft with `useParagraphDialog()`. The Font dialog has the same parts (`Field`: `allText`, `effects`) and `useFontDialog()`. The element sets `data-invalid` while an error shows.

```tsx
<PptxEditor.Root
  popups={{
    paragraph: definePopup((props) => (
      <ParagraphDialog {...props} className="my-dialog">
        <ParagraphDialog.Field name="indentation" hidden />
        <ParagraphDialog.Apply className="my-primary" />
      </ParagraphDialog>
    )),
  }}
>
  {/* … */}
</PptxEditor.Root>
```

Open a dialog from your own chrome with `useEditorDialog(key).open()`. A dialog of your own reads and writes with the matching hook. For example, `useParagraphFormat()` returns the selected paragraphs' format and `apply(update)`, which sends one `setParagraphFormat` command and returns whether the engine accepted it; `useFontFormat()` does the same for the Font dialog's fields with one `setFontFormat` command, and `useBulletFormat()` for the Bullets and Numbering dialog with one `setBulletFormat` command. Its `format` is `null` when the paragraphs have no list and `undefined` when they differ or no text is selected; `apply({ sizePercent: 150 })`, with no `kind`, changes the size and keeps each paragraph's own bullet or number. For the other two hooks, `format.disagrees` says which fields the paragraphs disagree about. A null field means that, or that the value is in the other unit (`spaceAfterPt` is null for a space in lines):

```tsx
import { useParagraphFormat } from '@pptx-editor.dev/react';

function SpaceAfter() {
  const { format, isEnabled, apply } = useParagraphFormat();
  return (
    <button disabled={!isEnabled} onClick={() => apply({ spaceAfterPt: 12 })}>
      After: {format?.disagrees.spaceAfter ? 'mixed' : (format?.spaceAfterPt ?? '–')} pt
    </button>
  );
}
```

## Fonts

Pass font bytes through the `fonts` prop of `PptxEditor` or `PptxEditor.Root`, as in DOCX Editor. `useFonts(...origins)` returns one stable value, so a render never reloads the fonts. A new `fonts` value replaces the faces in place with `setFonts`, and `onFontError` receives each failure:

```tsx
import { PptxEditor, useFonts } from '@pptx-editor.dev/react';
import { packagedFonts } from '@docx-editor.dev/fonts';

const fonts = useFonts(packagedFonts());
<PptxEditor document={bytes} fonts={fonts} onFontError={(error) => report(error)} />;
```

See [Fonts](https://github.com/eigenpal/pptx-editor/blob/main/docs/site/content/customization/fonts.mdx).

## License

Apache-2.0. See [LICENSE](LICENSE).
