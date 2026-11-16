# Document refresh design

## Reference

The DOCX Editor implementation is the reference for names, types, and behavior:

- Guide: `docx-editor/docs/site/content/guides/document-refresh.mdx`
- Types: `docx-editor/packages/core/src/editor/document-refresh-types.ts`
- Controller: `docx-editor/packages/core/src/editor/document-refresh.ts`
- Undoable replace: `docx-editor/packages/core/src/editor/replace-document-edit.ts`
- Command contract: `ReplaceDocumentRequest`, `ReplaceDocumentOutcome`, and `ReplaceDocumentRefusal` in `docs/api/docx-editor-core/contracts-editor.api.md`
- Revision-derived changes: `docx-editor/packages/core/src/editor/document-refresh-changes.ts`
- Pro comparison: `compareDocuments` in `@docx-editor.dev/pro`, `docx-editor/docs/site/content/pro/compare-documents.mdx`
- Example: `docx-editor/examples/document-refresh/`

Copy each shared type from DOCX Editor with its doc comment. Change a comment only where the format makes the behavior different. Name the source path and commit in the file header, as `@eigenpal/ooxml-core` does.

## Public API

### Shared with DOCX Editor, unchanged

| Name | Kind |
| --- | --- |
| `createDocumentRefresh(editor)` | Function. Repeated calls for one editor return the same controller. |
| `DocumentRefresh` | Controller interface. Same 12 methods, same signatures. |
| `DocumentRefreshState` | `phase`, `result`, `changes`, `highlightsVisible`, `recoveryAvailable`. |
| `DocumentRefreshError` | Capture error with a stable `code`. |
| `RefreshSubmission` | `{ id, bytes }`. An opaque token. |
| `RefreshUpdate` | `{ submission, sequence, bytes, mode?, changes?, failures? }`. |
| `RefreshUpdateMode` | `'reload' \| 'edit'`. |
| `RefreshResult`, `RefreshRefusal`, `RefreshRefusalFields` | Same shapes. |
| `RefreshFailureCode` | Same 15 codes. |
| `RefreshRefusalReason` | Same 6 reasons. See [Refusal mapping](#refusal-mapping). |
| `RefreshChangeInput` | `{ id, description?, location?, unavailableReason? }`. |
| `RefreshHighlightOptions`, `ClearRefreshHighlightsOptions`, `RefreshHighlightAnimation` | Same options and defaults. |
| `NavigateToChangeOptions` | `block`, `behavior`, `offsetPx`, `focus`. |
| `DocumentHandle` | `{ loadId, revision }`. New for PPTX Editor. |
| `DocumentChange` | `{ source?, revision }`. New for PPTX Editor. |
| `AnchorHighlightOptions`, `AnchorHighlightAnimation`, `ClearAnchorHighlightOptions`, `ScrollPlacement`, `ScrollToAnchorOptions` | New for PPTX Editor, copied from DOCX Editor's contracts. |

Both adapters re-export `createDocumentRefresh` and the types, as DOCX Editor's adapters do. They add no hook. A host reads the controller state through `subscribe()` and `snapshot()`.

### Different from DOCX Editor

#### `RefreshLocation`

```ts
/** A location in the returned presentation. @public */
export interface RefreshLocation {
  /** `p:sldId/@id` of the slide. Use this or `slideIndex`, not both. */
  readonly slideId?: number;
  /** Zero-based slide index in presentation order, hidden slides included. */
  readonly slideIndex?: number;
  /** `p:cNvPr/@id` of a shape on that slide. Omit to locate the whole slide. */
  readonly shapeId?: number;
  /** Zero-based paragraph index in the shape's text body. Requires `shapeId`. */
  readonly paragraphIndex?: number;
  /** UTF-16 offsets in that paragraph. Require `paragraphIndex`. */
  readonly start?: number;
  readonly end?: number;
  /** Exact returned text at these offsets. Required when offsets are present. */
  readonly text?: string;
  /** The notes slide of the slide instead of the slide itself. Default: false. */
  readonly notes?: boolean;
}
```

A location has three levels: a slide, a shape, or a text range. Each level is valid alone. A table cell is out of reach of `shapeId` and `paragraphIndex`. A change in a table locates the graphic frame (the shape level).

#### `RefreshChange`

`RefreshChange` keeps every DOCX Editor field and adds one optional field:

```ts
export interface RefreshChange {
  // id, description, resultId, isNew, status, location, diagnostic: as DOCX Editor.
  /** Present when the controller computed this change. Absent for processor changes. */
  readonly computed?: RefreshComputedChange;
}

/** A computed change, in fields that an adapter can describe in any language. @public */
export interface RefreshComputedChange {
  readonly scope: 'slide' | 'shape' | 'text' | 'notes' | 'layout' | 'master' | 'theme' | 'package';
  readonly action: 'added' | 'removed' | 'moved' | 'changed';
  /** For `changed`: what changed. Several entries are possible. */
  readonly aspects?: readonly (
    | 'position'
    | 'size'
    | 'rotation'
    | 'geometry'
    | 'fill'
    | 'line'
    | 'effects'
    | 'text'
    | 'textFormat'
    | 'picture'
    | 'table'
    | 'chart'
    | 'diagram'
    | 'transition'
    | 'animation'
    | 'background'
    | 'layout'
    | 'other'
  )[];
  /** For `layout`, `master`, and `theme`: zero-based indexes of the slides that show the change. */
  readonly affectedSlides?: readonly number[];
  /** For `removed` and `moved` slides: the zero-based slide index in the captured presentation. */
  readonly previousSlideIndex?: number;
}
```

A computed change has no `description`. The adapters provide localized text from `computed` through the `refresh.change.*` strings of `@pptx-editor.dev/i18n`. A host can replace that text.

#### Change sources

The controller takes changes from one source for each result, in this order:

1. `update.changes`, when present. The controller validates them as DOCX Editor does.
2. The `refresh` hook of a registered module, when present. `reviewModule()` from `@pptx-editor.dev/pro` provides it. The hook compares the captured package with the returned package.
3. None. The result reports `changeInformation: 'unavailable'` and an empty `changes` list.

The hook is part of `EditorModule`, as DOCX Editor's `review.collectReviewItems` is. `core` holds the package matcher but no change reporting. See [License boundary](#license-boundary).

#### Comparison baseline

The comparison always runs from the captured submission to the returned file, because each result is cumulative. A change that two results share keeps its id, and `isNew` marks new ids and changed changes, as in DOCX Editor.

`replaceDocument` in edit mode runs from the open package to the returned file. After the first accepted result, the open package is that result, not the captured submission. The controller advances its revision baseline after its own commit, so its own update never causes `local-edits`.

The `status` union and `diagnostic.code` union keep the DOCX Editor members, so shared host code compiles. PPTX Editor adds diagnostic codes and uses each DOCX Editor code as follows:

| `diagnostic.code` | PPTX Editor meaning |
| --- | --- |
| `invalid-selector` | Both `slideId` and `slideIndex`, neither, a bad index, or a level without its parent level. |
| `invalid-slide-id` | Added. `slideId` is outside 256 through 2147483647 (ECMA-376 Part 1, 19.2.1.33). |
| `slide-not-found` | Added. The returned presentation has no such slide. |
| `shape-not-found` | Added. The slide has no shape with that `p:cNvPr/@id`. |
| `ambiguous-shape-id` | Added. Several shapes on the slide have that id. |
| `invalid-offsets` | The offsets do not form a nonempty range in the paragraph, or the paragraph does not exist. |
| `text-mismatch` | The returned text differs from `text`. |
| `not-rendered` | The location has no layout position, for example a shape in a hidden group. |
| `missing-location` | The change has no location. |
| `deleted` | The deleted content has no surviving location. |
| `unsupported-story` | The location is in a part that no slide view shows: a master, a layout, a chart part, a diagram data part, or a comment. |
| `multiple-paragraphs` | Not produced. Kept for type compatibility. |
| `invalid-paragraph-id`, `paragraph-not-found`, `ambiguous-paragraph-id` | Not produced. Kept for type compatibility. |

#### `comparePresentations`

```ts
// @pptx-editor.dev/pro
export function comparePresentations(
  original: Uint8Array | ArrayBuffer,
  revised: Uint8Array | ArrayBuffer
): ComparePresentationsResult;

export type ComparePresentationsResult =
  | {
      readonly ok: true;
      /** Changes in the `RefreshChangeInput` form, with stable ids. */
      readonly changes: readonly (RefreshChangeInput & {
        readonly computed: RefreshComputedChange;
      })[];
      /** True when a limit stopped the comparison. Remaining slides report at slide level. */
      readonly truncated: boolean;
    }
  | {
      readonly ok: false;
      readonly code: 'invalidDocument' | 'tooLarge';
      /** Diagnostic detail, in English; not for display. */
      readonly message: string;
    };
```

It runs without a DOM, on Node.js and in a browser, and reads both files through `readPptx` with its limits. Its place, license, and result style follow DOCX Editor's `compareDocuments`. Its name and its change list differ, because a presentation has no tracked changes to write. The module hook calls the same comparison on the packages that the controller already parsed, so it does not parse the files twice.

#### `replaceDocument`

```ts
export interface ReplaceDocumentRequest {
  readonly pptx: Uint8Array | ArrayBuffer;
  readonly expected?: DocumentHandle;
}
// editor.exec({ type: 'replaceDocument', pptx, expected })
```

`CommandResult.replaceDocument` carries `ReplaceDocumentOutcome`, the DOCX Editor type, with camelCase codes: `staleDocument`, `invalidRequest`, `invalidDocument`, `unavailable`, `busy`, `editNotPermitted`, `editUnsupported`.

#### `change` event

`PptxEditorEvents.change` gains a second argument: `change: (snapshot: EditorSnapshot, change: DocumentChange) => void`. Existing listeners that take one argument keep working. `source` is absent for user edits, undo, and redo. A remote collaboration edit also has no `source`, as in DOCX Editor.

## Comparison

The comparison matches parts and nodes by keys from the file, not by store node ids. Store node ids come from document order, so two separate parses give different ids to the same shape.

### Steps

1. **Parts.** Compare each part by SHA-256 of its bytes. Skip equal parts. A part that exists in one package only is added or removed. A changed binary part is one `package` change, or a `picture` aspect of each shape that shows it. `[Content_Types].xml` and `.rels` parts are not change sources. They report through the parts that they describe.
2. **Slides.** Match slides by `p14:creationId`, then by `p:sldId/@id`, then by slide part name. An unmatched captured slide is removed. An unmatched returned slide is added. A matched slide at a new index is moved.
3. **Shapes.** On each matched slide, match the shapes of the shape tree and of each group (`p:grpSp`) by `p:cNvPr/@id`. A shape that moves into or out of a group is removed and added. If an id is duplicated or missing, match by `p:cNvPr/@name` and element name, in order. If matching is still ambiguous, report one `slide` change with aspect `other`. Do not guess at shape level.
4. **Shape aspects.** For a matched shape, compare `a:xfrm` (position, size, rotation), geometry, `p:spPr` fill, line, and effects, the picture relationship target bytes, and the text body. Report each changed aspect.
5. **Text.** Compare paragraphs by index after an LCS match on paragraph text. Report the changed range of each changed paragraph with `text-diff.ts`. Run formatting changes without text changes are aspect `textFormat`, located at the paragraph.
6. **Notes, layouts, masters, themes.** A changed notes slide is a `notes` change. A changed layout, master, or theme is one change with `affectedSlides`, not one change for each slide.
7. **Other parts.** Report changes to presentation properties, comments, and parts without a slide view as `package` changes without a location.

### Stable ids

A computed change id is a hash of its keys: the slide `creationId` or `sldId`, the shape id, the paragraph index, and the scope. The same change in two cumulative results has the same id, so `isNew` works as in DOCX Editor.

### Limits

- Part count, part size, and XML depth: the `readPptx` limits.
- Text comparison: the bounds of `text-diff.ts`. Over the bound, the paragraph reports one change for the whole paragraph.
- Changes: at most 10000, as DOCX Editor's `changeIds` limit. Over the limit, set `truncated` and report the remaining slides at slide level.
- Time: in the editor, the comparison runs one slide at a time and yields to the page between slides, because `applyUpdate()` is asynchronous. A benchmark sets the time budget for each step and for the whole comparison. Past the whole budget, the comparison reports the remaining slides at slide level and sets `truncated`.

## Edit mode

`replaceDocument` builds one transaction from the comparison:

- An added part is `addPart`. A removed part is `removePart`.
- A changed XML part is a node-level edit: `setAttribute`, `setText`, `insertNode`, `removeNode`, and `moveNode`, by the matches above. Unmatched subtrees are removed and inserted whole.
- If a node edit of one part exceeds a budget, or a part check refuses it, replace that part with `removePart` and `addPart` in the same transaction.
- A changed binary part is `removePart` and `addPart`.

The store checks the transaction (`part-checks.ts`) before it commits. A refusal leaves the presentation and its history unchanged and returns `editUnsupported` with `unsupportedDifference`.

After the commit:

- One undo returns to the captured presentation. One redo applies the result again.
- The current slide stays if it survives. Otherwise the editor shows the slide at the same index, clamped.
- A selected shape stays selected if it survives. A text caret maps through the paragraph change with `text-diff.ts`. Otherwise the caret moves to the end of the nearest earlier paragraph in the same shape, or the selection clears.
- The `change` event has `source: 'refresh'`.
- A file that equals the open package succeeds with `changed: false` and adds no undo step.

The saved result must equal a save of the returned file opened with `load()`. Equal means the same part names, the same bytes for each binary part, and the same canonical XML for each XML part: expanded names, attribute sets, child order, and text. Namespace prefixes and declaration order can differ, because the store keeps the prefixes of the open file.

#### License boundary

Decision: change reporting is Pro. The maintainers confirmed this boundary.

The package matcher (steps 1 to 3 and the node match of each part) is in `core`, because DOCX Editor's `replaceDocument` and its difference are in `core`. `replaceDocument` and edit mode need no module.

Change reporting is in `@pptx-editor.dev/pro`: the aspects, text ranges, `affectedSlides`, stable change ids, `RefreshComputedChange`, and `comparePresentations`. It reads the matches of the `core` matcher. This follows DOCX Editor, where change derivation in `reviewModule()` and `compareDocuments` are Pro.

## Refusal mapping

| Reason | PPTX Editor condition |
| --- | --- |
| `viewing` | The editor mode is `'view'`, or the snapshot is read-only. |
| `protected` | Not produced. PresentationML has no enforced edit protection. Kept for type compatibility. |
| `collaboration-unavailable` | The collaboration session cannot accept edits now. |
| `different-main-part` | The returned package's `officeDocument` relationship names a different part than the open package. |
| `unsupported-difference` | The store refuses the transaction, or a part cannot be replaced. |
| `collaboration-refused` | The collaboration room refuses the edit. |

## Smooth updates

An accepted update never blinks, flashes, or stalls the page. This applies to both modes. The user sees the old presentation until the new one is ready, then sees the new one in one frame.

### Rules

1. **No loading screen.** A refresh never sets `snapshot.status` to `'loading'`, so `PptxEditor.Loading` never shows. The controller reports progress through `DocumentRefreshState.phase` only.
2. **Swap in one frame.** The editor paints the new current slide off the page. It replaces the old paint in one animation frame. No frame shows an empty canvas, a partial slide, or the old and new slides together.
3. **Fonts before the swap.** When the returned file names fonts that the current slide needs, the editor waits for those faces before the swap, within the font loading budget. Text does not appear in a fallback font and then change.
4. **Unchanged thumbnails stay.** A thumbnail whose slide, layout, master, theme, and media parts have equal bytes keeps its paint. A changed thumbnail paints off the page, then replaces the old paint. A refresh never empties a painted thumbnail container. `thumbnail-queue.ts` empties a container that showed another slide; matched slides are the same slide for this rule.
5. **Stable rail items.** Rail and sorter items take their keys from the slide match (`p14:creationId`, then `p:sldId/@id`), so React and Vue do not mount matched items again. Added slides insert, and removed slides leave, without a layout animation. The rail keeps its scroll position. If the current slide's item moves, the rail scrolls only as far as needed to keep it in view.
6. **View state stays.** The zoom, the canvas scroll, the current slide, the notes pane scroll, and open panes stay. The current slide follows its slide part, also to a new index. In both modes, the selection stays when the nodes it names stay; else it clears.
7. **No long tasks.** Parsing, comparison, layout, and paint run in steps that yield to the page. No step of a refresh takes more than 50 milliseconds, the long task threshold, on the benchmark decks. Input stays responsive during a refresh.
8. **Current slide first.** Before `applyUpdate()` resolves, the editor paints the current slide and the thumbnails in view. Other thumbnails paint in view or when the page is idle, as `thumbnail-queue.ts` already does.

### Modes

- **Edit mode** is incremental. The store commits one transaction. The cascade re-resolves only the elements whose keys changed (`resolve-keys.ts`), and only changed slides paint again. It is the recommended mode for agent workflows.
- **Reload mode** applies the difference to the open session in place, as one commit with no undo step, and gives the document a new load id. Parts with equal bytes keep their paint, and rules 1 through 8 apply. It resets the undo history. The current slide follows its slide part, and the selection stays when the nodes it names stay. A removed current slide leaves the index clamped. When the difference cannot apply as an edit, or embedded fonts change, the editor opens the file with a quiet load that shows no loading screen. It refuses with `collaboration` while a collaboration session is attached.

### Failure

If the new presentation cannot open or paint, the old paint stays on screen. `applyUpdate()` returns `load-failed` with `recovered: true`. The user never sees an empty editor.

## Change flash

After a successful update, `highlightChanges()` flashes what changed. The host calls it, as in DOCX Editor. The guide shows the call directly after `applyUpdate()`.

### What flashes

| Change | Slide rail and slide sorter | Canvas (current slide) | Notes pane |
| --- | --- | --- | --- |
| Slide added or changed | The thumbnail flashes. | A ring inside the slide edge. | None. |
| Shape added or changed | The thumbnail flashes. | A ring around the shape bounds. | None. |
| Text changed | The thumbnail flashes. | A band over each line of the range. | None. |
| Notes changed | The thumbnail flashes. | None. | A ring around the notes pane. |
| Layout, master, or theme changed | Each affected thumbnail flashes. | A ring inside the slide edge, if the current slide is affected. | None. |
| Slide removed or moved | None. The change is `deleted` or reports its new place. | None. | None. |

A thumbnail flash has three parts: a ring around the thumbnail, a light wash over it, and a marker in its corner. The marker has a shape and the accessible label "Changed", so color is not the only signal.

### Motion

A flash uses the DOCX Editor highlight timing: a 180-millisecond fade in, a hold until `timeoutMs` (default 3000), and a fade out of `exitDurationMs`. All surfaces flash together, from one start time.

- A thumbnail that scrolls into view during the hold shows the flash for the remaining time. It does not start again.
- The user can go to another slide during the hold. The canvas then shows the flash for that slide's changes, if it has any.
- Reduced motion caps the fades at 125 milliseconds, as in DOCX Editor. The ring and marker still show.
- A highlight request during the fade out keeps the places that show: they return to shown from the opacity on screen, and only new places enter. The exit options of `clearHighlights()` set both the fade and the removal time.
- The flash runs in CSS on the overlay layer and on the thumbnail container, not in the slide paint. It never paints a slide or a thumbnail again. The canvas rings and bands reuse the overlay of `comment-highlights.ts`.

### Options

`RefreshHighlightOptions` keeps every DOCX Editor option. The color, opacity, border, padding, `className`, and animation options apply to every surface. One optional member is added:

```ts
export interface RefreshHighlightOptions extends AnchorHighlightOptions {
  // includePrevious, changeIds: as DOCX Editor.
  /** Where changes flash. Default: all three. */
  readonly surfaces?: readonly ('thumbnails' | 'canvas' | 'notes')[];
}
```

Hosts style the flash with `--doc-refresh-highlight-color` and the classes `pptx-refresh-thumbnail`, `pptx-refresh-ring`, `pptx-refresh-band`, and `pptx-refresh-marker`.

### Without Pro

A host without `reviewModule()` can send `changes` with slide locations. Thumbnails and the canvas then flash at slide level. Shape and text locations from the server flash at their own level too. Smooth updates are in `core` and need no module.

### Rules that stay

Highlights are not drawn in exported thumbnails and images, the slide show, PDF exports, or saved files. They add no undo entries. Any document edit removes them.

`highlightChanges()` returns the count of selected available changes, on every slide. `navigateToChange(id)` goes to the slide, scrolls the rail so its thumbnail is visible, and scrolls the canvas when the zoom hides the location. With `focus: true`, it selects the shape, or puts the caret at the start of the text range.

## Security

- The returned bytes are untrusted. They pass through `readPptx` and its limits in every mode.
- Descriptions are at most 1000 characters and render as text through DOM APIs.
- `changeIds` holds at most 10000 nonempty strings.
- The comparison never fetches external relationships and never runs embedded content.

## Open questions

1. Should DOCX Editor highlight computed changes without tracked changes? DOCX Editor hosts get locations today from tracked revisions, which a server `compareDocuments` call can write. A highlight-only mode would add `RefreshChange.computed` and a comparison path to DOCX Editor's `reviewModule()`. This change does not depend on it. Decide after PPTX Editor stage 1 ships and hosts use it.
2. Does the store accept `removePart` and `addPart` of one part name in one transaction? How does `addPart` update `[Content_Types].xml` and the `.rels` parts? Stage 2 verifies this first. If needed, add a `replacePart` op and update the collaboration baseline.
3. Should `RefreshLocation` address a table cell? This proposal keeps the graphic frame level. Add a cell level later only if hosts need it.
