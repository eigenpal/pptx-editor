# @pptx-editor.dev/editor-api

Drive a `.pptx` from code with an object model in the shape of the Office.js presentation API: `Presentation`, `Slide`, `Shape`, `TextFrame`, `TextRange`, `Table`, and their collections. The package runs on the PPTX Editor engine. It works on a server from PPTX bytes, or in a page against an editor that a reader already has open.

The package is not Office.js and does not run in an add-in host. It implements a documented subset of that API, and every type in it is authored here. It uses the EigenPal Pro License (`LICENSE.md`). Evaluation is permitted. Production use requires a commercial agreement.

For task guides and API details, see the [developer documentation](https://pptx-editor.dev/docs/editor-api).

## Edit a presentation on a server

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { PptxEditor } from '@pptx-editor.dev/editor-api';

const runtime = await PptxEditor.createServer(await readFile('deck.pptx'));
try {
  await runtime.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0);
    const box = slide.shapes.addTextBox('Quarterly results', { left: 40, top: 40, width: 400 });
    box.textFrame.textRange.font.bold = true;
    await context.sync();
  });
  await writeFile('edited.pptx', await runtime.save());
} finally {
  runtime.dispose();
}
```

`createServer` takes the font options of the engine: `fonts` (faces to measure text with) and `substitutes` (load the metric-compatible substitute fonts, default `true`). It also takes `rasterize`, a function that turns the SVG of a slide or shape picture into PNG bytes, so `getImageAsBase64` works on a server. The engine paints the SVG with the editor's own renderer, and your function must answer a PNG of exactly the asked size. The package ships no rasterizer.

## Drive an open editor

Import from `@pptx-editor.dev/editor-api/browser` to get `createBrowser`. It borrows a live `PptxEditorInstance`. Each `sync()` that writes is one undo step, and the selection members work.

```ts
import { PptxEditor } from '@pptx-editor.dev/editor-api/browser';

const runtime = PptxEditor.createBrowser(editor);
await runtime.run(async (context) => {
  const range = context.presentation.getSelectedTextRangeOrNullObject();
  range.load('text');
  await context.sync();
  if (!range.isNullObject) range.font.color = '#C00000';
  await context.sync();
});
```

Dispose the runtime before the editor. A browser runtime has no `save()`: save through the editor.

## Batches, loads, and null objects

- Calls on proxy objects are queued. Nothing reaches the presentation until `await context.sync()`.
- A host batch of independent writes is atomic. A `sync()` can split into several host batches for dependent reads; a later refusal keeps writes that already committed.
- Every read and lookup sees the calls queued before it in the same `sync()`, for example `getItemOrNullObject` after an `add`, or `getSelectedShapes()` after `setSelectedShapes`. When a later call needs the answer of such a read to go on, for example a write to a `getItemOrNullObject` result or to the selected text, the sync splits before that read: the calls before it apply first. On a live editor, the whole sync is still one undo step. A refusal after the split keeps its own diagnostic, and its hint says that the writes before the read applied.
- A call with an argument that the API does not accept does not throw. The next `sync()` refuses with `InvalidArgument`, `NotImplemented`, `NotSupported`, or `ItemNotFound`, and no call of that batch applies, as in Office.js. For example, a write of `ShadeFromTitle` or a bad `addTable` option refuses at the sync. A refused call that returns an object returns one that takes more calls and queues nothing.
- Only three errors throw at the call: a read of a property that you did not load (`PropertyNotLoaded`), a bad `load(...)` argument, and an object or context that you use after its run. A bad `load(...)` argument includes a wrong name anywhere in a path, also below the items of a collection: `shapes.load('itmes/name')` throws `InvalidArgument` with the hint "Did you mean 'items'?". Such a call queues nothing, and the calls before it stay queued.
- A refused sync drops only its own calls. Proxies stay usable, as in Office.js: a retry with the same objects works while the objects exist. An object that the refused sync would have created does not exist, and refuses with `InvalidObjectPath`.
- An object that a call creates is usable in the same batch, for example `addTextBox(...)` followed by `box.name = ...`.
- A `sync()` that you call before the previous one settles takes the calls queued so far. It sends them when the previous sync settles. Syncs apply in the order you call them. A refused sync rejects alone: the syncs after it still apply. A sync that you do not await still applies before the run ends.
- Read a property only after its `load(...)` completed in a `sync()`. An early read throws `PropertyNotLoaded`.
- `load(...)` takes paths through navigation properties and collections, such as `load('shapes/items/textFrame/textRange/text')`, the property tree `load({ shapes: { items: { name: true } } })`, `expand`, and `$all`. Collections take `load('items/name')` and `load({ select, top, skip })`; `load()` and `load('items')` load every scalar property of each item. Each level of collections on a path costs one more host batch; the items of one level are read together. In a sync that writes, put such a load before the writes. An item read after the writes that fails has the diagnostic `edit.read-after-commit`: the writes applied.
- `getItemOrNullObject` answers an object with `isNullObject` set when there is no such item. `getItem` refuses with `ItemNotFound`.
- Proxies end with their run. Use `context.trackedObjects.add(...)` to keep one, and pass it to a later `runtime.run(object, callback)`.
- A write that follows a read in the same run is conditional on that read. If another writer changed the presentation in between, the batch refuses with `StaleDocument` and applies nothing.

## Common mistakes

These four mistakes cause most first errors.

**A read with no load.** A proxy holds no values until a `sync()` fills them. Load the properties first, then sync, then read:

```ts
const box = slide.shapes.getItem('2');
box.load('name');
await context.sync();
console.log(box.name); // A read before the sync throws PropertyNotLoaded.
```

**A sync inside a loop.** Each `sync()` is one round trip and, on a live editor, one undo step. Queue the work for every item, then sync once:

```ts
const shapes = slide.shapes.load('items/name');
await context.sync();
for (const shape of shapes.items) shape.name = shape.name.trim();
await context.sync(); // One sync for all the writes.
```

**A retry after a refusal.** A refused host batch drops its own calls. If the sync split or reports `edit.read-after-commit`, earlier writes remain applied. Objects that you looked up before it still work, so queue the corrected calls again and sync. An object that the refused sync would have created was never created: create it again.

```ts
// Import isPptxEditorError from @pptx-editor.dev/editor-api.
// box and context belong to the active run.
try {
  box.fill.setSolidColor('notacolor');
  await context.sync();
} catch (error) {
  if (!isPptxEditorError(error) || error.code !== 'InvalidArgument') throw error;
  box.fill.setSolidColor('#1F4E79'); // The same proxy works in the retry.
  await context.sync();
}
```

**Bindings in a saved file.** A binding lives in memory, for the lifetime of the runtime. It is not saved, so a presentation that you open again has no bindings. Add them again after you open it. A binding whose shape is deleted stays listed, and `getShape()` refuses with `ItemNotFound`.

## Errors

Every refusal is a `PptxEditorError`. Test for one with `isPptxEditorError(error)`, and branch on `error.code`:

| Code | Cause |
| --- | --- |
| `PropertyNotLoaded`, `ValueNotLoaded` | A read before the `sync()` that fills it |
| `InvalidArgument` | An argument or load option the API does not accept |
| `ItemNotFound` | No such item in the collection |
| `NotSupported` | The call is outside the subset for this host or content, for example a selection on a server |
| `NotImplemented` | The member is in the Office.js shape but this package does not implement it |
| `StaleDocument` | Another writer changed the presentation after this context read it |
| `InvalidObjectPath`, `ObjectInUse`, `InvalidRequestContext`, `RuntimeDisposed` | A proxy or context used outside its lifetime |
| `ResourceLimitExceeded`, `DocumentUnavailable`, `GeneralException` | The presentation or the host refused the request |

`error.debugInfo.innerError` gives a fixed engine diagnostic, for example `host.no-selection` or `edit.read-only`. Messages never contain presentation content.

## Office.js compatibility

`compat/` measures the public surface against a pinned `@types/office-js` release. The report counts every editing method and property write of the selected subset, and compares their exact signatures. Run it from the repository root:

```bash
bun run --filter '@pptx-editor.dev/editor-api' compat:report
```

A signature match does not prove the same runtime behavior. `compat/runtime-notes.json` records each known difference, for example a member that refuses with `NotImplemented`. `compat/README.md` describes the files and the update process.
