# PPTX document refresh example

Receive complete PPTX files from a processing server without replacing the open editor. Keep the current slide and zoom, and flash the changes with text summaries.

## Run the example

From the repository root, run:

```bash
bun install
bun run dev:refresh
```

Open `http://localhost:5279`.

The example builds against the workspace source, so no package build runs first.

## Try the workflow

1. Go to slide 2 and select **Simulate server updates**. Two cumulative files change the revenue figure, add a summary slide, rewrite the notes of slide 3, and remove the draft label from slide 1.
2. Watch the slide rail. Changed thumbnails flash, and the current slide shows a band behind the changed text. Unchanged thumbnails stay painted, and the loading screen does not show.
3. Inspect **Review changes**. The deletion summary stays visible without a location. The mock also reports one failed processor operation.
4. Select **Show change** to go to a change. Highlights disappear after five seconds.
5. Select **Undoable edit** and request another update. Each accepted result adds one undo step. With **Reload**, each result resets undo history.
6. Select **Compute changes with Pro** and request another update. The client omits the processor changes, and `reviewModule()` computes them from the captured and returned files.
7. Select **Deliver an older result after the latest update**. The controller refuses that result with `out-of-order`.
8. Reset the example, request updates, and select **Cancel updates**. Accepted changes stay in the presentation.
9. Reset again and type in a slide during processing. The controller refuses the result with `local-edits` and keeps your edit.

## How it works

`runRefreshJob()` in `src/refresh-job.ts` captures the presentation and submits its bytes. It checks the response's `documentId` and `submissionId` before it applies each cumulative result. It stops after a refusal, except for `out-of-order` results. Its `finally` block finishes the submission after completion or transport failure. Cancellation stops the network request and invalidates pending controller work.

The two modes put a result on screen in different ways:

- **Reload** (`mode: 'reload'`, the default) shows the returned file as a reload does. It resets selection and undo history. Slides whose parts did not change keep their painted thumbnails and canvas content.
- **Undoable edit** (`mode: 'edit'`) applies the difference as one undoable edit. Undo returns to the captured presentation or the previous result. If the difference cannot apply as an edit, `runRefreshJob()` retries the same result as a reload.

Both modes keep the current slide and zoom, and never show the loading screen.

The Vite development server acts as the processing server. It generates controlled sample files from `sample-presentation.ts`, discards the uploaded bytes, and does not edit arbitrary presentations. Reset after manual edits. A production processor must edit the captured file and keep earlier successful changes.

The review list shows the processor descriptions, or `describeRefreshChange()` summaries for computed changes, with the status and location diagnostics. An invalid location does not reject a valid file. Keep unavailable summaries visible. Enable **Show change** only for available locations.

Highlights change presentation only. They add no saved content and no undo steps.

The review module requires the EigenPal Pro License. Evaluation use needs no key. To use your key, set `VITE_PPTX_LICENSE` before you start the server.

A reload refuses collaboration sessions. An undoable edit works in a collaboration session and reaches other participants as an ordinary edit.

For options and failure handling, see [Document refresh API](../../docs/site/content/guides/document-refresh.mdx).
