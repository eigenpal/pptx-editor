// The store over a package that was just read, for the editor and for the server host: without
// the editor's own table marks, and with the editor's count of text for the ranges of threaded
// comments (`store/store/comment-anchors.ts`).

import { PptxDocumentStore, type PptxPackage } from '../../store/index.ts';
import { editorParagraphText } from './comment-text-range.ts';
import { withoutEditorMarks } from './table-edit-marks.ts';

/** A store over `pkg`, a package that `readPptx` just read. */
export function openStore(pkg: PptxPackage): PptxDocumentStore {
  return new PptxDocumentStore(withoutEditorMarks(pkg), { paragraphText: editorParagraphText });
}
