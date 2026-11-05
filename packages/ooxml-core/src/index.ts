// The format-neutral OOXML layer: everything between the bytes of a `.docx`, `.pptx`, or `.xlsx`
// and a bounded XML tree, the text units, and the DrawingML color and geometry math, with no
// PresentationML or WordprocessingML knowledge. Internal for now: `@pptx-editor.dev/core`
// inlines it at build time. See README.md for the sharing plan.

export * from './limits.ts';
export * from './opc-names.ts';
export * from './zip.ts';
export * from './zip-write.ts';
export * from './sinks.ts';
export * from './xml-text.ts';
export * from './xml-reader.ts';
export * from './xml-fast-read.ts';
export * from './sha256.ts';
export * from './image-size.ts';
export * from './grapheme.ts';
export * from './word-segment.ts';
export * from './caret-word.ts';
export * from './drawingml-color.ts';
export * from './drawingml-geometry.ts';
