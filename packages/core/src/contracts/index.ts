// The `contracts` lane: declaration-only public types shared by the editor and its adapters.
// Adapters program against these; nothing here has runtime behavior.
//
// The DOM-free types are in `neutral.ts`; the DOM-typed `Editor` contract is in `editor.ts`. The
// automation lane imports `neutral.ts` directly, so it compiles without the DOM library.

export type * from './neutral.ts';
export type { Editor, RenderThumbnailOptions } from './editor.ts';
