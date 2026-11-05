// GENERATED FILE. Do not edit by hand.
//
// packages/editor-api/scripts/generate-conformance.mjs writes this file from
// compat/manifest.json (the selection and the conformance exceptions) and
// compat/reference/presentation.full-inventory.json (the reference types). Run
// `bun run compat:generate` to write it again.
//
// Each `_assert_*` alias compares one selected Office.js member with the real public type
// of the package. IsExact checks both directions, so a narrower or wider type fails, not
// only an unrelated one. A reference type names the package export of the same name, so a
// missing export fails as a real compiler error. Each `_write_*` function checks that the
// reference write type assigns to the real setter.

import type { Expect, IsExact } from '../type-assert.ts';
import type * as PptxEditor from '../../src/index.ts';

// PowerPoint.run
// Not asserted: manifest.json conformanceExceptions "PowerPoint.run".
// The package has no global host to run against, so it exports no top-level run function.
// PptxEditorRuntime#run takes the same batch callback and the same object or object-array first
// argument. The source-compatibility fixtures call it in place of PowerPoint.run.

// PowerPoint.RequestContext

// PowerPoint.RequestContext#presentation
type Ref_RequestContext_presentation = { readonly presentation: PptxEditor.Presentation };
type Auth_RequestContext_presentation = Pick<PptxEditor.RequestContext, 'presentation'>;
type _check_RequestContext_presentation = IsExact<Ref_RequestContext_presentation, Auth_RequestContext_presentation>;
type _assert_RequestContext_presentation = Expect<_check_RequestContext_presentation>;

// PowerPoint.RequestContext#application
type Ref_RequestContext_application = { readonly application: PptxEditor.Application };
type Auth_RequestContext_application = Pick<PptxEditor.RequestContext, 'application'>;
type _check_RequestContext_application = IsExact<Ref_RequestContext_application, Auth_RequestContext_application>;
type _assert_RequestContext_application = Expect<_check_RequestContext_application>;

// OfficeExtension.ClientRequestContext

// OfficeExtension.ClientRequestContext#debugInfo
type Ref_ClientRequestContext_debugInfo = { readonly debugInfo: PptxEditor.RequestContextDebugInfo };
type Auth_ClientRequestContext_debugInfo = Pick<PptxEditor.ClientRequestContext, 'debugInfo'>;
type _check_ClientRequestContext_debugInfo = IsExact<Ref_ClientRequestContext_debugInfo, Auth_ClientRequestContext_debugInfo>;
type _assert_ClientRequestContext_debugInfo = Expect<_check_ClientRequestContext_debugInfo>;

// OfficeExtension.ClientRequestContext#load
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientRequestContext#load".
// The package returns the object it loads, typed as that object (load<T extends
// ClientObject>(object: T, option?: LoadOption): T), so a call can chain; Office.js returns void.
// Every Office.js call compiles, because LoadOption already takes a string and a string array, and
// a void result is not used.

// OfficeExtension.ClientRequestContext#loadRecursive
type Ref_ClientRequestContext_loadRecursive = { (object: PptxEditor.ClientObject, options: { [typeName: string]: PptxEditor.LoadOption | string | string[]; }, maxDepth?: number | undefined): void };
type Auth_ClientRequestContext_loadRecursive = PptxEditor.ClientRequestContext['loadRecursive'];
type _check_ClientRequestContext_loadRecursive = IsExact<Ref_ClientRequestContext_loadRecursive, Auth_ClientRequestContext_loadRecursive>;
type _assert_ClientRequestContext_loadRecursive = Expect<_check_ClientRequestContext_loadRecursive>;

// OfficeExtension.ClientRequestContext#sync
type Ref_ClientRequestContext_sync = { <T>(passThroughValue?: T | undefined): Promise<T> };
type Auth_ClientRequestContext_sync = PptxEditor.ClientRequestContext['sync'];
type _check_ClientRequestContext_sync = IsExact<Ref_ClientRequestContext_sync, Auth_ClientRequestContext_sync>;
type _assert_ClientRequestContext_sync = Expect<_check_ClientRequestContext_sync>;

// OfficeExtension.ClientRequestContext#trace
type Ref_ClientRequestContext_trace = { (message: string): void };
type Auth_ClientRequestContext_trace = PptxEditor.ClientRequestContext['trace'];
type _check_ClientRequestContext_trace = IsExact<Ref_ClientRequestContext_trace, Auth_ClientRequestContext_trace>;
type _assert_ClientRequestContext_trace = Expect<_check_ClientRequestContext_trace>;

// OfficeExtension.ClientRequestContext#trackedObjects
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientRequestContext#trackedObjects".
// The package declares trackedObjects read-only. Each context has one TrackedObjects that the
// runtime makes; a write from a consumer would detach the tracked objects from the context that
// owns them.

// OfficeExtension.RequestContextDebugInfo

// OfficeExtension.RequestContextDebugInfo#pendingStatements
type Ref_RequestContextDebugInfo_pendingStatements = { pendingStatements: string[] };
type Auth_RequestContextDebugInfo_pendingStatements = Pick<PptxEditor.RequestContextDebugInfo, 'pendingStatements'>;
type _check_RequestContextDebugInfo_pendingStatements = IsExact<Ref_RequestContextDebugInfo_pendingStatements, Auth_RequestContextDebugInfo_pendingStatements>;
type _assert_RequestContextDebugInfo_pendingStatements = Expect<_check_RequestContextDebugInfo_pendingStatements>;
export function _write_RequestContextDebugInfo_pendingStatements(target: PptxEditor.RequestContextDebugInfo, value: string[]): void {
  target['pendingStatements'] = value;
}

// OfficeExtension.ClientObject

// OfficeExtension.ClientObject#context
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientObject#context".
// The package declares context read-only and types it as RequestContext. The runtime sets the
// context of a proxy, and a write from a consumer would detach the proxy from its batch. The
// package has one context class, RequestContext; ClientRequestContext is a type alias of it.

// OfficeExtension.ClientObject#isNullObject
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientObject#isNullObject".
// The package declares isNullObject read-only. The runtime sets it after a sync, and a write from a
// consumer would make a null object look real.

// PowerPoint.Presentation

// PowerPoint.Presentation#slides
type Ref_Presentation_slides = { readonly slides: PptxEditor.SlideCollection };
type Auth_Presentation_slides = Pick<PptxEditor.Presentation, 'slides'>;
type _check_Presentation_slides = IsExact<Ref_Presentation_slides, Auth_Presentation_slides>;
type _assert_Presentation_slides = Expect<_check_Presentation_slides>;

// PowerPoint.Presentation#slideMasters
type Ref_Presentation_slideMasters = { readonly slideMasters: PptxEditor.SlideMasterCollection };
type Auth_Presentation_slideMasters = Pick<PptxEditor.Presentation, 'slideMasters'>;
type _check_Presentation_slideMasters = IsExact<Ref_Presentation_slideMasters, Auth_Presentation_slideMasters>;
type _assert_Presentation_slideMasters = Expect<_check_Presentation_slideMasters>;

// PowerPoint.Presentation#tags
type Ref_Presentation_tags = { readonly tags: PptxEditor.TagCollection };
type Auth_Presentation_tags = Pick<PptxEditor.Presentation, 'tags'>;
type _check_Presentation_tags = IsExact<Ref_Presentation_tags, Auth_Presentation_tags>;
type _assert_Presentation_tags = Expect<_check_Presentation_tags>;

// PowerPoint.Presentation#properties
type Ref_Presentation_properties = { readonly properties: PptxEditor.DocumentProperties };
type Auth_Presentation_properties = Pick<PptxEditor.Presentation, 'properties'>;
type _check_Presentation_properties = IsExact<Ref_Presentation_properties, Auth_Presentation_properties>;
type _assert_Presentation_properties = Expect<_check_Presentation_properties>;

// PowerPoint.Presentation#title
type Ref_Presentation_title = { readonly title: string };
type Auth_Presentation_title = Pick<PptxEditor.Presentation, 'title'>;
type _check_Presentation_title = IsExact<Ref_Presentation_title, Auth_Presentation_title>;
type _assert_Presentation_title = Expect<_check_Presentation_title>;

// PowerPoint.Presentation#id
type Ref_Presentation_id = { readonly id: string };
type Auth_Presentation_id = Pick<PptxEditor.Presentation, 'id'>;
type _check_Presentation_id = IsExact<Ref_Presentation_id, Auth_Presentation_id>;
type _assert_Presentation_id = Expect<_check_Presentation_id>;

// PowerPoint.Presentation#pageSetup
type Ref_Presentation_pageSetup = { readonly pageSetup: PptxEditor.PageSetup };
type Auth_Presentation_pageSetup = Pick<PptxEditor.Presentation, 'pageSetup'>;
type _check_Presentation_pageSetup = IsExact<Ref_Presentation_pageSetup, Auth_Presentation_pageSetup>;
type _assert_Presentation_pageSetup = Expect<_check_Presentation_pageSetup>;

// PowerPoint.Presentation#getSelectedSlides
type Ref_Presentation_getSelectedSlides = { (): PptxEditor.SlideScopedCollection };
type Auth_Presentation_getSelectedSlides = PptxEditor.Presentation['getSelectedSlides'];
type _check_Presentation_getSelectedSlides = IsExact<Ref_Presentation_getSelectedSlides, Auth_Presentation_getSelectedSlides>;
type _assert_Presentation_getSelectedSlides = Expect<_check_Presentation_getSelectedSlides>;

// PowerPoint.Presentation#getSelectedShapes
type Ref_Presentation_getSelectedShapes = { (): PptxEditor.ShapeScopedCollection };
type Auth_Presentation_getSelectedShapes = PptxEditor.Presentation['getSelectedShapes'];
type _check_Presentation_getSelectedShapes = IsExact<Ref_Presentation_getSelectedShapes, Auth_Presentation_getSelectedShapes>;
type _assert_Presentation_getSelectedShapes = Expect<_check_Presentation_getSelectedShapes>;

// PowerPoint.Presentation#getSelectedTextRange
type Ref_Presentation_getSelectedTextRange = { (): PptxEditor.TextRange };
type Auth_Presentation_getSelectedTextRange = PptxEditor.Presentation['getSelectedTextRange'];
type _check_Presentation_getSelectedTextRange = IsExact<Ref_Presentation_getSelectedTextRange, Auth_Presentation_getSelectedTextRange>;
type _assert_Presentation_getSelectedTextRange = Expect<_check_Presentation_getSelectedTextRange>;

// PowerPoint.Presentation#getSelectedTextRangeOrNullObject
type Ref_Presentation_getSelectedTextRangeOrNullObject = { (): PptxEditor.TextRange };
type Auth_Presentation_getSelectedTextRangeOrNullObject = PptxEditor.Presentation['getSelectedTextRangeOrNullObject'];
type _check_Presentation_getSelectedTextRangeOrNullObject = IsExact<Ref_Presentation_getSelectedTextRangeOrNullObject, Auth_Presentation_getSelectedTextRangeOrNullObject>;
type _assert_Presentation_getSelectedTextRangeOrNullObject = Expect<_check_Presentation_getSelectedTextRangeOrNullObject>;

// PowerPoint.Presentation#setSelectedSlides
type Ref_Presentation_setSelectedSlides = { (slideIds: string[]): void };
type Auth_Presentation_setSelectedSlides = PptxEditor.Presentation['setSelectedSlides'];
type _check_Presentation_setSelectedSlides = IsExact<Ref_Presentation_setSelectedSlides, Auth_Presentation_setSelectedSlides>;
type _assert_Presentation_setSelectedSlides = Expect<_check_Presentation_setSelectedSlides>;

// PowerPoint.Presentation#insertSlidesFromBase64
type Ref_Presentation_insertSlidesFromBase64 = { (base64File: string, options?: PptxEditor.InsertSlideOptions | undefined): void };
type Auth_Presentation_insertSlidesFromBase64 = PptxEditor.Presentation['insertSlidesFromBase64'];
type _check_Presentation_insertSlidesFromBase64 = IsExact<Ref_Presentation_insertSlidesFromBase64, Auth_Presentation_insertSlidesFromBase64>;
type _assert_Presentation_insertSlidesFromBase64 = Expect<_check_Presentation_insertSlidesFromBase64>;

// PowerPoint.Presentation#customXmlParts
type Ref_Presentation_customXmlParts = { readonly customXmlParts: PptxEditor.CustomXmlPartCollection };
type Auth_Presentation_customXmlParts = Pick<PptxEditor.Presentation, 'customXmlParts'>;
type _check_Presentation_customXmlParts = IsExact<Ref_Presentation_customXmlParts, Auth_Presentation_customXmlParts>;
type _assert_Presentation_customXmlParts = Expect<_check_Presentation_customXmlParts>;

// PowerPoint.Presentation#bindings
type Ref_Presentation_bindings = { readonly bindings: PptxEditor.BindingCollection };
type Auth_Presentation_bindings = Pick<PptxEditor.Presentation, 'bindings'>;
type _check_Presentation_bindings = IsExact<Ref_Presentation_bindings, Auth_Presentation_bindings>;
type _assert_Presentation_bindings = Expect<_check_Presentation_bindings>;

// PowerPoint.SlideCollection

// PowerPoint.SlideCollection#add
type Ref_SlideCollection_add = { (options?: PptxEditor.AddSlideOptions | undefined): void };
type Auth_SlideCollection_add = PptxEditor.SlideCollection['add'];
type _check_SlideCollection_add = IsExact<Ref_SlideCollection_add, Auth_SlideCollection_add>;
type _assert_SlideCollection_add = Expect<_check_SlideCollection_add>;

// PowerPoint.SlideCollection#exportAsBase64Presentation
type Ref_SlideCollection_exportAsBase64Presentation = { (values: (PptxEditor.Slide | string)[]): PptxEditor.ClientResult<string> };
type Auth_SlideCollection_exportAsBase64Presentation = PptxEditor.SlideCollection['exportAsBase64Presentation'];
type _check_SlideCollection_exportAsBase64Presentation = IsExact<Ref_SlideCollection_exportAsBase64Presentation, Auth_SlideCollection_exportAsBase64Presentation>;
type _assert_SlideCollection_exportAsBase64Presentation = Expect<_check_SlideCollection_exportAsBase64Presentation>;

// PowerPoint.SlideCollection#getCount
type Ref_SlideCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_SlideCollection_getCount = PptxEditor.SlideCollection['getCount'];
type _check_SlideCollection_getCount = IsExact<Ref_SlideCollection_getCount, Auth_SlideCollection_getCount>;
type _assert_SlideCollection_getCount = Expect<_check_SlideCollection_getCount>;

// PowerPoint.SlideCollection#getItem
type Ref_SlideCollection_getItem = { (key: string): PptxEditor.Slide };
type Auth_SlideCollection_getItem = PptxEditor.SlideCollection['getItem'];
type _check_SlideCollection_getItem = IsExact<Ref_SlideCollection_getItem, Auth_SlideCollection_getItem>;
type _assert_SlideCollection_getItem = Expect<_check_SlideCollection_getItem>;

// PowerPoint.SlideCollection#getItemAt
type Ref_SlideCollection_getItemAt = { (index: number): PptxEditor.Slide };
type Auth_SlideCollection_getItemAt = PptxEditor.SlideCollection['getItemAt'];
type _check_SlideCollection_getItemAt = IsExact<Ref_SlideCollection_getItemAt, Auth_SlideCollection_getItemAt>;
type _assert_SlideCollection_getItemAt = Expect<_check_SlideCollection_getItemAt>;

// PowerPoint.SlideCollection#getItemOrNullObject
type Ref_SlideCollection_getItemOrNullObject = { (id: string): PptxEditor.Slide };
type Auth_SlideCollection_getItemOrNullObject = PptxEditor.SlideCollection['getItemOrNullObject'];
type _check_SlideCollection_getItemOrNullObject = IsExact<Ref_SlideCollection_getItemOrNullObject, Auth_SlideCollection_getItemOrNullObject>;
type _assert_SlideCollection_getItemOrNullObject = Expect<_check_SlideCollection_getItemOrNullObject>;

// PowerPoint.SlideCollection#items
type Ref_SlideCollection_items = { readonly items: PptxEditor.Slide[] };
type Auth_SlideCollection_items = Pick<PptxEditor.SlideCollection, 'items'>;
type _check_SlideCollection_items = IsExact<Ref_SlideCollection_items, Auth_SlideCollection_items>;
type _assert_SlideCollection_items = Expect<_check_SlideCollection_items>;

// PowerPoint.SlideScopedCollection

// PowerPoint.SlideScopedCollection#exportAsBase64Presentation
type Ref_SlideScopedCollection_exportAsBase64Presentation = { (): PptxEditor.ClientResult<string> };
type Auth_SlideScopedCollection_exportAsBase64Presentation = PptxEditor.SlideScopedCollection['exportAsBase64Presentation'];
type _check_SlideScopedCollection_exportAsBase64Presentation = IsExact<Ref_SlideScopedCollection_exportAsBase64Presentation, Auth_SlideScopedCollection_exportAsBase64Presentation>;
type _assert_SlideScopedCollection_exportAsBase64Presentation = Expect<_check_SlideScopedCollection_exportAsBase64Presentation>;

// PowerPoint.SlideScopedCollection#getCount
type Ref_SlideScopedCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_SlideScopedCollection_getCount = PptxEditor.SlideScopedCollection['getCount'];
type _check_SlideScopedCollection_getCount = IsExact<Ref_SlideScopedCollection_getCount, Auth_SlideScopedCollection_getCount>;
type _assert_SlideScopedCollection_getCount = Expect<_check_SlideScopedCollection_getCount>;

// PowerPoint.SlideScopedCollection#getItem
type Ref_SlideScopedCollection_getItem = { (key: string): PptxEditor.Slide };
type Auth_SlideScopedCollection_getItem = PptxEditor.SlideScopedCollection['getItem'];
type _check_SlideScopedCollection_getItem = IsExact<Ref_SlideScopedCollection_getItem, Auth_SlideScopedCollection_getItem>;
type _assert_SlideScopedCollection_getItem = Expect<_check_SlideScopedCollection_getItem>;

// PowerPoint.SlideScopedCollection#getItemAt
type Ref_SlideScopedCollection_getItemAt = { (index: number): PptxEditor.Slide };
type Auth_SlideScopedCollection_getItemAt = PptxEditor.SlideScopedCollection['getItemAt'];
type _check_SlideScopedCollection_getItemAt = IsExact<Ref_SlideScopedCollection_getItemAt, Auth_SlideScopedCollection_getItemAt>;
type _assert_SlideScopedCollection_getItemAt = Expect<_check_SlideScopedCollection_getItemAt>;

// PowerPoint.SlideScopedCollection#getItemOrNullObject
type Ref_SlideScopedCollection_getItemOrNullObject = { (id: string): PptxEditor.Slide };
type Auth_SlideScopedCollection_getItemOrNullObject = PptxEditor.SlideScopedCollection['getItemOrNullObject'];
type _check_SlideScopedCollection_getItemOrNullObject = IsExact<Ref_SlideScopedCollection_getItemOrNullObject, Auth_SlideScopedCollection_getItemOrNullObject>;
type _assert_SlideScopedCollection_getItemOrNullObject = Expect<_check_SlideScopedCollection_getItemOrNullObject>;

// PowerPoint.SlideScopedCollection#items
type Ref_SlideScopedCollection_items = { readonly items: PptxEditor.Slide[] };
type Auth_SlideScopedCollection_items = Pick<PptxEditor.SlideScopedCollection, 'items'>;
type _check_SlideScopedCollection_items = IsExact<Ref_SlideScopedCollection_items, Auth_SlideScopedCollection_items>;
type _assert_SlideScopedCollection_items = Expect<_check_SlideScopedCollection_items>;

// PowerPoint.Slide

// PowerPoint.Slide#id
type Ref_Slide_id = { readonly id: string };
type Auth_Slide_id = Pick<PptxEditor.Slide, 'id'>;
type _check_Slide_id = IsExact<Ref_Slide_id, Auth_Slide_id>;
type _assert_Slide_id = Expect<_check_Slide_id>;

// PowerPoint.Slide#index
type Ref_Slide_index = { readonly index: number };
type Auth_Slide_index = Pick<PptxEditor.Slide, 'index'>;
type _check_Slide_index = IsExact<Ref_Slide_index, Auth_Slide_index>;
type _assert_Slide_index = Expect<_check_Slide_index>;

// PowerPoint.Slide#layout
type Ref_Slide_layout = { readonly layout: PptxEditor.SlideLayout };
type Auth_Slide_layout = Pick<PptxEditor.Slide, 'layout'>;
type _check_Slide_layout = IsExact<Ref_Slide_layout, Auth_Slide_layout>;
type _assert_Slide_layout = Expect<_check_Slide_layout>;

// PowerPoint.Slide#slideMaster
type Ref_Slide_slideMaster = { readonly slideMaster: PptxEditor.SlideMaster };
type Auth_Slide_slideMaster = Pick<PptxEditor.Slide, 'slideMaster'>;
type _check_Slide_slideMaster = IsExact<Ref_Slide_slideMaster, Auth_Slide_slideMaster>;
type _assert_Slide_slideMaster = Expect<_check_Slide_slideMaster>;

// PowerPoint.Slide#shapes
type Ref_Slide_shapes = { readonly shapes: PptxEditor.ShapeCollection };
type Auth_Slide_shapes = Pick<PptxEditor.Slide, 'shapes'>;
type _check_Slide_shapes = IsExact<Ref_Slide_shapes, Auth_Slide_shapes>;
type _assert_Slide_shapes = Expect<_check_Slide_shapes>;

// PowerPoint.Slide#tags
type Ref_Slide_tags = { readonly tags: PptxEditor.TagCollection };
type Auth_Slide_tags = Pick<PptxEditor.Slide, 'tags'>;
type _check_Slide_tags = IsExact<Ref_Slide_tags, Auth_Slide_tags>;
type _assert_Slide_tags = Expect<_check_Slide_tags>;

// PowerPoint.Slide#hyperlinks
type Ref_Slide_hyperlinks = { readonly hyperlinks: PptxEditor.HyperlinkCollection };
type Auth_Slide_hyperlinks = Pick<PptxEditor.Slide, 'hyperlinks'>;
type _check_Slide_hyperlinks = IsExact<Ref_Slide_hyperlinks, Auth_Slide_hyperlinks>;
type _assert_Slide_hyperlinks = Expect<_check_Slide_hyperlinks>;

// PowerPoint.Slide#delete
type Ref_Slide_delete = { (): void };
type Auth_Slide_delete = PptxEditor.Slide['delete'];
type _check_Slide_delete = IsExact<Ref_Slide_delete, Auth_Slide_delete>;
type _assert_Slide_delete = Expect<_check_Slide_delete>;

// PowerPoint.Slide#moveTo
type Ref_Slide_moveTo = { (slideIndex: number): void };
type Auth_Slide_moveTo = PptxEditor.Slide['moveTo'];
type _check_Slide_moveTo = IsExact<Ref_Slide_moveTo, Auth_Slide_moveTo>;
type _assert_Slide_moveTo = Expect<_check_Slide_moveTo>;

// PowerPoint.Slide#exportAsBase64
type Ref_Slide_exportAsBase64 = { (): PptxEditor.ClientResult<string> };
type Auth_Slide_exportAsBase64 = PptxEditor.Slide['exportAsBase64'];
type _check_Slide_exportAsBase64 = IsExact<Ref_Slide_exportAsBase64, Auth_Slide_exportAsBase64>;
type _assert_Slide_exportAsBase64 = Expect<_check_Slide_exportAsBase64>;

// PowerPoint.Slide#applyLayout
type Ref_Slide_applyLayout = { (slideLayout: PptxEditor.SlideLayout): void };
type Auth_Slide_applyLayout = PptxEditor.Slide['applyLayout'];
type _check_Slide_applyLayout = IsExact<Ref_Slide_applyLayout, Auth_Slide_applyLayout>;
type _assert_Slide_applyLayout = Expect<_check_Slide_applyLayout>;

// PowerPoint.Slide#setSelectedShapes
type Ref_Slide_setSelectedShapes = { (shapeIds: string[]): void };
type Auth_Slide_setSelectedShapes = PptxEditor.Slide['setSelectedShapes'];
type _check_Slide_setSelectedShapes = IsExact<Ref_Slide_setSelectedShapes, Auth_Slide_setSelectedShapes>;
type _assert_Slide_setSelectedShapes = Expect<_check_Slide_setSelectedShapes>;

// PowerPoint.Slide#background
type Ref_Slide_background = { readonly background: PptxEditor.SlideBackground };
type Auth_Slide_background = Pick<PptxEditor.Slide, 'background'>;
type _check_Slide_background = IsExact<Ref_Slide_background, Auth_Slide_background>;
type _assert_Slide_background = Expect<_check_Slide_background>;

// PowerPoint.Slide#themeColorScheme
type Ref_Slide_themeColorScheme = { readonly themeColorScheme: PptxEditor.ThemeColorScheme };
type Auth_Slide_themeColorScheme = Pick<PptxEditor.Slide, 'themeColorScheme'>;
type _check_Slide_themeColorScheme = IsExact<Ref_Slide_themeColorScheme, Auth_Slide_themeColorScheme>;
type _assert_Slide_themeColorScheme = Expect<_check_Slide_themeColorScheme>;

// PowerPoint.Slide#customXmlParts
type Ref_Slide_customXmlParts = { readonly customXmlParts: PptxEditor.CustomXmlPartCollection };
type Auth_Slide_customXmlParts = Pick<PptxEditor.Slide, 'customXmlParts'>;
type _check_Slide_customXmlParts = IsExact<Ref_Slide_customXmlParts, Auth_Slide_customXmlParts>;
type _assert_Slide_customXmlParts = Expect<_check_Slide_customXmlParts>;

// PowerPoint.Slide#getImageAsBase64
type Ref_Slide_getImageAsBase64 = { (options?: PptxEditor.SlideGetImageOptions | undefined): PptxEditor.ClientResult<string> };
type Auth_Slide_getImageAsBase64 = PptxEditor.Slide['getImageAsBase64'];
type _check_Slide_getImageAsBase64 = IsExact<Ref_Slide_getImageAsBase64, Auth_Slide_getImageAsBase64>;
type _assert_Slide_getImageAsBase64 = Expect<_check_Slide_getImageAsBase64>;

// PowerPoint.SlideLayout

// PowerPoint.SlideLayout#id
type Ref_SlideLayout_id = { readonly id: string };
type Auth_SlideLayout_id = Pick<PptxEditor.SlideLayout, 'id'>;
type _check_SlideLayout_id = IsExact<Ref_SlideLayout_id, Auth_SlideLayout_id>;
type _assert_SlideLayout_id = Expect<_check_SlideLayout_id>;

// PowerPoint.SlideLayout#name
type Ref_SlideLayout_name = { readonly name: string };
type Auth_SlideLayout_name = Pick<PptxEditor.SlideLayout, 'name'>;
type _check_SlideLayout_name = IsExact<Ref_SlideLayout_name, Auth_SlideLayout_name>;
type _assert_SlideLayout_name = Expect<_check_SlideLayout_name>;

// PowerPoint.SlideLayout#shapes
type Ref_SlideLayout_shapes = { readonly shapes: PptxEditor.ShapeCollection };
type Auth_SlideLayout_shapes = Pick<PptxEditor.SlideLayout, 'shapes'>;
type _check_SlideLayout_shapes = IsExact<Ref_SlideLayout_shapes, Auth_SlideLayout_shapes>;
type _assert_SlideLayout_shapes = Expect<_check_SlideLayout_shapes>;

// PowerPoint.SlideLayout#type
type Ref_SlideLayout_type = { readonly type: "Blank" | "Chart" | "ChartAndText" | "ClipArtAndText" | "ClipArtAndVerticalText" | "Comparison" | "ContentWithCaption" | "Custom" | "FourObjects" | "LargeObject" | "MediaClipAndText" | "Mixed" | "Object" | "ObjectAndText" | "ObjectAndTwoObjects" | "ObjectOverText" | "OrganizationChart" | "PictureWithCaption" | "SectionHeader" | "Table" | "Text" | "TextAndChart" | "TextAndClipArt" | "TextAndMediaClip" | "TextAndObject" | "TextAndTwoObjects" | "TextOverObject" | "Title" | "TitleOnly" | "TwoColumnText" | "TwoObjects" | "TwoObjectsAndObject" | "TwoObjectsAndText" | "TwoObjectsOverText" | "VerticalText" | "VerticalTitleAndText" | "VerticalTitleAndTextOverChart" | PptxEditor.SlideLayoutType };
type Auth_SlideLayout_type = Pick<PptxEditor.SlideLayout, 'type'>;
type _check_SlideLayout_type = IsExact<Ref_SlideLayout_type, Auth_SlideLayout_type>;
type _assert_SlideLayout_type = Expect<_check_SlideLayout_type>;

// PowerPoint.SlideLayout#background
type Ref_SlideLayout_background = { readonly background: PptxEditor.SlideLayoutBackground };
type Auth_SlideLayout_background = Pick<PptxEditor.SlideLayout, 'background'>;
type _check_SlideLayout_background = IsExact<Ref_SlideLayout_background, Auth_SlideLayout_background>;
type _assert_SlideLayout_background = Expect<_check_SlideLayout_background>;

// PowerPoint.SlideLayout#themeColorScheme
type Ref_SlideLayout_themeColorScheme = { readonly themeColorScheme: PptxEditor.ThemeColorScheme };
type Auth_SlideLayout_themeColorScheme = Pick<PptxEditor.SlideLayout, 'themeColorScheme'>;
type _check_SlideLayout_themeColorScheme = IsExact<Ref_SlideLayout_themeColorScheme, Auth_SlideLayout_themeColorScheme>;
type _assert_SlideLayout_themeColorScheme = Expect<_check_SlideLayout_themeColorScheme>;

// PowerPoint.SlideLayout#customXmlParts
type Ref_SlideLayout_customXmlParts = { readonly customXmlParts: PptxEditor.CustomXmlPartCollection };
type Auth_SlideLayout_customXmlParts = Pick<PptxEditor.SlideLayout, 'customXmlParts'>;
type _check_SlideLayout_customXmlParts = IsExact<Ref_SlideLayout_customXmlParts, Auth_SlideLayout_customXmlParts>;
type _assert_SlideLayout_customXmlParts = Expect<_check_SlideLayout_customXmlParts>;

// PowerPoint.SlideLayoutCollection

// PowerPoint.SlideLayoutCollection#getCount
type Ref_SlideLayoutCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_SlideLayoutCollection_getCount = PptxEditor.SlideLayoutCollection['getCount'];
type _check_SlideLayoutCollection_getCount = IsExact<Ref_SlideLayoutCollection_getCount, Auth_SlideLayoutCollection_getCount>;
type _assert_SlideLayoutCollection_getCount = Expect<_check_SlideLayoutCollection_getCount>;

// PowerPoint.SlideLayoutCollection#getItem
type Ref_SlideLayoutCollection_getItem = { (key: string): PptxEditor.SlideLayout };
type Auth_SlideLayoutCollection_getItem = PptxEditor.SlideLayoutCollection['getItem'];
type _check_SlideLayoutCollection_getItem = IsExact<Ref_SlideLayoutCollection_getItem, Auth_SlideLayoutCollection_getItem>;
type _assert_SlideLayoutCollection_getItem = Expect<_check_SlideLayoutCollection_getItem>;

// PowerPoint.SlideLayoutCollection#getItemAt
type Ref_SlideLayoutCollection_getItemAt = { (index: number): PptxEditor.SlideLayout };
type Auth_SlideLayoutCollection_getItemAt = PptxEditor.SlideLayoutCollection['getItemAt'];
type _check_SlideLayoutCollection_getItemAt = IsExact<Ref_SlideLayoutCollection_getItemAt, Auth_SlideLayoutCollection_getItemAt>;
type _assert_SlideLayoutCollection_getItemAt = Expect<_check_SlideLayoutCollection_getItemAt>;

// PowerPoint.SlideLayoutCollection#getItemOrNullObject
type Ref_SlideLayoutCollection_getItemOrNullObject = { (id: string): PptxEditor.SlideLayout };
type Auth_SlideLayoutCollection_getItemOrNullObject = PptxEditor.SlideLayoutCollection['getItemOrNullObject'];
type _check_SlideLayoutCollection_getItemOrNullObject = IsExact<Ref_SlideLayoutCollection_getItemOrNullObject, Auth_SlideLayoutCollection_getItemOrNullObject>;
type _assert_SlideLayoutCollection_getItemOrNullObject = Expect<_check_SlideLayoutCollection_getItemOrNullObject>;

// PowerPoint.SlideLayoutCollection#items
type Ref_SlideLayoutCollection_items = { readonly items: PptxEditor.SlideLayout[] };
type Auth_SlideLayoutCollection_items = Pick<PptxEditor.SlideLayoutCollection, 'items'>;
type _check_SlideLayoutCollection_items = IsExact<Ref_SlideLayoutCollection_items, Auth_SlideLayoutCollection_items>;
type _assert_SlideLayoutCollection_items = Expect<_check_SlideLayoutCollection_items>;

// PowerPoint.SlideMaster

// PowerPoint.SlideMaster#id
type Ref_SlideMaster_id = { readonly id: string };
type Auth_SlideMaster_id = Pick<PptxEditor.SlideMaster, 'id'>;
type _check_SlideMaster_id = IsExact<Ref_SlideMaster_id, Auth_SlideMaster_id>;
type _assert_SlideMaster_id = Expect<_check_SlideMaster_id>;

// PowerPoint.SlideMaster#layouts
type Ref_SlideMaster_layouts = { readonly layouts: PptxEditor.SlideLayoutCollection };
type Auth_SlideMaster_layouts = Pick<PptxEditor.SlideMaster, 'layouts'>;
type _check_SlideMaster_layouts = IsExact<Ref_SlideMaster_layouts, Auth_SlideMaster_layouts>;
type _assert_SlideMaster_layouts = Expect<_check_SlideMaster_layouts>;

// PowerPoint.SlideMaster#name
type Ref_SlideMaster_name = { readonly name: string };
type Auth_SlideMaster_name = Pick<PptxEditor.SlideMaster, 'name'>;
type _check_SlideMaster_name = IsExact<Ref_SlideMaster_name, Auth_SlideMaster_name>;
type _assert_SlideMaster_name = Expect<_check_SlideMaster_name>;

// PowerPoint.SlideMaster#shapes
type Ref_SlideMaster_shapes = { readonly shapes: PptxEditor.ShapeCollection };
type Auth_SlideMaster_shapes = Pick<PptxEditor.SlideMaster, 'shapes'>;
type _check_SlideMaster_shapes = IsExact<Ref_SlideMaster_shapes, Auth_SlideMaster_shapes>;
type _assert_SlideMaster_shapes = Expect<_check_SlideMaster_shapes>;

// PowerPoint.SlideMaster#background
type Ref_SlideMaster_background = { readonly background: PptxEditor.SlideMasterBackground };
type Auth_SlideMaster_background = Pick<PptxEditor.SlideMaster, 'background'>;
type _check_SlideMaster_background = IsExact<Ref_SlideMaster_background, Auth_SlideMaster_background>;
type _assert_SlideMaster_background = Expect<_check_SlideMaster_background>;

// PowerPoint.SlideMaster#themeColorScheme
type Ref_SlideMaster_themeColorScheme = { readonly themeColorScheme: PptxEditor.ThemeColorScheme };
type Auth_SlideMaster_themeColorScheme = Pick<PptxEditor.SlideMaster, 'themeColorScheme'>;
type _check_SlideMaster_themeColorScheme = IsExact<Ref_SlideMaster_themeColorScheme, Auth_SlideMaster_themeColorScheme>;
type _assert_SlideMaster_themeColorScheme = Expect<_check_SlideMaster_themeColorScheme>;

// PowerPoint.SlideMaster#customXmlParts
type Ref_SlideMaster_customXmlParts = { readonly customXmlParts: PptxEditor.CustomXmlPartCollection };
type Auth_SlideMaster_customXmlParts = Pick<PptxEditor.SlideMaster, 'customXmlParts'>;
type _check_SlideMaster_customXmlParts = IsExact<Ref_SlideMaster_customXmlParts, Auth_SlideMaster_customXmlParts>;
type _assert_SlideMaster_customXmlParts = Expect<_check_SlideMaster_customXmlParts>;

// PowerPoint.SlideMasterCollection

// PowerPoint.SlideMasterCollection#getCount
type Ref_SlideMasterCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_SlideMasterCollection_getCount = PptxEditor.SlideMasterCollection['getCount'];
type _check_SlideMasterCollection_getCount = IsExact<Ref_SlideMasterCollection_getCount, Auth_SlideMasterCollection_getCount>;
type _assert_SlideMasterCollection_getCount = Expect<_check_SlideMasterCollection_getCount>;

// PowerPoint.SlideMasterCollection#getItem
type Ref_SlideMasterCollection_getItem = { (key: string): PptxEditor.SlideMaster };
type Auth_SlideMasterCollection_getItem = PptxEditor.SlideMasterCollection['getItem'];
type _check_SlideMasterCollection_getItem = IsExact<Ref_SlideMasterCollection_getItem, Auth_SlideMasterCollection_getItem>;
type _assert_SlideMasterCollection_getItem = Expect<_check_SlideMasterCollection_getItem>;

// PowerPoint.SlideMasterCollection#getItemAt
type Ref_SlideMasterCollection_getItemAt = { (index: number): PptxEditor.SlideMaster };
type Auth_SlideMasterCollection_getItemAt = PptxEditor.SlideMasterCollection['getItemAt'];
type _check_SlideMasterCollection_getItemAt = IsExact<Ref_SlideMasterCollection_getItemAt, Auth_SlideMasterCollection_getItemAt>;
type _assert_SlideMasterCollection_getItemAt = Expect<_check_SlideMasterCollection_getItemAt>;

// PowerPoint.SlideMasterCollection#getItemOrNullObject
type Ref_SlideMasterCollection_getItemOrNullObject = { (id: string): PptxEditor.SlideMaster };
type Auth_SlideMasterCollection_getItemOrNullObject = PptxEditor.SlideMasterCollection['getItemOrNullObject'];
type _check_SlideMasterCollection_getItemOrNullObject = IsExact<Ref_SlideMasterCollection_getItemOrNullObject, Auth_SlideMasterCollection_getItemOrNullObject>;
type _assert_SlideMasterCollection_getItemOrNullObject = Expect<_check_SlideMasterCollection_getItemOrNullObject>;

// PowerPoint.SlideMasterCollection#items
type Ref_SlideMasterCollection_items = { readonly items: PptxEditor.SlideMaster[] };
type Auth_SlideMasterCollection_items = Pick<PptxEditor.SlideMasterCollection, 'items'>;
type _check_SlideMasterCollection_items = IsExact<Ref_SlideMasterCollection_items, Auth_SlideMasterCollection_items>;
type _assert_SlideMasterCollection_items = Expect<_check_SlideMasterCollection_items>;

// PowerPoint.SlideBackground

// PowerPoint.SlideBackground#areBackgroundGraphicsHidden
type Ref_SlideBackground_areBackgroundGraphicsHidden = { areBackgroundGraphicsHidden: boolean };
type Auth_SlideBackground_areBackgroundGraphicsHidden = Pick<PptxEditor.SlideBackground, 'areBackgroundGraphicsHidden'>;
type _check_SlideBackground_areBackgroundGraphicsHidden = IsExact<Ref_SlideBackground_areBackgroundGraphicsHidden, Auth_SlideBackground_areBackgroundGraphicsHidden>;
type _assert_SlideBackground_areBackgroundGraphicsHidden = Expect<_check_SlideBackground_areBackgroundGraphicsHidden>;
export function _write_SlideBackground_areBackgroundGraphicsHidden(target: PptxEditor.SlideBackground, value: boolean): void {
  target['areBackgroundGraphicsHidden'] = value;
}

// PowerPoint.SlideBackground#fill
type Ref_SlideBackground_fill = { readonly fill: PptxEditor.SlideBackgroundFill };
type Auth_SlideBackground_fill = Pick<PptxEditor.SlideBackground, 'fill'>;
type _check_SlideBackground_fill = IsExact<Ref_SlideBackground_fill, Auth_SlideBackground_fill>;
type _assert_SlideBackground_fill = Expect<_check_SlideBackground_fill>;

// PowerPoint.SlideBackground#isMasterBackgroundFollowed
type Ref_SlideBackground_isMasterBackgroundFollowed = { isMasterBackgroundFollowed: boolean };
type Auth_SlideBackground_isMasterBackgroundFollowed = Pick<PptxEditor.SlideBackground, 'isMasterBackgroundFollowed'>;
type _check_SlideBackground_isMasterBackgroundFollowed = IsExact<Ref_SlideBackground_isMasterBackgroundFollowed, Auth_SlideBackground_isMasterBackgroundFollowed>;
type _assert_SlideBackground_isMasterBackgroundFollowed = Expect<_check_SlideBackground_isMasterBackgroundFollowed>;
export function _write_SlideBackground_isMasterBackgroundFollowed(target: PptxEditor.SlideBackground, value: boolean): void {
  target['isMasterBackgroundFollowed'] = value;
}

// PowerPoint.SlideBackground#reset
type Ref_SlideBackground_reset = { (): void };
type Auth_SlideBackground_reset = PptxEditor.SlideBackground['reset'];
type _check_SlideBackground_reset = IsExact<Ref_SlideBackground_reset, Auth_SlideBackground_reset>;
type _assert_SlideBackground_reset = Expect<_check_SlideBackground_reset>;

// PowerPoint.SlideLayoutBackground

// PowerPoint.SlideLayoutBackground#areBackgroundGraphicsHidden
type Ref_SlideLayoutBackground_areBackgroundGraphicsHidden = { areBackgroundGraphicsHidden: boolean };
type Auth_SlideLayoutBackground_areBackgroundGraphicsHidden = Pick<PptxEditor.SlideLayoutBackground, 'areBackgroundGraphicsHidden'>;
type _check_SlideLayoutBackground_areBackgroundGraphicsHidden = IsExact<Ref_SlideLayoutBackground_areBackgroundGraphicsHidden, Auth_SlideLayoutBackground_areBackgroundGraphicsHidden>;
type _assert_SlideLayoutBackground_areBackgroundGraphicsHidden = Expect<_check_SlideLayoutBackground_areBackgroundGraphicsHidden>;
export function _write_SlideLayoutBackground_areBackgroundGraphicsHidden(target: PptxEditor.SlideLayoutBackground, value: boolean): void {
  target['areBackgroundGraphicsHidden'] = value;
}

// PowerPoint.SlideLayoutBackground#fill
type Ref_SlideLayoutBackground_fill = { readonly fill: PptxEditor.SlideBackgroundFill };
type Auth_SlideLayoutBackground_fill = Pick<PptxEditor.SlideLayoutBackground, 'fill'>;
type _check_SlideLayoutBackground_fill = IsExact<Ref_SlideLayoutBackground_fill, Auth_SlideLayoutBackground_fill>;
type _assert_SlideLayoutBackground_fill = Expect<_check_SlideLayoutBackground_fill>;

// PowerPoint.SlideLayoutBackground#isMasterBackgroundFollowed
type Ref_SlideLayoutBackground_isMasterBackgroundFollowed = { isMasterBackgroundFollowed: boolean };
type Auth_SlideLayoutBackground_isMasterBackgroundFollowed = Pick<PptxEditor.SlideLayoutBackground, 'isMasterBackgroundFollowed'>;
type _check_SlideLayoutBackground_isMasterBackgroundFollowed = IsExact<Ref_SlideLayoutBackground_isMasterBackgroundFollowed, Auth_SlideLayoutBackground_isMasterBackgroundFollowed>;
type _assert_SlideLayoutBackground_isMasterBackgroundFollowed = Expect<_check_SlideLayoutBackground_isMasterBackgroundFollowed>;
export function _write_SlideLayoutBackground_isMasterBackgroundFollowed(target: PptxEditor.SlideLayoutBackground, value: boolean): void {
  target['isMasterBackgroundFollowed'] = value;
}

// PowerPoint.SlideLayoutBackground#reset
type Ref_SlideLayoutBackground_reset = { (): void };
type Auth_SlideLayoutBackground_reset = PptxEditor.SlideLayoutBackground['reset'];
type _check_SlideLayoutBackground_reset = IsExact<Ref_SlideLayoutBackground_reset, Auth_SlideLayoutBackground_reset>;
type _assert_SlideLayoutBackground_reset = Expect<_check_SlideLayoutBackground_reset>;

// PowerPoint.SlideMasterBackground

// PowerPoint.SlideMasterBackground#fill
type Ref_SlideMasterBackground_fill = { readonly fill: PptxEditor.SlideBackgroundFill };
type Auth_SlideMasterBackground_fill = Pick<PptxEditor.SlideMasterBackground, 'fill'>;
type _check_SlideMasterBackground_fill = IsExact<Ref_SlideMasterBackground_fill, Auth_SlideMasterBackground_fill>;
type _assert_SlideMasterBackground_fill = Expect<_check_SlideMasterBackground_fill>;

// PowerPoint.SlideBackgroundFill

// PowerPoint.SlideBackgroundFill#getGradientFillOrNullObject
type Ref_SlideBackgroundFill_getGradientFillOrNullObject = { (): PptxEditor.SlideBackgroundGradientFill };
type Auth_SlideBackgroundFill_getGradientFillOrNullObject = PptxEditor.SlideBackgroundFill['getGradientFillOrNullObject'];
type _check_SlideBackgroundFill_getGradientFillOrNullObject = IsExact<Ref_SlideBackgroundFill_getGradientFillOrNullObject, Auth_SlideBackgroundFill_getGradientFillOrNullObject>;
type _assert_SlideBackgroundFill_getGradientFillOrNullObject = Expect<_check_SlideBackgroundFill_getGradientFillOrNullObject>;

// PowerPoint.SlideBackgroundFill#getPatternFillOrNullObject
type Ref_SlideBackgroundFill_getPatternFillOrNullObject = { (): PptxEditor.SlideBackgroundPatternFill };
type Auth_SlideBackgroundFill_getPatternFillOrNullObject = PptxEditor.SlideBackgroundFill['getPatternFillOrNullObject'];
type _check_SlideBackgroundFill_getPatternFillOrNullObject = IsExact<Ref_SlideBackgroundFill_getPatternFillOrNullObject, Auth_SlideBackgroundFill_getPatternFillOrNullObject>;
type _assert_SlideBackgroundFill_getPatternFillOrNullObject = Expect<_check_SlideBackgroundFill_getPatternFillOrNullObject>;

// PowerPoint.SlideBackgroundFill#getPictureOrTextureFillOrNullObject
type Ref_SlideBackgroundFill_getPictureOrTextureFillOrNullObject = { (): PptxEditor.SlideBackgroundPictureOrTextureFill };
type Auth_SlideBackgroundFill_getPictureOrTextureFillOrNullObject = PptxEditor.SlideBackgroundFill['getPictureOrTextureFillOrNullObject'];
type _check_SlideBackgroundFill_getPictureOrTextureFillOrNullObject = IsExact<Ref_SlideBackgroundFill_getPictureOrTextureFillOrNullObject, Auth_SlideBackgroundFill_getPictureOrTextureFillOrNullObject>;
type _assert_SlideBackgroundFill_getPictureOrTextureFillOrNullObject = Expect<_check_SlideBackgroundFill_getPictureOrTextureFillOrNullObject>;

// PowerPoint.SlideBackgroundFill#getSolidFillOrNullObject
type Ref_SlideBackgroundFill_getSolidFillOrNullObject = { (): PptxEditor.SlideBackgroundSolidFill };
type Auth_SlideBackgroundFill_getSolidFillOrNullObject = PptxEditor.SlideBackgroundFill['getSolidFillOrNullObject'];
type _check_SlideBackgroundFill_getSolidFillOrNullObject = IsExact<Ref_SlideBackgroundFill_getSolidFillOrNullObject, Auth_SlideBackgroundFill_getSolidFillOrNullObject>;
type _assert_SlideBackgroundFill_getSolidFillOrNullObject = Expect<_check_SlideBackgroundFill_getSolidFillOrNullObject>;

// PowerPoint.SlideBackgroundFill#setGradientFill
type Ref_SlideBackgroundFill_setGradientFill = { (options?: PptxEditor.SlideBackgroundGradientFillOptions | undefined): void };
type Auth_SlideBackgroundFill_setGradientFill = PptxEditor.SlideBackgroundFill['setGradientFill'];
type _check_SlideBackgroundFill_setGradientFill = IsExact<Ref_SlideBackgroundFill_setGradientFill, Auth_SlideBackgroundFill_setGradientFill>;
type _assert_SlideBackgroundFill_setGradientFill = Expect<_check_SlideBackgroundFill_setGradientFill>;

// PowerPoint.SlideBackgroundFill#setPatternFill
type Ref_SlideBackgroundFill_setPatternFill = { (options?: PptxEditor.SlideBackgroundPatternFillOptions | undefined): void };
type Auth_SlideBackgroundFill_setPatternFill = PptxEditor.SlideBackgroundFill['setPatternFill'];
type _check_SlideBackgroundFill_setPatternFill = IsExact<Ref_SlideBackgroundFill_setPatternFill, Auth_SlideBackgroundFill_setPatternFill>;
type _assert_SlideBackgroundFill_setPatternFill = Expect<_check_SlideBackgroundFill_setPatternFill>;

// PowerPoint.SlideBackgroundFill#setPictureOrTextureFill
type Ref_SlideBackgroundFill_setPictureOrTextureFill = { (options?: PptxEditor.SlideBackgroundPictureOrTextureFillOptions | undefined): void };
type Auth_SlideBackgroundFill_setPictureOrTextureFill = PptxEditor.SlideBackgroundFill['setPictureOrTextureFill'];
type _check_SlideBackgroundFill_setPictureOrTextureFill = IsExact<Ref_SlideBackgroundFill_setPictureOrTextureFill, Auth_SlideBackgroundFill_setPictureOrTextureFill>;
type _assert_SlideBackgroundFill_setPictureOrTextureFill = Expect<_check_SlideBackgroundFill_setPictureOrTextureFill>;

// PowerPoint.SlideBackgroundFill#setSolidFill
type Ref_SlideBackgroundFill_setSolidFill = { (options?: PptxEditor.SlideBackgroundSolidFillOptions | undefined): void };
type Auth_SlideBackgroundFill_setSolidFill = PptxEditor.SlideBackgroundFill['setSolidFill'];
type _check_SlideBackgroundFill_setSolidFill = IsExact<Ref_SlideBackgroundFill_setSolidFill, Auth_SlideBackgroundFill_setSolidFill>;
type _assert_SlideBackgroundFill_setSolidFill = Expect<_check_SlideBackgroundFill_setSolidFill>;

// PowerPoint.SlideBackgroundFill#type
type Ref_SlideBackgroundFill_type = { readonly type: "Gradient" | "Pattern" | "PictureOrTexture" | "Solid" | "Unsupported" | PptxEditor.SlideBackgroundFillType };
type Auth_SlideBackgroundFill_type = Pick<PptxEditor.SlideBackgroundFill, 'type'>;
type _check_SlideBackgroundFill_type = IsExact<Ref_SlideBackgroundFill_type, Auth_SlideBackgroundFill_type>;
type _assert_SlideBackgroundFill_type = Expect<_check_SlideBackgroundFill_type>;

// PowerPoint.SlideBackgroundGradientFill

// PowerPoint.SlideBackgroundGradientFill#type
type Ref_SlideBackgroundGradientFill_type = { type: "Linear" | "Path" | "Radial" | "Rectangular" | "ShadeFromTitle" | "Unsupported" | PptxEditor.SlideBackgroundGradientFillType };
type Auth_SlideBackgroundGradientFill_type = Pick<PptxEditor.SlideBackgroundGradientFill, 'type'>;
type _check_SlideBackgroundGradientFill_type = IsExact<Ref_SlideBackgroundGradientFill_type, Auth_SlideBackgroundGradientFill_type>;
type _assert_SlideBackgroundGradientFill_type = Expect<_check_SlideBackgroundGradientFill_type>;
export function _write_SlideBackgroundGradientFill_type(target: PptxEditor.SlideBackgroundGradientFill, value: "Linear" | "Path" | "Radial" | "Rectangular" | "ShadeFromTitle" | "Unsupported" | PptxEditor.SlideBackgroundGradientFillType): void {
  target['type'] = value;
}

// PowerPoint.SlideBackgroundPatternFill

// PowerPoint.SlideBackgroundPatternFill#backgroundColor
type Ref_SlideBackgroundPatternFill_backgroundColor = { backgroundColor: string };
type Auth_SlideBackgroundPatternFill_backgroundColor = Pick<PptxEditor.SlideBackgroundPatternFill, 'backgroundColor'>;
type _check_SlideBackgroundPatternFill_backgroundColor = IsExact<Ref_SlideBackgroundPatternFill_backgroundColor, Auth_SlideBackgroundPatternFill_backgroundColor>;
type _assert_SlideBackgroundPatternFill_backgroundColor = Expect<_check_SlideBackgroundPatternFill_backgroundColor>;
export function _write_SlideBackgroundPatternFill_backgroundColor(target: PptxEditor.SlideBackgroundPatternFill, value: string): void {
  target['backgroundColor'] = value;
}

// PowerPoint.SlideBackgroundPatternFill#foregroundColor
type Ref_SlideBackgroundPatternFill_foregroundColor = { foregroundColor: string };
type Auth_SlideBackgroundPatternFill_foregroundColor = Pick<PptxEditor.SlideBackgroundPatternFill, 'foregroundColor'>;
type _check_SlideBackgroundPatternFill_foregroundColor = IsExact<Ref_SlideBackgroundPatternFill_foregroundColor, Auth_SlideBackgroundPatternFill_foregroundColor>;
type _assert_SlideBackgroundPatternFill_foregroundColor = Expect<_check_SlideBackgroundPatternFill_foregroundColor>;
export function _write_SlideBackgroundPatternFill_foregroundColor(target: PptxEditor.SlideBackgroundPatternFill, value: string): void {
  target['foregroundColor'] = value;
}

// PowerPoint.SlideBackgroundPatternFill#pattern
type Ref_SlideBackgroundPatternFill_pattern = { pattern: "Cross" | "DarkDownwardDiagonal" | "DarkHorizontal" | "DarkUpwardDiagonal" | "DarkVertical" | "DashedDownwardDiagonal" | "DashedHorizontal" | "DashedUpwardDiagonal" | "DashedVertical" | "DiagonalBrick" | "DiagonalCross" | "Divot" | "DottedDiamond" | "DottedGrid" | "DownwardDiagonal" | "Horizontal" | "HorizontalBrick" | "LargeCheckerBoard" | "LargeConfetti" | "LargeGrid" | "LightDownwardDiagonal" | "LightHorizontal" | "LightUpwardDiagonal" | "LightVertical" | "NarrowHorizontal" | "NarrowVertical" | "OutlinedDiamond" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent5" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "Plaid" | "Shingle" | "SmallCheckerBoard" | "SmallConfetti" | "SmallGrid" | "SolidDiamond" | "Sphere" | "Trellis" | "Unsupported" | "UpwardDiagonal" | "Vertical" | "Wave" | "Weave" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "ZigZag" | PptxEditor.SlideBackgroundPatternFillType };
type Auth_SlideBackgroundPatternFill_pattern = Pick<PptxEditor.SlideBackgroundPatternFill, 'pattern'>;
type _check_SlideBackgroundPatternFill_pattern = IsExact<Ref_SlideBackgroundPatternFill_pattern, Auth_SlideBackgroundPatternFill_pattern>;
type _assert_SlideBackgroundPatternFill_pattern = Expect<_check_SlideBackgroundPatternFill_pattern>;
export function _write_SlideBackgroundPatternFill_pattern(target: PptxEditor.SlideBackgroundPatternFill, value: "Cross" | "DarkDownwardDiagonal" | "DarkHorizontal" | "DarkUpwardDiagonal" | "DarkVertical" | "DashedDownwardDiagonal" | "DashedHorizontal" | "DashedUpwardDiagonal" | "DashedVertical" | "DiagonalBrick" | "DiagonalCross" | "Divot" | "DottedDiamond" | "DottedGrid" | "DownwardDiagonal" | "Horizontal" | "HorizontalBrick" | "LargeCheckerBoard" | "LargeConfetti" | "LargeGrid" | "LightDownwardDiagonal" | "LightHorizontal" | "LightUpwardDiagonal" | "LightVertical" | "NarrowHorizontal" | "NarrowVertical" | "OutlinedDiamond" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent5" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "Plaid" | "Shingle" | "SmallCheckerBoard" | "SmallConfetti" | "SmallGrid" | "SolidDiamond" | "Sphere" | "Trellis" | "Unsupported" | "UpwardDiagonal" | "Vertical" | "Wave" | "Weave" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "ZigZag" | PptxEditor.SlideBackgroundPatternFillType): void {
  target['pattern'] = value;
}

// PowerPoint.SlideBackgroundPictureOrTextureFill

// PowerPoint.SlideBackgroundPictureOrTextureFill#setImage
type Ref_SlideBackgroundPictureOrTextureFill_setImage = { (base64EncodedImage: string): void };
type Auth_SlideBackgroundPictureOrTextureFill_setImage = PptxEditor.SlideBackgroundPictureOrTextureFill['setImage'];
type _check_SlideBackgroundPictureOrTextureFill_setImage = IsExact<Ref_SlideBackgroundPictureOrTextureFill_setImage, Auth_SlideBackgroundPictureOrTextureFill_setImage>;
type _assert_SlideBackgroundPictureOrTextureFill_setImage = Expect<_check_SlideBackgroundPictureOrTextureFill_setImage>;

// PowerPoint.SlideBackgroundPictureOrTextureFill#transparency
type Ref_SlideBackgroundPictureOrTextureFill_transparency = { transparency: number };
type Auth_SlideBackgroundPictureOrTextureFill_transparency = Pick<PptxEditor.SlideBackgroundPictureOrTextureFill, 'transparency'>;
type _check_SlideBackgroundPictureOrTextureFill_transparency = IsExact<Ref_SlideBackgroundPictureOrTextureFill_transparency, Auth_SlideBackgroundPictureOrTextureFill_transparency>;
type _assert_SlideBackgroundPictureOrTextureFill_transparency = Expect<_check_SlideBackgroundPictureOrTextureFill_transparency>;
export function _write_SlideBackgroundPictureOrTextureFill_transparency(target: PptxEditor.SlideBackgroundPictureOrTextureFill, value: number): void {
  target['transparency'] = value;
}

// PowerPoint.SlideBackgroundSolidFill

// PowerPoint.SlideBackgroundSolidFill#color
type Ref_SlideBackgroundSolidFill_color = { color: string };
type Auth_SlideBackgroundSolidFill_color = Pick<PptxEditor.SlideBackgroundSolidFill, 'color'>;
type _check_SlideBackgroundSolidFill_color = IsExact<Ref_SlideBackgroundSolidFill_color, Auth_SlideBackgroundSolidFill_color>;
type _assert_SlideBackgroundSolidFill_color = Expect<_check_SlideBackgroundSolidFill_color>;
export function _write_SlideBackgroundSolidFill_color(target: PptxEditor.SlideBackgroundSolidFill, value: string): void {
  target['color'] = value;
}

// PowerPoint.SlideBackgroundSolidFill#transparency
type Ref_SlideBackgroundSolidFill_transparency = { transparency: number };
type Auth_SlideBackgroundSolidFill_transparency = Pick<PptxEditor.SlideBackgroundSolidFill, 'transparency'>;
type _check_SlideBackgroundSolidFill_transparency = IsExact<Ref_SlideBackgroundSolidFill_transparency, Auth_SlideBackgroundSolidFill_transparency>;
type _assert_SlideBackgroundSolidFill_transparency = Expect<_check_SlideBackgroundSolidFill_transparency>;
export function _write_SlideBackgroundSolidFill_transparency(target: PptxEditor.SlideBackgroundSolidFill, value: number): void {
  target['transparency'] = value;
}

// PowerPoint.ThemeColorScheme

// PowerPoint.ThemeColorScheme#getThemeColor
type Ref_ThemeColorScheme_getThemeColor = { (color: PptxEditor.ThemeColor): PptxEditor.ClientResult<string>; (color: "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Dark1" | "Dark2" | "FollowedHyperlink" | "Hyperlink" | "Light1" | "Light2" | "Mixed" | "None"): PptxEditor.ClientResult<string> };
type Auth_ThemeColorScheme_getThemeColor = PptxEditor.ThemeColorScheme['getThemeColor'];
type _check_ThemeColorScheme_getThemeColor = IsExact<Ref_ThemeColorScheme_getThemeColor, Auth_ThemeColorScheme_getThemeColor>;
type _assert_ThemeColorScheme_getThemeColor = Expect<_check_ThemeColorScheme_getThemeColor>;

// PowerPoint.ThemeColorScheme#setThemeColor
type Ref_ThemeColorScheme_setThemeColor = { (color: PptxEditor.ThemeColor, rgbColor: string): void; (color: "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Dark1" | "Dark2" | "FollowedHyperlink" | "Hyperlink" | "Light1" | "Light2" | "Mixed" | "None", rgbColor: string): void };
type Auth_ThemeColorScheme_setThemeColor = PptxEditor.ThemeColorScheme['setThemeColor'];
type _check_ThemeColorScheme_setThemeColor = IsExact<Ref_ThemeColorScheme_setThemeColor, Auth_ThemeColorScheme_setThemeColor>;
type _assert_ThemeColorScheme_setThemeColor = Expect<_check_ThemeColorScheme_setThemeColor>;

// PowerPoint.Shape

// PowerPoint.Shape#adjustments
type Ref_Shape_adjustments = { readonly adjustments: PptxEditor.Adjustments };
type Auth_Shape_adjustments = Pick<PptxEditor.Shape, 'adjustments'>;
type _check_Shape_adjustments = IsExact<Ref_Shape_adjustments, Auth_Shape_adjustments>;
type _assert_Shape_adjustments = Expect<_check_Shape_adjustments>;

// PowerPoint.Shape#altTextDescription
type Ref_Shape_altTextDescription = { altTextDescription: string };
type Auth_Shape_altTextDescription = Pick<PptxEditor.Shape, 'altTextDescription'>;
type _check_Shape_altTextDescription = IsExact<Ref_Shape_altTextDescription, Auth_Shape_altTextDescription>;
type _assert_Shape_altTextDescription = Expect<_check_Shape_altTextDescription>;
export function _write_Shape_altTextDescription(target: PptxEditor.Shape, value: string): void {
  target['altTextDescription'] = value;
}

// PowerPoint.Shape#altTextTitle
type Ref_Shape_altTextTitle = { altTextTitle: string };
type Auth_Shape_altTextTitle = Pick<PptxEditor.Shape, 'altTextTitle'>;
type _check_Shape_altTextTitle = IsExact<Ref_Shape_altTextTitle, Auth_Shape_altTextTitle>;
type _assert_Shape_altTextTitle = Expect<_check_Shape_altTextTitle>;
export function _write_Shape_altTextTitle(target: PptxEditor.Shape, value: string): void {
  target['altTextTitle'] = value;
}

// PowerPoint.Shape#creationId
type Ref_Shape_creationId = { readonly creationId: null | string };
type Auth_Shape_creationId = Pick<PptxEditor.Shape, 'creationId'>;
type _check_Shape_creationId = IsExact<Ref_Shape_creationId, Auth_Shape_creationId>;
type _assert_Shape_creationId = Expect<_check_Shape_creationId>;

// PowerPoint.Shape#delete
type Ref_Shape_delete = { (): void };
type Auth_Shape_delete = PptxEditor.Shape['delete'];
type _check_Shape_delete = IsExact<Ref_Shape_delete, Auth_Shape_delete>;
type _assert_Shape_delete = Expect<_check_Shape_delete>;

// PowerPoint.Shape#fill
type Ref_Shape_fill = { readonly fill: PptxEditor.ShapeFill };
type Auth_Shape_fill = Pick<PptxEditor.Shape, 'fill'>;
type _check_Shape_fill = IsExact<Ref_Shape_fill, Auth_Shape_fill>;
type _assert_Shape_fill = Expect<_check_Shape_fill>;

// PowerPoint.Shape#getParentSlide
type Ref_Shape_getParentSlide = { (): PptxEditor.Slide };
type Auth_Shape_getParentSlide = PptxEditor.Shape['getParentSlide'];
type _check_Shape_getParentSlide = IsExact<Ref_Shape_getParentSlide, Auth_Shape_getParentSlide>;
type _assert_Shape_getParentSlide = Expect<_check_Shape_getParentSlide>;

// PowerPoint.Shape#getParentSlideLayout
type Ref_Shape_getParentSlideLayout = { (): PptxEditor.SlideLayout };
type Auth_Shape_getParentSlideLayout = PptxEditor.Shape['getParentSlideLayout'];
type _check_Shape_getParentSlideLayout = IsExact<Ref_Shape_getParentSlideLayout, Auth_Shape_getParentSlideLayout>;
type _assert_Shape_getParentSlideLayout = Expect<_check_Shape_getParentSlideLayout>;

// PowerPoint.Shape#getParentSlideLayoutOrNullObject
type Ref_Shape_getParentSlideLayoutOrNullObject = { (): PptxEditor.SlideLayout };
type Auth_Shape_getParentSlideLayoutOrNullObject = PptxEditor.Shape['getParentSlideLayoutOrNullObject'];
type _check_Shape_getParentSlideLayoutOrNullObject = IsExact<Ref_Shape_getParentSlideLayoutOrNullObject, Auth_Shape_getParentSlideLayoutOrNullObject>;
type _assert_Shape_getParentSlideLayoutOrNullObject = Expect<_check_Shape_getParentSlideLayoutOrNullObject>;

// PowerPoint.Shape#getParentSlideMaster
type Ref_Shape_getParentSlideMaster = { (): PptxEditor.SlideMaster };
type Auth_Shape_getParentSlideMaster = PptxEditor.Shape['getParentSlideMaster'];
type _check_Shape_getParentSlideMaster = IsExact<Ref_Shape_getParentSlideMaster, Auth_Shape_getParentSlideMaster>;
type _assert_Shape_getParentSlideMaster = Expect<_check_Shape_getParentSlideMaster>;

// PowerPoint.Shape#getParentSlideMasterOrNullObject
type Ref_Shape_getParentSlideMasterOrNullObject = { (): PptxEditor.SlideMaster };
type Auth_Shape_getParentSlideMasterOrNullObject = PptxEditor.Shape['getParentSlideMasterOrNullObject'];
type _check_Shape_getParentSlideMasterOrNullObject = IsExact<Ref_Shape_getParentSlideMasterOrNullObject, Auth_Shape_getParentSlideMasterOrNullObject>;
type _assert_Shape_getParentSlideMasterOrNullObject = Expect<_check_Shape_getParentSlideMasterOrNullObject>;

// PowerPoint.Shape#getParentSlideOrNullObject
type Ref_Shape_getParentSlideOrNullObject = { (): PptxEditor.Slide };
type Auth_Shape_getParentSlideOrNullObject = PptxEditor.Shape['getParentSlideOrNullObject'];
type _check_Shape_getParentSlideOrNullObject = IsExact<Ref_Shape_getParentSlideOrNullObject, Auth_Shape_getParentSlideOrNullObject>;
type _assert_Shape_getParentSlideOrNullObject = Expect<_check_Shape_getParentSlideOrNullObject>;

// PowerPoint.Shape#getTable
type Ref_Shape_getTable = { (): PptxEditor.Table };
type Auth_Shape_getTable = PptxEditor.Shape['getTable'];
type _check_Shape_getTable = IsExact<Ref_Shape_getTable, Auth_Shape_getTable>;
type _assert_Shape_getTable = Expect<_check_Shape_getTable>;

// PowerPoint.Shape#getTextFrameOrNullObject
type Ref_Shape_getTextFrameOrNullObject = { (): PptxEditor.TextFrame };
type Auth_Shape_getTextFrameOrNullObject = PptxEditor.Shape['getTextFrameOrNullObject'];
type _check_Shape_getTextFrameOrNullObject = IsExact<Ref_Shape_getTextFrameOrNullObject, Auth_Shape_getTextFrameOrNullObject>;
type _assert_Shape_getTextFrameOrNullObject = Expect<_check_Shape_getTextFrameOrNullObject>;

// PowerPoint.Shape#group
type Ref_Shape_group = { readonly group: PptxEditor.ShapeGroup };
type Auth_Shape_group = Pick<PptxEditor.Shape, 'group'>;
type _check_Shape_group = IsExact<Ref_Shape_group, Auth_Shape_group>;
type _assert_Shape_group = Expect<_check_Shape_group>;

// PowerPoint.Shape#height
type Ref_Shape_height = { height: number };
type Auth_Shape_height = Pick<PptxEditor.Shape, 'height'>;
type _check_Shape_height = IsExact<Ref_Shape_height, Auth_Shape_height>;
type _assert_Shape_height = Expect<_check_Shape_height>;
export function _write_Shape_height(target: PptxEditor.Shape, value: number): void {
  target['height'] = value;
}

// PowerPoint.Shape#id
type Ref_Shape_id = { readonly id: string };
type Auth_Shape_id = Pick<PptxEditor.Shape, 'id'>;
type _check_Shape_id = IsExact<Ref_Shape_id, Auth_Shape_id>;
type _assert_Shape_id = Expect<_check_Shape_id>;

// PowerPoint.Shape#isDecorative
type Ref_Shape_isDecorative = { isDecorative: boolean };
type Auth_Shape_isDecorative = Pick<PptxEditor.Shape, 'isDecorative'>;
type _check_Shape_isDecorative = IsExact<Ref_Shape_isDecorative, Auth_Shape_isDecorative>;
type _assert_Shape_isDecorative = Expect<_check_Shape_isDecorative>;
export function _write_Shape_isDecorative(target: PptxEditor.Shape, value: boolean): void {
  target['isDecorative'] = value;
}

// PowerPoint.Shape#left
type Ref_Shape_left = { left: number };
type Auth_Shape_left = Pick<PptxEditor.Shape, 'left'>;
type _check_Shape_left = IsExact<Ref_Shape_left, Auth_Shape_left>;
type _assert_Shape_left = Expect<_check_Shape_left>;
export function _write_Shape_left(target: PptxEditor.Shape, value: number): void {
  target['left'] = value;
}

// PowerPoint.Shape#level
type Ref_Shape_level = { readonly level: number };
type Auth_Shape_level = Pick<PptxEditor.Shape, 'level'>;
type _check_Shape_level = IsExact<Ref_Shape_level, Auth_Shape_level>;
type _assert_Shape_level = Expect<_check_Shape_level>;

// PowerPoint.Shape#lineFormat
type Ref_Shape_lineFormat = { readonly lineFormat: PptxEditor.ShapeLineFormat };
type Auth_Shape_lineFormat = Pick<PptxEditor.Shape, 'lineFormat'>;
type _check_Shape_lineFormat = IsExact<Ref_Shape_lineFormat, Auth_Shape_lineFormat>;
type _assert_Shape_lineFormat = Expect<_check_Shape_lineFormat>;

// PowerPoint.Shape#name
type Ref_Shape_name = { name: string };
type Auth_Shape_name = Pick<PptxEditor.Shape, 'name'>;
type _check_Shape_name = IsExact<Ref_Shape_name, Auth_Shape_name>;
type _assert_Shape_name = Expect<_check_Shape_name>;
export function _write_Shape_name(target: PptxEditor.Shape, value: string): void {
  target['name'] = value;
}

// PowerPoint.Shape#parentGroup
type Ref_Shape_parentGroup = { readonly parentGroup: PptxEditor.Shape };
type Auth_Shape_parentGroup = Pick<PptxEditor.Shape, 'parentGroup'>;
type _check_Shape_parentGroup = IsExact<Ref_Shape_parentGroup, Auth_Shape_parentGroup>;
type _assert_Shape_parentGroup = Expect<_check_Shape_parentGroup>;

// PowerPoint.Shape#placeholderFormat
type Ref_Shape_placeholderFormat = { readonly placeholderFormat: PptxEditor.PlaceholderFormat };
type Auth_Shape_placeholderFormat = Pick<PptxEditor.Shape, 'placeholderFormat'>;
type _check_Shape_placeholderFormat = IsExact<Ref_Shape_placeholderFormat, Auth_Shape_placeholderFormat>;
type _assert_Shape_placeholderFormat = Expect<_check_Shape_placeholderFormat>;

// PowerPoint.Shape#rotation
type Ref_Shape_rotation = { rotation: number };
type Auth_Shape_rotation = Pick<PptxEditor.Shape, 'rotation'>;
type _check_Shape_rotation = IsExact<Ref_Shape_rotation, Auth_Shape_rotation>;
type _assert_Shape_rotation = Expect<_check_Shape_rotation>;
export function _write_Shape_rotation(target: PptxEditor.Shape, value: number): void {
  target['rotation'] = value;
}

// PowerPoint.Shape#setHyperlink
type Ref_Shape_setHyperlink = { (options?: PptxEditor.HyperlinkAddOptions | undefined): PptxEditor.Hyperlink };
type Auth_Shape_setHyperlink = PptxEditor.Shape['setHyperlink'];
type _check_Shape_setHyperlink = IsExact<Ref_Shape_setHyperlink, Auth_Shape_setHyperlink>;
type _assert_Shape_setHyperlink = Expect<_check_Shape_setHyperlink>;

// PowerPoint.Shape#setZOrder
type Ref_Shape_setZOrder = { (position: PptxEditor.ShapeZOrder): void; (position: "BringForward" | "BringToFront" | "SendBackward" | "SendToBack"): void };
type Auth_Shape_setZOrder = PptxEditor.Shape['setZOrder'];
type _check_Shape_setZOrder = IsExact<Ref_Shape_setZOrder, Auth_Shape_setZOrder>;
type _assert_Shape_setZOrder = Expect<_check_Shape_setZOrder>;

// PowerPoint.Shape#tags
type Ref_Shape_tags = { readonly tags: PptxEditor.TagCollection };
type Auth_Shape_tags = Pick<PptxEditor.Shape, 'tags'>;
type _check_Shape_tags = IsExact<Ref_Shape_tags, Auth_Shape_tags>;
type _assert_Shape_tags = Expect<_check_Shape_tags>;

// PowerPoint.Shape#textFrame
type Ref_Shape_textFrame = { readonly textFrame: PptxEditor.TextFrame };
type Auth_Shape_textFrame = Pick<PptxEditor.Shape, 'textFrame'>;
type _check_Shape_textFrame = IsExact<Ref_Shape_textFrame, Auth_Shape_textFrame>;
type _assert_Shape_textFrame = Expect<_check_Shape_textFrame>;

// PowerPoint.Shape#top
type Ref_Shape_top = { top: number };
type Auth_Shape_top = Pick<PptxEditor.Shape, 'top'>;
type _check_Shape_top = IsExact<Ref_Shape_top, Auth_Shape_top>;
type _assert_Shape_top = Expect<_check_Shape_top>;
export function _write_Shape_top(target: PptxEditor.Shape, value: number): void {
  target['top'] = value;
}

// PowerPoint.Shape#type
// Not asserted: manifest.json conformanceExceptions "PowerPoint.Shape#type".
// The enum leaves out the one value that names a trademarked feature. The file stores that content
// as a diagram (dgm:), and the package reports it as Diagram. Every other value matches, and
// compat:report does not count this read-only member.

// PowerPoint.Shape#visible
type Ref_Shape_visible = { visible: boolean };
type Auth_Shape_visible = Pick<PptxEditor.Shape, 'visible'>;
type _check_Shape_visible = IsExact<Ref_Shape_visible, Auth_Shape_visible>;
type _assert_Shape_visible = Expect<_check_Shape_visible>;
export function _write_Shape_visible(target: PptxEditor.Shape, value: boolean): void {
  target['visible'] = value;
}

// PowerPoint.Shape#width
type Ref_Shape_width = { width: number };
type Auth_Shape_width = Pick<PptxEditor.Shape, 'width'>;
type _check_Shape_width = IsExact<Ref_Shape_width, Auth_Shape_width>;
type _assert_Shape_width = Expect<_check_Shape_width>;
export function _write_Shape_width(target: PptxEditor.Shape, value: number): void {
  target['width'] = value;
}

// PowerPoint.Shape#zOrderPosition
type Ref_Shape_zOrderPosition = { readonly zOrderPosition: number };
type Auth_Shape_zOrderPosition = Pick<PptxEditor.Shape, 'zOrderPosition'>;
type _check_Shape_zOrderPosition = IsExact<Ref_Shape_zOrderPosition, Auth_Shape_zOrderPosition>;
type _assert_Shape_zOrderPosition = Expect<_check_Shape_zOrderPosition>;

// PowerPoint.Shape#customXmlParts
type Ref_Shape_customXmlParts = { readonly customXmlParts: PptxEditor.CustomXmlPartCollection };
type Auth_Shape_customXmlParts = Pick<PptxEditor.Shape, 'customXmlParts'>;
type _check_Shape_customXmlParts = IsExact<Ref_Shape_customXmlParts, Auth_Shape_customXmlParts>;
type _assert_Shape_customXmlParts = Expect<_check_Shape_customXmlParts>;

// PowerPoint.Shape#getImageAsBase64
type Ref_Shape_getImageAsBase64 = { (options?: PptxEditor.ShapeGetImageOptions | undefined): PptxEditor.ClientResult<string> };
type Auth_Shape_getImageAsBase64 = PptxEditor.Shape['getImageAsBase64'];
type _check_Shape_getImageAsBase64 = IsExact<Ref_Shape_getImageAsBase64, Auth_Shape_getImageAsBase64>;
type _assert_Shape_getImageAsBase64 = Expect<_check_Shape_getImageAsBase64>;

// PowerPoint.ShapeCollection

// PowerPoint.ShapeCollection#addGeometricShape
type Ref_ShapeCollection_addGeometricShape = { (geometricShapeType: PptxEditor.GeometricShapeType, options?: PptxEditor.ShapeAddOptions | undefined): PptxEditor.Shape; (geometricShapeType: "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "ActionButtonBackPrevious" | "ActionButtonBeginning" | "ActionButtonBlank" | "ActionButtonDocument" | "ActionButtonEnd" | "ActionButtonForwardNext" | "ActionButtonHelp" | "ActionButtonHome" | "ActionButtonInformation" | "ActionButtonMovie" | "ActionButtonReturn" | "ActionButtonSound" | "Arc" | "BentArrow" | "BentUpArrow" | "Bevel" | "BlockArc" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "BracePair" | "BracketPair" | "Callout1" | "Callout2" | "Callout3" | "Can" | "ChartPlus" | "ChartStar" | "ChartX" | "Chevron" | "Chord" | "CircularArrow" | "Cloud" | "CloudCallout" | "Corner" | "CornerTabs" | "Cube" | "CurvedDownArrow" | "CurvedLeftArrow" | "CurvedRightArrow" | "CurvedUpArrow" | "Decagon" | "DiagonalStripe" | "Diamond" | "Dodecagon" | "Donut" | "DoubleWave" | "DownArrow" | "DownArrowCallout" | "Ellipse" | "EllipseRibbon" | "EllipseRibbon2" | "FlowChartAlternateProcess" | "FlowChartCollate" | "FlowChartConnector" | "FlowChartDecision" | "FlowChartDelay" | "FlowChartDisplay" | "FlowChartDocument" | "FlowChartExtract" | "FlowChartInputOutput" | "FlowChartInternalStorage" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartMagneticTape" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartMerge" | "FlowChartMultidocument" | "FlowChartOfflineStorage" | "FlowChartOffpageConnector" | "FlowChartOnlineStorage" | "FlowChartOr" | "FlowChartPredefinedProcess" | "FlowChartPreparation" | "FlowChartProcess" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSort" | "FlowChartSummingJunction" | "FlowChartTerminator" | "FoldedCorner" | "Frame" | "Funnel" | "Gear6" | "Gear9" | "HalfFrame" | "Heart" | "Heptagon" | "Hexagon" | "HomePlate" | "HorizontalScroll" | "IrregularSeal1" | "IrregularSeal2" | "LeftArrow" | "LeftArrowCallout" | "LeftBrace" | "LeftBracket" | "LeftCircularArrow" | "LeftRightArrow" | "LeftRightArrowCallout" | "LeftRightCircularArrow" | "LeftRightRibbon" | "LeftRightUpArrow" | "LeftUpArrow" | "LightningBolt" | "LineInverse" | "MathDivide" | "MathEqual" | "MathMinus" | "MathMultiply" | "MathNotEqual" | "MathPlus" | "Moon" | "NonIsoscelesTrapezoid" | "NoSmoking" | "NotchedRightArrow" | "Octagon" | "Parallelogram" | "Pentagon" | "Pie" | "PieWedge" | "Plaque" | "PlaqueTabs" | "Plus" | "QuadArrow" | "QuadArrowCallout" | "Rectangle" | "Ribbon" | "Ribbon2" | "RightArrow" | "RightArrowCallout" | "RightBrace" | "RightBracket" | "RightTriangle" | "Round1Rectangle" | "Round2DiagonalRectangle" | "Round2SameRectangle" | "RoundRectangle" | "SmileyFace" | "Snip1Rectangle" | "Snip2DiagonalRectangle" | "Snip2SameRectangle" | "SnipRoundRectangle" | "SquareTabs" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "StripedRightArrow" | "Sun" | "SwooshArrow" | "Teardrop" | "Trapezoid" | "Triangle" | "UpArrow" | "UpArrowCallout" | "UpDownArrow" | "UpDownArrowCallout" | "UturnArrow" | "VerticalScroll" | "Wave" | "WedgeEllipseCallout" | "WedgeRectCallout" | "WedgeRRectCallout", options?: PptxEditor.ShapeAddOptions | undefined): PptxEditor.Shape };
type Auth_ShapeCollection_addGeometricShape = PptxEditor.ShapeCollection['addGeometricShape'];
type _check_ShapeCollection_addGeometricShape = IsExact<Ref_ShapeCollection_addGeometricShape, Auth_ShapeCollection_addGeometricShape>;
type _assert_ShapeCollection_addGeometricShape = Expect<_check_ShapeCollection_addGeometricShape>;

// PowerPoint.ShapeCollection#addLine
type Ref_ShapeCollection_addLine = { (connectorType?: PptxEditor.ConnectorType | undefined, options?: PptxEditor.ShapeAddOptions | undefined): PptxEditor.Shape; (connectorType?: "Curve" | "Elbow" | "Straight" | undefined, options?: PptxEditor.ShapeAddOptions | undefined): PptxEditor.Shape };
type Auth_ShapeCollection_addLine = PptxEditor.ShapeCollection['addLine'];
type _check_ShapeCollection_addLine = IsExact<Ref_ShapeCollection_addLine, Auth_ShapeCollection_addLine>;
type _assert_ShapeCollection_addLine = Expect<_check_ShapeCollection_addLine>;

// PowerPoint.ShapeCollection#addTextBox
type Ref_ShapeCollection_addTextBox = { (text: string, options?: PptxEditor.ShapeAddOptions | undefined): PptxEditor.Shape };
type Auth_ShapeCollection_addTextBox = PptxEditor.ShapeCollection['addTextBox'];
type _check_ShapeCollection_addTextBox = IsExact<Ref_ShapeCollection_addTextBox, Auth_ShapeCollection_addTextBox>;
type _assert_ShapeCollection_addTextBox = Expect<_check_ShapeCollection_addTextBox>;

// PowerPoint.ShapeCollection#addTable
type Ref_ShapeCollection_addTable = { (rowCount: number, columnCount: number, options?: PptxEditor.TableAddOptions | undefined): PptxEditor.Shape };
type Auth_ShapeCollection_addTable = PptxEditor.ShapeCollection['addTable'];
type _check_ShapeCollection_addTable = IsExact<Ref_ShapeCollection_addTable, Auth_ShapeCollection_addTable>;
type _assert_ShapeCollection_addTable = Expect<_check_ShapeCollection_addTable>;

// PowerPoint.ShapeCollection#addGroup
type Ref_ShapeCollection_addGroup = { (values: (PptxEditor.Shape | string)[]): PptxEditor.Shape };
type Auth_ShapeCollection_addGroup = PptxEditor.ShapeCollection['addGroup'];
type _check_ShapeCollection_addGroup = IsExact<Ref_ShapeCollection_addGroup, Auth_ShapeCollection_addGroup>;
type _assert_ShapeCollection_addGroup = Expect<_check_ShapeCollection_addGroup>;

// PowerPoint.ShapeCollection#getItem
type Ref_ShapeCollection_getItem = { (key: string): PptxEditor.Shape };
type Auth_ShapeCollection_getItem = PptxEditor.ShapeCollection['getItem'];
type _check_ShapeCollection_getItem = IsExact<Ref_ShapeCollection_getItem, Auth_ShapeCollection_getItem>;
type _assert_ShapeCollection_getItem = Expect<_check_ShapeCollection_getItem>;

// PowerPoint.ShapeCollection#getItemAt
type Ref_ShapeCollection_getItemAt = { (index: number): PptxEditor.Shape };
type Auth_ShapeCollection_getItemAt = PptxEditor.ShapeCollection['getItemAt'];
type _check_ShapeCollection_getItemAt = IsExact<Ref_ShapeCollection_getItemAt, Auth_ShapeCollection_getItemAt>;
type _assert_ShapeCollection_getItemAt = Expect<_check_ShapeCollection_getItemAt>;

// PowerPoint.ShapeCollection#getItemOrNullObject
type Ref_ShapeCollection_getItemOrNullObject = { (id: string): PptxEditor.Shape };
type Auth_ShapeCollection_getItemOrNullObject = PptxEditor.ShapeCollection['getItemOrNullObject'];
type _check_ShapeCollection_getItemOrNullObject = IsExact<Ref_ShapeCollection_getItemOrNullObject, Auth_ShapeCollection_getItemOrNullObject>;
type _assert_ShapeCollection_getItemOrNullObject = Expect<_check_ShapeCollection_getItemOrNullObject>;

// PowerPoint.ShapeCollection#getCount
type Ref_ShapeCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_ShapeCollection_getCount = PptxEditor.ShapeCollection['getCount'];
type _check_ShapeCollection_getCount = IsExact<Ref_ShapeCollection_getCount, Auth_ShapeCollection_getCount>;
type _assert_ShapeCollection_getCount = Expect<_check_ShapeCollection_getCount>;

// PowerPoint.ShapeCollection#items
type Ref_ShapeCollection_items = { readonly items: PptxEditor.Shape[] };
type Auth_ShapeCollection_items = Pick<PptxEditor.ShapeCollection, 'items'>;
type _check_ShapeCollection_items = IsExact<Ref_ShapeCollection_items, Auth_ShapeCollection_items>;
type _assert_ShapeCollection_items = Expect<_check_ShapeCollection_items>;

// PowerPoint.ShapeScopedCollection

// PowerPoint.ShapeScopedCollection#getCount
type Ref_ShapeScopedCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_ShapeScopedCollection_getCount = PptxEditor.ShapeScopedCollection['getCount'];
type _check_ShapeScopedCollection_getCount = IsExact<Ref_ShapeScopedCollection_getCount, Auth_ShapeScopedCollection_getCount>;
type _assert_ShapeScopedCollection_getCount = Expect<_check_ShapeScopedCollection_getCount>;

// PowerPoint.ShapeScopedCollection#getItem
type Ref_ShapeScopedCollection_getItem = { (key: string): PptxEditor.Shape };
type Auth_ShapeScopedCollection_getItem = PptxEditor.ShapeScopedCollection['getItem'];
type _check_ShapeScopedCollection_getItem = IsExact<Ref_ShapeScopedCollection_getItem, Auth_ShapeScopedCollection_getItem>;
type _assert_ShapeScopedCollection_getItem = Expect<_check_ShapeScopedCollection_getItem>;

// PowerPoint.ShapeScopedCollection#getItemAt
type Ref_ShapeScopedCollection_getItemAt = { (index: number): PptxEditor.Shape };
type Auth_ShapeScopedCollection_getItemAt = PptxEditor.ShapeScopedCollection['getItemAt'];
type _check_ShapeScopedCollection_getItemAt = IsExact<Ref_ShapeScopedCollection_getItemAt, Auth_ShapeScopedCollection_getItemAt>;
type _assert_ShapeScopedCollection_getItemAt = Expect<_check_ShapeScopedCollection_getItemAt>;

// PowerPoint.ShapeScopedCollection#getItemOrNullObject
type Ref_ShapeScopedCollection_getItemOrNullObject = { (id: string): PptxEditor.Shape };
type Auth_ShapeScopedCollection_getItemOrNullObject = PptxEditor.ShapeScopedCollection['getItemOrNullObject'];
type _check_ShapeScopedCollection_getItemOrNullObject = IsExact<Ref_ShapeScopedCollection_getItemOrNullObject, Auth_ShapeScopedCollection_getItemOrNullObject>;
type _assert_ShapeScopedCollection_getItemOrNullObject = Expect<_check_ShapeScopedCollection_getItemOrNullObject>;

// PowerPoint.ShapeScopedCollection#group
type Ref_ShapeScopedCollection_group = { (): PptxEditor.Shape };
type Auth_ShapeScopedCollection_group = PptxEditor.ShapeScopedCollection['group'];
type _check_ShapeScopedCollection_group = IsExact<Ref_ShapeScopedCollection_group, Auth_ShapeScopedCollection_group>;
type _assert_ShapeScopedCollection_group = Expect<_check_ShapeScopedCollection_group>;

// PowerPoint.ShapeScopedCollection#items
type Ref_ShapeScopedCollection_items = { readonly items: PptxEditor.Shape[] };
type Auth_ShapeScopedCollection_items = Pick<PptxEditor.ShapeScopedCollection, 'items'>;
type _check_ShapeScopedCollection_items = IsExact<Ref_ShapeScopedCollection_items, Auth_ShapeScopedCollection_items>;
type _assert_ShapeScopedCollection_items = Expect<_check_ShapeScopedCollection_items>;

// PowerPoint.ShapeGroup

// PowerPoint.ShapeGroup#creationId
type Ref_ShapeGroup_creationId = { readonly creationId: null | string };
type Auth_ShapeGroup_creationId = Pick<PptxEditor.ShapeGroup, 'creationId'>;
type _check_ShapeGroup_creationId = IsExact<Ref_ShapeGroup_creationId, Auth_ShapeGroup_creationId>;
type _assert_ShapeGroup_creationId = Expect<_check_ShapeGroup_creationId>;

// PowerPoint.ShapeGroup#id
type Ref_ShapeGroup_id = { readonly id: string };
type Auth_ShapeGroup_id = Pick<PptxEditor.ShapeGroup, 'id'>;
type _check_ShapeGroup_id = IsExact<Ref_ShapeGroup_id, Auth_ShapeGroup_id>;
type _assert_ShapeGroup_id = Expect<_check_ShapeGroup_id>;

// PowerPoint.ShapeGroup#shape
type Ref_ShapeGroup_shape = { readonly shape: PptxEditor.Shape };
type Auth_ShapeGroup_shape = Pick<PptxEditor.ShapeGroup, 'shape'>;
type _check_ShapeGroup_shape = IsExact<Ref_ShapeGroup_shape, Auth_ShapeGroup_shape>;
type _assert_ShapeGroup_shape = Expect<_check_ShapeGroup_shape>;

// PowerPoint.ShapeGroup#shapes
type Ref_ShapeGroup_shapes = { readonly shapes: PptxEditor.ShapeScopedCollection };
type Auth_ShapeGroup_shapes = Pick<PptxEditor.ShapeGroup, 'shapes'>;
type _check_ShapeGroup_shapes = IsExact<Ref_ShapeGroup_shapes, Auth_ShapeGroup_shapes>;
type _assert_ShapeGroup_shapes = Expect<_check_ShapeGroup_shapes>;

// PowerPoint.ShapeGroup#ungroup
type Ref_ShapeGroup_ungroup = { (): void };
type Auth_ShapeGroup_ungroup = PptxEditor.ShapeGroup['ungroup'];
type _check_ShapeGroup_ungroup = IsExact<Ref_ShapeGroup_ungroup, Auth_ShapeGroup_ungroup>;
type _assert_ShapeGroup_ungroup = Expect<_check_ShapeGroup_ungroup>;

// PowerPoint.ShapeFill

// PowerPoint.ShapeFill#clear
type Ref_ShapeFill_clear = { (): void };
type Auth_ShapeFill_clear = PptxEditor.ShapeFill['clear'];
type _check_ShapeFill_clear = IsExact<Ref_ShapeFill_clear, Auth_ShapeFill_clear>;
type _assert_ShapeFill_clear = Expect<_check_ShapeFill_clear>;

// PowerPoint.ShapeFill#foregroundColor
type Ref_ShapeFill_foregroundColor = { foregroundColor: string };
type Auth_ShapeFill_foregroundColor = Pick<PptxEditor.ShapeFill, 'foregroundColor'>;
type _check_ShapeFill_foregroundColor = IsExact<Ref_ShapeFill_foregroundColor, Auth_ShapeFill_foregroundColor>;
type _assert_ShapeFill_foregroundColor = Expect<_check_ShapeFill_foregroundColor>;
export function _write_ShapeFill_foregroundColor(target: PptxEditor.ShapeFill, value: string): void {
  target['foregroundColor'] = value;
}

// PowerPoint.ShapeFill#setImage
type Ref_ShapeFill_setImage = { (base64EncodedImage: string): void };
type Auth_ShapeFill_setImage = PptxEditor.ShapeFill['setImage'];
type _check_ShapeFill_setImage = IsExact<Ref_ShapeFill_setImage, Auth_ShapeFill_setImage>;
type _assert_ShapeFill_setImage = Expect<_check_ShapeFill_setImage>;

// PowerPoint.ShapeFill#setSolidColor
type Ref_ShapeFill_setSolidColor = { (color: string): void };
type Auth_ShapeFill_setSolidColor = PptxEditor.ShapeFill['setSolidColor'];
type _check_ShapeFill_setSolidColor = IsExact<Ref_ShapeFill_setSolidColor, Auth_ShapeFill_setSolidColor>;
type _assert_ShapeFill_setSolidColor = Expect<_check_ShapeFill_setSolidColor>;

// PowerPoint.ShapeFill#transparency
type Ref_ShapeFill_transparency = { transparency: number };
type Auth_ShapeFill_transparency = Pick<PptxEditor.ShapeFill, 'transparency'>;
type _check_ShapeFill_transparency = IsExact<Ref_ShapeFill_transparency, Auth_ShapeFill_transparency>;
type _assert_ShapeFill_transparency = Expect<_check_ShapeFill_transparency>;
export function _write_ShapeFill_transparency(target: PptxEditor.ShapeFill, value: number): void {
  target['transparency'] = value;
}

// PowerPoint.ShapeFill#type
type Ref_ShapeFill_type = { readonly type: "Gradient" | "NoFill" | "Pattern" | "PictureAndTexture" | "SlideBackground" | "Solid" | PptxEditor.ShapeFillType };
type Auth_ShapeFill_type = Pick<PptxEditor.ShapeFill, 'type'>;
type _check_ShapeFill_type = IsExact<Ref_ShapeFill_type, Auth_ShapeFill_type>;
type _assert_ShapeFill_type = Expect<_check_ShapeFill_type>;

// PowerPoint.ShapeLineFormat

// PowerPoint.ShapeLineFormat#color
type Ref_ShapeLineFormat_color = { color: string };
type Auth_ShapeLineFormat_color = Pick<PptxEditor.ShapeLineFormat, 'color'>;
type _check_ShapeLineFormat_color = IsExact<Ref_ShapeLineFormat_color, Auth_ShapeLineFormat_color>;
type _assert_ShapeLineFormat_color = Expect<_check_ShapeLineFormat_color>;
export function _write_ShapeLineFormat_color(target: PptxEditor.ShapeLineFormat, value: string): void {
  target['color'] = value;
}

// PowerPoint.ShapeLineFormat#dashStyle
type Ref_ShapeLineFormat_dashStyle = { dashStyle: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle };
type Auth_ShapeLineFormat_dashStyle = Pick<PptxEditor.ShapeLineFormat, 'dashStyle'>;
type _check_ShapeLineFormat_dashStyle = IsExact<Ref_ShapeLineFormat_dashStyle, Auth_ShapeLineFormat_dashStyle>;
type _assert_ShapeLineFormat_dashStyle = Expect<_check_ShapeLineFormat_dashStyle>;
export function _write_ShapeLineFormat_dashStyle(target: PptxEditor.ShapeLineFormat, value: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle): void {
  target['dashStyle'] = value;
}

// PowerPoint.ShapeLineFormat#style
type Ref_ShapeLineFormat_style = { style: "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin" | PptxEditor.ShapeLineStyle };
type Auth_ShapeLineFormat_style = Pick<PptxEditor.ShapeLineFormat, 'style'>;
type _check_ShapeLineFormat_style = IsExact<Ref_ShapeLineFormat_style, Auth_ShapeLineFormat_style>;
type _assert_ShapeLineFormat_style = Expect<_check_ShapeLineFormat_style>;
export function _write_ShapeLineFormat_style(target: PptxEditor.ShapeLineFormat, value: "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin" | PptxEditor.ShapeLineStyle): void {
  target['style'] = value;
}

// PowerPoint.ShapeLineFormat#transparency
type Ref_ShapeLineFormat_transparency = { transparency: number };
type Auth_ShapeLineFormat_transparency = Pick<PptxEditor.ShapeLineFormat, 'transparency'>;
type _check_ShapeLineFormat_transparency = IsExact<Ref_ShapeLineFormat_transparency, Auth_ShapeLineFormat_transparency>;
type _assert_ShapeLineFormat_transparency = Expect<_check_ShapeLineFormat_transparency>;
export function _write_ShapeLineFormat_transparency(target: PptxEditor.ShapeLineFormat, value: number): void {
  target['transparency'] = value;
}

// PowerPoint.ShapeLineFormat#visible
type Ref_ShapeLineFormat_visible = { visible: boolean };
type Auth_ShapeLineFormat_visible = Pick<PptxEditor.ShapeLineFormat, 'visible'>;
type _check_ShapeLineFormat_visible = IsExact<Ref_ShapeLineFormat_visible, Auth_ShapeLineFormat_visible>;
type _assert_ShapeLineFormat_visible = Expect<_check_ShapeLineFormat_visible>;
export function _write_ShapeLineFormat_visible(target: PptxEditor.ShapeLineFormat, value: boolean): void {
  target['visible'] = value;
}

// PowerPoint.ShapeLineFormat#weight
type Ref_ShapeLineFormat_weight = { weight: number };
type Auth_ShapeLineFormat_weight = Pick<PptxEditor.ShapeLineFormat, 'weight'>;
type _check_ShapeLineFormat_weight = IsExact<Ref_ShapeLineFormat_weight, Auth_ShapeLineFormat_weight>;
type _assert_ShapeLineFormat_weight = Expect<_check_ShapeLineFormat_weight>;
export function _write_ShapeLineFormat_weight(target: PptxEditor.ShapeLineFormat, value: number): void {
  target['weight'] = value;
}

// PowerPoint.ShapeFont

// PowerPoint.ShapeFont#allCaps
type Ref_ShapeFont_allCaps = { allCaps: boolean | null };
type Auth_ShapeFont_allCaps = Pick<PptxEditor.ShapeFont, 'allCaps'>;
type _check_ShapeFont_allCaps = IsExact<Ref_ShapeFont_allCaps, Auth_ShapeFont_allCaps>;
type _assert_ShapeFont_allCaps = Expect<_check_ShapeFont_allCaps>;
export function _write_ShapeFont_allCaps(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['allCaps'] = value;
}

// PowerPoint.ShapeFont#bold
type Ref_ShapeFont_bold = { bold: boolean | null };
type Auth_ShapeFont_bold = Pick<PptxEditor.ShapeFont, 'bold'>;
type _check_ShapeFont_bold = IsExact<Ref_ShapeFont_bold, Auth_ShapeFont_bold>;
type _assert_ShapeFont_bold = Expect<_check_ShapeFont_bold>;
export function _write_ShapeFont_bold(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['bold'] = value;
}

// PowerPoint.ShapeFont#color
type Ref_ShapeFont_color = { color: null | string };
type Auth_ShapeFont_color = Pick<PptxEditor.ShapeFont, 'color'>;
type _check_ShapeFont_color = IsExact<Ref_ShapeFont_color, Auth_ShapeFont_color>;
type _assert_ShapeFont_color = Expect<_check_ShapeFont_color>;
export function _write_ShapeFont_color(target: PptxEditor.ShapeFont, value: null | string): void {
  target['color'] = value;
}

// PowerPoint.ShapeFont#doubleStrikethrough
type Ref_ShapeFont_doubleStrikethrough = { doubleStrikethrough: boolean | null };
type Auth_ShapeFont_doubleStrikethrough = Pick<PptxEditor.ShapeFont, 'doubleStrikethrough'>;
type _check_ShapeFont_doubleStrikethrough = IsExact<Ref_ShapeFont_doubleStrikethrough, Auth_ShapeFont_doubleStrikethrough>;
type _assert_ShapeFont_doubleStrikethrough = Expect<_check_ShapeFont_doubleStrikethrough>;
export function _write_ShapeFont_doubleStrikethrough(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['doubleStrikethrough'] = value;
}

// PowerPoint.ShapeFont#italic
type Ref_ShapeFont_italic = { italic: boolean | null };
type Auth_ShapeFont_italic = Pick<PptxEditor.ShapeFont, 'italic'>;
type _check_ShapeFont_italic = IsExact<Ref_ShapeFont_italic, Auth_ShapeFont_italic>;
type _assert_ShapeFont_italic = Expect<_check_ShapeFont_italic>;
export function _write_ShapeFont_italic(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['italic'] = value;
}

// PowerPoint.ShapeFont#name
type Ref_ShapeFont_name = { name: null | string };
type Auth_ShapeFont_name = Pick<PptxEditor.ShapeFont, 'name'>;
type _check_ShapeFont_name = IsExact<Ref_ShapeFont_name, Auth_ShapeFont_name>;
type _assert_ShapeFont_name = Expect<_check_ShapeFont_name>;
export function _write_ShapeFont_name(target: PptxEditor.ShapeFont, value: null | string): void {
  target['name'] = value;
}

// PowerPoint.ShapeFont#size
type Ref_ShapeFont_size = { size: null | number };
type Auth_ShapeFont_size = Pick<PptxEditor.ShapeFont, 'size'>;
type _check_ShapeFont_size = IsExact<Ref_ShapeFont_size, Auth_ShapeFont_size>;
type _assert_ShapeFont_size = Expect<_check_ShapeFont_size>;
export function _write_ShapeFont_size(target: PptxEditor.ShapeFont, value: null | number): void {
  target['size'] = value;
}

// PowerPoint.ShapeFont#smallCaps
type Ref_ShapeFont_smallCaps = { smallCaps: boolean | null };
type Auth_ShapeFont_smallCaps = Pick<PptxEditor.ShapeFont, 'smallCaps'>;
type _check_ShapeFont_smallCaps = IsExact<Ref_ShapeFont_smallCaps, Auth_ShapeFont_smallCaps>;
type _assert_ShapeFont_smallCaps = Expect<_check_ShapeFont_smallCaps>;
export function _write_ShapeFont_smallCaps(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['smallCaps'] = value;
}

// PowerPoint.ShapeFont#strikethrough
type Ref_ShapeFont_strikethrough = { strikethrough: boolean | null };
type Auth_ShapeFont_strikethrough = Pick<PptxEditor.ShapeFont, 'strikethrough'>;
type _check_ShapeFont_strikethrough = IsExact<Ref_ShapeFont_strikethrough, Auth_ShapeFont_strikethrough>;
type _assert_ShapeFont_strikethrough = Expect<_check_ShapeFont_strikethrough>;
export function _write_ShapeFont_strikethrough(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['strikethrough'] = value;
}

// PowerPoint.ShapeFont#subscript
type Ref_ShapeFont_subscript = { subscript: boolean | null };
type Auth_ShapeFont_subscript = Pick<PptxEditor.ShapeFont, 'subscript'>;
type _check_ShapeFont_subscript = IsExact<Ref_ShapeFont_subscript, Auth_ShapeFont_subscript>;
type _assert_ShapeFont_subscript = Expect<_check_ShapeFont_subscript>;
export function _write_ShapeFont_subscript(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['subscript'] = value;
}

// PowerPoint.ShapeFont#superscript
type Ref_ShapeFont_superscript = { superscript: boolean | null };
type Auth_ShapeFont_superscript = Pick<PptxEditor.ShapeFont, 'superscript'>;
type _check_ShapeFont_superscript = IsExact<Ref_ShapeFont_superscript, Auth_ShapeFont_superscript>;
type _assert_ShapeFont_superscript = Expect<_check_ShapeFont_superscript>;
export function _write_ShapeFont_superscript(target: PptxEditor.ShapeFont, value: boolean | null): void {
  target['superscript'] = value;
}

// PowerPoint.ShapeFont#underline
type Ref_ShapeFont_underline = { underline: "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Dotted" | "DottedHeavy" | "Double" | "Heavy" | "None" | "Single" | "Wavy" | "WavyDouble" | "WavyHeavy" | null | PptxEditor.ShapeFontUnderlineStyle };
type Auth_ShapeFont_underline = Pick<PptxEditor.ShapeFont, 'underline'>;
type _check_ShapeFont_underline = IsExact<Ref_ShapeFont_underline, Auth_ShapeFont_underline>;
type _assert_ShapeFont_underline = Expect<_check_ShapeFont_underline>;
export function _write_ShapeFont_underline(target: PptxEditor.ShapeFont, value: "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Dotted" | "DottedHeavy" | "Double" | "Heavy" | "None" | "Single" | "Wavy" | "WavyDouble" | "WavyHeavy" | null | PptxEditor.ShapeFontUnderlineStyle): void {
  target['underline'] = value;
}

// PowerPoint.PlaceholderFormat

// PowerPoint.PlaceholderFormat#containedType
// Not asserted: manifest.json conformanceExceptions "PowerPoint.PlaceholderFormat#containedType".
// The enum leaves out the one value that names a trademarked feature. The file stores that content
// as a diagram (dgm:), and the package reports it as Diagram. Every other value matches, and
// compat:report does not count this read-only member.

// PowerPoint.PlaceholderFormat#type
// Not asserted: manifest.json conformanceExceptions "PowerPoint.PlaceholderFormat#type".
// The enum leaves out the one value that names a trademarked feature. The file stores that content
// as a diagram (dgm:), and the package reports it as Diagram. Every other value matches, and
// compat:report does not count this read-only member.

// PowerPoint.TextFrame

// PowerPoint.TextFrame#autoSizeSetting
type Ref_TextFrame_autoSizeSetting = { autoSizeSetting: "AutoSizeMixed" | "AutoSizeNone" | "AutoSizeShapeToFitText" | "AutoSizeTextToFitShape" | PptxEditor.ShapeAutoSize };
type Auth_TextFrame_autoSizeSetting = Pick<PptxEditor.TextFrame, 'autoSizeSetting'>;
type _check_TextFrame_autoSizeSetting = IsExact<Ref_TextFrame_autoSizeSetting, Auth_TextFrame_autoSizeSetting>;
type _assert_TextFrame_autoSizeSetting = Expect<_check_TextFrame_autoSizeSetting>;
export function _write_TextFrame_autoSizeSetting(target: PptxEditor.TextFrame, value: "AutoSizeMixed" | "AutoSizeNone" | "AutoSizeShapeToFitText" | "AutoSizeTextToFitShape" | PptxEditor.ShapeAutoSize): void {
  target['autoSizeSetting'] = value;
}

// PowerPoint.TextFrame#bottomMargin
type Ref_TextFrame_bottomMargin = { bottomMargin: number };
type Auth_TextFrame_bottomMargin = Pick<PptxEditor.TextFrame, 'bottomMargin'>;
type _check_TextFrame_bottomMargin = IsExact<Ref_TextFrame_bottomMargin, Auth_TextFrame_bottomMargin>;
type _assert_TextFrame_bottomMargin = Expect<_check_TextFrame_bottomMargin>;
export function _write_TextFrame_bottomMargin(target: PptxEditor.TextFrame, value: number): void {
  target['bottomMargin'] = value;
}

// PowerPoint.TextFrame#deleteText
type Ref_TextFrame_deleteText = { (): void };
type Auth_TextFrame_deleteText = PptxEditor.TextFrame['deleteText'];
type _check_TextFrame_deleteText = IsExact<Ref_TextFrame_deleteText, Auth_TextFrame_deleteText>;
type _assert_TextFrame_deleteText = Expect<_check_TextFrame_deleteText>;

// PowerPoint.TextFrame#getParentShape
type Ref_TextFrame_getParentShape = { (): PptxEditor.Shape };
type Auth_TextFrame_getParentShape = PptxEditor.TextFrame['getParentShape'];
type _check_TextFrame_getParentShape = IsExact<Ref_TextFrame_getParentShape, Auth_TextFrame_getParentShape>;
type _assert_TextFrame_getParentShape = Expect<_check_TextFrame_getParentShape>;

// PowerPoint.TextFrame#hasText
type Ref_TextFrame_hasText = { readonly hasText: boolean };
type Auth_TextFrame_hasText = Pick<PptxEditor.TextFrame, 'hasText'>;
type _check_TextFrame_hasText = IsExact<Ref_TextFrame_hasText, Auth_TextFrame_hasText>;
type _assert_TextFrame_hasText = Expect<_check_TextFrame_hasText>;

// PowerPoint.TextFrame#leftMargin
type Ref_TextFrame_leftMargin = { leftMargin: number };
type Auth_TextFrame_leftMargin = Pick<PptxEditor.TextFrame, 'leftMargin'>;
type _check_TextFrame_leftMargin = IsExact<Ref_TextFrame_leftMargin, Auth_TextFrame_leftMargin>;
type _assert_TextFrame_leftMargin = Expect<_check_TextFrame_leftMargin>;
export function _write_TextFrame_leftMargin(target: PptxEditor.TextFrame, value: number): void {
  target['leftMargin'] = value;
}

// PowerPoint.TextFrame#rightMargin
type Ref_TextFrame_rightMargin = { rightMargin: number };
type Auth_TextFrame_rightMargin = Pick<PptxEditor.TextFrame, 'rightMargin'>;
type _check_TextFrame_rightMargin = IsExact<Ref_TextFrame_rightMargin, Auth_TextFrame_rightMargin>;
type _assert_TextFrame_rightMargin = Expect<_check_TextFrame_rightMargin>;
export function _write_TextFrame_rightMargin(target: PptxEditor.TextFrame, value: number): void {
  target['rightMargin'] = value;
}

// PowerPoint.TextFrame#textRange
type Ref_TextFrame_textRange = { readonly textRange: PptxEditor.TextRange };
type Auth_TextFrame_textRange = Pick<PptxEditor.TextFrame, 'textRange'>;
type _check_TextFrame_textRange = IsExact<Ref_TextFrame_textRange, Auth_TextFrame_textRange>;
type _assert_TextFrame_textRange = Expect<_check_TextFrame_textRange>;

// PowerPoint.TextFrame#topMargin
type Ref_TextFrame_topMargin = { topMargin: number };
type Auth_TextFrame_topMargin = Pick<PptxEditor.TextFrame, 'topMargin'>;
type _check_TextFrame_topMargin = IsExact<Ref_TextFrame_topMargin, Auth_TextFrame_topMargin>;
type _assert_TextFrame_topMargin = Expect<_check_TextFrame_topMargin>;
export function _write_TextFrame_topMargin(target: PptxEditor.TextFrame, value: number): void {
  target['topMargin'] = value;
}

// PowerPoint.TextFrame#verticalAlignment
type Ref_TextFrame_verticalAlignment = { verticalAlignment: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment };
type Auth_TextFrame_verticalAlignment = Pick<PptxEditor.TextFrame, 'verticalAlignment'>;
type _check_TextFrame_verticalAlignment = IsExact<Ref_TextFrame_verticalAlignment, Auth_TextFrame_verticalAlignment>;
type _assert_TextFrame_verticalAlignment = Expect<_check_TextFrame_verticalAlignment>;
export function _write_TextFrame_verticalAlignment(target: PptxEditor.TextFrame, value: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment): void {
  target['verticalAlignment'] = value;
}

// PowerPoint.TextFrame#wordWrap
type Ref_TextFrame_wordWrap = { wordWrap: boolean };
type Auth_TextFrame_wordWrap = Pick<PptxEditor.TextFrame, 'wordWrap'>;
type _check_TextFrame_wordWrap = IsExact<Ref_TextFrame_wordWrap, Auth_TextFrame_wordWrap>;
type _assert_TextFrame_wordWrap = Expect<_check_TextFrame_wordWrap>;
export function _write_TextFrame_wordWrap(target: PptxEditor.TextFrame, value: boolean): void {
  target['wordWrap'] = value;
}

// PowerPoint.TextRange

// PowerPoint.TextRange#font
type Ref_TextRange_font = { readonly font: PptxEditor.ShapeFont };
type Auth_TextRange_font = Pick<PptxEditor.TextRange, 'font'>;
type _check_TextRange_font = IsExact<Ref_TextRange_font, Auth_TextRange_font>;
type _assert_TextRange_font = Expect<_check_TextRange_font>;

// PowerPoint.TextRange#getParentTextFrame
type Ref_TextRange_getParentTextFrame = { (): PptxEditor.TextFrame };
type Auth_TextRange_getParentTextFrame = PptxEditor.TextRange['getParentTextFrame'];
type _check_TextRange_getParentTextFrame = IsExact<Ref_TextRange_getParentTextFrame, Auth_TextRange_getParentTextFrame>;
type _assert_TextRange_getParentTextFrame = Expect<_check_TextRange_getParentTextFrame>;

// PowerPoint.TextRange#getSubstring
type Ref_TextRange_getSubstring = { (start: number, length?: number | undefined): PptxEditor.TextRange };
type Auth_TextRange_getSubstring = PptxEditor.TextRange['getSubstring'];
type _check_TextRange_getSubstring = IsExact<Ref_TextRange_getSubstring, Auth_TextRange_getSubstring>;
type _assert_TextRange_getSubstring = Expect<_check_TextRange_getSubstring>;

// PowerPoint.TextRange#hyperlinks
type Ref_TextRange_hyperlinks = { readonly hyperlinks: PptxEditor.HyperlinkScopedCollection };
type Auth_TextRange_hyperlinks = Pick<PptxEditor.TextRange, 'hyperlinks'>;
type _check_TextRange_hyperlinks = IsExact<Ref_TextRange_hyperlinks, Auth_TextRange_hyperlinks>;
type _assert_TextRange_hyperlinks = Expect<_check_TextRange_hyperlinks>;

// PowerPoint.TextRange#length
type Ref_TextRange_length = { length: number };
type Auth_TextRange_length = Pick<PptxEditor.TextRange, 'length'>;
type _check_TextRange_length = IsExact<Ref_TextRange_length, Auth_TextRange_length>;
type _assert_TextRange_length = Expect<_check_TextRange_length>;
export function _write_TextRange_length(target: PptxEditor.TextRange, value: number): void {
  target['length'] = value;
}

// PowerPoint.TextRange#paragraphFormat
type Ref_TextRange_paragraphFormat = { readonly paragraphFormat: PptxEditor.ParagraphFormat };
type Auth_TextRange_paragraphFormat = Pick<PptxEditor.TextRange, 'paragraphFormat'>;
type _check_TextRange_paragraphFormat = IsExact<Ref_TextRange_paragraphFormat, Auth_TextRange_paragraphFormat>;
type _assert_TextRange_paragraphFormat = Expect<_check_TextRange_paragraphFormat>;

// PowerPoint.TextRange#setHyperlink
type Ref_TextRange_setHyperlink = { (options?: PptxEditor.HyperlinkAddOptions | undefined): PptxEditor.Hyperlink };
type Auth_TextRange_setHyperlink = PptxEditor.TextRange['setHyperlink'];
type _check_TextRange_setHyperlink = IsExact<Ref_TextRange_setHyperlink, Auth_TextRange_setHyperlink>;
type _assert_TextRange_setHyperlink = Expect<_check_TextRange_setHyperlink>;

// PowerPoint.TextRange#setSelected
type Ref_TextRange_setSelected = { (): void };
type Auth_TextRange_setSelected = PptxEditor.TextRange['setSelected'];
type _check_TextRange_setSelected = IsExact<Ref_TextRange_setSelected, Auth_TextRange_setSelected>;
type _assert_TextRange_setSelected = Expect<_check_TextRange_setSelected>;

// PowerPoint.TextRange#start
type Ref_TextRange_start = { start: number };
type Auth_TextRange_start = Pick<PptxEditor.TextRange, 'start'>;
type _check_TextRange_start = IsExact<Ref_TextRange_start, Auth_TextRange_start>;
type _assert_TextRange_start = Expect<_check_TextRange_start>;
export function _write_TextRange_start(target: PptxEditor.TextRange, value: number): void {
  target['start'] = value;
}

// PowerPoint.TextRange#text
type Ref_TextRange_text = { text: string };
type Auth_TextRange_text = Pick<PptxEditor.TextRange, 'text'>;
type _check_TextRange_text = IsExact<Ref_TextRange_text, Auth_TextRange_text>;
type _assert_TextRange_text = Expect<_check_TextRange_text>;
export function _write_TextRange_text(target: PptxEditor.TextRange, value: string): void {
  target['text'] = value;
}

// PowerPoint.ParagraphFormat

// PowerPoint.ParagraphFormat#bulletFormat
type Ref_ParagraphFormat_bulletFormat = { readonly bulletFormat: PptxEditor.BulletFormat };
type Auth_ParagraphFormat_bulletFormat = Pick<PptxEditor.ParagraphFormat, 'bulletFormat'>;
type _check_ParagraphFormat_bulletFormat = IsExact<Ref_ParagraphFormat_bulletFormat, Auth_ParagraphFormat_bulletFormat>;
type _assert_ParagraphFormat_bulletFormat = Expect<_check_ParagraphFormat_bulletFormat>;

// PowerPoint.ParagraphFormat#horizontalAlignment
type Ref_ParagraphFormat_horizontalAlignment = { horizontalAlignment: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | null | PptxEditor.ParagraphHorizontalAlignment };
type Auth_ParagraphFormat_horizontalAlignment = Pick<PptxEditor.ParagraphFormat, 'horizontalAlignment'>;
type _check_ParagraphFormat_horizontalAlignment = IsExact<Ref_ParagraphFormat_horizontalAlignment, Auth_ParagraphFormat_horizontalAlignment>;
type _assert_ParagraphFormat_horizontalAlignment = Expect<_check_ParagraphFormat_horizontalAlignment>;
export function _write_ParagraphFormat_horizontalAlignment(target: PptxEditor.ParagraphFormat, value: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | null | PptxEditor.ParagraphHorizontalAlignment): void {
  target['horizontalAlignment'] = value;
}

// PowerPoint.ParagraphFormat#indentLevel
type Ref_ParagraphFormat_indentLevel = { indentLevel: number };
type Auth_ParagraphFormat_indentLevel = Pick<PptxEditor.ParagraphFormat, 'indentLevel'>;
type _check_ParagraphFormat_indentLevel = IsExact<Ref_ParagraphFormat_indentLevel, Auth_ParagraphFormat_indentLevel>;
type _assert_ParagraphFormat_indentLevel = Expect<_check_ParagraphFormat_indentLevel>;
export function _write_ParagraphFormat_indentLevel(target: PptxEditor.ParagraphFormat, value: number): void {
  target['indentLevel'] = value;
}

// PowerPoint.BulletFormat

// PowerPoint.BulletFormat#style
type Ref_BulletFormat_style = { style: "AlphabetLowercaseParenthesesBoth" | "AlphabetLowercaseParenthesisRight" | "AlphabetLowercasePeriod" | "AlphabetUppercaseParenthesesBoth" | "AlphabetUppercaseParenthesisRight" | "AlphabetUppercasePeriod" | "ArabicAbjadDash" | "ArabicAlphabetDash" | "ArabicDoubleBytePeriod" | "ArabicDoubleBytePlain" | "ArabicNumeralParenthesesBoth" | "ArabicNumeralParenthesisRight" | "ArabicNumeralPeriod" | "ArabicNumeralPlain" | "CircleNumberDoubleBytePlain" | "CircleNumberWideDoubleByteBlackPlain" | "CircleNumberWideDoubleByteWhitePlain" | "HebrewAlphabetDash" | "HindiAlphabet1Period" | "HindiAlphabetPeriod" | "HindiNumeralParenthesisRight" | "HindiNumeralPeriod" | "KanjiKoreanPeriod" | "KanjiKoreanPlain" | "KanjiSimplifiedChineseDoubleBytePeriod" | "RomanLowercaseParenthesesBoth" | "RomanLowercaseParenthesisRight" | "RomanLowercasePeriod" | "RomanUppercaseParenthesesBoth" | "RomanUppercaseParenthesisRight" | "RomanUppercasePeriod" | "SimplifiedChinesePeriod" | "SimplifiedChinesePlain" | "ThaiAlphabetParenthesesBoth" | "ThaiAlphabetParenthesisRight" | "ThaiAlphabetPeriod" | "ThaiNumeralParenthesesBoth" | "ThaiNumeralParenthesisRight" | "ThaiNumeralPeriod" | "TraditionalChinesePeriod" | "TraditionalChinesePlain" | "Unsupported" | PptxEditor.BulletStyle | null };
type Auth_BulletFormat_style = Pick<PptxEditor.BulletFormat, 'style'>;
type _check_BulletFormat_style = IsExact<Ref_BulletFormat_style, Auth_BulletFormat_style>;
type _assert_BulletFormat_style = Expect<_check_BulletFormat_style>;
export function _write_BulletFormat_style(target: PptxEditor.BulletFormat, value: "AlphabetLowercaseParenthesesBoth" | "AlphabetLowercaseParenthesisRight" | "AlphabetLowercasePeriod" | "AlphabetUppercaseParenthesesBoth" | "AlphabetUppercaseParenthesisRight" | "AlphabetUppercasePeriod" | "ArabicAbjadDash" | "ArabicAlphabetDash" | "ArabicDoubleBytePeriod" | "ArabicDoubleBytePlain" | "ArabicNumeralParenthesesBoth" | "ArabicNumeralParenthesisRight" | "ArabicNumeralPeriod" | "ArabicNumeralPlain" | "CircleNumberDoubleBytePlain" | "CircleNumberWideDoubleByteBlackPlain" | "CircleNumberWideDoubleByteWhitePlain" | "HebrewAlphabetDash" | "HindiAlphabet1Period" | "HindiAlphabetPeriod" | "HindiNumeralParenthesisRight" | "HindiNumeralPeriod" | "KanjiKoreanPeriod" | "KanjiKoreanPlain" | "KanjiSimplifiedChineseDoubleBytePeriod" | "RomanLowercaseParenthesesBoth" | "RomanLowercaseParenthesisRight" | "RomanLowercasePeriod" | "RomanUppercaseParenthesesBoth" | "RomanUppercaseParenthesisRight" | "RomanUppercasePeriod" | "SimplifiedChinesePeriod" | "SimplifiedChinesePlain" | "ThaiAlphabetParenthesesBoth" | "ThaiAlphabetParenthesisRight" | "ThaiAlphabetPeriod" | "ThaiNumeralParenthesesBoth" | "ThaiNumeralParenthesisRight" | "ThaiNumeralPeriod" | "TraditionalChinesePeriod" | "TraditionalChinesePlain" | "Unsupported" | PptxEditor.BulletStyle | null): void {
  target['style'] = value;
}

// PowerPoint.BulletFormat#type
type Ref_BulletFormat_type = { type: "None" | "Numbered" | "Unnumbered" | "Unsupported" | PptxEditor.BulletType | null };
type Auth_BulletFormat_type = Pick<PptxEditor.BulletFormat, 'type'>;
type _check_BulletFormat_type = IsExact<Ref_BulletFormat_type, Auth_BulletFormat_type>;
type _assert_BulletFormat_type = Expect<_check_BulletFormat_type>;
export function _write_BulletFormat_type(target: PptxEditor.BulletFormat, value: "None" | "Numbered" | "Unnumbered" | "Unsupported" | PptxEditor.BulletType | null): void {
  target['type'] = value;
}

// PowerPoint.BulletFormat#visible
type Ref_BulletFormat_visible = { visible: boolean | null };
type Auth_BulletFormat_visible = Pick<PptxEditor.BulletFormat, 'visible'>;
type _check_BulletFormat_visible = IsExact<Ref_BulletFormat_visible, Auth_BulletFormat_visible>;
type _assert_BulletFormat_visible = Expect<_check_BulletFormat_visible>;
export function _write_BulletFormat_visible(target: PptxEditor.BulletFormat, value: boolean | null): void {
  target['visible'] = value;
}

// PowerPoint.Table

// PowerPoint.Table#clear
type Ref_Table_clear = { (options?: PptxEditor.TableClearOptions | undefined): void };
type Auth_Table_clear = PptxEditor.Table['clear'];
type _check_Table_clear = IsExact<Ref_Table_clear, Auth_Table_clear>;
type _assert_Table_clear = Expect<_check_Table_clear>;

// PowerPoint.Table#columnCount
type Ref_Table_columnCount = { readonly columnCount: number };
type Auth_Table_columnCount = Pick<PptxEditor.Table, 'columnCount'>;
type _check_Table_columnCount = IsExact<Ref_Table_columnCount, Auth_Table_columnCount>;
type _assert_Table_columnCount = Expect<_check_Table_columnCount>;

// PowerPoint.Table#columns
type Ref_Table_columns = { readonly columns: PptxEditor.TableColumnCollection };
type Auth_Table_columns = Pick<PptxEditor.Table, 'columns'>;
type _check_Table_columns = IsExact<Ref_Table_columns, Auth_Table_columns>;
type _assert_Table_columns = Expect<_check_Table_columns>;

// PowerPoint.Table#getCellOrNullObject
type Ref_Table_getCellOrNullObject = { (rowIndex: number, columnIndex: number): PptxEditor.TableCell };
type Auth_Table_getCellOrNullObject = PptxEditor.Table['getCellOrNullObject'];
type _check_Table_getCellOrNullObject = IsExact<Ref_Table_getCellOrNullObject, Auth_Table_getCellOrNullObject>;
type _assert_Table_getCellOrNullObject = Expect<_check_Table_getCellOrNullObject>;

// PowerPoint.Table#getMergedAreas
type Ref_Table_getMergedAreas = { (): PptxEditor.TableCellCollection };
type Auth_Table_getMergedAreas = PptxEditor.Table['getMergedAreas'];
type _check_Table_getMergedAreas = IsExact<Ref_Table_getMergedAreas, Auth_Table_getMergedAreas>;
type _assert_Table_getMergedAreas = Expect<_check_Table_getMergedAreas>;

// PowerPoint.Table#getShape
type Ref_Table_getShape = { (): PptxEditor.Shape };
type Auth_Table_getShape = PptxEditor.Table['getShape'];
type _check_Table_getShape = IsExact<Ref_Table_getShape, Auth_Table_getShape>;
type _assert_Table_getShape = Expect<_check_Table_getShape>;

// PowerPoint.Table#mergeCells
type Ref_Table_mergeCells = { (rowIndex: number, columnIndex: number, rowCount: number, columnCount: number): void };
type Auth_Table_mergeCells = PptxEditor.Table['mergeCells'];
type _check_Table_mergeCells = IsExact<Ref_Table_mergeCells, Auth_Table_mergeCells>;
type _assert_Table_mergeCells = Expect<_check_Table_mergeCells>;

// PowerPoint.Table#rowCount
type Ref_Table_rowCount = { readonly rowCount: number };
type Auth_Table_rowCount = Pick<PptxEditor.Table, 'rowCount'>;
type _check_Table_rowCount = IsExact<Ref_Table_rowCount, Auth_Table_rowCount>;
type _assert_Table_rowCount = Expect<_check_Table_rowCount>;

// PowerPoint.Table#rows
type Ref_Table_rows = { readonly rows: PptxEditor.TableRowCollection };
type Auth_Table_rows = Pick<PptxEditor.Table, 'rows'>;
type _check_Table_rows = IsExact<Ref_Table_rows, Auth_Table_rows>;
type _assert_Table_rows = Expect<_check_Table_rows>;

// PowerPoint.Table#styleSettings
type Ref_Table_styleSettings = { readonly styleSettings: PptxEditor.TableStyleSettings };
type Auth_Table_styleSettings = Pick<PptxEditor.Table, 'styleSettings'>;
type _check_Table_styleSettings = IsExact<Ref_Table_styleSettings, Auth_Table_styleSettings>;
type _assert_Table_styleSettings = Expect<_check_Table_styleSettings>;

// PowerPoint.Table#values
type Ref_Table_values = { readonly values: string[][] };
type Auth_Table_values = Pick<PptxEditor.Table, 'values'>;
type _check_Table_values = IsExact<Ref_Table_values, Auth_Table_values>;
type _assert_Table_values = Expect<_check_Table_values>;

// PowerPoint.TableCell

// PowerPoint.TableCell#borders
type Ref_TableCell_borders = { readonly borders: PptxEditor.Borders };
type Auth_TableCell_borders = Pick<PptxEditor.TableCell, 'borders'>;
type _check_TableCell_borders = IsExact<Ref_TableCell_borders, Auth_TableCell_borders>;
type _assert_TableCell_borders = Expect<_check_TableCell_borders>;

// PowerPoint.TableCell#columnCount
type Ref_TableCell_columnCount = { readonly columnCount: number };
type Auth_TableCell_columnCount = Pick<PptxEditor.TableCell, 'columnCount'>;
type _check_TableCell_columnCount = IsExact<Ref_TableCell_columnCount, Auth_TableCell_columnCount>;
type _assert_TableCell_columnCount = Expect<_check_TableCell_columnCount>;

// PowerPoint.TableCell#columnIndex
type Ref_TableCell_columnIndex = { readonly columnIndex: number };
type Auth_TableCell_columnIndex = Pick<PptxEditor.TableCell, 'columnIndex'>;
type _check_TableCell_columnIndex = IsExact<Ref_TableCell_columnIndex, Auth_TableCell_columnIndex>;
type _assert_TableCell_columnIndex = Expect<_check_TableCell_columnIndex>;

// PowerPoint.TableCell#fill
type Ref_TableCell_fill = { readonly fill: PptxEditor.ShapeFill };
type Auth_TableCell_fill = Pick<PptxEditor.TableCell, 'fill'>;
type _check_TableCell_fill = IsExact<Ref_TableCell_fill, Auth_TableCell_fill>;
type _assert_TableCell_fill = Expect<_check_TableCell_fill>;

// PowerPoint.TableCell#font
type Ref_TableCell_font = { readonly font: PptxEditor.ShapeFont };
type Auth_TableCell_font = Pick<PptxEditor.TableCell, 'font'>;
type _check_TableCell_font = IsExact<Ref_TableCell_font, Auth_TableCell_font>;
type _assert_TableCell_font = Expect<_check_TableCell_font>;

// PowerPoint.TableCell#horizontalAlignment
type Ref_TableCell_horizontalAlignment = { horizontalAlignment: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | null | PptxEditor.ParagraphHorizontalAlignment };
type Auth_TableCell_horizontalAlignment = Pick<PptxEditor.TableCell, 'horizontalAlignment'>;
type _check_TableCell_horizontalAlignment = IsExact<Ref_TableCell_horizontalAlignment, Auth_TableCell_horizontalAlignment>;
type _assert_TableCell_horizontalAlignment = Expect<_check_TableCell_horizontalAlignment>;
export function _write_TableCell_horizontalAlignment(target: PptxEditor.TableCell, value: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | null | PptxEditor.ParagraphHorizontalAlignment): void {
  target['horizontalAlignment'] = value;
}

// PowerPoint.TableCell#indentLevel
type Ref_TableCell_indentLevel = { indentLevel: null | number };
type Auth_TableCell_indentLevel = Pick<PptxEditor.TableCell, 'indentLevel'>;
type _check_TableCell_indentLevel = IsExact<Ref_TableCell_indentLevel, Auth_TableCell_indentLevel>;
type _assert_TableCell_indentLevel = Expect<_check_TableCell_indentLevel>;
export function _write_TableCell_indentLevel(target: PptxEditor.TableCell, value: null | number): void {
  target['indentLevel'] = value;
}

// PowerPoint.TableCell#margins
type Ref_TableCell_margins = { readonly margins: PptxEditor.Margins };
type Auth_TableCell_margins = Pick<PptxEditor.TableCell, 'margins'>;
type _check_TableCell_margins = IsExact<Ref_TableCell_margins, Auth_TableCell_margins>;
type _assert_TableCell_margins = Expect<_check_TableCell_margins>;

// PowerPoint.TableCell#resize
type Ref_TableCell_resize = { (rowCount: number, columnCount: number): void };
type Auth_TableCell_resize = PptxEditor.TableCell['resize'];
type _check_TableCell_resize = IsExact<Ref_TableCell_resize, Auth_TableCell_resize>;
type _assert_TableCell_resize = Expect<_check_TableCell_resize>;

// PowerPoint.TableCell#rowCount
type Ref_TableCell_rowCount = { readonly rowCount: number };
type Auth_TableCell_rowCount = Pick<PptxEditor.TableCell, 'rowCount'>;
type _check_TableCell_rowCount = IsExact<Ref_TableCell_rowCount, Auth_TableCell_rowCount>;
type _assert_TableCell_rowCount = Expect<_check_TableCell_rowCount>;

// PowerPoint.TableCell#rowIndex
type Ref_TableCell_rowIndex = { readonly rowIndex: number };
type Auth_TableCell_rowIndex = Pick<PptxEditor.TableCell, 'rowIndex'>;
type _check_TableCell_rowIndex = IsExact<Ref_TableCell_rowIndex, Auth_TableCell_rowIndex>;
type _assert_TableCell_rowIndex = Expect<_check_TableCell_rowIndex>;

// PowerPoint.TableCell#split
type Ref_TableCell_split = { (rowCount: number, columnCount: number): void };
type Auth_TableCell_split = PptxEditor.TableCell['split'];
type _check_TableCell_split = IsExact<Ref_TableCell_split, Auth_TableCell_split>;
type _assert_TableCell_split = Expect<_check_TableCell_split>;

// PowerPoint.TableCell#text
type Ref_TableCell_text = { text: string };
type Auth_TableCell_text = Pick<PptxEditor.TableCell, 'text'>;
type _check_TableCell_text = IsExact<Ref_TableCell_text, Auth_TableCell_text>;
type _assert_TableCell_text = Expect<_check_TableCell_text>;
export function _write_TableCell_text(target: PptxEditor.TableCell, value: string): void {
  target['text'] = value;
}

// PowerPoint.TableCell#textRuns
type Ref_TableCell_textRuns = { textRuns: PptxEditor.TextRun[] };
type Auth_TableCell_textRuns = Pick<PptxEditor.TableCell, 'textRuns'>;
type _check_TableCell_textRuns = IsExact<Ref_TableCell_textRuns, Auth_TableCell_textRuns>;
type _assert_TableCell_textRuns = Expect<_check_TableCell_textRuns>;
export function _write_TableCell_textRuns(target: PptxEditor.TableCell, value: PptxEditor.TextRun[]): void {
  target['textRuns'] = value;
}

// PowerPoint.TableCell#verticalAlignment
type Ref_TableCell_verticalAlignment = { verticalAlignment: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment };
type Auth_TableCell_verticalAlignment = Pick<PptxEditor.TableCell, 'verticalAlignment'>;
type _check_TableCell_verticalAlignment = IsExact<Ref_TableCell_verticalAlignment, Auth_TableCell_verticalAlignment>;
type _assert_TableCell_verticalAlignment = Expect<_check_TableCell_verticalAlignment>;
export function _write_TableCell_verticalAlignment(target: PptxEditor.TableCell, value: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment): void {
  target['verticalAlignment'] = value;
}

// PowerPoint.TableRow

// PowerPoint.TableRow#currentHeight
type Ref_TableRow_currentHeight = { readonly currentHeight: number };
type Auth_TableRow_currentHeight = Pick<PptxEditor.TableRow, 'currentHeight'>;
type _check_TableRow_currentHeight = IsExact<Ref_TableRow_currentHeight, Auth_TableRow_currentHeight>;
type _assert_TableRow_currentHeight = Expect<_check_TableRow_currentHeight>;

// PowerPoint.TableRow#delete
type Ref_TableRow_delete = { (): void };
type Auth_TableRow_delete = PptxEditor.TableRow['delete'];
type _check_TableRow_delete = IsExact<Ref_TableRow_delete, Auth_TableRow_delete>;
type _assert_TableRow_delete = Expect<_check_TableRow_delete>;

// PowerPoint.TableRow#height
type Ref_TableRow_height = { height: number };
type Auth_TableRow_height = Pick<PptxEditor.TableRow, 'height'>;
type _check_TableRow_height = IsExact<Ref_TableRow_height, Auth_TableRow_height>;
type _assert_TableRow_height = Expect<_check_TableRow_height>;
export function _write_TableRow_height(target: PptxEditor.TableRow, value: number): void {
  target['height'] = value;
}

// PowerPoint.TableRow#rowIndex
type Ref_TableRow_rowIndex = { readonly rowIndex: number };
type Auth_TableRow_rowIndex = Pick<PptxEditor.TableRow, 'rowIndex'>;
type _check_TableRow_rowIndex = IsExact<Ref_TableRow_rowIndex, Auth_TableRow_rowIndex>;
type _assert_TableRow_rowIndex = Expect<_check_TableRow_rowIndex>;

// PowerPoint.TableRowCollection

// PowerPoint.TableRowCollection#add
type Ref_TableRowCollection_add = { (index?: null | number | undefined, count?: number | undefined): void };
type Auth_TableRowCollection_add = PptxEditor.TableRowCollection['add'];
type _check_TableRowCollection_add = IsExact<Ref_TableRowCollection_add, Auth_TableRowCollection_add>;
type _assert_TableRowCollection_add = Expect<_check_TableRowCollection_add>;

// PowerPoint.TableRowCollection#deleteRows
type Ref_TableRowCollection_deleteRows = { (rows: PptxEditor.TableRow[]): void };
type Auth_TableRowCollection_deleteRows = PptxEditor.TableRowCollection['deleteRows'];
type _check_TableRowCollection_deleteRows = IsExact<Ref_TableRowCollection_deleteRows, Auth_TableRowCollection_deleteRows>;
type _assert_TableRowCollection_deleteRows = Expect<_check_TableRowCollection_deleteRows>;

// PowerPoint.TableRowCollection#getCount
type Ref_TableRowCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_TableRowCollection_getCount = PptxEditor.TableRowCollection['getCount'];
type _check_TableRowCollection_getCount = IsExact<Ref_TableRowCollection_getCount, Auth_TableRowCollection_getCount>;
type _assert_TableRowCollection_getCount = Expect<_check_TableRowCollection_getCount>;

// PowerPoint.TableRowCollection#getItemAt
type Ref_TableRowCollection_getItemAt = { (index: number): PptxEditor.TableRow };
type Auth_TableRowCollection_getItemAt = PptxEditor.TableRowCollection['getItemAt'];
type _check_TableRowCollection_getItemAt = IsExact<Ref_TableRowCollection_getItemAt, Auth_TableRowCollection_getItemAt>;
type _assert_TableRowCollection_getItemAt = Expect<_check_TableRowCollection_getItemAt>;

// PowerPoint.TableRowCollection#items
type Ref_TableRowCollection_items = { readonly items: PptxEditor.TableRow[] };
type Auth_TableRowCollection_items = Pick<PptxEditor.TableRowCollection, 'items'>;
type _check_TableRowCollection_items = IsExact<Ref_TableRowCollection_items, Auth_TableRowCollection_items>;
type _assert_TableRowCollection_items = Expect<_check_TableRowCollection_items>;

// PowerPoint.TableColumn

// PowerPoint.TableColumn#columnIndex
type Ref_TableColumn_columnIndex = { readonly columnIndex: number };
type Auth_TableColumn_columnIndex = Pick<PptxEditor.TableColumn, 'columnIndex'>;
type _check_TableColumn_columnIndex = IsExact<Ref_TableColumn_columnIndex, Auth_TableColumn_columnIndex>;
type _assert_TableColumn_columnIndex = Expect<_check_TableColumn_columnIndex>;

// PowerPoint.TableColumn#delete
type Ref_TableColumn_delete = { (): void };
type Auth_TableColumn_delete = PptxEditor.TableColumn['delete'];
type _check_TableColumn_delete = IsExact<Ref_TableColumn_delete, Auth_TableColumn_delete>;
type _assert_TableColumn_delete = Expect<_check_TableColumn_delete>;

// PowerPoint.TableColumn#width
type Ref_TableColumn_width = { width: number };
type Auth_TableColumn_width = Pick<PptxEditor.TableColumn, 'width'>;
type _check_TableColumn_width = IsExact<Ref_TableColumn_width, Auth_TableColumn_width>;
type _assert_TableColumn_width = Expect<_check_TableColumn_width>;
export function _write_TableColumn_width(target: PptxEditor.TableColumn, value: number): void {
  target['width'] = value;
}

// PowerPoint.TableColumnCollection

// PowerPoint.TableColumnCollection#add
type Ref_TableColumnCollection_add = { (index?: null | number | undefined, count?: number | undefined): void };
type Auth_TableColumnCollection_add = PptxEditor.TableColumnCollection['add'];
type _check_TableColumnCollection_add = IsExact<Ref_TableColumnCollection_add, Auth_TableColumnCollection_add>;
type _assert_TableColumnCollection_add = Expect<_check_TableColumnCollection_add>;

// PowerPoint.TableColumnCollection#deleteColumns
type Ref_TableColumnCollection_deleteColumns = { (columns: PptxEditor.TableColumn[]): void };
type Auth_TableColumnCollection_deleteColumns = PptxEditor.TableColumnCollection['deleteColumns'];
type _check_TableColumnCollection_deleteColumns = IsExact<Ref_TableColumnCollection_deleteColumns, Auth_TableColumnCollection_deleteColumns>;
type _assert_TableColumnCollection_deleteColumns = Expect<_check_TableColumnCollection_deleteColumns>;

// PowerPoint.TableColumnCollection#getCount
type Ref_TableColumnCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_TableColumnCollection_getCount = PptxEditor.TableColumnCollection['getCount'];
type _check_TableColumnCollection_getCount = IsExact<Ref_TableColumnCollection_getCount, Auth_TableColumnCollection_getCount>;
type _assert_TableColumnCollection_getCount = Expect<_check_TableColumnCollection_getCount>;

// PowerPoint.TableColumnCollection#getItemAt
type Ref_TableColumnCollection_getItemAt = { (index: number): PptxEditor.TableColumn };
type Auth_TableColumnCollection_getItemAt = PptxEditor.TableColumnCollection['getItemAt'];
type _check_TableColumnCollection_getItemAt = IsExact<Ref_TableColumnCollection_getItemAt, Auth_TableColumnCollection_getItemAt>;
type _assert_TableColumnCollection_getItemAt = Expect<_check_TableColumnCollection_getItemAt>;

// PowerPoint.TableColumnCollection#items
type Ref_TableColumnCollection_items = { readonly items: PptxEditor.TableColumn[] };
type Auth_TableColumnCollection_items = Pick<PptxEditor.TableColumnCollection, 'items'>;
type _check_TableColumnCollection_items = IsExact<Ref_TableColumnCollection_items, Auth_TableColumnCollection_items>;
type _assert_TableColumnCollection_items = Expect<_check_TableColumnCollection_items>;

// PowerPoint.Hyperlink

// PowerPoint.Hyperlink#address
type Ref_Hyperlink_address = { address: string };
type Auth_Hyperlink_address = Pick<PptxEditor.Hyperlink, 'address'>;
type _check_Hyperlink_address = IsExact<Ref_Hyperlink_address, Auth_Hyperlink_address>;
type _assert_Hyperlink_address = Expect<_check_Hyperlink_address>;
export function _write_Hyperlink_address(target: PptxEditor.Hyperlink, value: string): void {
  target['address'] = value;
}

// PowerPoint.Hyperlink#delete
type Ref_Hyperlink_delete = { (): void };
type Auth_Hyperlink_delete = PptxEditor.Hyperlink['delete'];
type _check_Hyperlink_delete = IsExact<Ref_Hyperlink_delete, Auth_Hyperlink_delete>;
type _assert_Hyperlink_delete = Expect<_check_Hyperlink_delete>;

// PowerPoint.Hyperlink#getLinkedShapeOrNullObject
type Ref_Hyperlink_getLinkedShapeOrNullObject = { (): PptxEditor.Shape };
type Auth_Hyperlink_getLinkedShapeOrNullObject = PptxEditor.Hyperlink['getLinkedShapeOrNullObject'];
type _check_Hyperlink_getLinkedShapeOrNullObject = IsExact<Ref_Hyperlink_getLinkedShapeOrNullObject, Auth_Hyperlink_getLinkedShapeOrNullObject>;
type _assert_Hyperlink_getLinkedShapeOrNullObject = Expect<_check_Hyperlink_getLinkedShapeOrNullObject>;

// PowerPoint.Hyperlink#getLinkedTextRangeOrNullObject
type Ref_Hyperlink_getLinkedTextRangeOrNullObject = { (): PptxEditor.TextRange };
type Auth_Hyperlink_getLinkedTextRangeOrNullObject = PptxEditor.Hyperlink['getLinkedTextRangeOrNullObject'];
type _check_Hyperlink_getLinkedTextRangeOrNullObject = IsExact<Ref_Hyperlink_getLinkedTextRangeOrNullObject, Auth_Hyperlink_getLinkedTextRangeOrNullObject>;
type _assert_Hyperlink_getLinkedTextRangeOrNullObject = Expect<_check_Hyperlink_getLinkedTextRangeOrNullObject>;

// PowerPoint.Hyperlink#screenTip
type Ref_Hyperlink_screenTip = { screenTip: string };
type Auth_Hyperlink_screenTip = Pick<PptxEditor.Hyperlink, 'screenTip'>;
type _check_Hyperlink_screenTip = IsExact<Ref_Hyperlink_screenTip, Auth_Hyperlink_screenTip>;
type _assert_Hyperlink_screenTip = Expect<_check_Hyperlink_screenTip>;
export function _write_Hyperlink_screenTip(target: PptxEditor.Hyperlink, value: string): void {
  target['screenTip'] = value;
}

// PowerPoint.Hyperlink#type
type Ref_Hyperlink_type = { readonly type: "Shape" | "TextRange" | PptxEditor.HyperlinkType };
type Auth_Hyperlink_type = Pick<PptxEditor.Hyperlink, 'type'>;
type _check_Hyperlink_type = IsExact<Ref_Hyperlink_type, Auth_Hyperlink_type>;
type _assert_Hyperlink_type = Expect<_check_Hyperlink_type>;

// PowerPoint.HyperlinkCollection

// PowerPoint.HyperlinkCollection#add
type Ref_HyperlinkCollection_add = { (target: PptxEditor.Shape | PptxEditor.TextRange, options?: PptxEditor.HyperlinkAddOptions | undefined): PptxEditor.Hyperlink };
type Auth_HyperlinkCollection_add = PptxEditor.HyperlinkCollection['add'];
type _check_HyperlinkCollection_add = IsExact<Ref_HyperlinkCollection_add, Auth_HyperlinkCollection_add>;
type _assert_HyperlinkCollection_add = Expect<_check_HyperlinkCollection_add>;

// PowerPoint.HyperlinkCollection#getCount
type Ref_HyperlinkCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_HyperlinkCollection_getCount = PptxEditor.HyperlinkCollection['getCount'];
type _check_HyperlinkCollection_getCount = IsExact<Ref_HyperlinkCollection_getCount, Auth_HyperlinkCollection_getCount>;
type _assert_HyperlinkCollection_getCount = Expect<_check_HyperlinkCollection_getCount>;

// PowerPoint.HyperlinkCollection#getItemAt
type Ref_HyperlinkCollection_getItemAt = { (index: number): PptxEditor.Hyperlink };
type Auth_HyperlinkCollection_getItemAt = PptxEditor.HyperlinkCollection['getItemAt'];
type _check_HyperlinkCollection_getItemAt = IsExact<Ref_HyperlinkCollection_getItemAt, Auth_HyperlinkCollection_getItemAt>;
type _assert_HyperlinkCollection_getItemAt = Expect<_check_HyperlinkCollection_getItemAt>;

// PowerPoint.HyperlinkCollection#items
type Ref_HyperlinkCollection_items = { readonly items: PptxEditor.Hyperlink[] };
type Auth_HyperlinkCollection_items = Pick<PptxEditor.HyperlinkCollection, 'items'>;
type _check_HyperlinkCollection_items = IsExact<Ref_HyperlinkCollection_items, Auth_HyperlinkCollection_items>;
type _assert_HyperlinkCollection_items = Expect<_check_HyperlinkCollection_items>;

// PowerPoint.HyperlinkScopedCollection

// PowerPoint.HyperlinkScopedCollection#getCount
type Ref_HyperlinkScopedCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_HyperlinkScopedCollection_getCount = PptxEditor.HyperlinkScopedCollection['getCount'];
type _check_HyperlinkScopedCollection_getCount = IsExact<Ref_HyperlinkScopedCollection_getCount, Auth_HyperlinkScopedCollection_getCount>;
type _assert_HyperlinkScopedCollection_getCount = Expect<_check_HyperlinkScopedCollection_getCount>;

// PowerPoint.HyperlinkScopedCollection#getItemAt
type Ref_HyperlinkScopedCollection_getItemAt = { (index: number): PptxEditor.Hyperlink };
type Auth_HyperlinkScopedCollection_getItemAt = PptxEditor.HyperlinkScopedCollection['getItemAt'];
type _check_HyperlinkScopedCollection_getItemAt = IsExact<Ref_HyperlinkScopedCollection_getItemAt, Auth_HyperlinkScopedCollection_getItemAt>;
type _assert_HyperlinkScopedCollection_getItemAt = Expect<_check_HyperlinkScopedCollection_getItemAt>;

// PowerPoint.HyperlinkScopedCollection#items
type Ref_HyperlinkScopedCollection_items = { readonly items: PptxEditor.Hyperlink[] };
type Auth_HyperlinkScopedCollection_items = Pick<PptxEditor.HyperlinkScopedCollection, 'items'>;
type _check_HyperlinkScopedCollection_items = IsExact<Ref_HyperlinkScopedCollection_items, Auth_HyperlinkScopedCollection_items>;
type _assert_HyperlinkScopedCollection_items = Expect<_check_HyperlinkScopedCollection_items>;

// PowerPoint.Tag

// PowerPoint.Tag#key
type Ref_Tag_key = { readonly key: string };
type Auth_Tag_key = Pick<PptxEditor.Tag, 'key'>;
type _check_Tag_key = IsExact<Ref_Tag_key, Auth_Tag_key>;
type _assert_Tag_key = Expect<_check_Tag_key>;

// PowerPoint.Tag#value
type Ref_Tag_value = { value: string };
type Auth_Tag_value = Pick<PptxEditor.Tag, 'value'>;
type _check_Tag_value = IsExact<Ref_Tag_value, Auth_Tag_value>;
type _assert_Tag_value = Expect<_check_Tag_value>;
export function _write_Tag_value(target: PptxEditor.Tag, value: string): void {
  target['value'] = value;
}

// PowerPoint.Binding

// PowerPoint.Binding#delete
type Ref_Binding_delete = { (): void };
type Auth_Binding_delete = PptxEditor.Binding['delete'];
type _check_Binding_delete = IsExact<Ref_Binding_delete, Auth_Binding_delete>;
type _assert_Binding_delete = Expect<_check_Binding_delete>;

// PowerPoint.Binding#getShape
type Ref_Binding_getShape = { (): PptxEditor.Shape };
type Auth_Binding_getShape = PptxEditor.Binding['getShape'];
type _check_Binding_getShape = IsExact<Ref_Binding_getShape, Auth_Binding_getShape>;
type _assert_Binding_getShape = Expect<_check_Binding_getShape>;

// PowerPoint.Binding#id
type Ref_Binding_id = { readonly id: string };
type Auth_Binding_id = Pick<PptxEditor.Binding, 'id'>;
type _check_Binding_id = IsExact<Ref_Binding_id, Auth_Binding_id>;
type _assert_Binding_id = Expect<_check_Binding_id>;

// PowerPoint.Binding#type
type Ref_Binding_type = { readonly type: "Shape" | PptxEditor.BindingType };
type Auth_Binding_type = Pick<PptxEditor.Binding, 'type'>;
type _check_Binding_type = IsExact<Ref_Binding_type, Auth_Binding_type>;
type _assert_Binding_type = Expect<_check_Binding_type>;

// PowerPoint.BindingCollection

// PowerPoint.BindingCollection#add
type Ref_BindingCollection_add = { (shape: PptxEditor.Shape, bindingType: PptxEditor.BindingType, id: string): PptxEditor.Binding; (shape: PptxEditor.Shape, bindingType: "Shape", id: string): PptxEditor.Binding };
type Auth_BindingCollection_add = PptxEditor.BindingCollection['add'];
type _check_BindingCollection_add = IsExact<Ref_BindingCollection_add, Auth_BindingCollection_add>;
type _assert_BindingCollection_add = Expect<_check_BindingCollection_add>;

// PowerPoint.BindingCollection#addFromSelection
type Ref_BindingCollection_addFromSelection = { (bindingType: PptxEditor.BindingType, id: string): PptxEditor.Binding; (bindingType: "Shape", id: string): PptxEditor.Binding };
type Auth_BindingCollection_addFromSelection = PptxEditor.BindingCollection['addFromSelection'];
type _check_BindingCollection_addFromSelection = IsExact<Ref_BindingCollection_addFromSelection, Auth_BindingCollection_addFromSelection>;
type _assert_BindingCollection_addFromSelection = Expect<_check_BindingCollection_addFromSelection>;

// PowerPoint.BindingCollection#getCount
type Ref_BindingCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_BindingCollection_getCount = PptxEditor.BindingCollection['getCount'];
type _check_BindingCollection_getCount = IsExact<Ref_BindingCollection_getCount, Auth_BindingCollection_getCount>;
type _assert_BindingCollection_getCount = Expect<_check_BindingCollection_getCount>;

// PowerPoint.BindingCollection#getItem
type Ref_BindingCollection_getItem = { (key: string): PptxEditor.Binding };
type Auth_BindingCollection_getItem = PptxEditor.BindingCollection['getItem'];
type _check_BindingCollection_getItem = IsExact<Ref_BindingCollection_getItem, Auth_BindingCollection_getItem>;
type _assert_BindingCollection_getItem = Expect<_check_BindingCollection_getItem>;

// PowerPoint.BindingCollection#getItemAt
type Ref_BindingCollection_getItemAt = { (index: number): PptxEditor.Binding };
type Auth_BindingCollection_getItemAt = PptxEditor.BindingCollection['getItemAt'];
type _check_BindingCollection_getItemAt = IsExact<Ref_BindingCollection_getItemAt, Auth_BindingCollection_getItemAt>;
type _assert_BindingCollection_getItemAt = Expect<_check_BindingCollection_getItemAt>;

// PowerPoint.BindingCollection#getItemOrNullObject
type Ref_BindingCollection_getItemOrNullObject = { (id: string): PptxEditor.Binding };
type Auth_BindingCollection_getItemOrNullObject = PptxEditor.BindingCollection['getItemOrNullObject'];
type _check_BindingCollection_getItemOrNullObject = IsExact<Ref_BindingCollection_getItemOrNullObject, Auth_BindingCollection_getItemOrNullObject>;
type _assert_BindingCollection_getItemOrNullObject = Expect<_check_BindingCollection_getItemOrNullObject>;

// PowerPoint.BindingCollection#items
type Ref_BindingCollection_items = { readonly items: PptxEditor.Binding[] };
type Auth_BindingCollection_items = Pick<PptxEditor.BindingCollection, 'items'>;
type _check_BindingCollection_items = IsExact<Ref_BindingCollection_items, Auth_BindingCollection_items>;
type _assert_BindingCollection_items = Expect<_check_BindingCollection_items>;

// PowerPoint.TagCollection

// PowerPoint.TagCollection#add
type Ref_TagCollection_add = { (key: string, value: string): void };
type Auth_TagCollection_add = PptxEditor.TagCollection['add'];
type _check_TagCollection_add = IsExact<Ref_TagCollection_add, Auth_TagCollection_add>;
type _assert_TagCollection_add = Expect<_check_TagCollection_add>;

// PowerPoint.TagCollection#delete
type Ref_TagCollection_delete = { (key: string): void };
type Auth_TagCollection_delete = PptxEditor.TagCollection['delete'];
type _check_TagCollection_delete = IsExact<Ref_TagCollection_delete, Auth_TagCollection_delete>;
type _assert_TagCollection_delete = Expect<_check_TagCollection_delete>;

// PowerPoint.TagCollection#getCount
type Ref_TagCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_TagCollection_getCount = PptxEditor.TagCollection['getCount'];
type _check_TagCollection_getCount = IsExact<Ref_TagCollection_getCount, Auth_TagCollection_getCount>;
type _assert_TagCollection_getCount = Expect<_check_TagCollection_getCount>;

// PowerPoint.TagCollection#getItem
type Ref_TagCollection_getItem = { (key: string): PptxEditor.Tag };
type Auth_TagCollection_getItem = PptxEditor.TagCollection['getItem'];
type _check_TagCollection_getItem = IsExact<Ref_TagCollection_getItem, Auth_TagCollection_getItem>;
type _assert_TagCollection_getItem = Expect<_check_TagCollection_getItem>;

// PowerPoint.TagCollection#getItemAt
type Ref_TagCollection_getItemAt = { (index: number): PptxEditor.Tag };
type Auth_TagCollection_getItemAt = PptxEditor.TagCollection['getItemAt'];
type _check_TagCollection_getItemAt = IsExact<Ref_TagCollection_getItemAt, Auth_TagCollection_getItemAt>;
type _assert_TagCollection_getItemAt = Expect<_check_TagCollection_getItemAt>;

// PowerPoint.TagCollection#getItemOrNullObject
type Ref_TagCollection_getItemOrNullObject = { (key: string): PptxEditor.Tag };
type Auth_TagCollection_getItemOrNullObject = PptxEditor.TagCollection['getItemOrNullObject'];
type _check_TagCollection_getItemOrNullObject = IsExact<Ref_TagCollection_getItemOrNullObject, Auth_TagCollection_getItemOrNullObject>;
type _assert_TagCollection_getItemOrNullObject = Expect<_check_TagCollection_getItemOrNullObject>;

// PowerPoint.TagCollection#items
type Ref_TagCollection_items = { readonly items: PptxEditor.Tag[] };
type Auth_TagCollection_items = Pick<PptxEditor.TagCollection, 'items'>;
type _check_TagCollection_items = IsExact<Ref_TagCollection_items, Auth_TagCollection_items>;
type _assert_TagCollection_items = Expect<_check_TagCollection_items>;

// PowerPoint.DocumentProperties

// PowerPoint.DocumentProperties#author
type Ref_DocumentProperties_author = { author: string };
type Auth_DocumentProperties_author = Pick<PptxEditor.DocumentProperties, 'author'>;
type _check_DocumentProperties_author = IsExact<Ref_DocumentProperties_author, Auth_DocumentProperties_author>;
type _assert_DocumentProperties_author = Expect<_check_DocumentProperties_author>;
export function _write_DocumentProperties_author(target: PptxEditor.DocumentProperties, value: string): void {
  target['author'] = value;
}

// PowerPoint.DocumentProperties#category
type Ref_DocumentProperties_category = { category: string };
type Auth_DocumentProperties_category = Pick<PptxEditor.DocumentProperties, 'category'>;
type _check_DocumentProperties_category = IsExact<Ref_DocumentProperties_category, Auth_DocumentProperties_category>;
type _assert_DocumentProperties_category = Expect<_check_DocumentProperties_category>;
export function _write_DocumentProperties_category(target: PptxEditor.DocumentProperties, value: string): void {
  target['category'] = value;
}

// PowerPoint.DocumentProperties#comments
type Ref_DocumentProperties_comments = { comments: string };
type Auth_DocumentProperties_comments = Pick<PptxEditor.DocumentProperties, 'comments'>;
type _check_DocumentProperties_comments = IsExact<Ref_DocumentProperties_comments, Auth_DocumentProperties_comments>;
type _assert_DocumentProperties_comments = Expect<_check_DocumentProperties_comments>;
export function _write_DocumentProperties_comments(target: PptxEditor.DocumentProperties, value: string): void {
  target['comments'] = value;
}

// PowerPoint.DocumentProperties#company
type Ref_DocumentProperties_company = { company: string };
type Auth_DocumentProperties_company = Pick<PptxEditor.DocumentProperties, 'company'>;
type _check_DocumentProperties_company = IsExact<Ref_DocumentProperties_company, Auth_DocumentProperties_company>;
type _assert_DocumentProperties_company = Expect<_check_DocumentProperties_company>;
export function _write_DocumentProperties_company(target: PptxEditor.DocumentProperties, value: string): void {
  target['company'] = value;
}

// PowerPoint.DocumentProperties#creationDate
type Ref_DocumentProperties_creationDate = { readonly creationDate: Date };
type Auth_DocumentProperties_creationDate = Pick<PptxEditor.DocumentProperties, 'creationDate'>;
type _check_DocumentProperties_creationDate = IsExact<Ref_DocumentProperties_creationDate, Auth_DocumentProperties_creationDate>;
type _assert_DocumentProperties_creationDate = Expect<_check_DocumentProperties_creationDate>;

// PowerPoint.DocumentProperties#keywords
type Ref_DocumentProperties_keywords = { keywords: string };
type Auth_DocumentProperties_keywords = Pick<PptxEditor.DocumentProperties, 'keywords'>;
type _check_DocumentProperties_keywords = IsExact<Ref_DocumentProperties_keywords, Auth_DocumentProperties_keywords>;
type _assert_DocumentProperties_keywords = Expect<_check_DocumentProperties_keywords>;
export function _write_DocumentProperties_keywords(target: PptxEditor.DocumentProperties, value: string): void {
  target['keywords'] = value;
}

// PowerPoint.DocumentProperties#lastAuthor
type Ref_DocumentProperties_lastAuthor = { readonly lastAuthor: string };
type Auth_DocumentProperties_lastAuthor = Pick<PptxEditor.DocumentProperties, 'lastAuthor'>;
type _check_DocumentProperties_lastAuthor = IsExact<Ref_DocumentProperties_lastAuthor, Auth_DocumentProperties_lastAuthor>;
type _assert_DocumentProperties_lastAuthor = Expect<_check_DocumentProperties_lastAuthor>;

// PowerPoint.DocumentProperties#manager
type Ref_DocumentProperties_manager = { manager: string };
type Auth_DocumentProperties_manager = Pick<PptxEditor.DocumentProperties, 'manager'>;
type _check_DocumentProperties_manager = IsExact<Ref_DocumentProperties_manager, Auth_DocumentProperties_manager>;
type _assert_DocumentProperties_manager = Expect<_check_DocumentProperties_manager>;
export function _write_DocumentProperties_manager(target: PptxEditor.DocumentProperties, value: string): void {
  target['manager'] = value;
}

// PowerPoint.DocumentProperties#revisionNumber
type Ref_DocumentProperties_revisionNumber = { revisionNumber: number };
type Auth_DocumentProperties_revisionNumber = Pick<PptxEditor.DocumentProperties, 'revisionNumber'>;
type _check_DocumentProperties_revisionNumber = IsExact<Ref_DocumentProperties_revisionNumber, Auth_DocumentProperties_revisionNumber>;
type _assert_DocumentProperties_revisionNumber = Expect<_check_DocumentProperties_revisionNumber>;
export function _write_DocumentProperties_revisionNumber(target: PptxEditor.DocumentProperties, value: number): void {
  target['revisionNumber'] = value;
}

// PowerPoint.DocumentProperties#subject
type Ref_DocumentProperties_subject = { subject: string };
type Auth_DocumentProperties_subject = Pick<PptxEditor.DocumentProperties, 'subject'>;
type _check_DocumentProperties_subject = IsExact<Ref_DocumentProperties_subject, Auth_DocumentProperties_subject>;
type _assert_DocumentProperties_subject = Expect<_check_DocumentProperties_subject>;
export function _write_DocumentProperties_subject(target: PptxEditor.DocumentProperties, value: string): void {
  target['subject'] = value;
}

// PowerPoint.DocumentProperties#title
type Ref_DocumentProperties_title = { title: string };
type Auth_DocumentProperties_title = Pick<PptxEditor.DocumentProperties, 'title'>;
type _check_DocumentProperties_title = IsExact<Ref_DocumentProperties_title, Auth_DocumentProperties_title>;
type _assert_DocumentProperties_title = Expect<_check_DocumentProperties_title>;
export function _write_DocumentProperties_title(target: PptxEditor.DocumentProperties, value: string): void {
  target['title'] = value;
}

// PowerPoint.DocumentProperties#customProperties
type Ref_DocumentProperties_customProperties = { readonly customProperties: PptxEditor.CustomPropertyCollection };
type Auth_DocumentProperties_customProperties = Pick<PptxEditor.DocumentProperties, 'customProperties'>;
type _check_DocumentProperties_customProperties = IsExact<Ref_DocumentProperties_customProperties, Auth_DocumentProperties_customProperties>;
type _assert_DocumentProperties_customProperties = Expect<_check_DocumentProperties_customProperties>;

// PowerPoint.PageSetup

// PowerPoint.PageSetup#slideHeight
type Ref_PageSetup_slideHeight = { slideHeight: number };
type Auth_PageSetup_slideHeight = Pick<PptxEditor.PageSetup, 'slideHeight'>;
type _check_PageSetup_slideHeight = IsExact<Ref_PageSetup_slideHeight, Auth_PageSetup_slideHeight>;
type _assert_PageSetup_slideHeight = Expect<_check_PageSetup_slideHeight>;
export function _write_PageSetup_slideHeight(target: PptxEditor.PageSetup, value: number): void {
  target['slideHeight'] = value;
}

// PowerPoint.PageSetup#slideWidth
type Ref_PageSetup_slideWidth = { slideWidth: number };
type Auth_PageSetup_slideWidth = Pick<PptxEditor.PageSetup, 'slideWidth'>;
type _check_PageSetup_slideWidth = IsExact<Ref_PageSetup_slideWidth, Auth_PageSetup_slideWidth>;
type _assert_PageSetup_slideWidth = Expect<_check_PageSetup_slideWidth>;
export function _write_PageSetup_slideWidth(target: PptxEditor.PageSetup, value: number): void {
  target['slideWidth'] = value;
}

// PowerPoint.Adjustments

// PowerPoint.Adjustments#count
type Ref_Adjustments_count = { readonly count: number };
type Auth_Adjustments_count = Pick<PptxEditor.Adjustments, 'count'>;
type _check_Adjustments_count = IsExact<Ref_Adjustments_count, Auth_Adjustments_count>;
type _assert_Adjustments_count = Expect<_check_Adjustments_count>;

// PowerPoint.Adjustments#get
type Ref_Adjustments_get = { (index: number): PptxEditor.ClientResult<number> };
type Auth_Adjustments_get = PptxEditor.Adjustments['get'];
type _check_Adjustments_get = IsExact<Ref_Adjustments_get, Auth_Adjustments_get>;
type _assert_Adjustments_get = Expect<_check_Adjustments_get>;

// PowerPoint.Adjustments#set
type Ref_Adjustments_set = { (index: number, value: number): void };
type Auth_Adjustments_set = PptxEditor.Adjustments['set'];
type _check_Adjustments_set = IsExact<Ref_Adjustments_set, Auth_Adjustments_set>;
type _assert_Adjustments_set = Expect<_check_Adjustments_set>;

// PowerPoint.ShapeAddOptions

// PowerPoint.ShapeAddOptions#height
type Ref_ShapeAddOptions_height = { height?: number | undefined };
type Auth_ShapeAddOptions_height = Pick<PptxEditor.ShapeAddOptions, 'height'>;
type _check_ShapeAddOptions_height = IsExact<Ref_ShapeAddOptions_height, Auth_ShapeAddOptions_height>;
type _assert_ShapeAddOptions_height = Expect<_check_ShapeAddOptions_height>;
export function _write_ShapeAddOptions_height(target: PptxEditor.ShapeAddOptions, value: number | undefined): void {
  target['height'] = value;
}

// PowerPoint.ShapeAddOptions#left
type Ref_ShapeAddOptions_left = { left?: number | undefined };
type Auth_ShapeAddOptions_left = Pick<PptxEditor.ShapeAddOptions, 'left'>;
type _check_ShapeAddOptions_left = IsExact<Ref_ShapeAddOptions_left, Auth_ShapeAddOptions_left>;
type _assert_ShapeAddOptions_left = Expect<_check_ShapeAddOptions_left>;
export function _write_ShapeAddOptions_left(target: PptxEditor.ShapeAddOptions, value: number | undefined): void {
  target['left'] = value;
}

// PowerPoint.ShapeAddOptions#top
type Ref_ShapeAddOptions_top = { top?: number | undefined };
type Auth_ShapeAddOptions_top = Pick<PptxEditor.ShapeAddOptions, 'top'>;
type _check_ShapeAddOptions_top = IsExact<Ref_ShapeAddOptions_top, Auth_ShapeAddOptions_top>;
type _assert_ShapeAddOptions_top = Expect<_check_ShapeAddOptions_top>;
export function _write_ShapeAddOptions_top(target: PptxEditor.ShapeAddOptions, value: number | undefined): void {
  target['top'] = value;
}

// PowerPoint.ShapeAddOptions#width
type Ref_ShapeAddOptions_width = { width?: number | undefined };
type Auth_ShapeAddOptions_width = Pick<PptxEditor.ShapeAddOptions, 'width'>;
type _check_ShapeAddOptions_width = IsExact<Ref_ShapeAddOptions_width, Auth_ShapeAddOptions_width>;
type _assert_ShapeAddOptions_width = Expect<_check_ShapeAddOptions_width>;
export function _write_ShapeAddOptions_width(target: PptxEditor.ShapeAddOptions, value: number | undefined): void {
  target['width'] = value;
}

// PowerPoint.TableCellCollection

// PowerPoint.TableCellCollection#getCount
type Ref_TableCellCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_TableCellCollection_getCount = PptxEditor.TableCellCollection['getCount'];
type _check_TableCellCollection_getCount = IsExact<Ref_TableCellCollection_getCount, Auth_TableCellCollection_getCount>;
type _assert_TableCellCollection_getCount = Expect<_check_TableCellCollection_getCount>;

// PowerPoint.TableCellCollection#getItemAtOrNullObject
type Ref_TableCellCollection_getItemAtOrNullObject = { (row: number, column: number): PptxEditor.TableCell };
type Auth_TableCellCollection_getItemAtOrNullObject = PptxEditor.TableCellCollection['getItemAtOrNullObject'];
type _check_TableCellCollection_getItemAtOrNullObject = IsExact<Ref_TableCellCollection_getItemAtOrNullObject, Auth_TableCellCollection_getItemAtOrNullObject>;
type _assert_TableCellCollection_getItemAtOrNullObject = Expect<_check_TableCellCollection_getItemAtOrNullObject>;

// PowerPoint.TableCellCollection#items
type Ref_TableCellCollection_items = { readonly items: PptxEditor.TableCell[] };
type Auth_TableCellCollection_items = Pick<PptxEditor.TableCellCollection, 'items'>;
type _check_TableCellCollection_items = IsExact<Ref_TableCellCollection_items, Auth_TableCellCollection_items>;
type _assert_TableCellCollection_items = Expect<_check_TableCellCollection_items>;

// PowerPoint.TableStyleSettings

// PowerPoint.TableStyleSettings#areColumnsBanded
type Ref_TableStyleSettings_areColumnsBanded = { areColumnsBanded: boolean };
type Auth_TableStyleSettings_areColumnsBanded = Pick<PptxEditor.TableStyleSettings, 'areColumnsBanded'>;
type _check_TableStyleSettings_areColumnsBanded = IsExact<Ref_TableStyleSettings_areColumnsBanded, Auth_TableStyleSettings_areColumnsBanded>;
type _assert_TableStyleSettings_areColumnsBanded = Expect<_check_TableStyleSettings_areColumnsBanded>;
export function _write_TableStyleSettings_areColumnsBanded(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['areColumnsBanded'] = value;
}

// PowerPoint.TableStyleSettings#areRowsBanded
type Ref_TableStyleSettings_areRowsBanded = { areRowsBanded: boolean };
type Auth_TableStyleSettings_areRowsBanded = Pick<PptxEditor.TableStyleSettings, 'areRowsBanded'>;
type _check_TableStyleSettings_areRowsBanded = IsExact<Ref_TableStyleSettings_areRowsBanded, Auth_TableStyleSettings_areRowsBanded>;
type _assert_TableStyleSettings_areRowsBanded = Expect<_check_TableStyleSettings_areRowsBanded>;
export function _write_TableStyleSettings_areRowsBanded(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['areRowsBanded'] = value;
}

// PowerPoint.TableStyleSettings#isFirstColumnHighlighted
type Ref_TableStyleSettings_isFirstColumnHighlighted = { isFirstColumnHighlighted: boolean };
type Auth_TableStyleSettings_isFirstColumnHighlighted = Pick<PptxEditor.TableStyleSettings, 'isFirstColumnHighlighted'>;
type _check_TableStyleSettings_isFirstColumnHighlighted = IsExact<Ref_TableStyleSettings_isFirstColumnHighlighted, Auth_TableStyleSettings_isFirstColumnHighlighted>;
type _assert_TableStyleSettings_isFirstColumnHighlighted = Expect<_check_TableStyleSettings_isFirstColumnHighlighted>;
export function _write_TableStyleSettings_isFirstColumnHighlighted(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['isFirstColumnHighlighted'] = value;
}

// PowerPoint.TableStyleSettings#isFirstRowHighlighted
type Ref_TableStyleSettings_isFirstRowHighlighted = { isFirstRowHighlighted: boolean };
type Auth_TableStyleSettings_isFirstRowHighlighted = Pick<PptxEditor.TableStyleSettings, 'isFirstRowHighlighted'>;
type _check_TableStyleSettings_isFirstRowHighlighted = IsExact<Ref_TableStyleSettings_isFirstRowHighlighted, Auth_TableStyleSettings_isFirstRowHighlighted>;
type _assert_TableStyleSettings_isFirstRowHighlighted = Expect<_check_TableStyleSettings_isFirstRowHighlighted>;
export function _write_TableStyleSettings_isFirstRowHighlighted(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['isFirstRowHighlighted'] = value;
}

// PowerPoint.TableStyleSettings#isLastColumnHighlighted
type Ref_TableStyleSettings_isLastColumnHighlighted = { isLastColumnHighlighted: boolean };
type Auth_TableStyleSettings_isLastColumnHighlighted = Pick<PptxEditor.TableStyleSettings, 'isLastColumnHighlighted'>;
type _check_TableStyleSettings_isLastColumnHighlighted = IsExact<Ref_TableStyleSettings_isLastColumnHighlighted, Auth_TableStyleSettings_isLastColumnHighlighted>;
type _assert_TableStyleSettings_isLastColumnHighlighted = Expect<_check_TableStyleSettings_isLastColumnHighlighted>;
export function _write_TableStyleSettings_isLastColumnHighlighted(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['isLastColumnHighlighted'] = value;
}

// PowerPoint.TableStyleSettings#isLastRowHighlighted
type Ref_TableStyleSettings_isLastRowHighlighted = { isLastRowHighlighted: boolean };
type Auth_TableStyleSettings_isLastRowHighlighted = Pick<PptxEditor.TableStyleSettings, 'isLastRowHighlighted'>;
type _check_TableStyleSettings_isLastRowHighlighted = IsExact<Ref_TableStyleSettings_isLastRowHighlighted, Auth_TableStyleSettings_isLastRowHighlighted>;
type _assert_TableStyleSettings_isLastRowHighlighted = Expect<_check_TableStyleSettings_isLastRowHighlighted>;
export function _write_TableStyleSettings_isLastRowHighlighted(target: PptxEditor.TableStyleSettings, value: boolean): void {
  target['isLastRowHighlighted'] = value;
}

// PowerPoint.TableStyleSettings#style
type Ref_TableStyleSettings_style = { style: "DarkStyle1" | "DarkStyle1Accent1" | "DarkStyle1Accent2" | "DarkStyle1Accent3" | "DarkStyle1Accent4" | "DarkStyle1Accent5" | "DarkStyle1Accent6" | "DarkStyle2" | "DarkStyle2Accent1" | "DarkStyle2Accent2" | "DarkStyle2Accent3" | "LightStyle1" | "LightStyle1Accent1" | "LightStyle1Accent2" | "LightStyle1Accent3" | "LightStyle1Accent4" | "LightStyle1Accent5" | "LightStyle1Accent6" | "LightStyle2" | "LightStyle2Accent1" | "LightStyle2Accent2" | "LightStyle2Accent3" | "LightStyle2Accent4" | "LightStyle2Accent5" | "LightStyle2Accent6" | "LightStyle3" | "LightStyle3Accent1" | "LightStyle3Accent2" | "LightStyle3Accent3" | "LightStyle3Accent4" | "LightStyle3Accent5" | "LightStyle3Accent6" | "MediumStyle1" | "MediumStyle1Accent1" | "MediumStyle1Accent2" | "MediumStyle1Accent3" | "MediumStyle1Accent4" | "MediumStyle1Accent5" | "MediumStyle1Accent6" | "MediumStyle2" | "MediumStyle2Accent1" | "MediumStyle2Accent2" | "MediumStyle2Accent3" | "MediumStyle2Accent4" | "MediumStyle2Accent5" | "MediumStyle2Accent6" | "MediumStyle3" | "MediumStyle3Accent1" | "MediumStyle3Accent2" | "MediumStyle3Accent3" | "MediumStyle3Accent4" | "MediumStyle3Accent5" | "MediumStyle3Accent6" | "MediumStyle4" | "MediumStyle4Accent1" | "MediumStyle4Accent2" | "MediumStyle4Accent3" | "MediumStyle4Accent4" | "MediumStyle4Accent5" | "MediumStyle4Accent6" | "NoStyleNoGrid" | "NoStyleTableGrid" | "ThemedStyle1Accent1" | "ThemedStyle1Accent2" | "ThemedStyle1Accent3" | "ThemedStyle1Accent4" | "ThemedStyle1Accent5" | "ThemedStyle1Accent6" | "ThemedStyle2Accent1" | "ThemedStyle2Accent2" | "ThemedStyle2Accent3" | "ThemedStyle2Accent4" | "ThemedStyle2Accent5" | "ThemedStyle2Accent6" | PptxEditor.TableStyle };
type Auth_TableStyleSettings_style = Pick<PptxEditor.TableStyleSettings, 'style'>;
type _check_TableStyleSettings_style = IsExact<Ref_TableStyleSettings_style, Auth_TableStyleSettings_style>;
type _assert_TableStyleSettings_style = Expect<_check_TableStyleSettings_style>;
export function _write_TableStyleSettings_style(target: PptxEditor.TableStyleSettings, value: "DarkStyle1" | "DarkStyle1Accent1" | "DarkStyle1Accent2" | "DarkStyle1Accent3" | "DarkStyle1Accent4" | "DarkStyle1Accent5" | "DarkStyle1Accent6" | "DarkStyle2" | "DarkStyle2Accent1" | "DarkStyle2Accent2" | "DarkStyle2Accent3" | "LightStyle1" | "LightStyle1Accent1" | "LightStyle1Accent2" | "LightStyle1Accent3" | "LightStyle1Accent4" | "LightStyle1Accent5" | "LightStyle1Accent6" | "LightStyle2" | "LightStyle2Accent1" | "LightStyle2Accent2" | "LightStyle2Accent3" | "LightStyle2Accent4" | "LightStyle2Accent5" | "LightStyle2Accent6" | "LightStyle3" | "LightStyle3Accent1" | "LightStyle3Accent2" | "LightStyle3Accent3" | "LightStyle3Accent4" | "LightStyle3Accent5" | "LightStyle3Accent6" | "MediumStyle1" | "MediumStyle1Accent1" | "MediumStyle1Accent2" | "MediumStyle1Accent3" | "MediumStyle1Accent4" | "MediumStyle1Accent5" | "MediumStyle1Accent6" | "MediumStyle2" | "MediumStyle2Accent1" | "MediumStyle2Accent2" | "MediumStyle2Accent3" | "MediumStyle2Accent4" | "MediumStyle2Accent5" | "MediumStyle2Accent6" | "MediumStyle3" | "MediumStyle3Accent1" | "MediumStyle3Accent2" | "MediumStyle3Accent3" | "MediumStyle3Accent4" | "MediumStyle3Accent5" | "MediumStyle3Accent6" | "MediumStyle4" | "MediumStyle4Accent1" | "MediumStyle4Accent2" | "MediumStyle4Accent3" | "MediumStyle4Accent4" | "MediumStyle4Accent5" | "MediumStyle4Accent6" | "NoStyleNoGrid" | "NoStyleTableGrid" | "ThemedStyle1Accent1" | "ThemedStyle1Accent2" | "ThemedStyle1Accent3" | "ThemedStyle1Accent4" | "ThemedStyle1Accent5" | "ThemedStyle1Accent6" | "ThemedStyle2Accent1" | "ThemedStyle2Accent2" | "ThemedStyle2Accent3" | "ThemedStyle2Accent4" | "ThemedStyle2Accent5" | "ThemedStyle2Accent6" | PptxEditor.TableStyle): void {
  target['style'] = value;
}

// PowerPoint.Borders

// PowerPoint.Borders#bottom
type Ref_Borders_bottom = { readonly bottom: PptxEditor.Border };
type Auth_Borders_bottom = Pick<PptxEditor.Borders, 'bottom'>;
type _check_Borders_bottom = IsExact<Ref_Borders_bottom, Auth_Borders_bottom>;
type _assert_Borders_bottom = Expect<_check_Borders_bottom>;

// PowerPoint.Borders#diagonalDown
type Ref_Borders_diagonalDown = { readonly diagonalDown: PptxEditor.Border };
type Auth_Borders_diagonalDown = Pick<PptxEditor.Borders, 'diagonalDown'>;
type _check_Borders_diagonalDown = IsExact<Ref_Borders_diagonalDown, Auth_Borders_diagonalDown>;
type _assert_Borders_diagonalDown = Expect<_check_Borders_diagonalDown>;

// PowerPoint.Borders#diagonalUp
type Ref_Borders_diagonalUp = { readonly diagonalUp: PptxEditor.Border };
type Auth_Borders_diagonalUp = Pick<PptxEditor.Borders, 'diagonalUp'>;
type _check_Borders_diagonalUp = IsExact<Ref_Borders_diagonalUp, Auth_Borders_diagonalUp>;
type _assert_Borders_diagonalUp = Expect<_check_Borders_diagonalUp>;

// PowerPoint.Borders#left
type Ref_Borders_left = { readonly left: PptxEditor.Border };
type Auth_Borders_left = Pick<PptxEditor.Borders, 'left'>;
type _check_Borders_left = IsExact<Ref_Borders_left, Auth_Borders_left>;
type _assert_Borders_left = Expect<_check_Borders_left>;

// PowerPoint.Borders#right
type Ref_Borders_right = { readonly right: PptxEditor.Border };
type Auth_Borders_right = Pick<PptxEditor.Borders, 'right'>;
type _check_Borders_right = IsExact<Ref_Borders_right, Auth_Borders_right>;
type _assert_Borders_right = Expect<_check_Borders_right>;

// PowerPoint.Borders#top
type Ref_Borders_top = { readonly top: PptxEditor.Border };
type Auth_Borders_top = Pick<PptxEditor.Borders, 'top'>;
type _check_Borders_top = IsExact<Ref_Borders_top, Auth_Borders_top>;
type _assert_Borders_top = Expect<_check_Borders_top>;

// PowerPoint.Border

// PowerPoint.Border#color
type Ref_Border_color = { color: string | undefined };
type Auth_Border_color = Pick<PptxEditor.Border, 'color'>;
type _check_Border_color = IsExact<Ref_Border_color, Auth_Border_color>;
type _assert_Border_color = Expect<_check_Border_color>;
export function _write_Border_color(target: PptxEditor.Border, value: string | undefined): void {
  target['color'] = value;
}

// PowerPoint.Border#dashStyle
type Ref_Border_dashStyle = { dashStyle: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle | undefined };
type Auth_Border_dashStyle = Pick<PptxEditor.Border, 'dashStyle'>;
type _check_Border_dashStyle = IsExact<Ref_Border_dashStyle, Auth_Border_dashStyle>;
type _assert_Border_dashStyle = Expect<_check_Border_dashStyle>;
export function _write_Border_dashStyle(target: PptxEditor.Border, value: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle | undefined): void {
  target['dashStyle'] = value;
}

// PowerPoint.Border#transparency
type Ref_Border_transparency = { transparency: number | undefined };
type Auth_Border_transparency = Pick<PptxEditor.Border, 'transparency'>;
type _check_Border_transparency = IsExact<Ref_Border_transparency, Auth_Border_transparency>;
type _assert_Border_transparency = Expect<_check_Border_transparency>;
export function _write_Border_transparency(target: PptxEditor.Border, value: number | undefined): void {
  target['transparency'] = value;
}

// PowerPoint.Border#weight
type Ref_Border_weight = { weight: number | undefined };
type Auth_Border_weight = Pick<PptxEditor.Border, 'weight'>;
type _check_Border_weight = IsExact<Ref_Border_weight, Auth_Border_weight>;
type _assert_Border_weight = Expect<_check_Border_weight>;
export function _write_Border_weight(target: PptxEditor.Border, value: number | undefined): void {
  target['weight'] = value;
}

// PowerPoint.Margins

// PowerPoint.Margins#bottom
type Ref_Margins_bottom = { bottom: number | undefined };
type Auth_Margins_bottom = Pick<PptxEditor.Margins, 'bottom'>;
type _check_Margins_bottom = IsExact<Ref_Margins_bottom, Auth_Margins_bottom>;
type _assert_Margins_bottom = Expect<_check_Margins_bottom>;
export function _write_Margins_bottom(target: PptxEditor.Margins, value: number | undefined): void {
  target['bottom'] = value;
}

// PowerPoint.Margins#left
type Ref_Margins_left = { left: number | undefined };
type Auth_Margins_left = Pick<PptxEditor.Margins, 'left'>;
type _check_Margins_left = IsExact<Ref_Margins_left, Auth_Margins_left>;
type _assert_Margins_left = Expect<_check_Margins_left>;
export function _write_Margins_left(target: PptxEditor.Margins, value: number | undefined): void {
  target['left'] = value;
}

// PowerPoint.Margins#right
type Ref_Margins_right = { right: number | undefined };
type Auth_Margins_right = Pick<PptxEditor.Margins, 'right'>;
type _check_Margins_right = IsExact<Ref_Margins_right, Auth_Margins_right>;
type _assert_Margins_right = Expect<_check_Margins_right>;
export function _write_Margins_right(target: PptxEditor.Margins, value: number | undefined): void {
  target['right'] = value;
}

// PowerPoint.Margins#top
type Ref_Margins_top = { top: number | undefined };
type Auth_Margins_top = Pick<PptxEditor.Margins, 'top'>;
type _check_Margins_top = IsExact<Ref_Margins_top, Auth_Margins_top>;
type _assert_Margins_top = Expect<_check_Margins_top>;
export function _write_Margins_top(target: PptxEditor.Margins, value: number | undefined): void {
  target['top'] = value;
}

// PowerPoint.TableAddOptions

// PowerPoint.TableAddOptions#columns
type Ref_TableAddOptions_columns = { columns?: PptxEditor.TableColumnProperties[] | undefined };
type Auth_TableAddOptions_columns = Pick<PptxEditor.TableAddOptions, 'columns'>;
type _check_TableAddOptions_columns = IsExact<Ref_TableAddOptions_columns, Auth_TableAddOptions_columns>;
type _assert_TableAddOptions_columns = Expect<_check_TableAddOptions_columns>;
export function _write_TableAddOptions_columns(target: PptxEditor.TableAddOptions, value: PptxEditor.TableColumnProperties[] | undefined): void {
  target['columns'] = value;
}

// PowerPoint.TableAddOptions#height
type Ref_TableAddOptions_height = { height?: number | undefined };
type Auth_TableAddOptions_height = Pick<PptxEditor.TableAddOptions, 'height'>;
type _check_TableAddOptions_height = IsExact<Ref_TableAddOptions_height, Auth_TableAddOptions_height>;
type _assert_TableAddOptions_height = Expect<_check_TableAddOptions_height>;
export function _write_TableAddOptions_height(target: PptxEditor.TableAddOptions, value: number | undefined): void {
  target['height'] = value;
}

// PowerPoint.TableAddOptions#left
type Ref_TableAddOptions_left = { left?: number | undefined };
type Auth_TableAddOptions_left = Pick<PptxEditor.TableAddOptions, 'left'>;
type _check_TableAddOptions_left = IsExact<Ref_TableAddOptions_left, Auth_TableAddOptions_left>;
type _assert_TableAddOptions_left = Expect<_check_TableAddOptions_left>;
export function _write_TableAddOptions_left(target: PptxEditor.TableAddOptions, value: number | undefined): void {
  target['left'] = value;
}

// PowerPoint.TableAddOptions#mergedAreas
type Ref_TableAddOptions_mergedAreas = { mergedAreas?: PptxEditor.TableMergedAreaProperties[] | undefined };
type Auth_TableAddOptions_mergedAreas = Pick<PptxEditor.TableAddOptions, 'mergedAreas'>;
type _check_TableAddOptions_mergedAreas = IsExact<Ref_TableAddOptions_mergedAreas, Auth_TableAddOptions_mergedAreas>;
type _assert_TableAddOptions_mergedAreas = Expect<_check_TableAddOptions_mergedAreas>;
export function _write_TableAddOptions_mergedAreas(target: PptxEditor.TableAddOptions, value: PptxEditor.TableMergedAreaProperties[] | undefined): void {
  target['mergedAreas'] = value;
}

// PowerPoint.TableAddOptions#rows
type Ref_TableAddOptions_rows = { rows?: PptxEditor.TableRowProperties[] | undefined };
type Auth_TableAddOptions_rows = Pick<PptxEditor.TableAddOptions, 'rows'>;
type _check_TableAddOptions_rows = IsExact<Ref_TableAddOptions_rows, Auth_TableAddOptions_rows>;
type _assert_TableAddOptions_rows = Expect<_check_TableAddOptions_rows>;
export function _write_TableAddOptions_rows(target: PptxEditor.TableAddOptions, value: PptxEditor.TableRowProperties[] | undefined): void {
  target['rows'] = value;
}

// PowerPoint.TableAddOptions#specificCellProperties
type Ref_TableAddOptions_specificCellProperties = { specificCellProperties?: PptxEditor.TableCellProperties[][] | undefined };
type Auth_TableAddOptions_specificCellProperties = Pick<PptxEditor.TableAddOptions, 'specificCellProperties'>;
type _check_TableAddOptions_specificCellProperties = IsExact<Ref_TableAddOptions_specificCellProperties, Auth_TableAddOptions_specificCellProperties>;
type _assert_TableAddOptions_specificCellProperties = Expect<_check_TableAddOptions_specificCellProperties>;
export function _write_TableAddOptions_specificCellProperties(target: PptxEditor.TableAddOptions, value: PptxEditor.TableCellProperties[][] | undefined): void {
  target['specificCellProperties'] = value;
}

// PowerPoint.TableAddOptions#style
type Ref_TableAddOptions_style = { style?: "DarkStyle1" | "DarkStyle1Accent1" | "DarkStyle1Accent2" | "DarkStyle1Accent3" | "DarkStyle1Accent4" | "DarkStyle1Accent5" | "DarkStyle1Accent6" | "DarkStyle2" | "DarkStyle2Accent1" | "DarkStyle2Accent2" | "DarkStyle2Accent3" | "LightStyle1" | "LightStyle1Accent1" | "LightStyle1Accent2" | "LightStyle1Accent3" | "LightStyle1Accent4" | "LightStyle1Accent5" | "LightStyle1Accent6" | "LightStyle2" | "LightStyle2Accent1" | "LightStyle2Accent2" | "LightStyle2Accent3" | "LightStyle2Accent4" | "LightStyle2Accent5" | "LightStyle2Accent6" | "LightStyle3" | "LightStyle3Accent1" | "LightStyle3Accent2" | "LightStyle3Accent3" | "LightStyle3Accent4" | "LightStyle3Accent5" | "LightStyle3Accent6" | "MediumStyle1" | "MediumStyle1Accent1" | "MediumStyle1Accent2" | "MediumStyle1Accent3" | "MediumStyle1Accent4" | "MediumStyle1Accent5" | "MediumStyle1Accent6" | "MediumStyle2" | "MediumStyle2Accent1" | "MediumStyle2Accent2" | "MediumStyle2Accent3" | "MediumStyle2Accent4" | "MediumStyle2Accent5" | "MediumStyle2Accent6" | "MediumStyle3" | "MediumStyle3Accent1" | "MediumStyle3Accent2" | "MediumStyle3Accent3" | "MediumStyle3Accent4" | "MediumStyle3Accent5" | "MediumStyle3Accent6" | "MediumStyle4" | "MediumStyle4Accent1" | "MediumStyle4Accent2" | "MediumStyle4Accent3" | "MediumStyle4Accent4" | "MediumStyle4Accent5" | "MediumStyle4Accent6" | "NoStyleNoGrid" | "NoStyleTableGrid" | "ThemedStyle1Accent1" | "ThemedStyle1Accent2" | "ThemedStyle1Accent3" | "ThemedStyle1Accent4" | "ThemedStyle1Accent5" | "ThemedStyle1Accent6" | "ThemedStyle2Accent1" | "ThemedStyle2Accent2" | "ThemedStyle2Accent3" | "ThemedStyle2Accent4" | "ThemedStyle2Accent5" | "ThemedStyle2Accent6" | PptxEditor.TableStyle | undefined };
type Auth_TableAddOptions_style = Pick<PptxEditor.TableAddOptions, 'style'>;
type _check_TableAddOptions_style = IsExact<Ref_TableAddOptions_style, Auth_TableAddOptions_style>;
type _assert_TableAddOptions_style = Expect<_check_TableAddOptions_style>;
export function _write_TableAddOptions_style(target: PptxEditor.TableAddOptions, value: "DarkStyle1" | "DarkStyle1Accent1" | "DarkStyle1Accent2" | "DarkStyle1Accent3" | "DarkStyle1Accent4" | "DarkStyle1Accent5" | "DarkStyle1Accent6" | "DarkStyle2" | "DarkStyle2Accent1" | "DarkStyle2Accent2" | "DarkStyle2Accent3" | "LightStyle1" | "LightStyle1Accent1" | "LightStyle1Accent2" | "LightStyle1Accent3" | "LightStyle1Accent4" | "LightStyle1Accent5" | "LightStyle1Accent6" | "LightStyle2" | "LightStyle2Accent1" | "LightStyle2Accent2" | "LightStyle2Accent3" | "LightStyle2Accent4" | "LightStyle2Accent5" | "LightStyle2Accent6" | "LightStyle3" | "LightStyle3Accent1" | "LightStyle3Accent2" | "LightStyle3Accent3" | "LightStyle3Accent4" | "LightStyle3Accent5" | "LightStyle3Accent6" | "MediumStyle1" | "MediumStyle1Accent1" | "MediumStyle1Accent2" | "MediumStyle1Accent3" | "MediumStyle1Accent4" | "MediumStyle1Accent5" | "MediumStyle1Accent6" | "MediumStyle2" | "MediumStyle2Accent1" | "MediumStyle2Accent2" | "MediumStyle2Accent3" | "MediumStyle2Accent4" | "MediumStyle2Accent5" | "MediumStyle2Accent6" | "MediumStyle3" | "MediumStyle3Accent1" | "MediumStyle3Accent2" | "MediumStyle3Accent3" | "MediumStyle3Accent4" | "MediumStyle3Accent5" | "MediumStyle3Accent6" | "MediumStyle4" | "MediumStyle4Accent1" | "MediumStyle4Accent2" | "MediumStyle4Accent3" | "MediumStyle4Accent4" | "MediumStyle4Accent5" | "MediumStyle4Accent6" | "NoStyleNoGrid" | "NoStyleTableGrid" | "ThemedStyle1Accent1" | "ThemedStyle1Accent2" | "ThemedStyle1Accent3" | "ThemedStyle1Accent4" | "ThemedStyle1Accent5" | "ThemedStyle1Accent6" | "ThemedStyle2Accent1" | "ThemedStyle2Accent2" | "ThemedStyle2Accent3" | "ThemedStyle2Accent4" | "ThemedStyle2Accent5" | "ThemedStyle2Accent6" | PptxEditor.TableStyle | undefined): void {
  target['style'] = value;
}

// PowerPoint.TableAddOptions#top
type Ref_TableAddOptions_top = { top?: number | undefined };
type Auth_TableAddOptions_top = Pick<PptxEditor.TableAddOptions, 'top'>;
type _check_TableAddOptions_top = IsExact<Ref_TableAddOptions_top, Auth_TableAddOptions_top>;
type _assert_TableAddOptions_top = Expect<_check_TableAddOptions_top>;
export function _write_TableAddOptions_top(target: PptxEditor.TableAddOptions, value: number | undefined): void {
  target['top'] = value;
}

// PowerPoint.TableAddOptions#uniformCellProperties
type Ref_TableAddOptions_uniformCellProperties = { uniformCellProperties?: PptxEditor.TableCellProperties | undefined };
type Auth_TableAddOptions_uniformCellProperties = Pick<PptxEditor.TableAddOptions, 'uniformCellProperties'>;
type _check_TableAddOptions_uniformCellProperties = IsExact<Ref_TableAddOptions_uniformCellProperties, Auth_TableAddOptions_uniformCellProperties>;
type _assert_TableAddOptions_uniformCellProperties = Expect<_check_TableAddOptions_uniformCellProperties>;
export function _write_TableAddOptions_uniformCellProperties(target: PptxEditor.TableAddOptions, value: PptxEditor.TableCellProperties | undefined): void {
  target['uniformCellProperties'] = value;
}

// PowerPoint.TableAddOptions#values
type Ref_TableAddOptions_values = { values?: string[][] | undefined };
type Auth_TableAddOptions_values = Pick<PptxEditor.TableAddOptions, 'values'>;
type _check_TableAddOptions_values = IsExact<Ref_TableAddOptions_values, Auth_TableAddOptions_values>;
type _assert_TableAddOptions_values = Expect<_check_TableAddOptions_values>;
export function _write_TableAddOptions_values(target: PptxEditor.TableAddOptions, value: string[][] | undefined): void {
  target['values'] = value;
}

// PowerPoint.TableAddOptions#width
type Ref_TableAddOptions_width = { width?: number | undefined };
type Auth_TableAddOptions_width = Pick<PptxEditor.TableAddOptions, 'width'>;
type _check_TableAddOptions_width = IsExact<Ref_TableAddOptions_width, Auth_TableAddOptions_width>;
type _assert_TableAddOptions_width = Expect<_check_TableAddOptions_width>;
export function _write_TableAddOptions_width(target: PptxEditor.TableAddOptions, value: number | undefined): void {
  target['width'] = value;
}

// PowerPoint.TableClearOptions

// PowerPoint.TableClearOptions#all
type Ref_TableClearOptions_all = { all?: boolean | undefined };
type Auth_TableClearOptions_all = Pick<PptxEditor.TableClearOptions, 'all'>;
type _check_TableClearOptions_all = IsExact<Ref_TableClearOptions_all, Auth_TableClearOptions_all>;
type _assert_TableClearOptions_all = Expect<_check_TableClearOptions_all>;
export function _write_TableClearOptions_all(target: PptxEditor.TableClearOptions, value: boolean | undefined): void {
  target['all'] = value;
}

// PowerPoint.TableClearOptions#format
type Ref_TableClearOptions_format = { format?: boolean | undefined };
type Auth_TableClearOptions_format = Pick<PptxEditor.TableClearOptions, 'format'>;
type _check_TableClearOptions_format = IsExact<Ref_TableClearOptions_format, Auth_TableClearOptions_format>;
type _assert_TableClearOptions_format = Expect<_check_TableClearOptions_format>;
export function _write_TableClearOptions_format(target: PptxEditor.TableClearOptions, value: boolean | undefined): void {
  target['format'] = value;
}

// PowerPoint.TableClearOptions#text
type Ref_TableClearOptions_text = { text?: boolean | undefined };
type Auth_TableClearOptions_text = Pick<PptxEditor.TableClearOptions, 'text'>;
type _check_TableClearOptions_text = IsExact<Ref_TableClearOptions_text, Auth_TableClearOptions_text>;
type _assert_TableClearOptions_text = Expect<_check_TableClearOptions_text>;
export function _write_TableClearOptions_text(target: PptxEditor.TableClearOptions, value: boolean | undefined): void {
  target['text'] = value;
}

// PowerPoint.TableCellProperties

// PowerPoint.TableCellProperties#borders
type Ref_TableCellProperties_borders = { borders?: PptxEditor.TableCellBorders | undefined };
type Auth_TableCellProperties_borders = Pick<PptxEditor.TableCellProperties, 'borders'>;
type _check_TableCellProperties_borders = IsExact<Ref_TableCellProperties_borders, Auth_TableCellProperties_borders>;
type _assert_TableCellProperties_borders = Expect<_check_TableCellProperties_borders>;
export function _write_TableCellProperties_borders(target: PptxEditor.TableCellProperties, value: PptxEditor.TableCellBorders | undefined): void {
  target['borders'] = value;
}

// PowerPoint.TableCellProperties#fill
type Ref_TableCellProperties_fill = { fill?: PptxEditor.FillProperties | undefined };
type Auth_TableCellProperties_fill = Pick<PptxEditor.TableCellProperties, 'fill'>;
type _check_TableCellProperties_fill = IsExact<Ref_TableCellProperties_fill, Auth_TableCellProperties_fill>;
type _assert_TableCellProperties_fill = Expect<_check_TableCellProperties_fill>;
export function _write_TableCellProperties_fill(target: PptxEditor.TableCellProperties, value: PptxEditor.FillProperties | undefined): void {
  target['fill'] = value;
}

// PowerPoint.TableCellProperties#font
type Ref_TableCellProperties_font = { font?: PptxEditor.FontProperties | undefined };
type Auth_TableCellProperties_font = Pick<PptxEditor.TableCellProperties, 'font'>;
type _check_TableCellProperties_font = IsExact<Ref_TableCellProperties_font, Auth_TableCellProperties_font>;
type _assert_TableCellProperties_font = Expect<_check_TableCellProperties_font>;
export function _write_TableCellProperties_font(target: PptxEditor.TableCellProperties, value: PptxEditor.FontProperties | undefined): void {
  target['font'] = value;
}

// PowerPoint.TableCellProperties#horizontalAlignment
type Ref_TableCellProperties_horizontalAlignment = { horizontalAlignment?: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | PptxEditor.ParagraphHorizontalAlignment | undefined };
type Auth_TableCellProperties_horizontalAlignment = Pick<PptxEditor.TableCellProperties, 'horizontalAlignment'>;
type _check_TableCellProperties_horizontalAlignment = IsExact<Ref_TableCellProperties_horizontalAlignment, Auth_TableCellProperties_horizontalAlignment>;
type _assert_TableCellProperties_horizontalAlignment = Expect<_check_TableCellProperties_horizontalAlignment>;
export function _write_TableCellProperties_horizontalAlignment(target: PptxEditor.TableCellProperties, value: "Center" | "Distributed" | "Justify" | "JustifyLow" | "Left" | "Right" | "ThaiDistributed" | PptxEditor.ParagraphHorizontalAlignment | undefined): void {
  target['horizontalAlignment'] = value;
}

// PowerPoint.TableCellProperties#indentLevel
type Ref_TableCellProperties_indentLevel = { indentLevel?: number | undefined };
type Auth_TableCellProperties_indentLevel = Pick<PptxEditor.TableCellProperties, 'indentLevel'>;
type _check_TableCellProperties_indentLevel = IsExact<Ref_TableCellProperties_indentLevel, Auth_TableCellProperties_indentLevel>;
type _assert_TableCellProperties_indentLevel = Expect<_check_TableCellProperties_indentLevel>;
export function _write_TableCellProperties_indentLevel(target: PptxEditor.TableCellProperties, value: number | undefined): void {
  target['indentLevel'] = value;
}

// PowerPoint.TableCellProperties#margins
type Ref_TableCellProperties_margins = { margins?: PptxEditor.TableCellMargins | undefined };
type Auth_TableCellProperties_margins = Pick<PptxEditor.TableCellProperties, 'margins'>;
type _check_TableCellProperties_margins = IsExact<Ref_TableCellProperties_margins, Auth_TableCellProperties_margins>;
type _assert_TableCellProperties_margins = Expect<_check_TableCellProperties_margins>;
export function _write_TableCellProperties_margins(target: PptxEditor.TableCellProperties, value: PptxEditor.TableCellMargins | undefined): void {
  target['margins'] = value;
}

// PowerPoint.TableCellProperties#text
type Ref_TableCellProperties_text = { text?: string | undefined };
type Auth_TableCellProperties_text = Pick<PptxEditor.TableCellProperties, 'text'>;
type _check_TableCellProperties_text = IsExact<Ref_TableCellProperties_text, Auth_TableCellProperties_text>;
type _assert_TableCellProperties_text = Expect<_check_TableCellProperties_text>;
export function _write_TableCellProperties_text(target: PptxEditor.TableCellProperties, value: string | undefined): void {
  target['text'] = value;
}

// PowerPoint.TableCellProperties#textRuns
type Ref_TableCellProperties_textRuns = { textRuns?: PptxEditor.TextRun[] | undefined };
type Auth_TableCellProperties_textRuns = Pick<PptxEditor.TableCellProperties, 'textRuns'>;
type _check_TableCellProperties_textRuns = IsExact<Ref_TableCellProperties_textRuns, Auth_TableCellProperties_textRuns>;
type _assert_TableCellProperties_textRuns = Expect<_check_TableCellProperties_textRuns>;
export function _write_TableCellProperties_textRuns(target: PptxEditor.TableCellProperties, value: PptxEditor.TextRun[] | undefined): void {
  target['textRuns'] = value;
}

// PowerPoint.TableCellProperties#verticalAlignment
type Ref_TableCellProperties_verticalAlignment = { verticalAlignment?: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment | undefined };
type Auth_TableCellProperties_verticalAlignment = Pick<PptxEditor.TableCellProperties, 'verticalAlignment'>;
type _check_TableCellProperties_verticalAlignment = IsExact<Ref_TableCellProperties_verticalAlignment, Auth_TableCellProperties_verticalAlignment>;
type _assert_TableCellProperties_verticalAlignment = Expect<_check_TableCellProperties_verticalAlignment>;
export function _write_TableCellProperties_verticalAlignment(target: PptxEditor.TableCellProperties, value: "Bottom" | "BottomCentered" | "Middle" | "MiddleCentered" | "Top" | "TopCentered" | PptxEditor.TextVerticalAlignment | undefined): void {
  target['verticalAlignment'] = value;
}

// PowerPoint.TableCellBorders

// PowerPoint.TableCellBorders#bottom
type Ref_TableCellBorders_bottom = { bottom?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_bottom = Pick<PptxEditor.TableCellBorders, 'bottom'>;
type _check_TableCellBorders_bottom = IsExact<Ref_TableCellBorders_bottom, Auth_TableCellBorders_bottom>;
type _assert_TableCellBorders_bottom = Expect<_check_TableCellBorders_bottom>;
export function _write_TableCellBorders_bottom(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['bottom'] = value;
}

// PowerPoint.TableCellBorders#diagonalDown
type Ref_TableCellBorders_diagonalDown = { diagonalDown?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_diagonalDown = Pick<PptxEditor.TableCellBorders, 'diagonalDown'>;
type _check_TableCellBorders_diagonalDown = IsExact<Ref_TableCellBorders_diagonalDown, Auth_TableCellBorders_diagonalDown>;
type _assert_TableCellBorders_diagonalDown = Expect<_check_TableCellBorders_diagonalDown>;
export function _write_TableCellBorders_diagonalDown(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['diagonalDown'] = value;
}

// PowerPoint.TableCellBorders#diagonalUp
type Ref_TableCellBorders_diagonalUp = { diagonalUp?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_diagonalUp = Pick<PptxEditor.TableCellBorders, 'diagonalUp'>;
type _check_TableCellBorders_diagonalUp = IsExact<Ref_TableCellBorders_diagonalUp, Auth_TableCellBorders_diagonalUp>;
type _assert_TableCellBorders_diagonalUp = Expect<_check_TableCellBorders_diagonalUp>;
export function _write_TableCellBorders_diagonalUp(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['diagonalUp'] = value;
}

// PowerPoint.TableCellBorders#left
type Ref_TableCellBorders_left = { left?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_left = Pick<PptxEditor.TableCellBorders, 'left'>;
type _check_TableCellBorders_left = IsExact<Ref_TableCellBorders_left, Auth_TableCellBorders_left>;
type _assert_TableCellBorders_left = Expect<_check_TableCellBorders_left>;
export function _write_TableCellBorders_left(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['left'] = value;
}

// PowerPoint.TableCellBorders#right
type Ref_TableCellBorders_right = { right?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_right = Pick<PptxEditor.TableCellBorders, 'right'>;
type _check_TableCellBorders_right = IsExact<Ref_TableCellBorders_right, Auth_TableCellBorders_right>;
type _assert_TableCellBorders_right = Expect<_check_TableCellBorders_right>;
export function _write_TableCellBorders_right(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['right'] = value;
}

// PowerPoint.TableCellBorders#top
type Ref_TableCellBorders_top = { top?: PptxEditor.BorderProperties | undefined };
type Auth_TableCellBorders_top = Pick<PptxEditor.TableCellBorders, 'top'>;
type _check_TableCellBorders_top = IsExact<Ref_TableCellBorders_top, Auth_TableCellBorders_top>;
type _assert_TableCellBorders_top = Expect<_check_TableCellBorders_top>;
export function _write_TableCellBorders_top(target: PptxEditor.TableCellBorders, value: PptxEditor.BorderProperties | undefined): void {
  target['top'] = value;
}

// PowerPoint.BorderProperties

// PowerPoint.BorderProperties#color
type Ref_BorderProperties_color = { color?: string | undefined };
type Auth_BorderProperties_color = Pick<PptxEditor.BorderProperties, 'color'>;
type _check_BorderProperties_color = IsExact<Ref_BorderProperties_color, Auth_BorderProperties_color>;
type _assert_BorderProperties_color = Expect<_check_BorderProperties_color>;
export function _write_BorderProperties_color(target: PptxEditor.BorderProperties, value: string | undefined): void {
  target['color'] = value;
}

// PowerPoint.BorderProperties#dashStyle
type Ref_BorderProperties_dashStyle = { dashStyle?: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle | undefined };
type Auth_BorderProperties_dashStyle = Pick<PptxEditor.BorderProperties, 'dashStyle'>;
type _check_BorderProperties_dashStyle = IsExact<Ref_BorderProperties_dashStyle, Auth_BorderProperties_dashStyle>;
type _assert_BorderProperties_dashStyle = Expect<_check_BorderProperties_dashStyle>;
export function _write_BorderProperties_dashStyle(target: PptxEditor.BorderProperties, value: "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "RoundDot" | "Solid" | "SquareDot" | "SystemDash" | "SystemDashDot" | "SystemDot" | PptxEditor.ShapeLineDashStyle | undefined): void {
  target['dashStyle'] = value;
}

// PowerPoint.BorderProperties#transparency
type Ref_BorderProperties_transparency = { transparency?: number | undefined };
type Auth_BorderProperties_transparency = Pick<PptxEditor.BorderProperties, 'transparency'>;
type _check_BorderProperties_transparency = IsExact<Ref_BorderProperties_transparency, Auth_BorderProperties_transparency>;
type _assert_BorderProperties_transparency = Expect<_check_BorderProperties_transparency>;
export function _write_BorderProperties_transparency(target: PptxEditor.BorderProperties, value: number | undefined): void {
  target['transparency'] = value;
}

// PowerPoint.BorderProperties#weight
type Ref_BorderProperties_weight = { weight?: number | undefined };
type Auth_BorderProperties_weight = Pick<PptxEditor.BorderProperties, 'weight'>;
type _check_BorderProperties_weight = IsExact<Ref_BorderProperties_weight, Auth_BorderProperties_weight>;
type _assert_BorderProperties_weight = Expect<_check_BorderProperties_weight>;
export function _write_BorderProperties_weight(target: PptxEditor.BorderProperties, value: number | undefined): void {
  target['weight'] = value;
}

// PowerPoint.TableCellMargins

// PowerPoint.TableCellMargins#bottom
type Ref_TableCellMargins_bottom = { bottom?: number | undefined };
type Auth_TableCellMargins_bottom = Pick<PptxEditor.TableCellMargins, 'bottom'>;
type _check_TableCellMargins_bottom = IsExact<Ref_TableCellMargins_bottom, Auth_TableCellMargins_bottom>;
type _assert_TableCellMargins_bottom = Expect<_check_TableCellMargins_bottom>;
export function _write_TableCellMargins_bottom(target: PptxEditor.TableCellMargins, value: number | undefined): void {
  target['bottom'] = value;
}

// PowerPoint.TableCellMargins#left
type Ref_TableCellMargins_left = { left?: number | undefined };
type Auth_TableCellMargins_left = Pick<PptxEditor.TableCellMargins, 'left'>;
type _check_TableCellMargins_left = IsExact<Ref_TableCellMargins_left, Auth_TableCellMargins_left>;
type _assert_TableCellMargins_left = Expect<_check_TableCellMargins_left>;
export function _write_TableCellMargins_left(target: PptxEditor.TableCellMargins, value: number | undefined): void {
  target['left'] = value;
}

// PowerPoint.TableCellMargins#right
type Ref_TableCellMargins_right = { right?: number | undefined };
type Auth_TableCellMargins_right = Pick<PptxEditor.TableCellMargins, 'right'>;
type _check_TableCellMargins_right = IsExact<Ref_TableCellMargins_right, Auth_TableCellMargins_right>;
type _assert_TableCellMargins_right = Expect<_check_TableCellMargins_right>;
export function _write_TableCellMargins_right(target: PptxEditor.TableCellMargins, value: number | undefined): void {
  target['right'] = value;
}

// PowerPoint.TableCellMargins#top
type Ref_TableCellMargins_top = { top?: number | undefined };
type Auth_TableCellMargins_top = Pick<PptxEditor.TableCellMargins, 'top'>;
type _check_TableCellMargins_top = IsExact<Ref_TableCellMargins_top, Auth_TableCellMargins_top>;
type _assert_TableCellMargins_top = Expect<_check_TableCellMargins_top>;
export function _write_TableCellMargins_top(target: PptxEditor.TableCellMargins, value: number | undefined): void {
  target['top'] = value;
}

// PowerPoint.FillProperties

// PowerPoint.FillProperties#color
type Ref_FillProperties_color = { color?: string | undefined };
type Auth_FillProperties_color = Pick<PptxEditor.FillProperties, 'color'>;
type _check_FillProperties_color = IsExact<Ref_FillProperties_color, Auth_FillProperties_color>;
type _assert_FillProperties_color = Expect<_check_FillProperties_color>;
export function _write_FillProperties_color(target: PptxEditor.FillProperties, value: string | undefined): void {
  target['color'] = value;
}

// PowerPoint.FillProperties#transparency
type Ref_FillProperties_transparency = { transparency?: number | undefined };
type Auth_FillProperties_transparency = Pick<PptxEditor.FillProperties, 'transparency'>;
type _check_FillProperties_transparency = IsExact<Ref_FillProperties_transparency, Auth_FillProperties_transparency>;
type _assert_FillProperties_transparency = Expect<_check_FillProperties_transparency>;
export function _write_FillProperties_transparency(target: PptxEditor.FillProperties, value: number | undefined): void {
  target['transparency'] = value;
}

// PowerPoint.TableColumnProperties

// PowerPoint.TableColumnProperties#columnWidth
type Ref_TableColumnProperties_columnWidth = { columnWidth?: number | undefined };
type Auth_TableColumnProperties_columnWidth = Pick<PptxEditor.TableColumnProperties, 'columnWidth'>;
type _check_TableColumnProperties_columnWidth = IsExact<Ref_TableColumnProperties_columnWidth, Auth_TableColumnProperties_columnWidth>;
type _assert_TableColumnProperties_columnWidth = Expect<_check_TableColumnProperties_columnWidth>;
export function _write_TableColumnProperties_columnWidth(target: PptxEditor.TableColumnProperties, value: number | undefined): void {
  target['columnWidth'] = value;
}

// PowerPoint.TableRowProperties

// PowerPoint.TableRowProperties#rowHeight
type Ref_TableRowProperties_rowHeight = { rowHeight?: number | undefined };
type Auth_TableRowProperties_rowHeight = Pick<PptxEditor.TableRowProperties, 'rowHeight'>;
type _check_TableRowProperties_rowHeight = IsExact<Ref_TableRowProperties_rowHeight, Auth_TableRowProperties_rowHeight>;
type _assert_TableRowProperties_rowHeight = Expect<_check_TableRowProperties_rowHeight>;
export function _write_TableRowProperties_rowHeight(target: PptxEditor.TableRowProperties, value: number | undefined): void {
  target['rowHeight'] = value;
}

// PowerPoint.TableMergedAreaProperties

// PowerPoint.TableMergedAreaProperties#columnCount
type Ref_TableMergedAreaProperties_columnCount = { columnCount: number };
type Auth_TableMergedAreaProperties_columnCount = Pick<PptxEditor.TableMergedAreaProperties, 'columnCount'>;
type _check_TableMergedAreaProperties_columnCount = IsExact<Ref_TableMergedAreaProperties_columnCount, Auth_TableMergedAreaProperties_columnCount>;
type _assert_TableMergedAreaProperties_columnCount = Expect<_check_TableMergedAreaProperties_columnCount>;
export function _write_TableMergedAreaProperties_columnCount(target: PptxEditor.TableMergedAreaProperties, value: number): void {
  target['columnCount'] = value;
}

// PowerPoint.TableMergedAreaProperties#columnIndex
type Ref_TableMergedAreaProperties_columnIndex = { columnIndex: number };
type Auth_TableMergedAreaProperties_columnIndex = Pick<PptxEditor.TableMergedAreaProperties, 'columnIndex'>;
type _check_TableMergedAreaProperties_columnIndex = IsExact<Ref_TableMergedAreaProperties_columnIndex, Auth_TableMergedAreaProperties_columnIndex>;
type _assert_TableMergedAreaProperties_columnIndex = Expect<_check_TableMergedAreaProperties_columnIndex>;
export function _write_TableMergedAreaProperties_columnIndex(target: PptxEditor.TableMergedAreaProperties, value: number): void {
  target['columnIndex'] = value;
}

// PowerPoint.TableMergedAreaProperties#rowCount
type Ref_TableMergedAreaProperties_rowCount = { rowCount: number };
type Auth_TableMergedAreaProperties_rowCount = Pick<PptxEditor.TableMergedAreaProperties, 'rowCount'>;
type _check_TableMergedAreaProperties_rowCount = IsExact<Ref_TableMergedAreaProperties_rowCount, Auth_TableMergedAreaProperties_rowCount>;
type _assert_TableMergedAreaProperties_rowCount = Expect<_check_TableMergedAreaProperties_rowCount>;
export function _write_TableMergedAreaProperties_rowCount(target: PptxEditor.TableMergedAreaProperties, value: number): void {
  target['rowCount'] = value;
}

// PowerPoint.TableMergedAreaProperties#rowIndex
type Ref_TableMergedAreaProperties_rowIndex = { rowIndex: number };
type Auth_TableMergedAreaProperties_rowIndex = Pick<PptxEditor.TableMergedAreaProperties, 'rowIndex'>;
type _check_TableMergedAreaProperties_rowIndex = IsExact<Ref_TableMergedAreaProperties_rowIndex, Auth_TableMergedAreaProperties_rowIndex>;
type _assert_TableMergedAreaProperties_rowIndex = Expect<_check_TableMergedAreaProperties_rowIndex>;
export function _write_TableMergedAreaProperties_rowIndex(target: PptxEditor.TableMergedAreaProperties, value: number): void {
  target['rowIndex'] = value;
}

// PowerPoint.TextRun

// PowerPoint.TextRun#font
type Ref_TextRun_font = { font?: PptxEditor.FontProperties | undefined };
type Auth_TextRun_font = Pick<PptxEditor.TextRun, 'font'>;
type _check_TextRun_font = IsExact<Ref_TextRun_font, Auth_TextRun_font>;
type _assert_TextRun_font = Expect<_check_TextRun_font>;
export function _write_TextRun_font(target: PptxEditor.TextRun, value: PptxEditor.FontProperties | undefined): void {
  target['font'] = value;
}

// PowerPoint.TextRun#text
type Ref_TextRun_text = { text: string };
type Auth_TextRun_text = Pick<PptxEditor.TextRun, 'text'>;
type _check_TextRun_text = IsExact<Ref_TextRun_text, Auth_TextRun_text>;
type _assert_TextRun_text = Expect<_check_TextRun_text>;
export function _write_TextRun_text(target: PptxEditor.TextRun, value: string): void {
  target['text'] = value;
}

// PowerPoint.FontProperties

// PowerPoint.FontProperties#allCaps
type Ref_FontProperties_allCaps = { allCaps?: boolean | undefined };
type Auth_FontProperties_allCaps = Pick<PptxEditor.FontProperties, 'allCaps'>;
type _check_FontProperties_allCaps = IsExact<Ref_FontProperties_allCaps, Auth_FontProperties_allCaps>;
type _assert_FontProperties_allCaps = Expect<_check_FontProperties_allCaps>;
export function _write_FontProperties_allCaps(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['allCaps'] = value;
}

// PowerPoint.FontProperties#bold
type Ref_FontProperties_bold = { bold?: boolean | undefined };
type Auth_FontProperties_bold = Pick<PptxEditor.FontProperties, 'bold'>;
type _check_FontProperties_bold = IsExact<Ref_FontProperties_bold, Auth_FontProperties_bold>;
type _assert_FontProperties_bold = Expect<_check_FontProperties_bold>;
export function _write_FontProperties_bold(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['bold'] = value;
}

// PowerPoint.FontProperties#color
type Ref_FontProperties_color = { color?: string | undefined };
type Auth_FontProperties_color = Pick<PptxEditor.FontProperties, 'color'>;
type _check_FontProperties_color = IsExact<Ref_FontProperties_color, Auth_FontProperties_color>;
type _assert_FontProperties_color = Expect<_check_FontProperties_color>;
export function _write_FontProperties_color(target: PptxEditor.FontProperties, value: string | undefined): void {
  target['color'] = value;
}

// PowerPoint.FontProperties#doubleStrikethrough
type Ref_FontProperties_doubleStrikethrough = { doubleStrikethrough?: boolean | undefined };
type Auth_FontProperties_doubleStrikethrough = Pick<PptxEditor.FontProperties, 'doubleStrikethrough'>;
type _check_FontProperties_doubleStrikethrough = IsExact<Ref_FontProperties_doubleStrikethrough, Auth_FontProperties_doubleStrikethrough>;
type _assert_FontProperties_doubleStrikethrough = Expect<_check_FontProperties_doubleStrikethrough>;
export function _write_FontProperties_doubleStrikethrough(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['doubleStrikethrough'] = value;
}

// PowerPoint.FontProperties#italic
type Ref_FontProperties_italic = { italic?: boolean | undefined };
type Auth_FontProperties_italic = Pick<PptxEditor.FontProperties, 'italic'>;
type _check_FontProperties_italic = IsExact<Ref_FontProperties_italic, Auth_FontProperties_italic>;
type _assert_FontProperties_italic = Expect<_check_FontProperties_italic>;
export function _write_FontProperties_italic(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['italic'] = value;
}

// PowerPoint.FontProperties#name
type Ref_FontProperties_name = { name?: string | undefined };
type Auth_FontProperties_name = Pick<PptxEditor.FontProperties, 'name'>;
type _check_FontProperties_name = IsExact<Ref_FontProperties_name, Auth_FontProperties_name>;
type _assert_FontProperties_name = Expect<_check_FontProperties_name>;
export function _write_FontProperties_name(target: PptxEditor.FontProperties, value: string | undefined): void {
  target['name'] = value;
}

// PowerPoint.FontProperties#size
type Ref_FontProperties_size = { size?: number | undefined };
type Auth_FontProperties_size = Pick<PptxEditor.FontProperties, 'size'>;
type _check_FontProperties_size = IsExact<Ref_FontProperties_size, Auth_FontProperties_size>;
type _assert_FontProperties_size = Expect<_check_FontProperties_size>;
export function _write_FontProperties_size(target: PptxEditor.FontProperties, value: number | undefined): void {
  target['size'] = value;
}

// PowerPoint.FontProperties#smallCaps
type Ref_FontProperties_smallCaps = { smallCaps?: boolean | undefined };
type Auth_FontProperties_smallCaps = Pick<PptxEditor.FontProperties, 'smallCaps'>;
type _check_FontProperties_smallCaps = IsExact<Ref_FontProperties_smallCaps, Auth_FontProperties_smallCaps>;
type _assert_FontProperties_smallCaps = Expect<_check_FontProperties_smallCaps>;
export function _write_FontProperties_smallCaps(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['smallCaps'] = value;
}

// PowerPoint.FontProperties#strikethrough
type Ref_FontProperties_strikethrough = { strikethrough?: boolean | undefined };
type Auth_FontProperties_strikethrough = Pick<PptxEditor.FontProperties, 'strikethrough'>;
type _check_FontProperties_strikethrough = IsExact<Ref_FontProperties_strikethrough, Auth_FontProperties_strikethrough>;
type _assert_FontProperties_strikethrough = Expect<_check_FontProperties_strikethrough>;
export function _write_FontProperties_strikethrough(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['strikethrough'] = value;
}

// PowerPoint.FontProperties#subscript
type Ref_FontProperties_subscript = { subscript?: boolean | undefined };
type Auth_FontProperties_subscript = Pick<PptxEditor.FontProperties, 'subscript'>;
type _check_FontProperties_subscript = IsExact<Ref_FontProperties_subscript, Auth_FontProperties_subscript>;
type _assert_FontProperties_subscript = Expect<_check_FontProperties_subscript>;
export function _write_FontProperties_subscript(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['subscript'] = value;
}

// PowerPoint.FontProperties#superscript
type Ref_FontProperties_superscript = { superscript?: boolean | undefined };
type Auth_FontProperties_superscript = Pick<PptxEditor.FontProperties, 'superscript'>;
type _check_FontProperties_superscript = IsExact<Ref_FontProperties_superscript, Auth_FontProperties_superscript>;
type _assert_FontProperties_superscript = Expect<_check_FontProperties_superscript>;
export function _write_FontProperties_superscript(target: PptxEditor.FontProperties, value: boolean | undefined): void {
  target['superscript'] = value;
}

// PowerPoint.FontProperties#underline
type Ref_FontProperties_underline = { underline?: "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Dotted" | "DottedHeavy" | "Double" | "Heavy" | "None" | "Single" | "Wavy" | "WavyDouble" | "WavyHeavy" | PptxEditor.ShapeFontUnderlineStyle | undefined };
type Auth_FontProperties_underline = Pick<PptxEditor.FontProperties, 'underline'>;
type _check_FontProperties_underline = IsExact<Ref_FontProperties_underline, Auth_FontProperties_underline>;
type _assert_FontProperties_underline = Expect<_check_FontProperties_underline>;
export function _write_FontProperties_underline(target: PptxEditor.FontProperties, value: "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Dotted" | "DottedHeavy" | "Double" | "Heavy" | "None" | "Single" | "Wavy" | "WavyDouble" | "WavyHeavy" | PptxEditor.ShapeFontUnderlineStyle | undefined): void {
  target['underline'] = value;
}

// PowerPoint.HyperlinkAddOptions

// PowerPoint.HyperlinkAddOptions#address
type Ref_HyperlinkAddOptions_address = { address?: string | undefined };
type Auth_HyperlinkAddOptions_address = Pick<PptxEditor.HyperlinkAddOptions, 'address'>;
type _check_HyperlinkAddOptions_address = IsExact<Ref_HyperlinkAddOptions_address, Auth_HyperlinkAddOptions_address>;
type _assert_HyperlinkAddOptions_address = Expect<_check_HyperlinkAddOptions_address>;
export function _write_HyperlinkAddOptions_address(target: PptxEditor.HyperlinkAddOptions, value: string | undefined): void {
  target['address'] = value;
}

// PowerPoint.HyperlinkAddOptions#screenTip
type Ref_HyperlinkAddOptions_screenTip = { screenTip?: string | undefined };
type Auth_HyperlinkAddOptions_screenTip = Pick<PptxEditor.HyperlinkAddOptions, 'screenTip'>;
type _check_HyperlinkAddOptions_screenTip = IsExact<Ref_HyperlinkAddOptions_screenTip, Auth_HyperlinkAddOptions_screenTip>;
type _assert_HyperlinkAddOptions_screenTip = Expect<_check_HyperlinkAddOptions_screenTip>;
export function _write_HyperlinkAddOptions_screenTip(target: PptxEditor.HyperlinkAddOptions, value: string | undefined): void {
  target['screenTip'] = value;
}

// PowerPoint.AddSlideOptions

// PowerPoint.AddSlideOptions#layoutId
type Ref_AddSlideOptions_layoutId = { layoutId?: string | undefined };
type Auth_AddSlideOptions_layoutId = Pick<PptxEditor.AddSlideOptions, 'layoutId'>;
type _check_AddSlideOptions_layoutId = IsExact<Ref_AddSlideOptions_layoutId, Auth_AddSlideOptions_layoutId>;
type _assert_AddSlideOptions_layoutId = Expect<_check_AddSlideOptions_layoutId>;
export function _write_AddSlideOptions_layoutId(target: PptxEditor.AddSlideOptions, value: string | undefined): void {
  target['layoutId'] = value;
}

// PowerPoint.AddSlideOptions#slideMasterId
type Ref_AddSlideOptions_slideMasterId = { slideMasterId?: string | undefined };
type Auth_AddSlideOptions_slideMasterId = Pick<PptxEditor.AddSlideOptions, 'slideMasterId'>;
type _check_AddSlideOptions_slideMasterId = IsExact<Ref_AddSlideOptions_slideMasterId, Auth_AddSlideOptions_slideMasterId>;
type _assert_AddSlideOptions_slideMasterId = Expect<_check_AddSlideOptions_slideMasterId>;
export function _write_AddSlideOptions_slideMasterId(target: PptxEditor.AddSlideOptions, value: string | undefined): void {
  target['slideMasterId'] = value;
}

// PowerPoint.InsertSlideOptions

// PowerPoint.InsertSlideOptions#formatting
type Ref_InsertSlideOptions_formatting = { formatting?: "KeepSourceFormatting" | "UseDestinationTheme" | PptxEditor.InsertSlideFormatting | undefined };
type Auth_InsertSlideOptions_formatting = Pick<PptxEditor.InsertSlideOptions, 'formatting'>;
type _check_InsertSlideOptions_formatting = IsExact<Ref_InsertSlideOptions_formatting, Auth_InsertSlideOptions_formatting>;
type _assert_InsertSlideOptions_formatting = Expect<_check_InsertSlideOptions_formatting>;
export function _write_InsertSlideOptions_formatting(target: PptxEditor.InsertSlideOptions, value: "KeepSourceFormatting" | "UseDestinationTheme" | PptxEditor.InsertSlideFormatting | undefined): void {
  target['formatting'] = value;
}

// PowerPoint.InsertSlideOptions#sourceSlideIds
type Ref_InsertSlideOptions_sourceSlideIds = { sourceSlideIds?: string[] | undefined };
type Auth_InsertSlideOptions_sourceSlideIds = Pick<PptxEditor.InsertSlideOptions, 'sourceSlideIds'>;
type _check_InsertSlideOptions_sourceSlideIds = IsExact<Ref_InsertSlideOptions_sourceSlideIds, Auth_InsertSlideOptions_sourceSlideIds>;
type _assert_InsertSlideOptions_sourceSlideIds = Expect<_check_InsertSlideOptions_sourceSlideIds>;
export function _write_InsertSlideOptions_sourceSlideIds(target: PptxEditor.InsertSlideOptions, value: string[] | undefined): void {
  target['sourceSlideIds'] = value;
}

// PowerPoint.InsertSlideOptions#targetSlideId
type Ref_InsertSlideOptions_targetSlideId = { targetSlideId?: string | undefined };
type Auth_InsertSlideOptions_targetSlideId = Pick<PptxEditor.InsertSlideOptions, 'targetSlideId'>;
type _check_InsertSlideOptions_targetSlideId = IsExact<Ref_InsertSlideOptions_targetSlideId, Auth_InsertSlideOptions_targetSlideId>;
type _assert_InsertSlideOptions_targetSlideId = Expect<_check_InsertSlideOptions_targetSlideId>;
export function _write_InsertSlideOptions_targetSlideId(target: PptxEditor.InsertSlideOptions, value: string | undefined): void {
  target['targetSlideId'] = value;
}

// PowerPoint.SlideBackgroundSolidFillOptions

// PowerPoint.SlideBackgroundSolidFillOptions#color
type Ref_SlideBackgroundSolidFillOptions_color = { color?: string | undefined };
type Auth_SlideBackgroundSolidFillOptions_color = Pick<PptxEditor.SlideBackgroundSolidFillOptions, 'color'>;
type _check_SlideBackgroundSolidFillOptions_color = IsExact<Ref_SlideBackgroundSolidFillOptions_color, Auth_SlideBackgroundSolidFillOptions_color>;
type _assert_SlideBackgroundSolidFillOptions_color = Expect<_check_SlideBackgroundSolidFillOptions_color>;
export function _write_SlideBackgroundSolidFillOptions_color(target: PptxEditor.SlideBackgroundSolidFillOptions, value: string | undefined): void {
  target['color'] = value;
}

// PowerPoint.SlideBackgroundSolidFillOptions#transparency
type Ref_SlideBackgroundSolidFillOptions_transparency = { transparency?: number | undefined };
type Auth_SlideBackgroundSolidFillOptions_transparency = Pick<PptxEditor.SlideBackgroundSolidFillOptions, 'transparency'>;
type _check_SlideBackgroundSolidFillOptions_transparency = IsExact<Ref_SlideBackgroundSolidFillOptions_transparency, Auth_SlideBackgroundSolidFillOptions_transparency>;
type _assert_SlideBackgroundSolidFillOptions_transparency = Expect<_check_SlideBackgroundSolidFillOptions_transparency>;
export function _write_SlideBackgroundSolidFillOptions_transparency(target: PptxEditor.SlideBackgroundSolidFillOptions, value: number | undefined): void {
  target['transparency'] = value;
}

// PowerPoint.SlideBackgroundGradientFillOptions

// PowerPoint.SlideBackgroundGradientFillOptions#type
type Ref_SlideBackgroundGradientFillOptions_type = { type?: "Linear" | "Path" | "Radial" | "Rectangular" | "ShadeFromTitle" | "Unsupported" | PptxEditor.SlideBackgroundGradientFillType | undefined };
type Auth_SlideBackgroundGradientFillOptions_type = Pick<PptxEditor.SlideBackgroundGradientFillOptions, 'type'>;
type _check_SlideBackgroundGradientFillOptions_type = IsExact<Ref_SlideBackgroundGradientFillOptions_type, Auth_SlideBackgroundGradientFillOptions_type>;
type _assert_SlideBackgroundGradientFillOptions_type = Expect<_check_SlideBackgroundGradientFillOptions_type>;
export function _write_SlideBackgroundGradientFillOptions_type(target: PptxEditor.SlideBackgroundGradientFillOptions, value: "Linear" | "Path" | "Radial" | "Rectangular" | "ShadeFromTitle" | "Unsupported" | PptxEditor.SlideBackgroundGradientFillType | undefined): void {
  target['type'] = value;
}

// PowerPoint.SlideBackgroundPatternFillOptions

// PowerPoint.SlideBackgroundPatternFillOptions#backgroundColor
type Ref_SlideBackgroundPatternFillOptions_backgroundColor = { backgroundColor?: string | undefined };
type Auth_SlideBackgroundPatternFillOptions_backgroundColor = Pick<PptxEditor.SlideBackgroundPatternFillOptions, 'backgroundColor'>;
type _check_SlideBackgroundPatternFillOptions_backgroundColor = IsExact<Ref_SlideBackgroundPatternFillOptions_backgroundColor, Auth_SlideBackgroundPatternFillOptions_backgroundColor>;
type _assert_SlideBackgroundPatternFillOptions_backgroundColor = Expect<_check_SlideBackgroundPatternFillOptions_backgroundColor>;
export function _write_SlideBackgroundPatternFillOptions_backgroundColor(target: PptxEditor.SlideBackgroundPatternFillOptions, value: string | undefined): void {
  target['backgroundColor'] = value;
}

// PowerPoint.SlideBackgroundPatternFillOptions#foregroundColor
type Ref_SlideBackgroundPatternFillOptions_foregroundColor = { foregroundColor?: string | undefined };
type Auth_SlideBackgroundPatternFillOptions_foregroundColor = Pick<PptxEditor.SlideBackgroundPatternFillOptions, 'foregroundColor'>;
type _check_SlideBackgroundPatternFillOptions_foregroundColor = IsExact<Ref_SlideBackgroundPatternFillOptions_foregroundColor, Auth_SlideBackgroundPatternFillOptions_foregroundColor>;
type _assert_SlideBackgroundPatternFillOptions_foregroundColor = Expect<_check_SlideBackgroundPatternFillOptions_foregroundColor>;
export function _write_SlideBackgroundPatternFillOptions_foregroundColor(target: PptxEditor.SlideBackgroundPatternFillOptions, value: string | undefined): void {
  target['foregroundColor'] = value;
}

// PowerPoint.SlideBackgroundPatternFillOptions#pattern
type Ref_SlideBackgroundPatternFillOptions_pattern = { pattern?: "Cross" | "DarkDownwardDiagonal" | "DarkHorizontal" | "DarkUpwardDiagonal" | "DarkVertical" | "DashedDownwardDiagonal" | "DashedHorizontal" | "DashedUpwardDiagonal" | "DashedVertical" | "DiagonalBrick" | "DiagonalCross" | "Divot" | "DottedDiamond" | "DottedGrid" | "DownwardDiagonal" | "Horizontal" | "HorizontalBrick" | "LargeCheckerBoard" | "LargeConfetti" | "LargeGrid" | "LightDownwardDiagonal" | "LightHorizontal" | "LightUpwardDiagonal" | "LightVertical" | "NarrowHorizontal" | "NarrowVertical" | "OutlinedDiamond" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent5" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "Plaid" | "Shingle" | "SmallCheckerBoard" | "SmallConfetti" | "SmallGrid" | "SolidDiamond" | "Sphere" | "Trellis" | "Unsupported" | "UpwardDiagonal" | "Vertical" | "Wave" | "Weave" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "ZigZag" | PptxEditor.SlideBackgroundPatternFillType | undefined };
type Auth_SlideBackgroundPatternFillOptions_pattern = Pick<PptxEditor.SlideBackgroundPatternFillOptions, 'pattern'>;
type _check_SlideBackgroundPatternFillOptions_pattern = IsExact<Ref_SlideBackgroundPatternFillOptions_pattern, Auth_SlideBackgroundPatternFillOptions_pattern>;
type _assert_SlideBackgroundPatternFillOptions_pattern = Expect<_check_SlideBackgroundPatternFillOptions_pattern>;
export function _write_SlideBackgroundPatternFillOptions_pattern(target: PptxEditor.SlideBackgroundPatternFillOptions, value: "Cross" | "DarkDownwardDiagonal" | "DarkHorizontal" | "DarkUpwardDiagonal" | "DarkVertical" | "DashedDownwardDiagonal" | "DashedHorizontal" | "DashedUpwardDiagonal" | "DashedVertical" | "DiagonalBrick" | "DiagonalCross" | "Divot" | "DottedDiamond" | "DottedGrid" | "DownwardDiagonal" | "Horizontal" | "HorizontalBrick" | "LargeCheckerBoard" | "LargeConfetti" | "LargeGrid" | "LightDownwardDiagonal" | "LightHorizontal" | "LightUpwardDiagonal" | "LightVertical" | "NarrowHorizontal" | "NarrowVertical" | "OutlinedDiamond" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent5" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "Plaid" | "Shingle" | "SmallCheckerBoard" | "SmallConfetti" | "SmallGrid" | "SolidDiamond" | "Sphere" | "Trellis" | "Unsupported" | "UpwardDiagonal" | "Vertical" | "Wave" | "Weave" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "ZigZag" | PptxEditor.SlideBackgroundPatternFillType | undefined): void {
  target['pattern'] = value;
}

// PowerPoint.SlideBackgroundPictureOrTextureFillOptions

// PowerPoint.SlideBackgroundPictureOrTextureFillOptions#imageBase64
type Ref_SlideBackgroundPictureOrTextureFillOptions_imageBase64 = { imageBase64?: string | undefined };
type Auth_SlideBackgroundPictureOrTextureFillOptions_imageBase64 = Pick<PptxEditor.SlideBackgroundPictureOrTextureFillOptions, 'imageBase64'>;
type _check_SlideBackgroundPictureOrTextureFillOptions_imageBase64 = IsExact<Ref_SlideBackgroundPictureOrTextureFillOptions_imageBase64, Auth_SlideBackgroundPictureOrTextureFillOptions_imageBase64>;
type _assert_SlideBackgroundPictureOrTextureFillOptions_imageBase64 = Expect<_check_SlideBackgroundPictureOrTextureFillOptions_imageBase64>;
export function _write_SlideBackgroundPictureOrTextureFillOptions_imageBase64(target: PptxEditor.SlideBackgroundPictureOrTextureFillOptions, value: string | undefined): void {
  target['imageBase64'] = value;
}

// PowerPoint.SlideBackgroundPictureOrTextureFillOptions#transparency
type Ref_SlideBackgroundPictureOrTextureFillOptions_transparency = { transparency: number };
type Auth_SlideBackgroundPictureOrTextureFillOptions_transparency = Pick<PptxEditor.SlideBackgroundPictureOrTextureFillOptions, 'transparency'>;
type _check_SlideBackgroundPictureOrTextureFillOptions_transparency = IsExact<Ref_SlideBackgroundPictureOrTextureFillOptions_transparency, Auth_SlideBackgroundPictureOrTextureFillOptions_transparency>;
type _assert_SlideBackgroundPictureOrTextureFillOptions_transparency = Expect<_check_SlideBackgroundPictureOrTextureFillOptions_transparency>;
export function _write_SlideBackgroundPictureOrTextureFillOptions_transparency(target: PptxEditor.SlideBackgroundPictureOrTextureFillOptions, value: number): void {
  target['transparency'] = value;
}

// PowerPoint.CustomPropertyCollection

// PowerPoint.CustomPropertyCollection#add
type Ref_CustomPropertyCollection_add = { (key: string, value: boolean | Date | number | string): PptxEditor.CustomProperty };
type Auth_CustomPropertyCollection_add = PptxEditor.CustomPropertyCollection['add'];
type _check_CustomPropertyCollection_add = IsExact<Ref_CustomPropertyCollection_add, Auth_CustomPropertyCollection_add>;
type _assert_CustomPropertyCollection_add = Expect<_check_CustomPropertyCollection_add>;

// PowerPoint.CustomPropertyCollection#deleteAll
type Ref_CustomPropertyCollection_deleteAll = { (): void };
type Auth_CustomPropertyCollection_deleteAll = PptxEditor.CustomPropertyCollection['deleteAll'];
type _check_CustomPropertyCollection_deleteAll = IsExact<Ref_CustomPropertyCollection_deleteAll, Auth_CustomPropertyCollection_deleteAll>;
type _assert_CustomPropertyCollection_deleteAll = Expect<_check_CustomPropertyCollection_deleteAll>;

// PowerPoint.CustomPropertyCollection#getCount
type Ref_CustomPropertyCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_CustomPropertyCollection_getCount = PptxEditor.CustomPropertyCollection['getCount'];
type _check_CustomPropertyCollection_getCount = IsExact<Ref_CustomPropertyCollection_getCount, Auth_CustomPropertyCollection_getCount>;
type _assert_CustomPropertyCollection_getCount = Expect<_check_CustomPropertyCollection_getCount>;

// PowerPoint.CustomPropertyCollection#getItem
type Ref_CustomPropertyCollection_getItem = { (key: string): PptxEditor.CustomProperty };
type Auth_CustomPropertyCollection_getItem = PptxEditor.CustomPropertyCollection['getItem'];
type _check_CustomPropertyCollection_getItem = IsExact<Ref_CustomPropertyCollection_getItem, Auth_CustomPropertyCollection_getItem>;
type _assert_CustomPropertyCollection_getItem = Expect<_check_CustomPropertyCollection_getItem>;

// PowerPoint.CustomPropertyCollection#getItemAt
type Ref_CustomPropertyCollection_getItemAt = { (index: number): PptxEditor.CustomProperty };
type Auth_CustomPropertyCollection_getItemAt = PptxEditor.CustomPropertyCollection['getItemAt'];
type _check_CustomPropertyCollection_getItemAt = IsExact<Ref_CustomPropertyCollection_getItemAt, Auth_CustomPropertyCollection_getItemAt>;
type _assert_CustomPropertyCollection_getItemAt = Expect<_check_CustomPropertyCollection_getItemAt>;

// PowerPoint.CustomPropertyCollection#getItemOrNullObject
type Ref_CustomPropertyCollection_getItemOrNullObject = { (key: string): PptxEditor.CustomProperty };
type Auth_CustomPropertyCollection_getItemOrNullObject = PptxEditor.CustomPropertyCollection['getItemOrNullObject'];
type _check_CustomPropertyCollection_getItemOrNullObject = IsExact<Ref_CustomPropertyCollection_getItemOrNullObject, Auth_CustomPropertyCollection_getItemOrNullObject>;
type _assert_CustomPropertyCollection_getItemOrNullObject = Expect<_check_CustomPropertyCollection_getItemOrNullObject>;

// PowerPoint.CustomPropertyCollection#items
type Ref_CustomPropertyCollection_items = { readonly items: PptxEditor.CustomProperty[] };
type Auth_CustomPropertyCollection_items = Pick<PptxEditor.CustomPropertyCollection, 'items'>;
type _check_CustomPropertyCollection_items = IsExact<Ref_CustomPropertyCollection_items, Auth_CustomPropertyCollection_items>;
type _assert_CustomPropertyCollection_items = Expect<_check_CustomPropertyCollection_items>;

// PowerPoint.CustomProperty

// PowerPoint.CustomProperty#delete
type Ref_CustomProperty_delete = { (): void };
type Auth_CustomProperty_delete = PptxEditor.CustomProperty['delete'];
type _check_CustomProperty_delete = IsExact<Ref_CustomProperty_delete, Auth_CustomProperty_delete>;
type _assert_CustomProperty_delete = Expect<_check_CustomProperty_delete>;

// PowerPoint.CustomProperty#key
type Ref_CustomProperty_key = { readonly key: string };
type Auth_CustomProperty_key = Pick<PptxEditor.CustomProperty, 'key'>;
type _check_CustomProperty_key = IsExact<Ref_CustomProperty_key, Auth_CustomProperty_key>;
type _assert_CustomProperty_key = Expect<_check_CustomProperty_key>;

// PowerPoint.CustomProperty#type
type Ref_CustomProperty_type = { readonly type: "Boolean" | "Date" | "Number" | "String" | PptxEditor.DocumentPropertyType };
type Auth_CustomProperty_type = Pick<PptxEditor.CustomProperty, 'type'>;
type _check_CustomProperty_type = IsExact<Ref_CustomProperty_type, Auth_CustomProperty_type>;
type _assert_CustomProperty_type = Expect<_check_CustomProperty_type>;

// PowerPoint.CustomProperty#value
type Ref_CustomProperty_value = { value: boolean | Date | number | string };
type Auth_CustomProperty_value = Pick<PptxEditor.CustomProperty, 'value'>;
type _check_CustomProperty_value = IsExact<Ref_CustomProperty_value, Auth_CustomProperty_value>;
type _assert_CustomProperty_value = Expect<_check_CustomProperty_value>;
export function _write_CustomProperty_value(target: PptxEditor.CustomProperty, value: boolean | Date | number | string): void {
  target['value'] = value;
}

// PowerPoint.CustomXmlPartCollection

// PowerPoint.CustomXmlPartCollection#add
type Ref_CustomXmlPartCollection_add = { (xml: string): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartCollection_add = PptxEditor.CustomXmlPartCollection['add'];
type _check_CustomXmlPartCollection_add = IsExact<Ref_CustomXmlPartCollection_add, Auth_CustomXmlPartCollection_add>;
type _assert_CustomXmlPartCollection_add = Expect<_check_CustomXmlPartCollection_add>;

// PowerPoint.CustomXmlPartCollection#getByNamespace
type Ref_CustomXmlPartCollection_getByNamespace = { (namespaceUri: string): PptxEditor.CustomXmlPartScopedCollection };
type Auth_CustomXmlPartCollection_getByNamespace = PptxEditor.CustomXmlPartCollection['getByNamespace'];
type _check_CustomXmlPartCollection_getByNamespace = IsExact<Ref_CustomXmlPartCollection_getByNamespace, Auth_CustomXmlPartCollection_getByNamespace>;
type _assert_CustomXmlPartCollection_getByNamespace = Expect<_check_CustomXmlPartCollection_getByNamespace>;

// PowerPoint.CustomXmlPartCollection#getCount
type Ref_CustomXmlPartCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_CustomXmlPartCollection_getCount = PptxEditor.CustomXmlPartCollection['getCount'];
type _check_CustomXmlPartCollection_getCount = IsExact<Ref_CustomXmlPartCollection_getCount, Auth_CustomXmlPartCollection_getCount>;
type _assert_CustomXmlPartCollection_getCount = Expect<_check_CustomXmlPartCollection_getCount>;

// PowerPoint.CustomXmlPartCollection#getItem
type Ref_CustomXmlPartCollection_getItem = { (id: string): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartCollection_getItem = PptxEditor.CustomXmlPartCollection['getItem'];
type _check_CustomXmlPartCollection_getItem = IsExact<Ref_CustomXmlPartCollection_getItem, Auth_CustomXmlPartCollection_getItem>;
type _assert_CustomXmlPartCollection_getItem = Expect<_check_CustomXmlPartCollection_getItem>;

// PowerPoint.CustomXmlPartCollection#getItemAt
type Ref_CustomXmlPartCollection_getItemAt = { (index: number): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartCollection_getItemAt = PptxEditor.CustomXmlPartCollection['getItemAt'];
type _check_CustomXmlPartCollection_getItemAt = IsExact<Ref_CustomXmlPartCollection_getItemAt, Auth_CustomXmlPartCollection_getItemAt>;
type _assert_CustomXmlPartCollection_getItemAt = Expect<_check_CustomXmlPartCollection_getItemAt>;

// PowerPoint.CustomXmlPartCollection#getItemOrNullObject
type Ref_CustomXmlPartCollection_getItemOrNullObject = { (id: string): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartCollection_getItemOrNullObject = PptxEditor.CustomXmlPartCollection['getItemOrNullObject'];
type _check_CustomXmlPartCollection_getItemOrNullObject = IsExact<Ref_CustomXmlPartCollection_getItemOrNullObject, Auth_CustomXmlPartCollection_getItemOrNullObject>;
type _assert_CustomXmlPartCollection_getItemOrNullObject = Expect<_check_CustomXmlPartCollection_getItemOrNullObject>;

// PowerPoint.CustomXmlPartCollection#items
type Ref_CustomXmlPartCollection_items = { readonly items: PptxEditor.CustomXmlPart[] };
type Auth_CustomXmlPartCollection_items = Pick<PptxEditor.CustomXmlPartCollection, 'items'>;
type _check_CustomXmlPartCollection_items = IsExact<Ref_CustomXmlPartCollection_items, Auth_CustomXmlPartCollection_items>;
type _assert_CustomXmlPartCollection_items = Expect<_check_CustomXmlPartCollection_items>;

// PowerPoint.CustomXmlPartScopedCollection

// PowerPoint.CustomXmlPartScopedCollection#getCount
type Ref_CustomXmlPartScopedCollection_getCount = { (): PptxEditor.ClientResult<number> };
type Auth_CustomXmlPartScopedCollection_getCount = PptxEditor.CustomXmlPartScopedCollection['getCount'];
type _check_CustomXmlPartScopedCollection_getCount = IsExact<Ref_CustomXmlPartScopedCollection_getCount, Auth_CustomXmlPartScopedCollection_getCount>;
type _assert_CustomXmlPartScopedCollection_getCount = Expect<_check_CustomXmlPartScopedCollection_getCount>;

// PowerPoint.CustomXmlPartScopedCollection#getItem
type Ref_CustomXmlPartScopedCollection_getItem = { (id: string): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartScopedCollection_getItem = PptxEditor.CustomXmlPartScopedCollection['getItem'];
type _check_CustomXmlPartScopedCollection_getItem = IsExact<Ref_CustomXmlPartScopedCollection_getItem, Auth_CustomXmlPartScopedCollection_getItem>;
type _assert_CustomXmlPartScopedCollection_getItem = Expect<_check_CustomXmlPartScopedCollection_getItem>;

// PowerPoint.CustomXmlPartScopedCollection#getItemAt
type Ref_CustomXmlPartScopedCollection_getItemAt = { (index: number): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartScopedCollection_getItemAt = PptxEditor.CustomXmlPartScopedCollection['getItemAt'];
type _check_CustomXmlPartScopedCollection_getItemAt = IsExact<Ref_CustomXmlPartScopedCollection_getItemAt, Auth_CustomXmlPartScopedCollection_getItemAt>;
type _assert_CustomXmlPartScopedCollection_getItemAt = Expect<_check_CustomXmlPartScopedCollection_getItemAt>;

// PowerPoint.CustomXmlPartScopedCollection#getItemOrNullObject
type Ref_CustomXmlPartScopedCollection_getItemOrNullObject = { (id: string): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartScopedCollection_getItemOrNullObject = PptxEditor.CustomXmlPartScopedCollection['getItemOrNullObject'];
type _check_CustomXmlPartScopedCollection_getItemOrNullObject = IsExact<Ref_CustomXmlPartScopedCollection_getItemOrNullObject, Auth_CustomXmlPartScopedCollection_getItemOrNullObject>;
type _assert_CustomXmlPartScopedCollection_getItemOrNullObject = Expect<_check_CustomXmlPartScopedCollection_getItemOrNullObject>;

// PowerPoint.CustomXmlPartScopedCollection#getOnlyItem
type Ref_CustomXmlPartScopedCollection_getOnlyItem = { (): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartScopedCollection_getOnlyItem = PptxEditor.CustomXmlPartScopedCollection['getOnlyItem'];
type _check_CustomXmlPartScopedCollection_getOnlyItem = IsExact<Ref_CustomXmlPartScopedCollection_getOnlyItem, Auth_CustomXmlPartScopedCollection_getOnlyItem>;
type _assert_CustomXmlPartScopedCollection_getOnlyItem = Expect<_check_CustomXmlPartScopedCollection_getOnlyItem>;

// PowerPoint.CustomXmlPartScopedCollection#getOnlyItemOrNullObject
type Ref_CustomXmlPartScopedCollection_getOnlyItemOrNullObject = { (): PptxEditor.CustomXmlPart };
type Auth_CustomXmlPartScopedCollection_getOnlyItemOrNullObject = PptxEditor.CustomXmlPartScopedCollection['getOnlyItemOrNullObject'];
type _check_CustomXmlPartScopedCollection_getOnlyItemOrNullObject = IsExact<Ref_CustomXmlPartScopedCollection_getOnlyItemOrNullObject, Auth_CustomXmlPartScopedCollection_getOnlyItemOrNullObject>;
type _assert_CustomXmlPartScopedCollection_getOnlyItemOrNullObject = Expect<_check_CustomXmlPartScopedCollection_getOnlyItemOrNullObject>;

// PowerPoint.CustomXmlPartScopedCollection#items
type Ref_CustomXmlPartScopedCollection_items = { readonly items: PptxEditor.CustomXmlPart[] };
type Auth_CustomXmlPartScopedCollection_items = Pick<PptxEditor.CustomXmlPartScopedCollection, 'items'>;
type _check_CustomXmlPartScopedCollection_items = IsExact<Ref_CustomXmlPartScopedCollection_items, Auth_CustomXmlPartScopedCollection_items>;
type _assert_CustomXmlPartScopedCollection_items = Expect<_check_CustomXmlPartScopedCollection_items>;

// PowerPoint.CustomXmlPart

// PowerPoint.CustomXmlPart#delete
type Ref_CustomXmlPart_delete = { (): void };
type Auth_CustomXmlPart_delete = PptxEditor.CustomXmlPart['delete'];
type _check_CustomXmlPart_delete = IsExact<Ref_CustomXmlPart_delete, Auth_CustomXmlPart_delete>;
type _assert_CustomXmlPart_delete = Expect<_check_CustomXmlPart_delete>;

// PowerPoint.CustomXmlPart#getXml
type Ref_CustomXmlPart_getXml = { (): PptxEditor.ClientResult<string> };
type Auth_CustomXmlPart_getXml = PptxEditor.CustomXmlPart['getXml'];
type _check_CustomXmlPart_getXml = IsExact<Ref_CustomXmlPart_getXml, Auth_CustomXmlPart_getXml>;
type _assert_CustomXmlPart_getXml = Expect<_check_CustomXmlPart_getXml>;

// PowerPoint.CustomXmlPart#id
type Ref_CustomXmlPart_id = { readonly id: string };
type Auth_CustomXmlPart_id = Pick<PptxEditor.CustomXmlPart, 'id'>;
type _check_CustomXmlPart_id = IsExact<Ref_CustomXmlPart_id, Auth_CustomXmlPart_id>;
type _assert_CustomXmlPart_id = Expect<_check_CustomXmlPart_id>;

// PowerPoint.CustomXmlPart#namespaceUri
type Ref_CustomXmlPart_namespaceUri = { readonly namespaceUri: string };
type Auth_CustomXmlPart_namespaceUri = Pick<PptxEditor.CustomXmlPart, 'namespaceUri'>;
type _check_CustomXmlPart_namespaceUri = IsExact<Ref_CustomXmlPart_namespaceUri, Auth_CustomXmlPart_namespaceUri>;
type _assert_CustomXmlPart_namespaceUri = Expect<_check_CustomXmlPart_namespaceUri>;

// PowerPoint.CustomXmlPart#setXml
type Ref_CustomXmlPart_setXml = { (xml: string): void };
type Auth_CustomXmlPart_setXml = PptxEditor.CustomXmlPart['setXml'];
type _check_CustomXmlPart_setXml = IsExact<Ref_CustomXmlPart_setXml, Auth_CustomXmlPart_setXml>;
type _assert_CustomXmlPart_setXml = Expect<_check_CustomXmlPart_setXml>;

// PowerPoint.SlideGetImageOptions

// PowerPoint.SlideGetImageOptions#height
type Ref_SlideGetImageOptions_height = { height?: number | undefined };
type Auth_SlideGetImageOptions_height = Pick<PptxEditor.SlideGetImageOptions, 'height'>;
type _check_SlideGetImageOptions_height = IsExact<Ref_SlideGetImageOptions_height, Auth_SlideGetImageOptions_height>;
type _assert_SlideGetImageOptions_height = Expect<_check_SlideGetImageOptions_height>;
export function _write_SlideGetImageOptions_height(target: PptxEditor.SlideGetImageOptions, value: number | undefined): void {
  target['height'] = value;
}

// PowerPoint.SlideGetImageOptions#width
type Ref_SlideGetImageOptions_width = { width?: number | undefined };
type Auth_SlideGetImageOptions_width = Pick<PptxEditor.SlideGetImageOptions, 'width'>;
type _check_SlideGetImageOptions_width = IsExact<Ref_SlideGetImageOptions_width, Auth_SlideGetImageOptions_width>;
type _assert_SlideGetImageOptions_width = Expect<_check_SlideGetImageOptions_width>;
export function _write_SlideGetImageOptions_width(target: PptxEditor.SlideGetImageOptions, value: number | undefined): void {
  target['width'] = value;
}

// PowerPoint.ShapeGetImageOptions

// PowerPoint.ShapeGetImageOptions#format
type Ref_ShapeGetImageOptions_format = { format?: "Png" | PptxEditor.ShapeGetImageFormatType | undefined };
type Auth_ShapeGetImageOptions_format = Pick<PptxEditor.ShapeGetImageOptions, 'format'>;
type _check_ShapeGetImageOptions_format = IsExact<Ref_ShapeGetImageOptions_format, Auth_ShapeGetImageOptions_format>;
type _assert_ShapeGetImageOptions_format = Expect<_check_ShapeGetImageOptions_format>;
export function _write_ShapeGetImageOptions_format(target: PptxEditor.ShapeGetImageOptions, value: "Png" | PptxEditor.ShapeGetImageFormatType | undefined): void {
  target['format'] = value;
}

// PowerPoint.ShapeGetImageOptions#height
type Ref_ShapeGetImageOptions_height = { height?: number | undefined };
type Auth_ShapeGetImageOptions_height = Pick<PptxEditor.ShapeGetImageOptions, 'height'>;
type _check_ShapeGetImageOptions_height = IsExact<Ref_ShapeGetImageOptions_height, Auth_ShapeGetImageOptions_height>;
type _assert_ShapeGetImageOptions_height = Expect<_check_ShapeGetImageOptions_height>;
export function _write_ShapeGetImageOptions_height(target: PptxEditor.ShapeGetImageOptions, value: number | undefined): void {
  target['height'] = value;
}

// PowerPoint.ShapeGetImageOptions#width
type Ref_ShapeGetImageOptions_width = { width?: number | undefined };
type Auth_ShapeGetImageOptions_width = Pick<PptxEditor.ShapeGetImageOptions, 'width'>;
type _check_ShapeGetImageOptions_width = IsExact<Ref_ShapeGetImageOptions_width, Auth_ShapeGetImageOptions_width>;
type _assert_ShapeGetImageOptions_width = Expect<_check_ShapeGetImageOptions_width>;
export function _write_ShapeGetImageOptions_width(target: PptxEditor.ShapeGetImageOptions, value: number | undefined): void {
  target['width'] = value;
}
