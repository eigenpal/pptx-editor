/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public diagnostics of the engine's edit refusals, beside `errors.ts` (which keeps the
// shape of DOCX Editor's module). Each row maps a host detail to a stable code and a fixed hint:
// the cause, then the action. `scripts/refusal-reasons.mjs` lists every reason the core can
// give through an automation host, and `scripts/__tests__/refusal-reasons.test.ts` fails when
// one has no diagnostic.

/** Rows of `[host detail, public code, hint]`. */
export const EDIT_DIAGNOSTICS: readonly (readonly [string, string, string])[] = [
  // Edit refusals of the engine (`disabled.*`): NotSupported.
  [
    'disabled.cropFillMode',
    'edit.picture-not-stretched',
    'The picture is tiled or inset in its frame, so it cannot be cropped. Use a stretched picture.',
  ],
  [
    'disabled.cropLocked',
    'edit.picture-locked',
    'The picture is locked against this change in the file. Remove the lock in the file first.',
  ],
  [
    'disabled.rotationLocked',
    'edit.picture-locked',
    'The picture is locked against rotation in the file. Remove the lock in the file first.',
  ],
  [
    'disabled.sizeLocked',
    'edit.picture-locked',
    'The picture is locked against resizing in the file. Remove the lock in the file first.',
  ],
  [
    'disabled.shapeLocked',
    'edit.picture-locked',
    'The picture is locked against a change of shape in the file. Remove the lock in the file first.',
  ],
  [
    'disabled.decorative',
    'edit.decorative',
    'The shape is decorative, so it has no alternative text. Set isDecorative to false first.',
  ],
  [
    'disabled.distributeNeedsThree',
    'edit.needs-three-shapes',
    'Distributing needs three or more shapes. Pass more shapes.',
  ],
  [
    'disabled.groupNeedsTwo',
    'edit.group-count',
    'A group needs two or more different shapes. Remove repeated shapes or add more.',
  ],
  [
    'disabled.groupScope',
    'edit.group-parent',
    'The shapes are not in the same parent. Use shapes of one slide, or of one group.',
  ],
  [
    'disabled.imageSizeUnknown',
    'edit.picture-size-unknown',
    'The size of the picture cannot be read from its bytes. Use a PNG, JPEG, or GIF picture.',
  ],
  [
    'disabled.mergeOneCell',
    'edit.merge-one-cell',
    'A merge needs two or more cells. Pass a row count or a column count above 1.',
  ],
  [
    'disabled.mixedSize',
    'edit.mixed-size',
    'The text has more than one size, so it cannot be stepped as one. Set font.size instead.',
  ],
  [
    'disabled.noArrangeChange',
    'edit.no-change',
    'The shapes are already in that position. Read zOrderPosition before you move a shape.',
  ],
  [
    'disabled.noEffectChange',
    'edit.no-change',
    'The shapes do not have that effect, so nothing changes.',
  ],
  [
    'disabled.noLockChange',
    'edit.no-change',
    'The setting already has that value, so nothing changes.',
  ],
  [
    'disabled.noNameChange',
    'edit.no-change',
    'The shape already has that name, so nothing changes.',
  ],
  [
    'disabled.noAdjustChange',
    'edit.no-change',
    'The shape already has those adjust values, so nothing changes.',
  ],
  [
    'disabled.noAnimationChange',
    'edit.no-change',
    'The effect or the transition already has those settings, so nothing changes.',
  ],
  [
    'disabled.animationsReadOnly',
    'edit.animations-read-only',
    'The slide keeps its effects in markup that the editor does not change. Edit them in the file.',
  ],
  [
    'disabled.noAnimation',
    'edit.no-animation',
    'The slide has no main sequence effect with that id. Read the effect ids of the slide first.',
  ],
  [
    'disabled.instantAnimation',
    'edit.instant-animation',
    'The effect is instant, so it has no duration. Change its delay instead.',
  ],
  [
    'disabled.noAdjustHandles',
    'edit.unsupported-element',
    'Only a preset shape with adjust handles has adjust values. Select such a shape.',
  ],
  [
    'disabled.noSizeChange',
    'edit.no-change',
    'The slides already have that size, so nothing changes.',
  ],
  [
    'disabled.noVisibilityChange',
    'edit.no-change',
    'The shapes already have that visibility, so nothing changes.',
  ],
  [
    'disabled.noEditor',
    'host.no-document',
    'The editor has no presentation open yet. Wait until it has loaded one.',
  ],
  [
    'disabled.noList',
    'edit.no-list',
    'The paragraphs have no bullets or numbers. Set bulletFormat.visible to true first.',
  ],
  ['disabled.noMatch', 'edit.no-match', 'The text was not found. Read the text again.'],
  ['disabled.noNotes', 'edit.no-notes', 'The slide has no speaker notes. Add notes first.'],
  ['disabled.noPicture', 'edit.not-a-picture', 'The shape is not a picture. Use a picture shape.'],
  [
    'disabled.onePicture',
    'edit.not-a-picture',
    'This change takes exactly one picture. Pass one picture shape.',
  ],
  [
    'disabled.pictureEffects',
    'edit.picture-effects',
    'The picture has effects this API cannot change. Change another property, or remove the effects in the file.',
  ],
  [
    'disabled.noSections',
    'edit.no-sections',
    'The presentation has no sections. Add a section first.',
  ],
  [
    'disabled.sectionEdge',
    'edit.section-edge',
    'The section is already at that edge of the presentation, so it cannot move further.',
  ],
  [
    'disabled.noSelection',
    'edit.no-target',
    'The call has no shape to act on. Pass a shape, or select one in the editor first.',
  ],
  ['disabled.selectOne', 'edit.no-target', 'This change takes exactly one shape. Pass one shape.'],
  [
    'disabled.selectText',
    'edit.no-target',
    'The call needs a text range. Pass a text range with text in it.',
  ],
  [
    'disabled.selectWord',
    'edit.no-target',
    'The call needs text. Pass a text range that is not empty.',
  ],
  [
    'disabled.selectNotesText',
    'edit.no-target',
    'The call needs the text of the notes. Address the text of the notes placeholder.',
  ],
  [
    'disabled.noShapeEffects',
    'edit.no-effects',
    'The shape cannot take effects. Use a shape, a line, or a picture.',
  ],
  [
    'disabled.noShapeFill',
    'edit.no-fill',
    'The shape cannot have a fill, for example a line or a connector. Use a shape with an area.',
  ],
  [
    'disabled.noShapeStyle',
    'edit.no-style',
    'The shape cannot take a shape style. Use a geometric shape.',
  ],
  [
    'disabled.noSlideNumberPlaceholder',
    'edit.no-slide-number-placeholder',
    'The layouts have no slide number placeholder. Add one to the layout first.',
  ],
  [
    'disabled.noFooterPlaceholder',
    'edit.no-footer-placeholder',
    'The layouts and the master have no footer placeholder. Add one to the layout first.',
  ],
  [
    'disabled.noDatePlaceholder',
    'edit.no-date-placeholder',
    'The layouts and the master have no date placeholder. Add one to the layout first.',
  ],
  [
    'disabled.noFooterText',
    'edit.no-footer-text',
    'No slide shows a footer, so there is no text to keep. Pass text.',
  ],
  ['disabled.noTable', 'edit.not-a-table', 'The shape is not a table. Use a table shape.'],
  [
    'disabled.notAGroup',
    'edit.not-a-group',
    'The shape is not a group. Check shape.type first, and use a group shape.',
  ],
  ['disabled.notALine', 'edit.not-a-line', 'The shape is not a line or a connector. Use a line.'],
  [
    'disabled.noText',
    'edit.no-text-frame',
    'The shape has no text frame. Check shape.type first, and read text only of a shape that holds text.',
  ],
  [
    'disabled.slideNotCopyable',
    'import.slide-not-copyable',
    'A slide to import has threaded comments or an embedded control. Leave it out with sourceSlideIds.',
  ],
  [
    'disabled.slideMaster',
    'host.slide-master',
    'The editor shows the slide master view, which takes no slide edits. Close the slide master first.',
  ],
  [
    'disabled.noSlideMaster',
    'host.no-slide-master',
    'The presentation has no slide master, so the slide master view cannot open. Stay in the normal view.',
  ],
  [
    'disabled.slideSorter',
    'host.slide-sorter',
    'The editor shows the slide sorter, which takes no shape edits. Switch the editor to the normal view.',
  ],
  [
    'disabled.splitOneCell',
    'edit.split-one-cell',
    'A split takes one cell. Use getCellOrNullObject for one cell, then split it.',
  ],
  [
    'disabled.splitUneven',
    'edit.split-uneven',
    'A merged cell splits only into parts that divide it evenly. Use a row and column count that divides its span.',
  ],
  [
    'disabled.tableText',
    'edit.table-text',
    'A table cell does not take this text body setting. Set it on a text box or a shape.',
  ],
  [
    'disabled.tableTooLarge',
    'edit.table-too-large',
    'A table has at most 75 columns and 1,000 rows. Use fewer rows or columns.',
  ],
  [
    'disabled.unsupportedElement',
    'edit.unsupported-element',
    'The object is of a kind this API cannot edit, for example a chart or an embedded object. Read it, or edit another object.',
  ],
  [
    'disabled.noCommentAuthor',
    'edit.no-comment-author',
    'A comment needs an author. Pass author, or set the author option of the editor.',
  ],
  [
    'disabled.commentNoStatus',
    'edit.comment-no-status',
    'Legacy comments have no resolved state. Reply to the thread instead.',
  ],
  [
    'disabled.commentNotEditable',
    'edit.comment-not-editable',
    'The comment holds content other than text, such as a mention. Reply instead of editing it.',
  ],
  [
    'disabled.noCommentTarget',
    'edit.no-comment-target',
    'A comment can point only at an element on the top level of the current slide. Pass its id, or none.',
  ],
  // Edit errors of the engine (`errors.*`): GeneralException, or InvalidArgument for a name.
  [
    'errors.commentText',
    'edit.comment-text',
    'Comment text has 1 to 10000 characters. Pass shorter text that is not empty.',
  ],
  [
    'errors.commentAuthor',
    'edit.comment-author',
    'An author name has up to 255 characters. Pass a shorter name.',
  ],
  [
    'errors.commentDate',
    'edit.comment-date',
    'A comment date is an ISO 8601 date and time, such as 2026-11-07T10:00:00Z. Pass one, or none.',
  ],
  [
    'errors.noComment',
    'edit.comment-not-found',
    'The slide has no comment with this id. Read the comments of the slide again.',
  ],
  [
    'errors.commentChanged',
    'edit.comment-changed',
    'The comment at this id no longer matches expect. Read the comments of the slide again.',
  ],
  [
    'errors.commentAmbiguous',
    'edit.comment-ambiguous',
    'Another comment of the slide has the same author, date, and text. Pass the comment identity in expect to name one of them.',
  ],
  [
    'errors.altTextLength',
    'edit.alt-text-length',
    'Alternative text can have up to 4000 characters. Shorten it.',
  ],
  [
    'errors.clipboard',
    'edit.clipboard',
    'The content is not shapes this API can insert. Insert shapes from a presentation instead.',
  ],
  [
    'errors.coordinateRange',
    'edit.coordinate-range',
    'The position or size is outside the range a slide can hold. Use values within the slide size limits.',
  ],
  [
    'errors.slideSizeRange',
    'edit.coordinate-range',
    'A slide side is from 1 to 56 inches (914400 to 51206400 EMU). Pass a size in that range.',
  ],
  [
    'errors.chartData.noChart',
    'edit.chart-data',
    'Select one chart frame, or pass the node id of a chart on the current slide.',
  ],
  [
    'errors.chartData.noData',
    'edit.chart-data',
    'The chart has no data series to edit. Edit a chart that plots at least one series.',
  ],
  [
    'errors.chartData.combination',
    'edit.chart-data-unsupported',
    'The chart combines chart types, and its data is read-only. Edit a chart with one chart type.',
  ],
  [
    'errors.chartData.unsupportedType',
    'edit.chart-data-unsupported',
    'Only bar, column, line, area, pie, doughnut, radar, and scatter chart data is editable.',
  ],
  [
    'errors.chartData.filtered',
    'edit.chart-data-unsupported',
    'The chart hides some of its series, and its data is read-only.',
  ],
  [
    'errors.chartData.multiLevel',
    'edit.chart-data-unsupported',
    'The chart has multi-level categories, and its data is read-only.',
  ],
  [
    'errors.chartData.mixedCategories',
    'edit.chart-data-unsupported',
    'The series of the chart use different categories, and its data is read-only.',
  ],
  [
    'errors.chartData.tooLarge',
    'edit.chart-data-range',
    'A chart edits up to 255 series and 4000 categories. Pass less data.',
  ],
  [
    'errors.chartData.empty',
    'edit.chart-data-range',
    'A chart needs at least one series and one category. Keep one of each.',
  ],
  [
    'errors.chartData.invalidText',
    'edit.chart-data-text',
    'A name or label holds characters that XML cannot keep, or is longer than 255 characters. Remove them.',
  ],
  [
    'errors.chartData.notNumber',
    'edit.chart-data-value',
    'Each value must be a finite number or null. Pass numbers, and finite X values for a scatter chart.',
  ],
  [
    'errors.chartData.workbook',
    'edit.chart-data-layout',
    'The embedded workbook cannot take the new table. Give each series a different name, and keep the cells beside the data empty.',
  ],
  [
    'errors.chartFormat.noAxes',
    'edit.chart-format-axes',
    'A pie or doughnut chart has no axes. Leave out gridlines, axes, and axis titles, or change the chart type first.',
  ],
  [
    'errors.chartFormat.sameType',
    'edit.chart-type-same',
    'The chart already has this type. Read chartFormat().type and send setChartType only for another type.',
  ],
  [
    'errors.chartData.layout',
    'edit.chart-data-layout',
    'The cell references of the chart do not form one table. Change only names and values.',
  ],
  [
    'errors.diagram.noDiagram',
    'edit.diagram',
    'Select one diagram frame, or pass the node id of a diagram on the current slide.',
  ],
  [
    'errors.diagram.noNode',
    'edit.diagram',
    'Put the caret in a diagram shape, or pass the model id of a content node as `node`.',
  ],
  [
    'errors.diagram.noText',
    'edit.diagram-text',
    'No single diagram shape with a paragraph shows this node. Edit the text of each node instead.',
  ],
  [
    'errors.diagram.textTooLong',
    'edit.diagram-range',
    'Pass up to 10,000 characters and 1,000 paragraphs as `text` of `setDiagramText`.',
  ],
  [
    'errors.diagram.lastShape',
    'edit.diagram-range',
    'A diagram keeps at least one top-level node. Add a node before you delete this one.',
  ],
  [
    'errors.diagram.layoutUnsupported',
    'edit.diagram-unsupported',
    'Only a diagram whose saved drawing is a line of blocks lays out again. Edit the text of its nodes instead.',
  ],
  [
    'errors.diagram.noChildren',
    'edit.diagram-unsupported',
    'The diagram draws no level below this node. Add a node at the same level instead.',
  ],
  [
    'errors.diagram.tooLarge',
    'edit.diagram-range',
    'A diagram data model holds up to 4096 points. Delete nodes before you add more.',
  ],
  [
    'errors.exportFailed',
    'image.failed',
    'The picture could not be made. Read the shape, then try again.',
  ],
  [
    'errors.exportSize',
    'image.invalid-size',
    'A picture is from 16 to 8192 pixels on a side. Pass a size in that range.',
  ],
  [
    'errors.findStale',
    'edit.stale-target',
    'An edit removed the object. Read the presentation again.',
  ],
  [
    'errors.findTooLong',
    'edit.text-too-long',
    'The text is too long for this call. Use shorter text.',
  ],
  [
    'errors.invalidName',
    'edit.invalid-name',
    'A name has 1 to 255 characters. Use a shorter name that is not empty.',
  ],
  [
    'errors.invalidText',
    'edit.invalid-text',
    'The text has no characters that a slide can hold, such as control characters only. Remove them.',
  ],
  [
    'errors.customShowNotFound',
    'edit.custom-show-not-found',
    'The presentation has no custom show with this id. Read the custom shows of the deck again.',
  ],
  [
    'errors.layoutNotFound',
    'edit.layout-not-found',
    'The presentation has no such slide layout. Read the layouts of the master again.',
  ],
  [
    'errors.malformed',
    'edit.malformed-part',
    'A part of the presentation is damaged, so the change cannot apply. Repair the file, then try again.',
  ],
  [
    'errors.sectionName',
    'edit.section-name',
    'A section name has 1 to 255 characters. Use a shorter name that is not empty.',
  ],
  [
    'errors.sectionNotFound',
    'edit.section-not-found',
    'The presentation has no such section. Read the sections again.',
  ],
  [
    'errors.slideNotFound',
    'edit.slide-not-found',
    'The presentation has no such slide. Read the slides again.',
  ],
  [
    'errors.tooLarge',
    'edit.too-large',
    'The content is larger than the presentation accepts. Use smaller content.',
  ],
  [
    'errors.unsafeLink',
    'argument.not-accepted',
    'Use an http, https, mailto, tel, or ftp address of up to 2048 characters.',
  ],
  [
    'errors.unsafeFileLink',
    'argument.not-accepted',
    'Use a relative path that starts with ./ or ../, with at most two ../ steps, to a document, picture, sound, or video file.',
  ],
  [
    'errors.unsupported',
    'edit.unsupported',
    'The presentation cannot take this change. Change another property.',
  ],
  [
    'partly-applied',
    'edit.partly-applied',
    'A read in this sync needed the result of the writes before it, so those writes applied first. The call in errorLocation and the calls after it did not apply.',
  ],
  // Host refusals with a detail of their own.
  ['document', 'host.no-document', 'The host holds no presentation. Open one before you run.'],
  [
    'damaged',
    'host.damaged-document',
    'A slide of the presentation did not open, so the editor holds it read-only. Repair the file and open it again.',
  ],
  [
    'internal',
    'host.internal',
    'The engine could not complete the request, and nothing changed. Report the call that failed.',
  ],
  [
    'no-theme',
    'edit.no-theme-color',
    'The master has no theme, or the theme has no such color. Use a theme color of the master.',
  ],
  ['part', 'edit.not-xml-part', 'That part of the presentation is not XML. Use an XML part.'],
  [
    'picture',
    'edit.picture-not-in-package',
    'The background picture is a theme fill or a link outside the file. Pass imageBase64 to set a picture.',
  ],
  [
    'rotation',
    'edit.no-rotation',
    'A table or another graphic frame does not rotate. Rotate another shape.',
  ],
  [
    'transparency',
    'edit.fill-not-solid',
    'Only a solid fill has a transparency. Call setSolidColor first.',
  ],
  [
    'result',
    'edit.stale-target',
    'The object is not available in this batch. Read the presentation again.',
  ],
  [
    'parentGroup',
    'edit.not-in-group',
    'The shape is not in a group. Check shape.level first: 0 is the top level of the slide.',
  ],
  [
    'group',
    'edit.not-a-group',
    'The shape is not a group. Check shape.type first, and use a group shape.',
  ],
];

/**
 * The diagnostic of a host refusal that has no detail of its own, by host code: an argument
 * that is out of range or of the wrong form, a missing item, a stale read, a change the store
 * refused as a whole, and an operation the host does not know.
 */
export const CODE_DIAGNOSTICS: Readonly<Record<string, readonly [string, string]>> = Object.freeze({
  'invalid-argument': [
    'argument.not-accepted',
    'An argument is out of range or of the wrong form. errorLocation names the call; its reference lists the values it takes.',
  ],
  'item-not-found': [
    'edit.item-not-found',
    'The collection has no item with that id, key, or index. Read the items again.',
  ],
  'stale-revision': [
    'edit.stale-document',
    'Another writer changed the presentation after this context read it. Read again before you write.',
  ],
  'transaction-refused': [
    'edit.refused',
    'The presentation refused the change as a whole, and nothing applied. Change one property at a time to find the cause.',
  ],
  disposed: ['host.disposed', 'The host was disposed. Create a new runtime.'],
  'document-unavailable': [
    'host.no-document',
    'The editor holds no presentation right now. Wait until it has loaded one.',
  ],
  'unknown-operation': [
    'host.unknown-operation',
    'The host does not know this call. Use a host of the same package version.',
  ],
});
