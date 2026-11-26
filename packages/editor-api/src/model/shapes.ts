/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Shapes: Office.js `Shape`, `ShapeCollection`, `ShapeScopedCollection`, `ShapeGroup`, and
// `PlaceholderFormat`. Positions and sizes are in points, as in Office.js; the host converts
// to EMU (12,700 per point).

import type {
  ShapeFormatPatch,
  ShapeFormatProperties,
  AutomationBox,
  AutomationHandle,
  AutomationItemKey,
  AutomationNewShape,
  AutomationOperation,
  AutomationShapeWrite,
  AutomationValue,
  AutomationZOrder,
} from '@pptx-editor.dev/core/automation';
import { ClientObject } from '../runtime/client-object.ts';
import { ClientResult } from '../runtime/client-result.ts';
import { diagnosticOf, fail, failArgument } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { Adjustments } from './adjustments.ts';
import { Chart, chartDataOf, chartTypeToFile, type ChartData, type ChartType } from './charts.ts';
import { ItemCollection } from './collection.ts';
import {
  CONNECTOR,
  PLACEHOLDER_TYPE,
  shapeTypeOf,
  Z_ORDER,
  type ConnectorType,
  type PlaceholderType,
  type ShapeType,
  type ShapeZOrder,
} from './enums.ts';
import { ShapeFill, ShapeLineFormat } from './fill-line.ts';
import { geometryOf, type GeometricShapeType } from './geometry.ts';
import { Hyperlink, type HyperlinkAddOptions } from './hyperlinks.ts';
import { imageBox, settleImage, type ShapeGetImageOptions } from './images.ts';
import { finite, flag, ModelObject, pathOf, text, type RefOf } from './model-object.ts';
import { SlideLayout, SlideMaster, Slide } from './slides.ts';
import { tablePlan, type TableAddOptions } from './table-options.ts';
import { Table } from './tables.ts';
import { CustomXmlPartCollection } from './custom-xml.ts';
import { TagCollection } from './tags.ts';
import { TextFrame } from './text.ts';

/**
 * Where a new shape goes, in points from the top left of the slide.
 *
 * @public
 */
export interface ShapeAddOptions {
  height?: number;
  left?: number;
  top?: number;
  width?: number;
}

const SHAPE_NAMES = [
  'id',
  'name',
  'type',
  'left',
  'top',
  'width',
  'height',
  'rotation',
  'visible',
  'zOrderPosition',
  'level',
  'altTextTitle',
  'altTextDescription',
  'isDecorative',
  'creationId',
] as const;

/** A box from `ShapeAddOptions`, checked. */
export function boxOf(options: ShapeAddOptions | undefined, target: string): AutomationBox {
  if (options === undefined) return {};
  if (typeof options !== 'object' || options === null)
    failArgument(target, 'a ShapeAddOptions object');
  const box: { left?: number; top?: number; width?: number; height?: number } = {};
  for (const key of ['left', 'top', 'width', 'height'] as const) {
    const value = options[key];
    if (value !== undefined) box[key] = finite(value, `${target}.${key}`);
  }
  return box;
}

/**
 * A shape on a slide, layout, or master: a geometric shape, a text box, a picture, a line, a
 * group, a table, or another graphic frame.
 *
 * @public
 */
export class Shape extends ModelObject {
  #fill: ShapeFill | undefined;
  #line: ShapeLineFormat | undefined;
  #textFrame: TextFrame | undefined;
  #tags: TagCollection | undefined;
  #customXml: CustomXmlPartCollection | undefined;
  #group: ShapeGroup | undefined;
  #placeholder: PlaceholderFormat | undefined;
  #parentGroup: Shape | undefined;
  #adjustments: Adjustments | undefined;
  #chart: Chart | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, nullable = false) {
    super(context, path, { nullable });
  }

  /** The shape id (`p:cNvPr/@id`), unique on its slide. */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The shape name (`p:cNvPr/@name`). */
  get name(): string {
    return this.loadedProperty('name');
  }
  set name(value: string) {
    this.#set('name', { name: text(value, this.at('name')) });
  }
  /** What kind of shape this is. */
  get type(): ShapeType | `${ShapeType}` {
    return this.loadedProperty('type');
  }
  /** The distance from the left edge of the slide, in points. */
  get left(): number {
    return this.loadedProperty('left');
  }
  set left(value: number) {
    this.#set('left', { left: finite(value, this.at('left')) });
  }
  /** The distance from the top edge of the slide, in points. */
  get top(): number {
    return this.loadedProperty('top');
  }
  set top(value: number) {
    this.#set('top', { top: finite(value, this.at('top')) });
  }
  /** The width, in points. */
  get width(): number {
    return this.loadedProperty('width');
  }
  set width(value: number) {
    this.#set('width', { width: finite(value, this.at('width')) });
  }
  /** The height, in points. */
  get height(): number {
    return this.loadedProperty('height');
  }
  set height(value: number) {
    this.#set('height', { height: finite(value, this.at('height')) });
  }
  /** The rotation, in degrees clockwise. */
  get rotation(): number {
    return this.loadedProperty('rotation');
  }
  set rotation(value: number) {
    this.#set('rotation', { rotation: finite(value, this.at('rotation')) });
  }
  /** Whether the shape shows (`p:cNvPr/@hidden` is not set). */
  get visible(): boolean {
    return this.loadedProperty('visible');
  }
  set visible(value: boolean) {
    this.#set('visible', { visible: flag(value, this.at('visible')) });
  }
  /** The zero-based position in the z-order of its parent. */
  get zOrderPosition(): number {
    return this.loadedProperty('zOrderPosition');
  }
  /** How deep in groups the shape is: 0 at the top level. */
  get level(): number {
    return this.loadedProperty('level');
  }
  /** The alternative text title (`p:cNvPr/@title`). */
  get altTextTitle(): string {
    return this.loadedProperty('altTextTitle');
  }
  set altTextTitle(value: string) {
    this.#set('altTextTitle', { altTextTitle: text(value, this.at('altTextTitle')) });
  }
  /** The alternative text description (`p:cNvPr/@descr`). */
  get altTextDescription(): string {
    return this.loadedProperty('altTextDescription');
  }
  set altTextDescription(value: string) {
    this.#set('altTextDescription', {
      altTextDescription: text(value, this.at('altTextDescription')),
    });
  }
  /** Whether the shape is decorative, so readers of alternative text skip it. */
  get isDecorative(): boolean {
    return this.loadedProperty('isDecorative');
  }
  set isDecorative(value: boolean) {
    this.#set('isDecorative', { isDecorative: flag(value, this.at('isDecorative')) });
  }
  /** The creation id of the shape. The file format has none, so this reads `null`. */
  get creationId(): string | null {
    return this.loadedProperty('creationId');
  }

  /** The adjust values of the shape's preset geometry (`a:prstGeom/a:avLst`). */
  get adjustments(): Adjustments {
    return (this.#adjustments ??= new Adjustments(
      this.context,
      ObjectPath.derived(`${this.path.label}.adjustments`, this.path),
      this.path
    ));
  }
  /** The fill of the shape. */
  get fill(): ShapeFill {
    return (this.#fill ??= new ShapeFill(
      this.context,
      ObjectPath.derived(`${this.path.label}.fill`, this.path),
      { kind: 'shape', shape: this.path }
    ));
  }
  /** The outline of the shape. */
  get lineFormat(): ShapeLineFormat {
    return (this.#line ??= new ShapeLineFormat(
      this.context,
      ObjectPath.derived(`${this.path.label}.lineFormat`, this.path),
      this.path
    ));
  }
  /** The text frame. `NotSupported` at the sync for a shape that holds no text. */
  get textFrame(): TextFrame {
    return (this.#textFrame ??= new TextFrame(
      this.context,
      ObjectPath.derived(`${this.path.label}.textFrame`, this.path),
      this.path
    ));
  }
  /** The tags of the shape. */
  get tags(): TagCollection {
    return (this.#tags ??= new TagCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.tags`, this.path)
    ));
  }
  /** The custom XML parts of the shape. */
  get customXmlParts(): CustomXmlPartCollection {
    return (this.#customXml ??= new CustomXmlPartCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.customXmlParts`, this.path)
    ));
  }
  /** The group this shape is, for a group shape. */
  get group(): ShapeGroup {
    return (this.#group ??= new ShapeGroup(
      this.context,
      ObjectPath.derived(`${this.path.label}.group`, this.path),
      this
    ));
  }
  /** The placeholder properties, for a placeholder. */
  get placeholderFormat(): PlaceholderFormat {
    return (this.#placeholder ??= new PlaceholderFormat(
      this.context,
      ObjectPath.derived(`${this.path.label}.placeholderFormat`, this.path),
      this.path
    ));
  }
  /**
   * The chart of a chart frame (`c:chart` in `a:graphicData`, ECMA-376 Part 1, §21.2.2.26), or a
   * null object for any other shape. Check `isNullObject` after a sync. A chart of an extension
   * namespace is no chart frame: its shape type reads `Unsupported`, and this is a null object.
   *
   * This property is an SDK extension. The pinned Office.js reference has no such member.
   *
   * @example
   * ```ts
   * const chart = shape.chart.load('chartType,categories,series');
   * await context.sync();
   * if (!chart.isNullObject) console.log(chart.chartType, chart.series.length);
   * ```
   */
  get chart(): Chart {
    return (this.#chart ??= this.#chartOf());
  }

  /** The chart object, with the read that tells a chart frame from another shape queued now. */
  #chartOf(): Chart {
    const path = ObjectPath.pending(`${this.path.label}.chart`);
    const chart = new Chart(this.context, path, this);
    // A shape that a sync found to be a null object has a null chart, as a getter answers one.
    if (this.path.isNull) {
      path.resolveNull();
      return chart;
    }
    const label = path.label;
    // The read only tells a null object: the chart's loads and writes name the shape itself, so
    // they wait for no read.
    this.enqueue(
      'read',
      'chart',
      (ref) => ({ op: 'readShape', shape: ref(this.path) }),
      (value) => {
        const read = value.kind === 'null' ? null : hydrated(value, 'shape', label).shape;
        if (read?.type === 'chart' && this.path.isAddressable) path.resolveTo(this.path.handle());
        else path.resolveNull();
      },
      path,
      false
    );
    return chart;
  }

  /** The group that holds this shape. `ItemNotFound` at the sync for a shape at the top level. */
  get parentGroup(): Shape {
    return this.#parentGroupOf(false);
  }

  /**
   * The parent group, made on first use. Through a collection's items (`orNull`), a shape at
   * the top level has a null parent group instead of refusing the sync.
   */
  #parentGroupOf(orNull: boolean): Shape {
    return (this.#parentGroup ??= this.navigation(
      'parentGroup',
      (path) => new Shape(this.context, path, orNull),
      (ref) => ({ op: 'getParentGroup', shape: ref(this.path), ...(orNull ? { orNull } : {}) }),
      orNull
    ));
  }

  /** Delete the shape. */
  delete(): void {
    this.command('delete', (ref) => ({ op: 'deleteShape', shape: ref(this.path) }));
  }

  /**
   * A PNG picture of the shape, without the slide behind it, as base64, drawn by the editor's
   * own renderer. A runtime over a live editor in a browser draws pictures, and so does a server
   * runtime created with a `rasterize` option; any other runtime refuses with `NotSupported`.
   *
   * @example
   * ```ts
   * const image = shape.getImageAsBase64({ width: 640 });
   * await context.sync();
   * const png = `data:image/png;base64,${image.value}`;
   * ```
   */
  getImageAsBase64(options?: ShapeGetImageOptions): ClientResult<string> {
    const target = this.at('getImageAsBase64');
    const box = imageBox(options, target, true);
    const { result, fill } = ClientResult.create<string>(target);
    this.enqueue(
      'read',
      'getImageAsBase64',
      (ref) => ({ op: 'prepareImage', target: ref(this.path), ...box }),
      settleImage(this.internals, box, target, fill)
    );
    return result;
  }

  /** Read detailed formatting. Read the result after `context.sync()`. */
  getFormat(): ClientResult<ShapeFormatProperties> {
    const target = this.at('getFormat');
    const { result, fill } = ClientResult.create<ShapeFormatProperties>(target);
    this.enqueue(
      'read',
      'getFormat',
      (ref) => ({ op: 'readShapeFormat', shape: ref(this.path) }),
      (value) => fill(hydrated(value, 'shapeFormat', target).format)
    );
    return result;
  }

  /** Patch existing formatting. Omitted fields retain their source XML. */
  patchFormat(patch: ShapeFormatPatch): void {
    this.command('patchFormat', (ref) => ({
      op: 'patchShapeFormat',
      shape: ref(this.path),
      patch,
    }));
  }

  /** The slide the shape is on; `ItemNotFound` at the sync for a layout or master shape. */
  getParentSlide(): Slide {
    return this.#parent('getParentSlide()', 'slide', false, (p) => new Slide(this.context, p));
  }
  /** The slide the shape is on, or a null object. */
  getParentSlideOrNullObject(): Slide {
    return this.#parent(
      'getParentSlideOrNullObject()',
      'slide',
      true,
      (p) => new Slide(this.context, p, true)
    );
  }
  /** The layout the shape is on; `ItemNotFound` at the sync for any other shape. */
  getParentSlideLayout(): SlideLayout {
    return this.#parent(
      'getParentSlideLayout()',
      'slideLayout',
      false,
      (p) => new SlideLayout(this.context, p)
    );
  }
  /** The layout the shape is on, or a null object. */
  getParentSlideLayoutOrNullObject(): SlideLayout {
    return this.#parent(
      'getParentSlideLayoutOrNullObject()',
      'slideLayout',
      true,
      (p) => new SlideLayout(this.context, p, true)
    );
  }
  /** The master the shape is on; `ItemNotFound` at the sync for any other shape. */
  getParentSlideMaster(): SlideMaster {
    return this.#parent(
      'getParentSlideMaster()',
      'slideMaster',
      false,
      (p) => new SlideMaster(this.context, p)
    );
  }
  /** The master the shape is on, or a null object. */
  getParentSlideMasterOrNullObject(): SlideMaster {
    return this.#parent(
      'getParentSlideMasterOrNullObject()',
      'slideMaster',
      true,
      (p) => new SlideMaster(this.context, p, true)
    );
  }

  /** The table of a table shape. `NotSupported` at the sync for any other shape. */
  getTable(): Table {
    return new Table(
      this.context,
      ObjectPath.derived(`${this.path.label}.getTable()`, this.path),
      this
    );
  }

  /** The text frame, or a null object for a shape that holds no text. */
  getTextFrameOrNullObject(): TextFrame {
    const path = ObjectPath.pending(`${this.path.label}.getTextFrameOrNullObject()`);
    const frame = new TextFrame(this.context, path, path, true);
    const label = path.label;
    // The read answers the shape, not a handle: a call on the frame waits for this read, and by
    // then the shape's own path has resolved.
    this.enqueue(
      'read',
      'getTextFrameOrNullObject()',
      (ref) => ({ op: 'readShape', shape: ref(this.path) }),
      (value) => {
        const read = value.kind === 'null' ? null : hydrated(value, 'shape', label).shape;
        if (read?.hasTextFrame && this.path.isAddressable) path.resolveTo(this.path.handle());
        else path.resolveNull();
      },
      path,
      false
    );
    return frame;
  }

  /** Set the click link of the whole shape. */
  setHyperlink(options?: HyperlinkAddOptions): Hyperlink {
    const address = options?.address ?? '';
    const screenTip = options?.screenTip;
    text(address, this.at('address'));
    if (screenTip !== undefined) text(screenTip, this.at('screenTip'));
    this.command('setHyperlink', (ref) => ({
      op: 'setShapeHyperlink',
      shape: ref(this.path),
      address,
      ...(screenTip === undefined ? {} : { screenTip }),
    }));
    return Hyperlink.created(this.context, this.path, null, address, screenTip ?? '');
  }

  /** Move the shape in the z-order of its parent. */
  setZOrder(position: ShapeZOrder): void;
  /** Move the shape in the z-order of its parent. */
  setZOrder(position: 'BringForward' | 'BringToFront' | 'SendBackward' | 'SendToBack'): void;
  setZOrder(position: ShapeZOrder | `${ShapeZOrder}`): void {
    const move = Z_ORDER.toFile(position, `${this.path.label}.setZOrder`) as AutomationZOrder;
    this.command('setZOrder', (ref) => ({
      op: 'setZOrder',
      shape: ref(this.path),
      position: move,
    }));
  }

  #parent<T>(
    name: string,
    kind: 'slide' | 'slideLayout' | 'slideMaster',
    orNull: boolean,
    make: (path: ObjectPath) => T
  ): T {
    return this.navigation(
      name,
      make,
      (ref) => ({ op: 'parentOf', shape: ref(this.path), kind, ...(orNull ? { orNull } : {}) }),
      orNull
    );
  }

  #set(name: string, props: AutomationShapeWrite): void {
    this.command(name, (ref) => ({ op: 'setShape', shape: ref(this.path), props }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, SHAPE_NAMES, {
      adjustments: () => this.adjustments,
      fill: () => this.fill,
      lineFormat: () => this.lineFormat,
      textFrame: () => this.textFrame,
      tags: () => this.tags,
      customXmlParts: () => this.customXmlParts,
      placeholderFormat: () => this.placeholderFormat,
      group: () => this.group,
      chart: () => this.chart,
      parentGroup: () => this.#parentGroupOf(this.loadingOrNull),
    });
    this.loadScalars(
      names,
      'shape',
      (ref) => ({ op: 'readShape', shape: ref(this.path) }),
      (v) => ({ ...v.shape, type: shapeTypeOf(v.shape.type), creationId: null })
    );
  }
}

/** The handles a value of shape handles carries. */
function shapeHandles(value: AutomationValue, label: string): readonly AutomationHandle[] {
  return hydrated(value, 'handles', label).handles;
}

/** What a shape collection lists: a container's shapes, or the reader's selected shapes. */
export type ShapeSource =
  { readonly kind: 'container'; readonly container: ObjectPath } | { readonly kind: 'selection' };

/** The members both shape collections share. */
abstract class ShapeList extends ItemCollection<Shape> {
  /** @internal */
  protected abstract readonly source: ShapeSource;

  /** A shape by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Shape {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }
  /** A shape by zero-based position in the z-order. */
  getItemAt(index: number): Shape {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }
  /** A shape by id, or a null object. */
  getItemOrNullObject(id: string): Shape {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(ref: RefOf): AutomationOperation {
    return this.source.kind === 'container'
      ? { op: 'listShapes', container: ref(this.source.container) }
      : { op: 'readSelection' };
  }
  /** @internal */
  protected override handlesOf(value: AutomationValue, label: string): readonly AutomationHandle[] {
    return this.source.kind === 'container'
      ? shapeHandles(value, label)
      : hydrated(value, 'selection', label).selection.shapes;
  }
  /** @internal */
  protected override get hasLookup(): boolean {
    return this.source.kind === 'container';
  }
  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): Shape {
    return new Shape(this.context, path, nullable);
  }
  /** @internal */
  protected lookup(
    ref: RefOf,
    key: AutomationItemKey,
    orNull: boolean
  ): AutomationOperation | undefined {
    if (this.source.kind !== 'container') return undefined;
    return {
      op: 'getShape',
      container: ref(this.source.container),
      key,
      ...(orNull ? { orNull: true } : {}),
    };
  }

  /** @internal Queue the group of `shapes` and answer the new group shape. */
  protected groupOf(shapes: readonly (() => ObjectPath[])[], name: string): Shape {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const group = new Shape(this.context, path);
    const label = path.label;
    this.enqueue(
      'write',
      name,
      (ref) => ({ op: 'group', shapes: shapes.flatMap((list) => list()).map((p) => ref(p)) }),
      (value) => path.resolveTo(hydrated(value, 'handle', label).handle),
      path
    );
    return group;
  }
}

/**
 * The shapes of a slide, layout, master, or group, in z-order.
 *
 * @public
 */
export class ShapeCollection extends ShapeList {
  /** @internal */
  protected readonly source: ShapeSource;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, container: ObjectPath) {
    super(context, path);
    this.source = { kind: 'container', container };
  }

  /** Add a preset geometric shape (`a:prstGeom`, ECMA-376 Part 1, 20.1.9.18). */
  addGeometricShape(geometricShapeType: GeometricShapeType, options?: ShapeAddOptions): Shape;
  /** Add a preset geometric shape (`a:prstGeom`, ECMA-376 Part 1, 20.1.9.18). */
  addGeometricShape(geometricShapeType: `${GeometricShapeType}`, options?: ShapeAddOptions): Shape;
  addGeometricShape(geometricShapeType: GeometricShapeType, options?: ShapeAddOptions): Shape {
    const geometry = geometryOf(geometricShapeType, `${this.path.label}.addGeometricShape`);
    const box = boxOf(options, `${this.path.label}.addGeometricShape`);
    return this.#add('addGeometricShape', { kind: 'geometric', geometry }, box);
  }

  /** Add a line or connector (`p:cxnSp`). */
  addLine(connectorType?: ConnectorType, options?: ShapeAddOptions): Shape;
  /** Add a line or connector (`p:cxnSp`). */
  addLine(connectorType?: 'Straight' | 'Elbow' | 'Curve', options?: ShapeAddOptions): Shape;
  addLine(
    connectorType: ConnectorType | `${ConnectorType}` = 'Straight',
    options?: ShapeAddOptions
  ): Shape {
    const connector = CONNECTOR.toFile(connectorType, `${this.path.label}.addLine`);
    const box = boxOf(options, `${this.path.label}.addLine`);
    return this.#add(
      'addLine',
      { kind: 'line', connector: connector as 'straight' | 'elbow' | 'curve' },
      box
    );
  }

  /** Add a text box with `text`. */
  addTextBox(text_: string, options?: ShapeAddOptions): Shape {
    const box = boxOf(options, `${this.path.label}.addTextBox`);
    return this.#add('addTextBox', { kind: 'textBox', text: text(text_, this.at('text')) }, box);
  }

  /** Add a table (`a:tbl` in a graphic frame) of `rowCount` rows and `columnCount` columns. */
  addTable(rowCount: number, columnCount: number, options?: TableAddOptions): Shape {
    const target = `${this.path.label}.addTable`;
    for (const value of [rowCount, columnCount])
      if (typeof value !== 'number' || !Number.isInteger(value) || value < 1)
        failArgument(target, 'a whole number of rows and of columns of 1 or more');
    const plan = tablePlan(options, rowCount, columnCount, target);
    const { left, top, width, height } = options ?? {};
    const box = boxOf({ left, top, width, height }, target);
    const values = plan.values;
    const shape = this.#add(
      'addTable',
      { kind: 'table', rows: rowCount, columns: columnCount, ...(values ? { values } : {}) },
      box
    );
    plan.apply(shape.getTable());
    return shape;
  }

  /**
   * Add a chart (`p:graphicFrame`, `CT_GraphicalObjectFrame`, with a new `c:chartSpace` part,
   * `CT_ChartSpace`, ECMA-376 Part 1, §19.3.1.21 and §21.2.2.29) of `chartType` over `data`, and
   * answer the new chart frame. The chart part caches the data, and an embedded workbook holds
   * the same table (`c:externalData`, `CT_ExternalData`, §21.2.2.63), so `chart.setData` and a
   * later edit of the workbook start from the same values.
   *
   * The chart has the editor's default look. It has no title and a legend at the bottom. Its
   * series take the theme colors in order. A `Pie` or `Doughnut` chart colors each category
   * (`c:varyColors`, §21.2.2.227).
   *
   * Positions and sizes are in points. Without a box, the chart is three fifths of the slide
   * wide. It is centered below the top fifth of the slide. A graphic frame or picture can cover
   * more than half of that place. Then the chart moves down and right by 5% of the slide. Each
   * new chart at that place moves one more step, but not off the slide. A box field that
   * `options` does not give keeps its default value. Thus, give `left` and `top` with a new
   * `width` or `height`.
   *
   * `data` takes the form of `Chart.setData`. Give at least one category and one series. Give at
   * most 4000 categories and 255 series, and one value for each category. Labels and names have
   * at most 255 characters. The chart XML must also fit in 16 MiB. That limit depends on the
   * length of the labels and values. The categories of an `XYScatter` or `XYScatterLines` chart
   * are its X values, so each is a number as text. The chart keeps them as numbers, so `'1e3'`
   * reads back as `'1000'`. Another chart type, data of another form, or a box out of range
   * refuses with `InvalidArgument`. Then the call adds nothing.
   *
   * This method is an SDK extension. The pinned Office.js reference has no such member.
   *
   * @example
   * ```ts
   * const shape = slide.shapes.addChart(
   *   'ColumnClustered',
   *   {
   *     categories: ['Q1', 'Q2', 'Q3', 'Q4'],
   *     series: [{ name: 'Revenue', values: [12, 15, 14, 18] }],
   *   },
   *   { left: 60, top: 100, width: 480, height: 300 }
   * );
   * await context.sync();
   * ```
   */
  addChart(
    chartType: ChartType | `${ChartType}`,
    data: ChartData,
    options?: ShapeAddOptions
  ): Shape {
    const target = `${this.path.label}.addChart`;
    const type = chartTypeToFile(chartType, `${target}.chartType`);
    const checked = chartDataOf(data, `${target}.data`);
    const box = boxOf(options, target);
    return this.#add('addChart', { kind: 'chart', chartType: type, data: checked }, box);
  }

  /** Group shapes of this collection, by shape or by id. Answers the new group. */
  addGroup(values: Array<string | Shape>): Shape {
    if (!Array.isArray(values))
      failArgument(`${this.path.label}.addGroup`, 'an array of shapes or shape ids');
    if (values.length < 2)
      fail({
        code: 'InvalidArgument',
        target: `${this.path.label}.addGroup`,
        diagnostic: diagnosticOf('group-count')!,
      });
    const shapes = values.map((v) => (typeof v === 'string' ? this.getItem(v) : v));
    if (shapes.some((s) => !(s instanceof ClientObject)))
      failArgument(`${this.path.label}.addGroup`, 'an array of shapes or shape ids');
    return this.groupOf([() => shapes.map((s) => pathOf(s))], 'addGroup');
  }

  /**
   * Add a picture shape (`p:pic`, ECMA-376 Part 1, §19.3.1.37) and return it. This is an SDK
   * extension: the pinned Office.js reference has no member that adds a picture.
   *
   * @param base64EncodedImage - The Base64 bytes of a PNG, JPEG, or GIF image of at most 50 MiB.
   * The bytes go into a new image part of the package. Other bytes, SVG among them, refuse with
   * `InvalidArgument` at the sync, and the call writes nothing. Send the bytes alone, with no
   * `data:` URL prefix.
   * @param options - The box in points. With no size, the picture keeps the aspect ratio of the
   * image and fits 60% of the slide, centered. It is never larger than its own size. When the
   * image states no resolution, the engine uses 96 pixels per inch. With only `width` or only
   * `height`, the other keeps the aspect ratio. When you give a size, the engine centers each
   * missing `left` or `top` on the slide.
   *
   * @remarks
   * `Shape.type` reads `Image`. The new shape has no alt text: set `altTextDescription` in the
   * same batch. The picture never fills an empty placeholder. The call and the writes queued
   * with it apply as one undo step at the sync.
   *
   * @example
   * ```ts
   * const logo = slide.shapes.addPicture(logoBase64, { left: 24, top: 24, width: 120 });
   * logo.altTextDescription = 'Company logo';
   * await context.sync();
   * ```
   */
  addPicture(base64EncodedImage: string, options?: ShapeAddOptions): Shape {
    const target = `${this.path.label}.addPicture`;
    const base64 = text(base64EncodedImage, target);
    return this.#add('addPicture', { kind: 'picture', base64 }, boxOf(options, target));
  }

  #add(name: string, shape: AutomationNewShape, box: AutomationBox): Shape {
    const path = ObjectPath.pending(`${this.path.label}.${name}()`);
    const created = new Shape(this.context, path);
    const label = path.label;
    const container = this.source.kind === 'container' ? this.source.container : this.path;
    this.enqueue(
      'write',
      name,
      (ref) => ({ op: 'addShape', container: ref(container), shape, box }),
      (value) => path.resolveTo(hydrated(value, 'handle', label).handle),
      path
    );
    return created;
  }
}

/**
 * A set of shapes: the reader's selection, or the members of a group.
 *
 * @public
 */
export class ShapeScopedCollection extends ShapeList {
  /** @internal */
  protected readonly source: ShapeSource;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, source: ShapeSource) {
    super(context, path);
    this.source = source;
  }

  /**
   * Group the shapes of this collection and answer the new group. The items need no load: when
   * a completed load has not filled them in, the sync reads them first, in an earlier round at
   * the same revision, and then groups them.
   */
  group(): Shape {
    if (this.hasLoadedProperty('items')) {
      const items = this.items;
      return this.groupOf([() => items.map((s) => pathOf(s))], 'group()');
    }
    // The listing read produces a path that no call can name by result reference, so the group
    // waits for it (`WaitForRead`) and is planned again with the members it answered.
    const listed = ObjectPath.pending(`${this.path.label}.items`);
    let members: readonly Shape[] | undefined;
    this.enqueue(
      'read',
      'items',
      (ref) => this.listing(ref),
      (value) => {
        members = this.handlesOf(value, listed.label).map((handle, i) =>
          this.itemOf(this.memberPath(handle, i), false)
        );
      },
      listed,
      false
    );
    return this.groupOf([() => (members ? members.map((s) => pathOf(s)) : [listed])], 'group()');
  }
}

/**
 * A group shape (`p:grpSp`).
 *
 * @public
 */
export class ShapeGroup extends ModelObject {
  readonly #shape: Shape;
  #shapes: ShapeScopedCollection | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: Shape) {
    super(context, path);
    this.#shape = shape;
  }

  /** The id of the group shape. */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The creation id. The file format has none, so this reads `null`. */
  get creationId(): string | null {
    return this.loadedProperty('creationId');
  }
  /** The group as a shape. */
  get shape(): Shape {
    return this.#shape;
  }
  /** The members of the group, in z-order. */
  get shapes(): ShapeScopedCollection {
    return (this.#shapes ??= new ShapeScopedCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.shapes`, this.path),
      { kind: 'container', container: this.path }
    ));
  }

  /** Ungroup: the members move to the group's parent, and the group is deleted. */
  ungroup(): void {
    // The host answers the former members; the object model's `ungroup` answers nothing.
    this.enqueue('write', 'ungroup', (ref) => ({ op: 'ungroup', group: ref(this.path) }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'creationId'], { shapes: () => this.shapes });
    this.loadScalars(
      names,
      'shape',
      (ref) => ({ op: 'readShape', shape: ref(this.path), group: true }),
      (v) => ({ id: v.shape.id, creationId: null })
    );
  }
}

/**
 * The placeholder properties of a shape (`p:ph`, ECMA-376 Part 1, 19.3.1.36).
 *
 * @public
 */
export class PlaceholderFormat extends ModelObject {
  readonly #shape: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: ObjectPath) {
    super(context, path);
    this.#shape = shape;
  }

  /** The placeholder type. A shape that is no placeholder reads `Unsupported`. */
  get type(): PlaceholderType | `${PlaceholderType}` {
    return this.loadedProperty('type');
  }
  /** The kind of content the placeholder holds, or null for a shape placeholder. */
  get containedType(): ShapeType | `${ShapeType}` | null {
    return this.loadedProperty('containedType');
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['type', 'containedType']);
    this.loadScalars(
      names,
      'shape',
      (ref) => ({ op: 'readShape', shape: ref(this.#shape) }),
      (v) => ({
        type: PLACEHOLDER_TYPE.fromFile(v.shape.placeholder?.type) ?? 'Unsupported',
        containedType: v.shape.placeholder?.containedType
          ? shapeTypeOf(v.shape.placeholder.containedType)
          : null,
      })
    );
  }
}
