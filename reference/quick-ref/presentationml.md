# PresentationML quick reference

Extracted from ECMA-376 5th edition, Part 1 (§19, PresentationML) and Part 2 (OPC). Schemas: `reference/ecma-376/part4/transitional-schemas/pml.xsd`. Where the reference PDFs differ from the text of the standard, the row says so.

Prefixes: `p:` PresentationML main, `a:` DrawingML main, `r:` relationships, `mc:` markup compatibility, `p14:`/`p15:` the 2010 and 2013 PresentationML extensions.

## Package structure

```
[Content_Types].xml
_rels/.rels                         → ppt/presentation.xml (officeDocument)
docProps/app.xml, core.xml, thumbnail.jpeg
ppt/
├── presentation.xml                p:presentation
├── _rels/presentation.xml.rels     → masters, slides, theme, presProps, viewProps, tableStyles
├── slides/slideN.xml               p:sld        (rels → layout, notesSlide, media, charts)
├── slideLayouts/slideLayoutN.xml   p:sldLayout  (rels → master)
├── slideMasters/slideMasterN.xml   p:sldMaster  (rels → layouts, theme)
├── notesSlides/notesSlideN.xml     p:notes      (rels → slide, notesMaster)
├── notesMasters/notesMaster1.xml   p:notesMaster
├── handoutMasters/                 p:handoutMaster
├── theme/themeN.xml                a:theme (one per master, plus notes and handout masters)
├── presProps.xml, viewProps.xml, tableStyles.xml
├── commentAuthors.xml, comments/commentN.xml   (legacy comments)
├── charts/, diagrams/, embeddings/, media/
└── fonts/fontN.fntdata             embedded fonts (obfuscated EOT/MTX)
```

Slide order is the order of `p:sldIdLst`, never the part name. Slide `id` values are ≥ 256 and unique. Master `id` and layout `id` values share one space and are ≥ 2147483648.

Every part has a content type: an `Override` for its name, or a `Default` for its extension (ECMA-376 Part 2). Every internal relationship target is a part of the package; an external target (`TargetMode="External"`) is a URI and is never fetched. The store refuses an edit that breaks one of these rules: a touched part that loses its type or targets a missing part, or a removed part that a relationship still targets (`store/store/part-checks.ts`).

## `p:presentation`

| Element | Purpose |
| --- | --- |
| `p:sldMasterIdLst/p:sldMasterId` | Masters, by `r:id` |
| `p:notesMasterIdLst`, `p:handoutMasterIdLst` | Notes and handout masters |
| `p:sldIdLst/p:sldId` | Slides in order: `id` and `r:id` |
| `p:sldSz` | Slide size: `cx`, `cy` in EMU, `type` (`screen16x9`, `custom`, …) |
| `p:notesSz` | Notes page size |
| `p:embeddedFontLst/p:embeddedFont` | Embedded fonts: `p:font typeface`, `p:regular`/`p:bold`/`p:italic`/`p:boldItalic` by `r:id` |
| `p:defaultTextStyle` | Lowest text style level (`a:defPPr`, `a:lvl1pPr` … `a:lvl9pPr`) |
| `p:extLst` → `p14:sectionLst` | Sections: `p14:section name id` with `p14:sldIdLst` |

Attributes: `saveSubsetFonts`, `embedTrueTypeFonts`, `firstSlideNum` (default 1), `rtl`.

### Slide list and sections

- `p:sldId/@id` is `ST_SlideId`: an integer from 256 to 2147483647 (ECMA-376 Part 1, 19.2.1.33). Each `id` is unique in the list, and each `r:id` names a slide part through a relationship of the presentation part.
- `CT_Presentation` has at most one `p:sldIdLst` (19.2.1.26). It is optional when the deck has no slides.
- `p14:sectionLst` (the `p14` extension schema, `CT_SectionList`) sits in a `p:ext` with `uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"` in the one `p:extLst`. Each `p14:section` has `name`, a GUID `id`, and a required `p14:sldIdLst` as its first child, which can be empty.
- The sections name every slide of `p:sldIdLst` once, in slide order, and only those slides. So each section holds one run of consecutive slides, and the section order follows the slide order.

The writer repairs a part that breaks these rules (`presentation-repair.ts`), so the file opens without a repair prompt:

- Two `p:sldIdLst` lists join into the first, in document order.
- An entry whose `r:id` names no part goes. Of entries that name one slide part, only the first stays.
- An entry that repeats the `id` of another slide part, or whose `id` is missing, not an integer, or outside 256 to 2147483647, gets a fresh `id`: one past the largest `id` that the part names.
- Two section lists or two `p:extLst` elements join into the first. A section id that repeats goes.
- A section without `p14:sldIdLst` gets an empty one. The sections move into slide order, a slide that two sections name stays in the first, and a slide that no section names joins the section of the slide before it.

`p:sldSz` need not be whole points: the ISO A4 paper size as a custom size is 10693400 × 7562850 EMU, 842 × 595.5 pt. Keep the exact size. In the reference PDFs, each page side is a whole number of points. Measured: a fraction of 0.375 pt or less rounds down, and 0.45 pt and every fraction from 0.5 pt round up (595.5 → 596, 400.25 → 400, 890.5 → 891, 400.45 → 401), on the fixtures `slide-size` and `slide-size-rounding`, a test deck, and corpus PDFs. So a page side is less than 0.625 pt off. Each page holds the slide at scale 1 from the page's top-left corner. Compare against the slide area of the page, never the whole page. When a side rounds down, the page ends before the slide, and that strip has no reference.

## Slide, layout, and master

```xml
<p:sld show="0">                          <!-- show="0": hidden slide -->
  <p:cSld name="…">
    <p:bg>…</p:bg>                         <!-- optional -->
    <p:spTree>
      <p:nvGrpSpPr>…</p:nvGrpSpPr>
      <p:grpSpPr>…</p:grpSpPr>
      <!-- p:sp | p:grpSp | p:graphicFrame | p:cxnSp | p:pic | p:contentPart | mc:AlternateContent -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:transition>…</p:transition>
  <p:timing>…</p:timing>
</p:sld>
```

| Part | Root | Specific content |
| --- | --- | --- |
| Slide | `p:sld` | `show`, `showMasterSp`, `p:clrMapOvr` |
| Layout | `p:sldLayout` | `type` (`title`, `obj`, `twoObj`, `blank`, …), `matchingName`, `showMasterSp`, `preserve`, `userDrawn`, `p:clrMapOvr` |
| Master | `p:sldMaster` | `p:clrMap` (required), `p:sldLayoutIdLst`, `p:txStyles`, `preserve` |
| Notes slide | `p:notes` | Slide image placeholder (`type="sldImg"`) and body placeholder |

`showMasterSp="0"` on a slide or layout hides the non-placeholder shapes of the level above.

### Background (`p:bg`)

| Form | Meaning |
| --- | --- |
| `p:bgPr` | Explicit fill (`a:solidFill`, `a:gradFill`, `a:blipFill`, `a:pattFill`, `a:noFill`) and optional effects |
| `p:bgRef idx="1001"` | Theme background fill: `idx` ≥ 1001 indexes `a:bgFillStyleLst` (1001 = first). A color child fills `phClr` |

The first `p:bg` found in slide → layout → master wins.

### Color map

`p:clrMap` on the master maps the 12 logical names to theme slots: `bg1 tx1 bg2 tx2 accent1 … accent6 hlink folHlink`. Typical: `bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"`. `a:schemeClr val="tx1"` goes through the map; `val="dk1"` reads the theme directly. A slide or layout `p:clrMapOvr/a:overrideClrMapping` replaces the map; `a:masterClrMapping` keeps it.

## Shapes in the shape tree

| Element | What | Non-visual properties |
| --- | --- | --- |
| `p:sp` | Shape or text box | `p:nvSpPr` → `p:cNvPr`, `p:cNvSpPr txBox="1"`, `p:nvPr` |
| `p:pic` | Picture | `p:nvPicPr`; fill in `p:blipFill` |
| `p:grpSp` | Group | `p:nvGrpSpPr`; `p:grpSpPr/a:xfrm` with `a:chOff`/`a:chExt` |
| `p:cxnSp` | Connector | `p:nvCxnSpPr` → `a:stCxn`/`a:endCxn` (`id`, `idx` of the site) |
| `p:graphicFrame` | Table, chart, diagram, OLE | `p:xfrm` (not `a:xfrm`), `a:graphic/a:graphicData uri="…"` |

`p:cNvPr`: `id` (`ST_DrawingElementId`, an `xsd:unsignedInt` from 0 to 4294967295; unique among the shape-tree members of one slide, layout, or master; the `mc:Fallback` of an `mc:AlternateContent` repeats the ids of its `mc:Choice`, and a picture inside `p:oleObj` is not a member), `name`, `descr` (alt text), `hidden`, `a:hlinkClick`. Hide and rename (`setShapesVisible`, `renameShape`) set `hidden="1"` and `name`; a hidden element does not draw, in the editor and in a slide show.

New ids that the editor writes (`p:cNvPr/@id`, `p:cm/@idx`) are one past the largest id in use and at most 2147483647, the `ST_SlideId` maximum, so a reader that stores ids in a signed 32-bit integer reads them. Past that bound the editor takes the lowest free id. A file value counts as an id when it is an `xsd:unsignedInt` in lexical form: digits with optional leading zeros, an optional `+`, and surrounding whitespace (`whiteSpace="collapse"`), so `" 7"`, `"+7"`, and `"007"` are all id 7. A value such as `1e3`, `-1`, or `0x10` is no id and sets no maximum.

The locks element is the first child of each `p:cNv…Pr`: `a:spLocks` (`p:cNvSpPr`), `a:picLocks` (`p:cNvPicPr`), `a:grpSpLocks` (`p:cNvGrpSpPr`), `a:cxnSpLocks` (`p:cNvCxnSpPr`), `a:graphicFrameLocks` (`p:cNvGraphicFramePr`). `noChangeAspect="1"` is Lock aspect ratio (`setLockAspectRatio`).

Child order in `p:sp`: `p:nvSpPr`, `p:spPr`, `p:style`, `p:txBody`, `p:extLst`. Writers must keep schema order.

`a:graphicData uri` values:

| URI ends with | Content |
| --- | --- |
| `/drawingml/2006/table` | `a:tbl` inline |
| `/drawingml/2006/chart` | `c:chart r:id` → chart part |
| `/drawingml/2006/diagram` | `dgm:relIds` → diagram data, layout, style, colors, drawing |
| `/presentationml/2006/ole` | `p:oleObj` (often inside `mc:AlternateContent`) |

## Placeholders

`p:nvPr/p:ph` attributes: `type` (default `obj`), `idx` (default 0), `orient`, `sz`, `hasCustomPrompt`.

| `type` | Use | Master text style |
| --- | --- | --- |
| `title`, `ctrTitle` | Title, centered title | `p:titleStyle` |
| `body`, `subTitle`, `obj` | Body, subtitle, content | `p:bodyStyle` |
| `dt`, `ftr`, `sldNum`, `hdr` | Date, footer, slide number, header | `p:otherStyle` |
| `chart`, `tbl`, `clipArt`, `dgm`, `media`, `pic` | Content placeholders for one kind of object | Not stated by the spec. The engine uses `p:otherStyle` (approximation `content-placeholders`, listed in `docs/site/content/pptx-fidelity.mdx`) |
| `sldImg` | Slide image on notes pages | none |

Matching (`packages/core/src/cascade/slide.ts`):

1. Slide → layout: by `idx` when `idx` is present and not 0; otherwise by `type`; otherwise by the master type family.
2. Layout → master: by type family. `ctrTitle` maps to `title`; `subTitle` and `obj` map to `body`.
3. A slide placeholder inherits `p:spPr` (`a:xfrm`, geometry, fill, line), `a:bodyPr`, and `a:lstStyle` from the matched layout placeholder, and that one from the master.
4. A shape without `p:ph` inherits nothing from the layout or master. Its text starts from `p:defaultTextStyle`.

## Text style cascade

Lowest precedence first. For each property, the last level that states it wins. Run defaults come from each level's `a:defRPr`.

```
p:presentation/p:defaultTextStyle
→ p:sldMaster/p:txStyles (titleStyle | bodyStyle | otherStyle, by placeholder type)
→ p:style/a:fontRef (font and color only)
→ master placeholder a:lstStyle → layout placeholder a:lstStyle
→ shape a:lstStyle
→ a:pPr (paragraph) → a:rPr (run)
```

Paragraph level: `a:pPr lvl` (0–8) selects `a:lvl{lvl+1}pPr`. The end-of-paragraph run style is `a:endParaRPr`.

## Notes

The notes text is the body placeholder (`type="body"`) of `p:notes`. The slide links to its notes slide; the notes slide links back to the slide and to the notes master.

## Comments

| Era | Parts |
| --- | --- |
| Legacy | `ppt/commentAuthors.xml` (`p:cmAuthor id name initials`), `ppt/comments/commentN.xml` (`p:cm authorId dt idx`, `p:pos`, `p:text`) |
| Modern | `ppt/authors.xml` and `ppt/comments/modernCommentN.xml` (`p188:` namespace), threaded |

Legacy `p:cm/@idx` is `ST_Index` and `p:cmAuthor/@lastIdx` is `xsd:unsignedInt` (`pml.xsd`). An index is unique per author in the presentation, and `lastIdx` is the author's last index.

## Customer data, custom XML, and custom properties

- `p:custDataLst` (`CT_CustomerDataList`, 19.3.1.18) holds `p:custData r:id` entries (19.3.1.17), then at most one `p:tags r:id`. It sits in `p:cSld` after `p:spTree` (slides, layouts, masters), in `p:nvPr` before `p:extLst` (shapes), and in `p:presentation` before `p:kinsoku`.
- A `p:custData` relationship has the type `customXml` and targets a data part, `/customXml/itemN.xml` (ECMA-376 Part 1, 22.5). The data part relates (`customXmlProps`) to `itemPropsN.xml`, a `ds:datastoreItem` with a GUID `ds:itemID` and `ds:schemaRefs`. The properties part needs an `Override` (`application/vnd.openxmlformats-officedocument.customXmlProperties+xml`); the data part uses the `xml` default.
- The presentation holds its own custom XML parts through `customXml` relationships of the presentation part, with no `p:custData` entry.
- A copy of a data part needs a new `ds:itemID`: two properties parts with one id are one data store to a reader.
- Custom document properties are `/docProps/custom.xml` (22.3), related from the package with the `custom-properties` type. Each `property` has `fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"`, a unique `pid` from 2 up, a `name`, and one `vt:` variant child (22.4): `vt:lpwstr`, `vt:i4`, `vt:r8`, `vt:bool`, `vt:filetime`, and others.

## Transitions and timing

`p:transition` (`spd`, `advClick`, `advTm`, one effect child) and `p:timing` (the animation time tree: `p:tnLst/p:par/p:cTn`, `p:bldLst`). Newer effects sit in `mc:AlternateContent` with a `p14:`/`p15:` Choice and a plain Fallback. The engine preserves both and plays neither.

`p:timing` names shapes by `p:cNvPr/@id` in `spid`: `p:spTgt` and its `p:subSp`, `p:inkTgt`, and the build entries `p:bldP`, `p:bldDgm`, `p:bldOleChart`, and `p:bldGraphic` (19.5). A shape trigger is a `p:cond` with `p:tgtEl/p:spTgt`; `p:tn` names another time node by `p:cTn/@id`. Rules that a removal must keep:

- `CT_TimeNodeList` (`p:tnLst`, `p:childTnLst`, `p:subTnLst`) and `CT_BuildList` (`p:bldLst`) need at least one child; `CT_TLTimeConditionList` needs one `p:cond`.
- `p:childTnLst` and `p:subTnLst` are optional in `CT_TLCommonTimeNodeData`, and `p:prevCondLst` and `p:nextCondLst` are optional in `CT_TLTimeNodeSequence`.
- A behavior (`p:cBhvr`) and a media node (`p:cMediaNode`) need exactly one `p:tgtEl` with one target, so a target cannot go without its behavior.
- A build entry (`AG_TLBuild`: `spid` and `grpId`, both required) pairs with the effects whose `p:cTn/@grpId` is its `grpId` and that target its `spid`.
- `CT_SlideTiming` children are all optional, in the order `p:tnLst`, `p:bldLst`, `p:extLst`.

Connector ends `a:stCxn` and `a:endCxn` (`CT_Connection`: `id`, `idx`) also name a `p:cNvPr/@id`. Both are optional in `CT_NonVisualConnectorProperties`; a connector without them is not glued.

`packages/core/src/store/package/timing.ts` applies the timing rules when shapes go, and `packages/core/src/automation/commands/shape-removal.ts` removes the connector ends.

## Units

| Unit | Value | Used for |
| --- | --- | --- |
| EMU | 914400 per inch, 12700 per point, 360000 per cm | Positions and sizes (`a:off`, `a:ext`, `p:sldSz`), line widths, insets |
| Hundredths of a point | 1200 = 12 pt | `a:rPr sz`, `a:spcPts val` |
| Thousandths of a percent | 100000 = 100 % | `a:spcPct`, `a:lumMod`, `fontScale`, alpha |
| 60000ths of a degree | 5400000 = 90° | `rot`, gradient `ang` |

Keep EMU in the model and in layout. Convert at property-read boundaries only.

## Common pitfalls

- Missing `p:ph type` means `obj`, not "no type". Missing `idx` means 0.
- `p:graphicFrame` uses `p:xfrm`, and a group uses `a:xfrm` with `a:chOff`/`a:chExt`. Child coordinates are scaled by `ext/chExt`.
- `a:schemeClr val="phClr"` takes the color passed by the referencing `p:style` or `p:bgRef`.
- `mc:AlternateContent` (ECMA-376 Part 3, 10.2.1): one or more `mc:Choice` (each with a required `Requires`, a space-separated list of prefixes), then an optional `mc:Fallback`. Render the first `mc:Choice` whose `Requires` prefixes all resolve, in the Choice's namespace scope, to namespaces you support; otherwise `mc:Fallback`. With neither, the spec ignores the element; the engine draws the first `mc:Choice`, so its shapes stay editable. The engine supports `p14` for a Choice whose children are all `p:sp`, `p:pic`, `p:grpSp`, `p:cxnSp`, or `p:graphicFrame`; a `p14:contentPart` (ink), `v` (VML), or `a14` Choice falls back. The fixture `preserve-unsupported` shows the `Requires="p14"` Choice drawn in the reference PDF. Save every branch. An edit of the drawn copy leaves the copies in the other branches out of date: hide, rename, and the aspect lock change every copy; a delete removes every copy; when a `mc:Choice` is drawn, other edits remove the copies in the other branches (`automation/commands/alternate-copies.ts`).
- A layout or master shape with `userDrawn="1"` is content, not a placeholder template.
