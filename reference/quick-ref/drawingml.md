# DrawingML quick reference

Extracted from ECMA-376 5th edition, Part 1 (§20.1, DrawingML main). Schema: `reference/ecma-376/part4/transitional-schemas/dml-main.xsd`. Preset geometry: `reference/ecma-376/part1/drawingml-geometries/presetShapeDefinitions.xml`.

DrawingML is shared by PresentationML, WordprocessingML, and SpreadsheetML. Code that reads only `a:` elements is a candidate for `packages/ooxml-core`. Colors and themes: [`themes-colors.md`](themes-colors.md). Slides and placeholders: [`presentationml.md`](presentationml.md).

## Shape properties (`p:spPr`)

Schema order: `a:xfrm`, geometry (`a:prstGeom` | `a:custGeom`), fill, `a:ln`, effects (`a:effectLst` | `a:effectDag`), `a:scene3d`, `a:sp3d`, `a:extLst`.

### Transform (`a:xfrm`)

```xml
<a:xfrm rot="5400000" flipH="1" flipV="0">
  <a:off x="914400" y="914400"/>    <!-- top-left, EMU -->
  <a:ext cx="1828800" cy="914400"/> <!-- size, EMU -->
</a:xfrm>
```

- `rot` is in 60000ths of a degree, clockwise, around the box center. Flips apply before rotation.
- Groups add `a:chOff` and `a:chExt`: child coordinates map from the `chOff`/`chExt` box onto the `off`/`ext` box.
- A missing `a:xfrm` on a placeholder means "inherit from the layout or master placeholder".

### Geometry

```xml
<a:prstGeom prst="roundRect">
  <a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst>
</a:prstGeom>
```

- 187 presets (`ST_ShapeType`). Each has adjust values (`avLst`), guides (`gdLst`), and paths. Evaluate the guide formulas from the spec file; do not draw presets by eye.
- A shape's `a:avLst` stores only the adjust values that differ from the preset; an absent guide takes the preset's `avLst` default. The preset's `a:ahLst` gives each adjust handle a range: `a:ahXY` (`gdRefX`/`minX`/`maxX`, `gdRefY`/`minY`/`maxY`) and `a:ahPolar` (`gdRefR`/`minR`/`maxR`, `gdRefAng`/`minAng`/`maxAng`). A bound is a literal or a guide name, which depends on the shape size and the other adjust values (`rightArrow` `maxAdj2`). A guide that `gdRefAng` moves is an angle in 60 000ths of a degree; the others are in 100 000ths.
- Guide operators: `*/ +- +/ ?: abs at2 cos max min mod pin sat2 sin sqrt tan val`. Angles are in 60000ths of a degree.
- `a:custGeom`: `a:avLst`, `a:gdLst`, `a:ahLst`, `a:cxnLst`, `a:rect` (text box), `a:pathLst`. Each `a:path` has `w`, `h`, `fill` (`norm`, `none`, `lighten`, `darken`, …), `stroke`, and commands `a:moveTo`, `a:lnTo`, `a:arcTo` (`wR hR stAng swAng`), `a:quadBezTo`, `a:cubicBezTo`, `a:close`.
- `CT_Path2D` lets a path start with `a:lnTo` or `a:arcTo`, but only `a:moveTo` sets the current point. A command before the first `a:moveTo` draws nothing: the reference PDF of fixture `custom-geometry` slide 2 shows no fill and no outline for a path that has no `a:moveTo`.
- The text rectangle comes from the geometry's `a:rect`, then `a:bodyPr` insets apply.

## Fills

| Element | Key content |
| --- | --- |
| `a:noFill` | No fill |
| `a:solidFill` | One color element |
| `a:gradFill` | `a:gsLst/a:gs pos` (0–100000), then `a:lin ang scaled` or `a:path path="circle|rect|shape"` with `a:fillToRect`, then `a:tileRect`; `rotWithShape`, `flip` |
| `a:blipFill` | `a:blip r:embed` (+ color effects), `a:srcRect` (crop, thousandths of a percent), `a:stretch/a:fillRect` or `a:tile tx ty sx sy flip algn`; `dpi` (the engine does not read it; saved decks hold `0`), `rotWithShape` |
| `a:pattFill` | `prst` (54 presets), `a:fgClr`, `a:bgClr`; each 8×8 preset pixel draws as 1 pt, and the 8 pt tiles sit on a grid from the slide's top-left corner, not the shape's (`fixtures/fills-lines.pdf`) |
| `a:grpFill` | Use the parent group's fill; a group with `a:grpFill` passes its own parent's fill on (§20.1.8.35) |

Gradient rules (`CT_GradientFillProperties`; behavior measured against reference PDFs):

- `a:lin/@ang` is the direction of color change, clockwise from +x, in 60000ths of a degree. `scaled="0"` (the default) keeps that angle in the fill box. `scaled="1"`: the gradient is defined on the unit square, and the square stretches to the box, so the isolines run along (−w sin a, h cos a). The spec text says the vector (w cos a, h sin a) instead; the reference PDFs win.
- The ends of a linear gradient are the projections of the farthest box corners onto the gradient line through the box center.
- Stops are sorted by `pos`; two stops at one `pos` make a hard edge. Three or more stops interpolate linearly in sRGB. A gradient with exactly two stops, at 0 and 100 %, bends (measured on reference PDFs): each sRGB channel follows (1 − u)² a + 2u(1 − u) max(a, b) + u² b; alpha stays linear.
- `a:path`: the first stop is at the focus `a:fillToRect` (insets from the box; for `rect` the whole focus rectangle, for `circle` and `shape` its center), the last stop is at the outer boundary, linear in between: the box edge for `rect`, the shape outline for `shape`. For `circle`, the fill is a focal radial gradient: the outer circle passes through the corners of the tile rectangle (the box by default) and is centered on it, and grows to 1.005 times the focus distance when the focus is at or outside it; the focal point is the center of `a:fillToRect`. The spec defines `a:fillToRect` relative to the tile rectangle (§20.1.8.31); reference PDFs place it on the shape's box.
- `a:tileRect` names the rectangle the gradient is computed in, as insets from the box, tiled and mirrored by `flip` (`ST_TileFlipMode`). Saved decks hold `<a:fillToRect r="100000" b="100000"/>` with `<a:tileRect r="-100000" b="-100000"/>` for a radial gradient from the top-left corner: the circle spans the doubled box, and the focus stays at the shape's corner (`a:fillToRect` applies to the shape's box).

Picture fill rules (`CT_BlipFillProperties`):

- `a:srcRect` crops the image (positive insets) or pads it (negative insets). `a:stretch/a:fillRect` places the cropped image in the box: positive insets shrink it, negative insets extend it past the box. The geometry clips the result.
- `a:tile`: the tile is the image's natural size × `sx`/`sy` (`ST_Percentage`), placed at the `algn` point (`ST_RectAlignment`, default `tl`) plus `tx`/`ty` (EMU), mirrored by `flip`. The natural size comes from the image header's resolution; it uses 72 DPI when the header states none (`fixtures/fills-lines.pdf`).
- Blip effects (`a:duotone`, `a:grayscl`, `a:alphaModFix`, `a:lum`, `a:biLevel`) apply to tiles as to stretched pictures.

## Lines (`a:ln`)

```xml
<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
  <a:prstDash val="dash"/>          <!-- or a:custDash/a:ds d sp -->
  <a:round/>                         <!-- or a:bevel, a:miter lim -->
  <a:headEnd type="triangle" w="med" len="med"/>
  <a:tailEnd type="none"/>
</a:ln>
```

- `w` in EMU. `cmpd`: `sng`, `dbl`, `thickThin`, `thinThick`, `tri`. `algn`: `ctr` or `in`.
- Arrowhead `type`: `none`, `triangle`, `stealth`, `diamond`, `oval`, `arrow`. `w`/`len`: `sm`, `med`, `lg`.

## Effects (`a:effectLst`)

Schema order: `a:blur`, `a:fillOverlay`, `a:glow`, `a:innerShdw`, `a:outerShdw`, `a:prstShdw`, `a:reflection`, `a:softEdge`.

| Effect | Attributes |
| --- | --- |
| `a:outerShdw` | `blurRad dist dir sx sy kx ky algn rotWithShape` + color |
| `a:innerShdw` | `blurRad dist dir` + color |
| `a:glow` | `rad` + color |
| `a:softEdge` | `rad` |
| `a:reflection` | `blurRad stA stPos endA endPos dist dir fadeDir sx sy kx ky algn rotWithShape` |
| `a:fillOverlay` | `blend` (`over`, `mult`, `screen`, `darken`, `lighten`) + one fill element; the fill blends with the shape's fill only |

Distances are in EMU; `dir` is in 60000ths of a degree. The "Known approximations" section of `docs/site/content/pptx-fidelity.mdx` lists the effects the engine does not draw.

Measured on the reference PDFs (`effects` slide 1, `shape-effects-presets` slides 1 and 2); the schema gives each radius but no kernel:

- A blur radius (`blurRad` of `a:outerShdw` and `a:innerShdw`, and the blur of `a:glow` and `a:softEdge`) is a Gaussian with σ = radius / 3: its tail ends at the radius.
- `a:glow rad`: full glow alpha on the edge, half alpha at `rad / 2` outside it, none at `rad`. That is the outline grown by `rad / 2` and blurred with radius `rad / 2`.
- `a:softEdge rad`: alpha 0 on the outline, half at about 0.84 `rad` inside it, full at about 1.84 `rad`. That is the outline shrunk by 0.84 `rad` and blurred with radius `rad`. Measured at 2.5 and 10 pt only, on opaque fills.
- `a:innerShdw dir`: the shadow shows on the side that `dir` points to. `dir="0"` shades the right edge, `5400000` the bottom edge, and `13500000` the top and left edges.

## Shape style (`p:style`)

```xml
<p:style>
  <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
  <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
  <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
  <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>
```

Explicit `p:spPr` values win over style references. The reference color replaces `phClr` in the theme style.

- `p:style` comes after `p:spPr` and before `p:txBody` (CT_Shape).
- An `a:effectLst` in `p:spPr` replaces the whole effect list of `a:effectRef`, so an empty `<a:effectLst/>` turns a theme shadow off. It is not merged effect by effect.

## Effect presets and Shape Styles

These were confirmed on 31 October 2026 against saved check decks (the same edits on `arrange-flip-hidden`, saved, and compared): Offset Bottom Right and Inside Top Left shadows, Half Reflection 4 pt (`endPos` 55500), Soft Edges 2.5 pt, Quick Style preset 38 (Intense Effect, Accent 2), flip, rotate, the aspect lock, hide, and rename. The other presets are the same pattern but are **not checked one by one**. That check wrote the glow color without `a:satMod` and `a:alpha`; the gallery value is not checked.

| Gallery | XML the editor writes |
| --- | --- |
| Shadow, Outer ("Offset: …") | `a:outerShdw blurRad="50800" dist="38100" dir algn rotWithShape="0"`, black at 40 % alpha (`a:prstClr val="black"`). A saved check deck of 2 November 2026 holds this, also on a shape turned 90°, whose shadow stays at the slide's bottom right. `dir`/`algn`: bottom right 2700000/`tl`, bottom 5400000/`t`, bottom left 8100000/`tr`, right 0/`l`, left 10800000/`r`, top right 18900000/`bl`, top 16200000/`b`, top left 13500000/`br` |
| Shadow, Offset: center | `a:outerShdw blurRad="63500" sx="102000" sy="102000" algn="ctr" rotWithShape="0"`, black at 40 % |
| Shadow, Inner ("Inside: …") | `a:innerShdw blurRad="63500" dist="50800" dir`, black at 50 % (`CT_InnerShadowEffect` has no `rotWithShape`); the same `dir` per side as the outer shadow from the opposite corner. Inside: center is `blurRad="114300"`, opaque black |
| Reflection | `a:reflection blurRad="6350" stA="50000" endA="300" endPos dist dir="5400000" sy="-100000" algn="bl" rotWithShape="0"`; `endPos` 35000 (tight), 55500 (half), 90000 (full); `dist` 0, 50800 (4 pt), 101600 (8 pt) |
| Glow | `a:glow rad` (5, 8, 11, 18 pt) with the color, `a:satMod val="175000"` for a theme color, and `a:alpha val="40000"` |
| Soft Edges | `a:softEdge rad` (1, 2.5, 5, 10, 25, 50 pt) |

Shape Styles gallery (theme styles), as `lnRef` / `fillRef` / `effectRef` index and color, and the `fontRef` color. `C` is the column color: `dk1` or `accent1` to `accent6`.

| Row | `lnRef` | `fillRef` | `effectRef` | `fontRef` |
| --- | --- | --- | --- | --- |
| Colored Outline | 2, `C` | 1, `lt1` | 0 | `dk1` |
| Colored Fill | 2, `C` with `a:shade val="15000"` | 1, `C` | 0 | `lt1` |
| Light 1 Outline, Colored Fill | 3, `lt1` | 1, `C` | 1 | `lt1` |
| Subtle Effect | 1, `C` | 2, `C` | 1 | `dk1` |
| Moderate Effect | 1, `C` | 3, `C` | 2 | `lt1` |
| Intense Effect | 0, `C` | 3, `C` | 3 | `lt1` |

Applying a style removes the shape's own fill, `a:ln`, and effects, so the style shows.

## Text body (`p:txBody`)

```xml
<p:txBody>
  <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"
            anchor="t" anchorCtr="0" vert="horz" numCol="1" spcCol="0" rot="0">
    <a:normAutofit fontScale="90000" lnSpcReduction="10000"/>
  </a:bodyPr>
  <a:lstStyle/>
  <a:p>…</a:p>
</p:txBody>
```

### Body properties (`a:bodyPr`)

| Attribute | Values | Default |
| --- | --- | --- |
| `lIns`, `rIns` / `tIns`, `bIns` | EMU | 91440 / 45720 (0.1 in / 0.05 in) |
| `anchor` | `t`, `ctr`, `b`, `just`, `dist` | `t` |
| `wrap` | `square`, `none` | `square` |
| `vert` | `horz`, `vert`, `vert270`, `eaVert`, `wordArtVert`, `mongolianVert`, `wordArtVertRtl` | `horz` |
| `numCol`, `spcCol` | Columns, 1 to 16 (`ST_TextColumnCount`), and the gap in EMU (`ST_PositiveCoordinate32`) | 1, 0 |
| `rtlCol`, `upright`, `rot`, `fromWordArt`, `forceAA` | | |

Autofit child (one of): `a:noAutofit`, `a:normAutofit fontScale lnSpcReduction` (shrink text on overflow; the stored values are what the last writer computed), `a:spAutoFit` (resize the shape to fit the text). `a:prstTxWarp` sets a text warp; `a:scene3d`/`a:sp3d` set 3-D.

Applying a stored `a:normAutofit` (measured on reference PDFs):

- Opening or exporting a deck does not fit the text again.
- Each run's size becomes `round(sz × fontScale)` in whole points, half up (28 pt × 62.5 % → 18 pt).
- `lnSpcReduction` is subtracted from the line spacing percentage (110 % − 10 % = 100 %). It does not reduce paragraph spacing.
- After an edit, the writer probably first reduces line spacing (10 %, then 20 %), then lowers `fontScale` in 7.5 % steps with the 20 % reduction. This order is inferred from the pairs stored in saved decks, not verified.

The defaults above are the schema defaults. Inherited `a:bodyPr` from the layout and master placeholders applies first.

### Paragraph (`a:p`)

```xml
<a:p>
  <a:pPr lvl="0" algn="l" marL="342900" indent="-342900" rtl="0">
    <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
    <a:spcBef><a:spcPts val="600"/></a:spcBef>
    <a:spcAft><a:spcPts val="0"/></a:spcAft>
    <a:buClr>…</a:buClr> <a:buSzPct val="100000"/> <a:buFont typeface="Arial"/>
    <a:buChar char="•"/>                   <!-- or a:buAutoNum type startAt, a:buBlip, a:buNone -->
    <a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst>
    <a:defRPr/>
  </a:pPr>
  <a:r><a:rPr lang="en-US" sz="1800" b="1"/><a:t>Text</a:t></a:r>
  <a:br/>                                   <!-- soft line break, carries a:rPr -->
  <a:fld id="{…}" type="slidenum"><a:rPr/><a:t>‹#›</a:t></a:fld>
  <a:endParaRPr lang="en-US"/>
</a:p>
```

- `algn`: `l`, `ctr`, `r`, `just`, `justLow`, `dist`, `thaiDist`. `lvl`: 0–8.
- `a:spcPct` is a percentage of the line, 0 to 13200000 in 1/1000 % (`ST_TextSpacingPercentOrPercentString`); `a:spcPts` is in hundredths of a point, 0 to 158400 (`ST_TextSpacingPoint`). The line model, measured against reference PDFs (`packages/core/src/layout/text/line-metrics.ts`):
  - A line advances `pct × 1.2 ×` its largest run size, for the 11 Latin faces measured. `a:spcPts` line spacing rounds to a whole point.
  - With the face's ascent share `a / (a + d)` and the single line `S = 1.2 × size`: a single line (`L = S`) puts the baseline at `L × share`. A taller line (percent and exact spacing alike) puts it at `min(0.75 × L, 0.75 × (L − S) + S × share)`. A shorter line keeps the face's descent share of `S` below the baseline, but never more than 25 % of the box (measured on Latin faces only). For a face with a share of 0.75 or more, every line taller than `S` sits at 75 % of the box. Cordia New (share 0.648) and Gabriola (0.684) show the rule: 0 of 178 single and 0 of 51 1.5-line probe lines missed, against all of them with a 75 % floor.
  - The face's ascent and descent come from `OS/2`: sTypoAscender and sTypoDescender when fsSelection bit 7 (USE_TYPO_METRICS) is set, in any table version; otherwise usWinAscent and usWinDescent. Never `hhea`. Measured on 2,231 probe lines in 10 fonts exported to reference PDFs, the rule misses 8 lines, all 0.0002 pt below a half-point rounding step. Aptos (bit 7, typo 1923 and 577 of 2048) misses 68 of 178 lines with its win values. Yu Gothic (no bit 7, win 2017 and 619) misses 64 of 240 with `hhea` (fixture `east-asian-line-breaking`). Cambria Math sets bit 7 in a version 3 table, and its lines follow the typo values. Most faces store the same values in every set.
  - Every baseline rounds to a whole point from the top of the text block (from the top of its column with `numCol`). Line tops do not round, so the rounding does not add up from line to line. Vertical anchoring (`ctr`, `b`) then moves the block by an offset that does not round. ECMA-376 does not say this; it puts each line-checked fixture line within 0.12 pt of the reference PDF's glyph origin (the PDF stores origins on a 1/300 in grid), against up to 0.61 pt without it. Table cells round from their own text top as well (39 of 43 measured cell lines within 0.12 pt).
  - A raised or lowered run (`a:rPr/@baseline`) does not move its line's baseline: the line keeps the baseline of its other runs (fixture `font-effects`).
  - `a:spcBef`/`a:spcAft` in points round to a whole point; in percent they are `pct × 1.2 ×` the first run size and do not round. They apply to the first and last paragraph only with `spcFirstLastPara="1"`.
  - A bullet larger than its text (`a:buSzPct` above 100 %) does not make the line taller.
- Bullet sub-elements may each be "follow text" (`a:buClrTx`, `a:buSzTx`, `a:buFontTx`).
- `a:buAutoNum type`: `arabicPeriod`, `arabicParenR`, `romanUcPeriod`, `alphaLcParenR`, and others (`ST_TextAutonumberScheme`).
- Field `type`: `slidenum`, `datetime`, `datetime1`–`datetime13`, and others. The cached `a:t` shows until a field is recomputed.

### Run properties (`a:rPr`, `a:defRPr`, `a:endParaRPr`)

| Attribute | Meaning |
| --- | --- |
| `sz` | Size in hundredths of a point (1800 = 18 pt) |
| `b`, `i` | Bold, italic (`1`/`0`) |
| `u` | `none`, `sng`, `dbl`, `heavy`, `dotted`, `dash`, `wavy`, … |
| `strike` | `noStrike`, `sngStrike`, `dblStrike` |
| `cap` | `none`, `small`, `all` |
| `baseline` | Thousandths of a percent; 30000 = superscript, -25000 = subscript |
| `spc` | Character spacing in hundredths of a point |
| `kern` | Kerning threshold in hundredths of a point. Saved decks hold `kern="0"` for "Kerning for fonts" off, so 0 means off, not a minimum size of 0 |
| `lang`, `altLang` | Language tags |

Children in schema order: `a:ln`, fill, effects, `a:highlight`, underline line/fill, `a:latin`, `a:ea`, `a:cs`, `a:sym`, `a:hlinkClick`, `a:hlinkMouseOver`, `a:rtl`. Font `typeface` may be a theme reference (`+mj-lt`, `+mn-ea`, …).

Hyperlinks: `a:hlinkClick r:id` points to a relationship with `TargetMode="External"` for URLs, or to a slide part. `action="ppaction://hlinksldjump"` jumps to a slide; `ppaction://hlinkshowjump?jump=nextslide` and others jump relatively.

### What the text controls write

The editor's text commands write these values. Values marked "not yet checked" come from saved decks and are not confirmed against a saved check deck. Confirm them when the fixtures `paragraph-format`, `font-effects`, `text-direction`, and `bullets-gallery` get reference PDFs and saves.

| Control | Writes |
| --- | --- |
| Paragraph dialog, spacing before and after | `a:spcBef` / `a:spcAft` with `a:spcPts` (hundredths of a point, 0 to 158400) or `a:spcPct` (lines; 100000 = one line) |
| Paragraph dialog, line spacing | `a:lnSpc` with `a:spcPct` (Single 100000, 1.5 lines 150000, Double 200000, Multiple) or `a:spcPts` (Exactly) |
| Paragraph dialog, indentation | `marL` (before text, 0 to 51206400) and `indent` (negative: hanging; positive: first line) |
| Superscript, Subscript | `baseline="30000"`, `baseline="-25000"` |
| Character Spacing presets | `spc` -300, -150, 0, 300, 600 (Very Tight to Very Loose; not yet checked) |
| Font dialog, kerning off | `kern="0"` (not yet checked) |
| Font dialog, underline color | `a:uFill` with `a:solidFill` after `a:highlight`; Automatic writes `a:uFillTx` (not yet checked) |
| Text Direction | `a:bodyPr/@vert`: `horz`, `vert`, `vert270`, `wordArtVert` (Stacked), `eaVert` |
| Columns, More Columns | `a:bodyPr/@numCol` (1 to 16) and `@spcCol` (EMU) |
| Bullets gallery | `a:buFont` + `a:buChar`: `•` Arial, `o` Courier New, and `§` `q` `v` `Ø` `ü` in Wingdings with `pitchFamily="2" charset="2"` (not yet checked) |
| Bullets dialog, custom character | `a:buChar` with the character, and `a:buFont` for a chosen font or `a:buFontTx` for the text font |
| Numbering gallery | `a:buFont typeface="+mj-lt"` + `a:buAutoNum type` (`startAt` only above 1) |

A Wingdings `a:buChar` holds the symbol font's character code (`§` is U+00A7), not the Unicode symbol. Without the font, the code draws as the plain character.

Underline color (`EG_TextUnderlineFill`): `a:uFill` holds the underline's own fill, and the empty `a:uFillTx` makes the underline follow the text fill again. The nearest level of the text style cascade that has either one wins. The underline line (`a:uLn`, `a:uLnTx`) is a separate choice. Strikethrough always takes the text fill.

## Pictures (`p:pic`)

```xml
<p:pic>
  <p:nvPicPr><p:cNvPr id="4" name="Picture 3" descr="Alt text"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"><a:alphaModFix amt="50000"/></a:blip>
    <a:srcRect l="10000" t="0" r="10000" b="0"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>…</p:spPr>
</p:pic>
```

- Blip effects: `a:alphaModFix`, `a:grayscl`, `a:duotone`, `a:lum bright contrast`, `a:biLevel thresh`, `a:clrChange`, and others.
- SVG pictures: `a:blip r:embed` points to a PNG fallback; `a:extLst` holds `asvg:svgBlip r:embed` with the SVG.
- Only load pictures from package parts. Never fetch `TargetMode="External"` targets.
- Order inside `p:blipFill` (CT_BlipFillProperties): `a:blip`, `a:srcRect`, then `a:tile` or `a:stretch`. Effects inside `a:blip` apply in document order, before `a:extLst`.
- Crop (`a:srcRect`, ECMA-376 Part 1, 20.1.8.55): each edge is a signed `ST_Percentage` of the image. A negative value adds empty space on that side. Crop > Fit writes such values. The visible image is `1 − l − r` of its width; paint scales the full image by `cx / (1 − l − r)`.
- Crop to shape is the picture's own geometry: `a:prstGeom` in `p:spPr` clips the image. Lines have no inside, so they are not offered.
- Alt text lives on `p:cNvPr`: `descr` (description) and `title`. "Mark as decorative" is an extension of `p:cNvPr`: `<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>`. The editor drops `descr` when it marks an object as decorative.
- Picture presets (read from saved decks, not yet checked against a reference PDF; `editor/picture-presets.ts`): Transparency is `a:alphaModFix amt` (100000 − transparency). Grayscale is `a:grayscl`. Black and white is `a:biLevel thresh` (25000, 50000, 75000). Sepia is `a:duotone` with `a:prstClr val="black"` and `a:srgbClr val="D9C3A5"` (`a:tint 50000`, `a:satMod 180000`). An accent's dark variation is `a:duotone` from the accent (`a:shade 45000`, `a:satMod 135000`) to white; its light variation is from black to the accent (`a:tint 45000`, `a:satMod 400000`).
- Corrections: ECMA-376 has `a:lum bright contrast` (20.1.8.42). The `a14` extension writes Brightness/Contrast as `a14:imgProps/a14:imgLayer/a14:imgEffect/a14:brightnessContrast` in the blip's `a:extLst` (`uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"`), with an HD Photo copy of the original image in `a14:imgLayer`; `a:blip/@r:embed` then holds the image with the effects applied. The editor writes `a:lum`; a reader that knows `a14` reads both. The editor refuses Reset Picture and Corrections on a picture with `a14:imgProps`. Saturation, color tone, and artistic effects exist only as `a14` effects.
- A picture's outline (`p:pic/p:spPr/a:ln`) is drawn outside the picture, not centered on its edge (a reference PDF of an editor-saved check deck).
- A JPEG's EXIF orientation (TIFF tag 0x0112) is ignored: a picture draws and sizes the stored pixels (check deck r54-1). A browser applies the tag to an SVG `<image>` unless `image-orientation: none` is set.
- Picture styles set the geometry, `a:ln`, and `a:effectLst` of `p:spPr`. Simple Frame, White is `a:ln w="88900" cap="sq"` in white with `a:miter lim="800000"`, and an `a:outerShdw` (`blurRad="55000" dist="18000" dir="5400000" algn="tl" rotWithShape="0"`, black at 40 % alpha). Soft Edge Rectangle is `a:softEdge rad="112500"`. Reflected Rounded Rectangle is `roundRect` (`adj` 8594) with `a:reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"`.

## Tables (`a:tbl`)

```xml
<a:tbl>
  <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-…}</a:tableStyleId></a:tblPr>
  <a:tblGrid><a:gridCol w="3048000"/>…</a:tblGrid>
  <a:tr h="370840">
    <a:tc gridSpan="2" rowSpan="1" hMerge="0" vMerge="0">
      <a:txBody>…</a:txBody>
      <a:tcPr marL="91440" anchor="ctr"><a:lnL>…</a:lnL>…<a:solidFill>…</a:solidFill></a:tcPr>
    </a:tc>
  </a:tr>
</a:tbl>
```

- `a:tblPr` flags: `firstRow`, `lastRow`, `firstCol`, `lastCol`, `bandRow`, `bandCol`, `rtl`.
- `a:tableStyleId` refers to `ppt/tableStyles.xml` or to a built-in style by GUID.
- Cell borders: `a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`, `a:lnTlToBr`, `a:lnBlToTr`. A row `h` is a minimum: rows grow to fit text.
- Every `a:tr` holds one `a:tc` per `a:gridCol`, covered cells too (21.1.3.16 tc). The schema allows any count (`CT_TableRow`: `tc` 0 to unbounded) and an `a:tc` without `a:txBody`, but saved decks hold both in every cell; the editor gives a covered cell a body when it becomes a drawn cell.
- A merged block, as saved decks hold it: every cell of the block's first row carries the block's `rowSpan`, and every cell of its first column the block's `gridSpan`. The other cells of the first row carry `hMerge="1"`, the other cells of the first column `vMerge="1"`, and the rest `hMerge="1" vMerge="1"`. A 2 × 2 block: `gridSpan="2" rowSpan="2"`, `hMerge="1" rowSpan="2"`; `gridSpan="2" vMerge="1"`, `hMerge="1" vMerge="1"`. A reader takes the spans from the top-left cell only.
- Merge markup that disagrees with itself is not defined by ECMA-376 (`CT_TableCell` gives only the defaults: `gridSpan` and `rowSpan` 1, `hMerge` and `vMerge` false). The engine reads coverage from spans only, in reading order (`store/package/table-grid.ts`): an `hMerge` or `vMerge` cell that no span reaches is a drawn cell with its own text and spans, a span past the grid ends at the grid, and a span into an earlier block ends before it. Not measured by a reference PDF yet.
- `a:tcPr` child order (CT_TableCellProperties): `lnL`, `lnR`, `lnT`, `lnB`, `lnTlToBr`, `lnBlToTr`, `cell3D`, one fill, `headers`, `extLst`. Attributes: `marL`, `marR` (default 91440), `marT`, `marB` (default 45720), `vert`, `anchor`, `anchorCtr`, `horzOverflow`.
- A cell border from the Borders menu, as saved decks hold it: `<a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>…</a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL>`. No Border is `<a:lnL w="12700" cmpd="sng"><a:noFill/></a:lnL>`.
- `a:tblStyleLst/@def` in `ppt/tableStyles.xml` is the style a new table gets. Insert > Table makes a table two thirds of the slide wide, with `a:tr h="370840"`, `firstRow="1" bandRow="1"`, and `a:graphicFrameLocks noGrp="1"`.

## Common pitfalls

- An absent attribute is not 0. Use the schema default, then the cascade.
- `a:xfrm` is optional on placeholders, and `a:ext` can be 0 on lines. Do not divide by a size without a check.
- Keep schema element order on write; a file with out-of-order children gets a repair prompt.
- Clamp every file number before it reaches a loop or an allocation: guide counts, path commands, `numCol`, gradient stops.
