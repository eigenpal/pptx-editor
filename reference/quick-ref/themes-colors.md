# DrawingML Themes & Colors Quick Reference

Extracted from ECMA-376 5th Edition - dml-main.xsd

After DOCX Editor's `reference/quick-ref/themes-colors.md` at adf0e178. The DrawingML sections are the same in both products. The WordprocessingML sections are replaced with PresentationML usage.

## Theme Structure

Located in: `ppt/theme/themeN.xml`, one per slide master (linked from the master's relationships). Notes and handout masters have their own themes.

```xml
<a:theme name="Default Theme">
    <a:themeElements>
        <a:clrScheme name="Default">...</a:clrScheme>
        <a:fontScheme name="Default">...</a:fontScheme>
        <a:fmtScheme name="Default">...</a:fmtScheme>
    </a:themeElements>
</a:theme>
```

---

## Color Scheme (a:clrScheme)

### Slots

| Slot        | Typical Use                        | Common Default |
| ----------- | ---------------------------------- | -------------- |
| `dk1`       | Dark text on light background      | Black (000000) |
| `lt1`       | Light background                   | White (FFFFFF) |
| `dk2`       | Secondary dark text                | Dark gray      |
| `lt2`       | Secondary light/accent background  | Light gray     |
| `accent1`   | Primary accent (links, highlights) | Blue           |
| `accent2`   | Secondary accent                   | Orange/Red     |
| `accent3`   | Tertiary accent                    | Green/Yellow   |
| `accent4-6` | Additional accents                 | Varies         |
| `hlink`     | Hyperlink color                    | Blue           |
| `folHlink`  | Followed hyperlink                 | Purple         |

### Color Definitions

```xml
<a:clrScheme name="Default">
    <!-- System colors (resolve at runtime) -->
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>

    <!-- Fixed RGB colors -->
    <a:dk2><a:srgbClr val="44546A"/></a:dk2>
    <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
    <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
    <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
    <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
```

---

## Color Types

### sRGB Color

```xml
<a:srgbClr val="FF0000"/>  <!-- Red -->
```

### System Color

```xml
<a:sysClr val="windowText" lastClr="000000"/>
```

System color values:

```
window, windowText, menuText, activeCaption, captionText,
highlight, highlightText, buttonFace, buttonText, etc.
```

### Scheme Color (reference to theme)

```xml
<a:schemeClr val="accent1"/>
```

With modifiers:

```xml
<a:schemeClr val="accent1">
    <a:tint val="50000"/>      <!-- 50% tint (lighter) -->
    <a:shade val="75000"/>     <!-- 75% shade (darker) -->
    <a:satMod val="120000"/>   <!-- 120% saturation -->
    <a:lumMod val="90000"/>    <!-- 90% luminance -->
</a:schemeClr>
```

### HSL Color

```xml
<a:hslClr hue="14400000" sat="100000" lum="50000"/>
<!-- Hue: 0-21600000 (0-360 degrees * 60000) -->
<!-- Sat/Lum: 0-100000 (0-100%) -->
```

---

## Tint and Shade Calculations

### In DrawingML

Values are in 1/1000ths of a percent (0-100000):

```xml
<a:tint val="50000"/>   <!-- 50% tint -->
<a:shade val="75000"/>  <!-- 75% shade -->
```

A percentage has a second spelling with a `%` sign. `a:ST_Percentage`, `a:ST_PositivePercentage`, `a:ST_FixedPercentage`, and `a:ST_PositiveFixedPercentage` are each a union of an `xsd:int` in thousandths of a percent and the matching `s:ST_*Percentage` pattern (`-?[0-9]+(\.[0-9]+)?%`). The union is in the Transitional schema (`dml-main.xsd`) too, so read both spellings in every file: `val="50%"` is `val="50000"`, and `val="12.5%"` is `val="12500"`. The `xsd:int` spelling can carry a leading `+`. Angles (`a:hue`, `a:hueOff`, and `a:hslClr/@hue`, `ST_Angle` and `ST_PositiveFixedAngle`) are `xsd:int` only, in 60000ths of a degree.

Apply the transforms in document order. `tint`, `shade`, `lumMod`, `lumOff`, `satMod`, and the other modifiers each have their own color space (linear RGB or HSL). The engine implements the spec math in `packages/ooxml-core/src/drawingml-color.ts` (`applyColorTransforms`); `packages/core/src/cascade/color.ts` resolves scheme names through the color map; the color space choices that are not calibrated are listed in the "Known approximations" section of `docs/site/content/pptx-fidelity.mdx`. Do not use the WordprocessingML `w:themeTint` formula here: it is a different, simpler rule.

Measured on the `colors` fixture (ECMA-376 Part 1, 20.1.2.3):

| Transform | Space | Example |
| --- | --- | --- |
| `a:inv` (20.1.2.3.17) | Linear RGB: each component becomes `linearToSrgb(1 - srgbToLinear(c))` | `4472C4` gives `F8EBB3`, not the sRGB inverse `BB8D3B` |
| `a:gray` (20.1.2.3.9) | Rec. 709 weights (0.2126, 0.7152, 0.0722) on sRGB components | `ED7D31` gives `8F8F8F`; Rec. 601 weights give `969696` |

---

## Font Scheme (a:fontScheme)

```xml
<a:fontScheme name="Default">
    <a:majorFont>  <!-- Headings -->
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>      <!-- East Asian -->
        <a:cs typeface=""/>      <!-- Complex script -->
        <!-- Script-specific overrides -->
        <a:font script="Jpan" typeface="Yu Gothic Light"/>
        <a:font script="Hans" typeface="DengXian Light"/>
    </a:majorFont>
    <a:minorFont>  <!-- Body text -->
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
    </a:minorFont>
</a:fontScheme>
```

### Referencing theme fonts in DrawingML

```xml
<a:rPr>
    <a:latin typeface="+mj-lt"/>  <!-- Major (heading) Latin -->
    <a:ea typeface="+mn-ea"/>     <!-- Minor (body) East Asian -->
    <a:cs typeface="+mn-cs"/>     <!-- Minor (body) complex script -->
</a:rPr>
```

Theme font references: `+mj-lt`, `+mj-ea`, `+mj-cs` (major) and `+mn-lt`, `+mn-ea`, `+mn-cs` (minor). `p:style/a:fontRef idx="major"` or `idx="minor"` sets the font through a shape style. Master title styles normally use `+mj-lt`; body styles use `+mn-lt`.

---

## Resolving Theme Colors

### Algorithm

1. Read `a:schemeClr val` (for example `tx1` or `accent1`).
2. If `val` is `phClr`, use the color passed by the referencing `p:style` entry or `p:bgRef`.
3. If `val` is a logical name (`bg1`, `tx1`, `bg2`, `tx2`), map it through the slide's color map: `p:clrMapOvr/a:overrideClrMapping` on the slide or layout, otherwise the master's `p:clrMap`.
4. Look up the slot in the theme's `a:clrScheme`.
5. If `a:sysClr`, use the `lastClr` attribute. If `a:srgbClr`, use `val`.
6. Apply the child transforms in order. The result is the final color, with alpha.

### Example Resolution

Slide has:

```xml
<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill>
```

Master has `<p:clrMap tx1="dk1" …/>`, and the theme has `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`.

Resolution:

1. `tx1` maps to `dk1` through the master color map.
2. Base color: #000000 (from `lastClr`).
3. `lumMod` 75 % then `lumOff` 25 % in HSL: luminance 0 × 0.75 + 0.25 = 25 %.
4. Final: dark gray #404040.

---

## Common default theme colors

| Slot     | Default color  | Hex    |
| -------- | -------------- | ------ |
| dk1      | Black          | 000000 |
| lt1      | White          | FFFFFF |
| dk2      | Dark Blue-Gray | 44546A |
| lt2      | Light Gray     | E7E6E6 |
| accent1  | Blue           | 4472C4 |
| accent2  | Orange         | ED7D31 |
| accent3  | Gray           | A5A5A5 |
| accent4  | Gold           | FFC000 |
| accent5  | Light Blue     | 5B9BD5 |
| accent6  | Green          | 70AD47 |
| hlink    | Blue           | 0563C1 |
| folHlink | Purple         | 954F72 |

---

## Format Scheme (a:fmtScheme)

Contains fill, line, and effect styles:

```xml
<a:fmtScheme name="Default">
    <a:fillStyleLst>
        <a:solidFill>...</a:solidFill>
        <a:gradFill>...</a:gradFill>
    </a:fillStyleLst>
    <a:lnStyleLst>
        <a:ln>...</a:ln>
    </a:lnStyleLst>
    <a:effectStyleLst>
        <a:effectStyle>...</a:effectStyle>
    </a:effectStyleLst>
    <a:bgFillStyleLst>
        <a:solidFill>...</a:solidFill>
    </a:bgFillStyleLst>
</a:fmtScheme>
```

Shapes reference the format scheme through `p:style`: `a:lnRef idx`, `a:fillRef idx`, `a:effectRef idx` (1-based into the lists; 0 means none), and `a:fontRef idx="major|minor"`. The color child of each reference replaces `phClr` in the referenced style. Backgrounds use `p:bgRef idx` ≥ 1001, which indexes `a:bgFillStyleLst`.

---

## References

- ECMA-376 Part 1, Section 20.1 (DrawingML - Main)
- dml-main.xsd: Color and theme schemas
- The implementer notes for ECMA-376 themes
