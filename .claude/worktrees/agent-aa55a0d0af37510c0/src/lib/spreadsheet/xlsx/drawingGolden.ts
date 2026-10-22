// The golden wire forms of drawings created on a sheet (SPEC-131 "Canonical
// new drawings"): hand-authored against ECMA-376 Part 1 §20.5 (DrawingML —
// SpreadsheetML Drawing) and written BEFORE the serialiser in
// xlsx/drawingXml.ts, which drawingXml.spec.ts holds to these strings byte
// for byte. THIS FILE is the normative source for every detail the spec left
// to the goldens:
// - an image is a `twoCellAnchor editAs="oneCell"` (Excel's own default for a
//   picture) holding an `xdr:pic`; `a:blip` declares `xmlns:r` itself (as
//   Excel writes it), so the anchor is valid inside a source part whose root
//   binds no relationships prefix;
// - a shape is a `twoCellAnchor` (no editAs: "twoCell" is the schema default)
//   holding an `xdr:sp macro="" textlink=""` styled explicitly in `xdr:spPr`
//   — no `xdr:style`, so nothing resolves through the theme's style matrix:
//   `a:solidFill` (none for a line), `a:ln w` (EMU) with its own solid fill
//   and `a:prstDash` only when not solid — and an empty `xdr:txBody` exactly
//   as Excel writes a fresh shape's (`vertOverflow="clip"
//   horzOverflow="clip" rtlCol="0" anchor="t"`, one paragraph holding
//   `a:endParaRPr lang="en-US" sz="1100"`);
// - both carry `<xdr:clientData/>`; colours are upper-case `srgbClr` hex;
// - a cNvPr id is k + 1 for the drawing named "Picture {k}" / "{label} {k}";
// - a new drawing part is the XML declaration and an `xdr:wsDr` binding
//   `xdr` and `a`; its `.rels` lists one `…/relationships/image` per media
//   part its pictures embed, rId1 upwards.
// Every golden below is for the default grid (64 px columns, 20 px rows —
// 609,600 and 190,500 EMU), so each `a:off`/`a:ext` is checkable by hand.
import type { SheetDrawing } from "../model";

/** ECMA-376 Part 1 §20.5.2.33 `wsDr`: a new drawing part's root, with the
 * prefixes every canonical anchor below uses. */
export const GOLDEN_DRAWING_PART_OPEN =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`;

/** The new drawing part's close tag. */
export const GOLDEN_DRAWING_PART_CLOSE = `</xdr:wsDr>`;

/** The input of GOLDEN_IMAGE_ANCHOR: a picture as insertImage makes it —
 * B3 (1, 2) to D7 + 32 px × 10 px, 2 columns + 32 px by 4 rows + 10 px —
 * alt text set, minted cNvPr id 2, embedding rId1. */
export const GOLDEN_IMAGE_DRAWING: SheetDrawing = {
  id: "dr1",
  name: "Picture 1",
  altText: "Company logo",
  anchor: {
    kind: "twoCell",
    editAs: "oneCell",
    from: { col: 1, colOffEmu: 0, row: 2, rowOffEmu: 0 },
    to: { col: 3, colOffEmu: 304800, row: 6, rowOffEmu: 95250 },
    extEmu: { cx: 1524000, cy: 857250 },
  },
  content: { kind: "image", mediaId: "xl/media/image1.png" },
};

/** ECMA-376 Part 1 §20.5.2.33 `twoCellAnchor` (editAs §20.5.3.2) holding
 * §20.5.2.25 `pic`: `nvPicPr` (§20.5.2.24) with `cNvPr` (§20.5.2.8) and
 * `cNvPicPr` (§20.5.2.7) locking the aspect ratio; `blipFill` (§20.5.2.2)
 * with `a:blip r:embed` and `a:stretch`/`a:fillRect`; `spPr` (§20.5.2.30)
 * with `a:xfrm` — the anchor's absolute box (x = 1 × 609600, y = 2 × 190500,
 * cx = 2 × 609600 + 304800, cy = 4 × 190500 + 95250) — and the `rect`
 * preset; then §20.5.2.3 `clientData`. */
export const GOLDEN_IMAGE_ANCHOR =
  `<xdr:twoCellAnchor editAs="oneCell">` +
  `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>3</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:to>` +
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="Company logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  `<xdr:spPr><a:xfrm><a:off x="609600" y="381000"/><a:ext cx="1524000" cy="857250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
  `<xdr:clientData/></xdr:twoCellAnchor>`;

/** The input of GOLDEN_SHAPE_ANCHOR: a rectangle as insertShape makes it —
 * E3 (4, 2) to G8 + 32 px, styled DEFAULT_NEW_SHAPE_FILL /
 * DEFAULT_NEW_SHAPE_OUTLINE, minted cNvPr id 3. */
export const GOLDEN_SHAPE_DRAWING: SheetDrawing = {
  id: "dr2",
  name: "Rectangle 2",
  anchor: {
    kind: "twoCell",
    editAs: "twoCell",
    from: { col: 4, colOffEmu: 0, row: 2, rowOffEmu: 0 },
    to: { col: 6, colOffEmu: 304800, row: 7, rowOffEmu: 0 },
  },
  content: { kind: "shape", visual: { geometry: "rect", fill: { kind: "solid", color: "#4472c4" }, outline: { color: "#2e5395", widthPt: 1, dash: "solid" } } },
};

/** ECMA-376 Part 1 §20.5.2.33 `twoCellAnchor` (no editAs) holding §20.5.2.29
 * `sp`: `nvSpPr` (§20.5.2.23) with `cNvPr` and `cNvSpPr` (§20.5.2.9);
 * `spPr` with `a:xfrm` (x = 4 × 609600, y = 2 × 190500, cx = 2 × 609600 +
 * 304800, cy = 5 × 190500), `a:prstGeom` + `a:avLst`, `a:solidFill` and
 * `a:ln w="12700"` (1 pt); and §20.5.2.35 `txBody` empty — then
 * `clientData`. */
export const GOLDEN_SHAPE_ANCHOR =
  `<xdr:twoCellAnchor>` +
  `<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>6</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Rectangle 2"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="2438400" y="381000"/><a:ext cx="1524000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="2E5395"/></a:solidFill></a:ln></xdr:spPr>` +
  `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody>` +
  `</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;

/** The input of GOLDEN_LINE_ANCHOR: a line (no fill) drawn from B10 to
 * D10, minted cNvPr id 5. */
export const GOLDEN_LINE_DRAWING: SheetDrawing = {
  id: "dr3",
  name: "Line 4",
  anchor: {
    kind: "twoCell",
    editAs: "twoCell",
    from: { col: 1, colOffEmu: 0, row: 9, rowOffEmu: 0 },
    to: { col: 3, colOffEmu: 0, row: 9, rowOffEmu: 0 },
  },
  content: { kind: "shape", visual: { geometry: "line", outline: { color: "#2e5395", widthPt: 1, dash: "solid" } } },
};

/** GOLDEN_SHAPE_ANCHOR's form for a line: the `line` preset, no fill
 * element, a zero-height box (x = 609600, y = 9 × 190500, cx = 2 ×
 * 609600). */
export const GOLDEN_LINE_ANCHOR =
  `<xdr:twoCellAnchor>` +
  `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="5" name="Line 4"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="609600" y="1714500"/><a:ext cx="1219200" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
  `<a:ln w="12700"><a:solidFill><a:srgbClr val="2E5395"/></a:solidFill></a:ln></xdr:spPr>` +
  `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody>` +
  `</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;

/** The input of GOLDEN_TEXT_SHAPE_ANCHOR: a shape another format carried
 * across (SPEC-121 "Download a copy as…"), with what it says — a one-cell
 * anchor at F2 + 10 px × 5 px, 96 × 40 px, a rounded rectangle with two
 * centred, middle-anchored paragraphs in 14 pt bold dark red, a dashed
 * 1.5 pt theme outline and alt text; minted cNvPr id 7. */
export const GOLDEN_TEXT_SHAPE_DRAWING: SheetDrawing = {
  id: "dr4",
  name: "Rounded Rectangle 6",
  altText: "Model assumptions",
  anchor: { kind: "oneCell", from: { col: 5, colOffEmu: 95250, row: 1, rowOffEmu: 47625 }, extEmu: { cx: 914400, cy: 381000 } },
  content: {
    kind: "shape",
    visual: { geometry: "roundRect", fill: { kind: "solid", color: "#fff2cc" }, outline: { color: { theme: "accent2" }, widthPt: 1.5, dash: "dash" } },
    text: { paragraphs: [{ text: "Assumptions", align: "center" }, { text: "Growth 5%", align: "center" }], anchor: "middle", sizePt: 14, color: "#c00000", bold: true },
  },
};

/** ECMA-376 Part 1 §20.5.2.24 `oneCellAnchor` (`from` + §20.5.2.14 `ext`)
 * holding an `sp` whose `txBody` carries text: `a:bodyPr anchor="ctr"`,
 * each paragraph's `a:pPr algn`, each run's `a:rPr` (sz in hundredths of a
 * point, `b`, a solid fill); the outline's `a:prstDash` and a scheme
 * colour; x = 5 × 609600 + 95250, y = 190500 + 47625. */
export const GOLDEN_TEXT_SHAPE_ANCHOR =
  `<xdr:oneCellAnchor>` +
  `<xdr:from><xdr:col>5</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:from>` +
  `<xdr:ext cx="914400" cy="381000"/>` +
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="7" name="Rounded Rectangle 6" descr="Model assumptions"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="3143250" y="238125"/><a:ext cx="914400" cy="381000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="dash"/></a:ln></xdr:spPr>` +
  `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>` +
  `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>Assumptions</a:t></a:r></a:p>` +
  `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>Growth 5%</a:t></a:r></a:p>` +
  `</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`;

/** The input of GOLDEN_ABSOLUTE_IMAGE_ANCHOR: a cropped picture fixed at
 * (640, 20) px, 96 × 64 px, minted cNvPr id 9, embedding rId2. */
export const GOLDEN_ABSOLUTE_IMAGE_DRAWING: SheetDrawing = {
  id: "dr5",
  name: "Picture 8",
  anchor: { kind: "absolute", posEmu: { x: 6096000, y: 190500 }, extEmu: { cx: 914400, cy: 609600 } },
  content: { kind: "image", mediaId: "xl/media/image2.jpeg", crop: { l: 0.1, t: 0, r: 0.25, b: 0.05 } },
};

/** ECMA-376 Part 1 §20.5.2.1 `absoluteAnchor` (§20.5.2.26 `pos` + `ext`)
 * holding a `pic` whose `blipFill` crops with `a:srcRect` (l/t/r/b in
 * 1/1000 of a percent, zero sides omitted). */
export const GOLDEN_ABSOLUTE_IMAGE_ANCHOR =
  `<xdr:absoluteAnchor><xdr:pos x="6096000" y="190500"/><xdr:ext cx="914400" cy="609600"/>` +
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="Picture 8"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/><a:srcRect l="10000" r="25000" b="5000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  `<xdr:spPr><a:xfrm><a:off x="6096000" y="190500"/><a:ext cx="914400" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
  `<xdr:clientData/></xdr:absoluteAnchor>`;

/** The new part holding GOLDEN_IMAGE_ANCHOR then GOLDEN_SHAPE_ANCHOR. */
export const GOLDEN_DRAWING_PART = GOLDEN_DRAWING_PART_OPEN + GOLDEN_IMAGE_ANCHOR + GOLDEN_SHAPE_ANCHOR + GOLDEN_DRAWING_PART_CLOSE;

/** ECMA-376 Part 2 §9.3 (OPC relationships) for that part: its one picture
 * embeds xl/media/image1.png as rId1. */
export const GOLDEN_DRAWING_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
  `</Relationships>`;

/** ECMA-376 Part 1 §18.3.1.36 `drawing`: what a worksheet that gains its
 * first drawing carries, between `smartTags` and `legacyDrawing`. */
export const GOLDEN_WORKSHEET_DRAWING = `<drawing r:id="rId1"/>`;

/** The worksheet's relationship to its new drawing part. */
export const GOLDEN_SHEET_DRAWING_RELATIONSHIP = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>`;

/** The `[Content_Types].xml` Override a new drawing part takes. */
export const GOLDEN_DRAWING_OVERRIDE = `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
