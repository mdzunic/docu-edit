// SPEC-131 golden wire forms for the ODS sheet drawings this product CREATES
// (ods/drawingXml.ts's canonical serialiser) — hand-authored against ODF 1.2
// before the serialiser was written, never screen-scraped from its output.
// An imported drawing is never written from these: its source element is
// re-emitted with only its anchor patched.
//
// Every golden is laid out on the GOLDEN_DRAWING_SHEET geometry: columns
// 96 px (1 in = 2.54 cm) and rows 18 pt (24 px = 0.635 cm), so each length
// is an exact decimal. Settled here, not guessed in code:
// - attribute order is LibreOffice Calc's: `table:end-cell-address`,
//   `table:end-x`, `table:end-y` (moveAndSize only), `draw:z-index`,
//   `draw:name`, `draw:style-name`, then `svg:width`, `svg:height`, `svg:x`,
//   `svg:y` (a line: `svg:x1`, `svg:y1`, `svg:x2`, `svg:y2`);
// - lengths are centimetres rounded to 4 decimals (`2.2225cm`, `0cm`), `svg:x`
//   / `svg:y` sheet-absolute (ODF 1.2 §9.2.15 — LibreOffice re-derives the
//   cell anchor from the absolute rectangle); `table:end-x`/`-y` are offsets
//   into the end cell;
// - the end cell is `Sheet.A1` as LibreOffice writes it — no `$`, the sheet
//   name quoted (`'My Sheet'.E6`, `''` for a quote) when a formula would
//   quote it or it holds a `.`;
// - a picture is `draw:frame` > `draw:image` (§10.4.2, §10.4.4) with
//   `xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"`; its
//   alt text an `svg:desc` after the image (§10.4.2's content order);
// - an area shape is `draw:custom-shape` > `draw:enhanced-geometry` (§10.6.1,
//   §19.145) with `svg:viewBox="0 0 21600 21600"` and the catalogue's
//   `draw:type`; stored adjustments as `draw:modifiers`, each in this
//   product's hundred-thousandths (shapeCatalogue.ts adjustmentToGuideValue)
//   and omitted when the shape stores none — a new shape stores none; its
//   alt text an `svg:desc` first (§9.2.21 shape content order), then any
//   `text:p`, then the geometry;
// - a line or arrow is `draw:line` (§10.3.3) from the box's top-left to its
//   bottom-right; an arrow's head is the graphic style's
//   `draw:marker-end="Arrow"` (LibreOffice's standard marker name);
// - each drawing names a generated automatic graphic style `de-gr{n}`
//   (§16.2 `style:style style:family="graphic"`, §17.21
//   `style:graphic-properties`): a picture's draws no fill and no line; a
//   shape's carries `draw:fill`/`draw:fill-color`, `draw:stroke`,
//   `svg:stroke-color` and `svg:stroke-width` in points — equal property
//   runs share one style.

/** The sheet every golden is laid out on (see the module header). */
export const GOLDEN_DRAWING_SHEET = { name: "Sheet1", defaultColWidthPx: 96, defaultRowHeightPt: 18 } as const;

/** A new picture (`move`: twoCell editAs "oneCell", Excel's picture default)
 * at B2, 192 × 96 px, at the back: no end cell — the anchor reads back
 * `oneCell` with its size from `svg:width`/`svg:height`. ODF 1.2 §10.4.2. */
export const GOLDEN_ODS_IMAGE =
  `<draw:frame draw:z-index="0" draw:name="Picture 1" draw:style-name="de-gr1" svg:width="5.08cm" svg:height="2.54cm" svg:x="2.54cm" svg:y="0.635cm">` +
  `<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;

/** The same picture with alt text "Company logo": `svg:desc` after the image
 * (ODF 1.2 §10.4.2, §9.2.21). */
export const GOLDEN_ODS_IMAGE_WITH_ALT_TEXT =
  `<draw:frame draw:z-index="0" draw:name="Picture 1" draw:style-name="de-gr1" svg:width="5.08cm" svg:height="2.54cm" svg:x="2.54cm" svg:y="0.635cm">` +
  `<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>Company logo</svg:desc></draw:frame>`;

/** A picture's graphic style: no fill, no line. ODF 1.2 §16.2, §17.21. */
export const GOLDEN_ODS_IMAGE_STYLE = `<style:style style:name="de-gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>`;

/** A new rectangle (`moveAndSize`: twoCell editAs "twoCell") from C3 + 48 px
 * to E6 + (0, 12 px), second from the back: the end cell and its offsets
 * name where it ends. ODF 1.2 §10.6.1 (draw:custom-shape), §19.145
 * (draw:type), §9.2.15 (table:end-cell-address). */
export const GOLDEN_ODS_SHAPE =
  `<draw:custom-shape table:end-cell-address="Sheet1.E6" table:end-x="0cm" table:end-y="0.3175cm" draw:z-index="1" draw:name="Rectangle 2" ` +
  `draw:style-name="de-gr2" svg:width="3.81cm" svg:height="2.2225cm" svg:x="6.35cm" svg:y="1.27cm">` +
  `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="rectangle"/></draw:custom-shape>`;

/** The rectangle on a sheet named "My Sheet", with alt text "Note": the end
 * cell's sheet quoted, `svg:desc` before the geometry. */
export const GOLDEN_ODS_SHAPE_QUOTED_WITH_ALT_TEXT =
  `<draw:custom-shape table:end-cell-address="'My Sheet'.E6" table:end-x="0cm" table:end-y="0.3175cm" draw:z-index="1" draw:name="Rectangle 2" ` +
  `draw:style-name="de-gr2" svg:width="3.81cm" svg:height="2.2225cm" svg:x="6.35cm" svg:y="1.27cm">` +
  `<svg:desc>Note</svg:desc><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="rectangle"/></draw:custom-shape>`;

/** A shape's graphic style, DEFAULT_NEW_SHAPE_FILL / DEFAULT_NEW_SHAPE_OUTLINE
 * (core/shapes.ts): a solid #4472c4 fill, a 1 pt solid #2e5395 line. ODF
 * 1.2 §17.21, §20.112 (draw:fill), §20.142 (draw:stroke). */
export const GOLDEN_ODS_SHAPE_STYLE =
  `<style:style style:name="de-gr2" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#4472c4" draw:stroke="solid" ` +
  `svg:stroke-color="#2e5395" svg:stroke-width="1pt"/></style:style>`;

/** A `fixed` rounded rectangle (absolute) at the sheet's top-left, 96 × 48
 * px, with one stored adjustment (0.25 → `draw:modifiers="25000"`): no end
 * cell — it sits in `table:shapes`. */
export const GOLDEN_ODS_FIXED_SHAPE =
  `<draw:custom-shape draw:z-index="2" draw:name="Rounded Rectangle 3" draw:style-name="de-gr2" svg:width="2.54cm" svg:height="1.27cm" svg:x="0cm" svg:y="0cm">` +
  `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="round-rectangle" draw:modifiers="25000"/></draw:custom-shape>`;

/** The `table:shapes` element a dirty table writes its fixed drawings in
 * (after `office:forms`, before the columns). ODF 1.2 §9.2.15 is the anchor
 * kind; §9.1.1 table:table's content order places it. */
export const GOLDEN_ODS_TABLE_SHAPES = `<table:shapes>${GOLDEN_ODS_FIXED_SHAPE}</table:shapes>`;

/** A new line (`moveAndSize`) from A1 to C2 (192 × 24 px), drawn
 * top-left → bottom-right. ODF 1.2 §10.3.3. */
export const GOLDEN_ODS_LINE =
  `<draw:line table:end-cell-address="Sheet1.C2" table:end-x="0cm" table:end-y="0cm" draw:z-index="3" draw:name="Line 4" draw:style-name="de-gr3" ` +
  `svg:x1="0cm" svg:y1="0cm" svg:x2="5.08cm" svg:y2="0.635cm"/>`;

/** A line's graphic style: no fill, the default outline. */
export const GOLDEN_ODS_LINE_STYLE =
  `<style:style style:name="de-gr3" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="solid" svg:stroke-color="#2e5395" svg:stroke-width="1pt"/></style:style>`;

/** An arrow's graphic style: the line's, plus the end marker. */
export const GOLDEN_ODS_ARROW_STYLE =
  `<style:style style:name="de-gr4" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="solid" svg:stroke-color="#2e5395" svg:stroke-width="1pt" ` +
  `draw:marker-end="Arrow" draw:marker-end-width="0.3cm"/></style:style>`;
