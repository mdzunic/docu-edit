// SPEC-096 golden-first, odp side: the authored `<draw:g>` and the
// `draw:modifiers` adjustment attribute, hand-authored against ODF 1.3
// §10.3.6 (draw:g) / §19.145 (draw:modifiers) — not screen-scraped. The
// MODEL is the pptx golden's, re-exported so both formats pin the same one.
//
// A `draw:g` has no box of its own: its children keep their OWN
// `svg:x`/`svg:y`/`svg:width`/`svg:height` in slide space (odp/reader.ts
// derives the group's bounds as the children's union). The `draw:modifiers`
// scale is OUR OWN convention — `adjustmentToGuideValue`'s 100000ths, as
// shapeCatalogue.ts documents for OOXML — so what these pin is a faithful
// round trip through our own reader/writer pair, not byte-parity with
// Impress.
export { GOLDEN_GROUP_FRAME, GOLDEN_ADJ_FRAME } from "../pptx/groupGolden";

/** An authored group as `<draw:g>`, ODF 1.3 §10.3.6: children in slide
 * space, the filled rect carrying its generated `de-G1` graphic style. */
export const GROUP_GOLDEN =
  `<draw:g>` +
  `<draw:rect draw:style-name="de-G1" svg:x="2cm" svg:y="2cm" svg:width="3cm" svg:height="2cm"></draw:rect>` +
  `<draw:ellipse svg:x="7cm" svg:y="4cm" svg:width="3cm" svg:height="2cm"></draw:ellipse>` +
  `</draw:g>`;

/** The automatic style the golden group's filled rect references. */
export const GROUP_GOLDEN_AUTO_STYLES =
  `<style:style style:name="de-G1" style:family="graphic">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000"/>` +
  `</style:style>`;

/** The group with attributes reordered — the reader must accept it too. */
export const GROUP_GOLDEN_REORDERED =
  `<draw:g>` +
  `<draw:rect svg:y="2cm" svg:x="2cm" svg:height="2cm" svg:width="3cm" draw:style-name="de-G1"></draw:rect>` +
  `<draw:ellipse svg:height="2cm" svg:width="3cm" svg:y="4cm" svg:x="7cm"></draw:ellipse>` +
  `</draw:g>`;

/** The tuned roundRect as `<draw:custom-shape>`, its `draw:modifiers`
 * carrying the 100000ths guide value (0.25 → "25000"), ODF 1.3 §19.145. */
export const ADJ_GOLDEN =
  `<draw:custom-shape svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm">` +
  `<draw:enhanced-geometry draw:type="round-rectangle" draw:modifiers="25000"/>` +
  `</draw:custom-shape>`;

/** The adjustment golden with attributes reordered, for the reader. */
export const ADJ_GOLDEN_REORDERED =
  `<draw:custom-shape svg:height="2cm" svg:width="4cm" svg:y="1cm" svg:x="1cm">` +
  `<draw:enhanced-geometry draw:modifiers="25000" draw:type="round-rectangle"/>` +
  `</draw:custom-shape>`;
