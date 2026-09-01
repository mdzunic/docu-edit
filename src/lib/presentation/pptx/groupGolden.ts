// SPEC-096 golden-first (SPEC-086's process): the authored-group `p:grpSp`
// scaffold and the real `a:avLst` adjustment list are hand-authored against
// ECMA-376 §19.3.1.22 (grpSp) / §20.1.9.11 (avLst/gd) — NOT screen-scraped
// from the serialiser. The serialiser must match these byte-for-byte; the
// reader must accept them plus a reordered-attribute variant.
//
// Geometry: 1cm = 360000 EMU. The golden group sits at (2cm, 2cm), 8×4cm —
// `a:off`/`a:ext` 720000/720000/2880000/1440000 — and `a:chOff`/`a:chExt`
// are IDENTICAL to `a:off`/`a:ext` (the identity child mapping, so
// `mapGroupChild`'s scale is 1 and children's slide-space EMU survive
// re-parse unchanged). Ids: the spTree scaffold is 1, the group claims 2,
// its two children 3 and 4 — the single-pass mint's order.
import type { GroupFrame, ShapeFrame } from "../model";

/** The model the goldens serialise from — re-exported by odp/groupGolden.ts
 * so both formats pin the SAME model. */
export const GOLDEN_GROUP_FRAME: GroupFrame = {
  id: "golden-group",
  kind: "group",
  xCm: 2,
  yCm: 2,
  widthCm: 8,
  heightCm: 4,
  children: [
    {
      id: "golden-rect",
      kind: "shape",
      geometry: "rect",
      xCm: 2,
      yCm: 2,
      widthCm: 3,
      heightCm: 2,
      fill: { type: "solid", color: "#ff0000" },
    },
    {
      id: "golden-ellipse",
      kind: "shape",
      geometry: "ellipse",
      xCm: 7,
      yCm: 4,
      widthCm: 3,
      heightCm: 2,
    },
  ],
};

/** An authored group inside `p:spTree`, ECMA-376 §19.3.1.22. */
export const GROUP_GOLDEN =
  `<p:grpSp>` +
  `<p:nvGrpSpPr><p:cNvPr id="2" name="Group 2"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm>` +
  `<a:off x="720000" y="720000"/><a:ext cx="2880000" cy="1440000"/>` +
  `<a:chOff x="720000" y="720000"/><a:chExt cx="2880000" cy="1440000"/>` +
  `</a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="720000" y="720000"/><a:ext cx="1080000" cy="720000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Shape 4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="2520000" y="1440000"/><a:ext cx="1080000" cy="720000"/></a:xfrm>` +
  `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>` +
  `</p:grpSp>`;

/** The same group with every element's attributes reordered — readers must
 * accept it identically (attribute order is not significant in XML). */
export const GROUP_GOLDEN_REORDERED =
  `<p:grpSp>` +
  `<p:nvGrpSpPr><p:cNvPr name="Group 2" id="2"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm>` +
  `<a:off y="720000" x="720000"/><a:ext cy="1440000" cx="2880000"/>` +
  `<a:chOff y="720000" x="720000"/><a:chExt cy="1440000" cx="2880000"/>` +
  `</a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr name="Shape 3" id="3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off y="720000" x="720000"/><a:ext cy="720000" cx="1080000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>` +
  `<p:sp><p:nvSpPr><p:cNvPr name="Shape 4" id="4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off y="1440000" x="2520000"/><a:ext cy="720000" cx="1080000"/></a:xfrm>` +
  `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>` +
  `</p:grpSp>`;

/** The shape model ADJ_GOLDEN serialises from: a roundRect whose one
 * adjustment (corner radius) is tuned to 0.25 — guide value 25000 in the
 * 100000ths scale `adjustmentToGuideValue` documents. */
export const GOLDEN_ADJ_FRAME: ShapeFrame = {
  id: "golden-adj",
  kind: "shape",
  geometry: "roundRect",
  xCm: 1,
  yCm: 1,
  widthCm: 4,
  heightCm: 2,
  adjustments: [0.25],
};

/** A shape carrying a REAL adjustment list, ECMA-376 §20.1.9.11: one
 * `a:gd name="adj" fmla="val N"` per stored adjustment (a lone adjustment is
 * named `adj`, several are `adj1`…). Ids: minted 2 as the slide's only
 * frame. */
export const ADJ_GOLDEN =
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="360000" y="360000"/><a:ext cx="1440000" cy="720000"/></a:xfrm>` +
  `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>` +
  `</p:spPr></p:sp>`;

/** The adjustment golden with attributes reordered, for the reader. */
export const ADJ_GOLDEN_REORDERED =
  `<p:sp><p:nvSpPr><p:cNvPr name="Shape 2" id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off y="360000" x="360000"/><a:ext cy="720000" cx="1440000"/></a:xfrm>` +
  `<a:prstGeom prst="roundRect"><a:avLst><a:gd fmla="val 25000" name="adj"/></a:avLst></a:prstGeom>` +
  `</p:spPr></p:sp>`;
