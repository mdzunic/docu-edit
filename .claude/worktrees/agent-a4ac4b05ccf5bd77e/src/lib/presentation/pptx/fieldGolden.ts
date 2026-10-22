// The NORMATIVE canonical pptx field placeholder markup (SPEC-093), hand-
// authored against ECMA-376 §19.3.1.36 (p:sp) / §21.1.2.2.4 (a:fld): a footer
// (plain a:r), a slide number (a:fld type="slidenum") and an auto date (a:fld
// type="datetime3" ⇄ d MMMM yyyy, the one sanctioned locale-dependent delta),
// each at an explicit 2×18cm / 8×1cm box so the EMU is fixed. serialize.spec.ts
// pins the serialiser to it both ways: canonicalFieldSp(RF, id) === the bytes,
// and the reader decodes the bytes (plus attribute-reordered variants) back to
// the entry. Cross-app verification is manual QA in the PR (AC-131).
import type { SlideFieldPlaceholder } from "../model";
import type { ResolvedField } from "./serialize";

const XFRM = `<a:xfrm><a:off x="720000" y="6480000"/><a:ext cx="2880000" cy="360000"/></a:xfrm>`;
const GEOM = { xCm: 2, yCm: 18, widthCm: 8, heightCm: 1 } as const;

// --- footer (plain a:r) -----------------------------------------------------
export const GOLDEN_FOOTER_SP =
  `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder"/>` +
  `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="ftr" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr>${XFRM}</p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
  `<a:r><a:rPr lang="en-US"/><a:t>Confidential</a:t></a:r></a:p></p:txBody></p:sp>`;
export const GOLDEN_FOOTER_RF: ResolvedField = { field: { kind: "footer", ...GEOM, align: "center" }, text: "Confidential" };
export const GOLDEN_FOOTER_ENTRY: SlideFieldPlaceholder = { kind: "footer", ...GEOM, align: "center", text: "Confidential" };

// --- slide number (a:fld slidenum) ------------------------------------------
export const GOLDEN_SLIDENUM_SP =
  `<p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder"/>` +
  `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr>${XFRM}</p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/>` +
  `<a:fld id="{CT5093A1-0000-4000-8000-0000000000EE}" type="slidenum"><a:t>3</a:t></a:fld></a:p></p:txBody></p:sp>`;
export const GOLDEN_SLIDENUM_RF: ResolvedField = { field: { kind: "slideNumber", ...GEOM, align: "right" }, text: "3", fldType: "slidenum" };
export const GOLDEN_SLIDENUM_ENTRY: SlideFieldPlaceholder = { kind: "slideNumber", ...GEOM, align: "right" };

// --- auto date (a:fld datetime3) --------------------------------------------
export const GOLDEN_DATE_SP =
  `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Date Placeholder"/>` +
  `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="dt" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr>${XFRM}</p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/>` +
  `<a:fld id="{CT5093A2-0000-4000-8000-0000000000EE}" type="datetime3"><a:t>14 May 2026</a:t></a:fld></a:p></p:txBody></p:sp>`;
export const GOLDEN_DATE_RF: ResolvedField = { field: { kind: "date", ...GEOM }, text: "14 May 2026", fldType: "datetime3" };
/** An in-subset auto date reads back with no wire/text (recomputed live). */
export const GOLDEN_DATE_ENTRY: SlideFieldPlaceholder = { kind: "date", ...GEOM };

/** A PowerPoint-flavoured footer variant: reordered `p:ph` attributes and an
 * `a:rPr` carrying extra attributes the reader ignores — must read identically. */
export const GOLDEN_FOOTER_SP_REORDERED =
  `<p:sp><p:nvSpPr><p:cNvPr name="Footer Placeholder" id="5"/>` +
  `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph idx="10" type="ftr" sz="quarter"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr>${XFRM}</p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
  `<a:r><a:rPr lang="en-GB" dirty="0"/><a:t>Confidential</a:t></a:r></a:p></p:txBody></p:sp>`;
