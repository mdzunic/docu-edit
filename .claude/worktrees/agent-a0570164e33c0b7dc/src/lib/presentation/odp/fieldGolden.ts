// The NORMATIVE canonical odp field declarations (SPEC-093), hand-authored
// against the ODF 1.2 presentation schema: a footer declaration, an auto
// date-time declaration referencing the catalogue's `d MMMM yyyy` data style
// (`FieldDT0`), and that `number:date-style` itself. odp/serialize.spec.ts pins
// `planOdpDeclarations` to these bytes, and the reader decodes them (plus an
// Impress-flavoured reordered variant) back to the deck values. Cross-app
// verification is manual QA in the PR (AC-131).
import type { PresentationModel } from "../model";

/** The footer-decl + date-time-decl block, in first-use order. */
export const GOLDEN_ODP_DECLS =
  `<presentation:footer-decl presentation:name="ftr1">Confidential</presentation:footer-decl>` +
  `<presentation:date-time-decl presentation:name="dtd1" presentation:source="current-date" style:data-style-name="FieldDT0"/>`;

/** The `number:date-style` the auto date's `d MMMM yyyy` needs (catalogue #0). */
export const GOLDEN_ODP_DATE_STYLE =
  `<number:date-style style:name="FieldDT0">` +
  `<number:day/><number:text> </number:text>` +
  `<number:month number:textual="true" number:style="long"/><number:text> </number:text>` +
  `<number:year number:style="long"/></number:date-style>`;

/** The deck the declarations encode: a footer, an auto date and a slide number
 * on the single body slide (the title slide index 0 is suppressed). */
export function goldenOdpModel(): PresentationModel {
  return {
    kind: "presentation",
    id: "golden",
    title: "Golden",
    format: "odp",
    slideWidthCm: 25.4,
    slideHeightCm: 19.05,
    headerFooter: { footerText: "Confidential" },
    headerFooterDirty: true,
    slides: [
      { id: "s0", frames: [] },
      {
        id: "s1",
        frames: [],
        fields: [
          { kind: "footer", xCm: 2, yCm: 18, widthCm: 8, heightCm: 1, align: "center" },
          { kind: "slideNumber", xCm: 15, yCm: 18, widthCm: 8, heightCm: 1, align: "right" },
          { kind: "date", xCm: 2, yCm: 18, widthCm: 8, heightCm: 1 },
        ],
      },
    ],
  };
}

/** An Impress-flavoured reordered declaration block (attributes swapped, an
 * extra `presentation:use-header-name`-style whitespace) — must read identically. */
export const GOLDEN_ODP_DECLS_REORDERED =
  `<presentation:footer-decl presentation:name="ftr1" >Confidential</presentation:footer-decl>` +
  `<presentation:date-time-decl style:data-style-name="FieldDT0" presentation:source="current-date" presentation:name="dtd1"/>`;
