// The NORMATIVE canonical odp wire form of a hidden slide (SPEC-104), pinned
// golden-first per SPEC-086's process rather than trusted: ODF 1.2 §15.36
// puts `presentation:visibility` (`visible` | `hidden`) on the page's
// `style:drawing-page-properties`, NOT on `draw:page`, so a hidden page
// references a minted `de-DP{n}` drawing-page style carrying it. The
// serialiser must emit exactly these bytes for a hidden from-scratch page
// (odp/serialize.spec.ts), and the reader must accept them plus an
// Impress-flavoured variant whose properties carry other attributes around
// it (hiddenRoundTrip.spec.ts). Cross-app verification in Impress is manual
// QA recorded in the PR, not CI.

/** The drawing-page property a hidden page's style carries. */
export const ODP_HIDDEN_VISIBILITY_PROP = `presentation:visibility="hidden"`;

/** The automatic style a hidden from-scratch page (no transition, no field
 * display flags) is given — the first one minted in content.xml. */
export const GOLDEN_ODP_HIDDEN_STYLE =
  `<style:style style:name="de-DP1" style:family="drawing-page">` +
  `<style:drawing-page-properties presentation:visibility="hidden"/></style:style>`;

/** The `draw:page` start tag of that page: its own attributes, then the
 * appended `draw:style-name` referencing the golden style. */
export const GOLDEN_ODP_HIDDEN_PAGE_ATTR = ` draw:style-name="de-DP1"`;

/** How Impress itself writes a hidden page's style: other drawing-page
 * properties around the visibility flag (the reader must not care). */
export const IMPRESS_ODP_HIDDEN_STYLE =
  `<style:style style:family="drawing-page" style:name="dp2">` +
  `<style:drawing-page-properties presentation:background-visible="true" ` +
  `presentation:background-objects-visible="true" presentation:visibility="hidden" ` +
  `presentation:display-footer="true"/></style:style>`;
