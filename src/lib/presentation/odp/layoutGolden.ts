// The NORMATIVE canonical odp layout markup (SPEC-089), hand-authored against
// ODF 1.2 §9.6.2 the way animGolden.ts authors ANIM_GOLDEN. The ODP
// `presentation:presentation-page-layout` vocabulary is not confirmed from a real
// file in this repo, so this golden is the source of truth: the reader is
// asserted to decode LAYOUT_GOLDEN to GOLDEN_LAYOUTS, and the styles.xml splice
// (`authorPageLayoutStyle`) is asserted to reproduce the standard entries'
// markup. Cross-application Impress verification is recorded in the PR, not CI.
import type { SlideLayout } from "../model";

/** The master page these golden layouts hang off. */
export const GOLDEN_MASTER = "Default";

/** A "Title" page layout: a title and a subtitle placeholder. */
export const GOLDEN_TITLE_MARKUP =
  `<style:presentation-page-layout style:name="AL-title">` +
  `<presentation:placeholder presentation:object="title" svg:x="2cm" svg:y="6cm" svg:width="29cm" svg:height="3cm"/>` +
  `<presentation:placeholder presentation:object="subtitle" svg:x="2cm" svg:y="10cm" svg:width="29cm" svg:height="2cm"/>` +
  `</style:presentation-page-layout>`;

/** A "Two Content" page layout: a title and two side-by-side outline bodies. */
export const GOLDEN_TWO_MARKUP =
  `<style:presentation-page-layout style:name="AL-two">` +
  `<presentation:placeholder presentation:object="title" svg:x="2cm" svg:y="1cm" svg:width="29cm" svg:height="3cm"/>` +
  `<presentation:placeholder presentation:object="outline" svg:x="2cm" svg:y="5cm" svg:width="13cm" svg:height="12cm"/>` +
  `<presentation:placeholder presentation:object="outline" svg:x="17cm" svg:y="5cm" svg:width="13cm" svg:height="12cm"/>` +
  `</style:presentation-page-layout>`;

/** An exotic layout whose second placeholder is a `graphic` object — outside the
 * recognised role vocabulary, so it classifies `"other"` (pins the fallback). */
export const GOLDEN_EXOTIC_MARKUP =
  `<style:presentation-page-layout style:name="AL-exotic">` +
  `<presentation:placeholder presentation:object="title" svg:x="2cm" svg:y="1cm" svg:width="29cm" svg:height="3cm"/>` +
  `<presentation:placeholder presentation:object="graphic" svg:x="2cm" svg:y="5cm" svg:width="29cm" svg:height="12cm"/>` +
  `</style:presentation-page-layout>`;

/** All three page-layout styles, the inner content of an `<office:styles>`. */
export const LAYOUT_GOLDEN = GOLDEN_TITLE_MARKUP + GOLDEN_TWO_MARKUP + GOLDEN_EXOTIC_MARKUP;

/** The page-layout names, in golden order (the order pages reference them). */
export const GOLDEN_PAGE_LAYOUT_NAMES = ["AL-title", "AL-two", "AL-exotic"] as const;

/** The `SlideLayout[]` the reader decodes LAYOUT_GOLDEN to, keyed
 * `${master}::${pageLayout}` — hand-authored against the markup above. ODF has
 * no `p:ph idx`, so slots carry no `phType`/`phIdx`. */
export const GOLDEN_LAYOUTS: SlideLayout[] = [
  {
    key: "Default::AL-title",
    kind: "title",
    name: "AL-title",
    masterName: "Default",
    slots: [
      { role: "title", xCm: 2, yCm: 6, widthCm: 29, heightCm: 3 },
      { role: "subtitle", xCm: 2, yCm: 10, widthCm: 29, heightCm: 2 },
    ],
  },
  {
    key: "Default::AL-two",
    kind: "twoContent",
    name: "AL-two",
    masterName: "Default",
    slots: [
      { role: "title", xCm: 2, yCm: 1, widthCm: 29, heightCm: 3 },
      { role: "body", xCm: 2, yCm: 5, widthCm: 13, heightCm: 12 },
      { role: "body", xCm: 17, yCm: 5, widthCm: 13, heightCm: 12 },
    ],
  },
  {
    key: "Default::AL-exotic",
    kind: "other",
    name: "AL-exotic",
    masterName: "Default",
    slots: [
      { role: "title", xCm: 2, yCm: 1, widthCm: 29, heightCm: 3 },
      { role: "other", xCm: 2, yCm: 5, widthCm: 29, heightCm: 12 },
    ],
  },
];
