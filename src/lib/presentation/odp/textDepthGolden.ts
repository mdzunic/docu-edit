// SPEC-097's golden-fixture-first record (SPEC-086's process) for the three
// ODF text-depth details that are NOT settled from memory: the
// shrink-on-overflow attribute, the numbering-restart attribute and the
// slide-jump link form. Each is PENDING — no golden has yet been hand-authored
// against ODF 1.2 §15.5/§16.31 and verified by opening it in Impress — so,
// never inventing bytes:
//
//  - the odp writer OMITS that one property (autofit "shrink" writes no shrink
//    attribute; a start-at that restarts numbering mid-run writes nothing —
//    `text:start-value` on the NUMBER LEVEL STYLE is the settled level-style
//    attribute and is still written; a `linkSlide` writes no link at all). The
//    field still round-trips on the pptx side.
//  - the odp reader PRESERVES any such source markup: both shrink candidates
//    ride `odpBodyProps`; an in-deck `#…` `text:a` keeps today's plain-text
//    reading; a `draw:a` to `#…` stays a verbatim foreign frame. A per-item
//    `text:start-value` is read tolerantly as that item's `startAt` (it is
//    unambiguous to read; only which form to WRITE is open).
//
// When a golden lands: add its exact bytes here (as the other `*Golden.ts`
// modules do), assert them byte-exact in both directions in
// textDepthGolden.spec.ts, record the Impress verification in the PR (manual
// QA, not CI), switch the entry's `status` and teach serialize.ts to write it.

/** The state of one golden-pending ODF detail. */
export type OdpGoldenStatus = "pending";

/** One unsettled ODF wire detail and what both directions do meanwhile. */
export interface PendingOdpGolden {
  status: OdpGoldenStatus;
  /** The model field the detail would carry. */
  models: string;
  /** The two candidate wire forms the golden must choose between. */
  candidates: readonly [string, string];
  /** The cross-app check the golden still needs. */
  verification: "Impress";
  /** What serialize.ts emits until the golden exists. */
  writer: string;
  /** What reader.ts does with source markup in either candidate form. */
  reader: string;
}

/** The shrink-on-overflow attribute on a frame's `style:graphic-properties`. */
export const ODP_SHRINK_ON_OVERFLOW: PendingOdpGolden = {
  status: "pending",
  models: 'TextBodyProps.autofit = "shrink" (with fontScalePct/lineScalePct)',
  candidates: ['draw:fit-to-size="shrink-to-fit"', 'style:shrink-to-fit="true"'],
  verification: "Impress",
  writer: "omitted — no shrink attribute is written; a shrink-only body mints no graphic style",
  reader: "either candidate is preserved verbatim in the frame's odpBodyProps; autofit stays unmodelled",
};

/** Restarting a list's numbering at an item that would otherwise continue. */
export const ODP_NUMBERING_RESTART: PendingOdpGolden = {
  status: "pending",
  models: "SlideParagraph.startAt on an item that continues its level's run",
  candidates: ['text:continue-numbering="false" on the text:list', "text:start-value on the text:list-item"],
  verification: "Impress",
  writer:
    "omitted — only the number level style's text:start-value (the settled level-style attribute) is written, from the level's first item",
  reader: "a per-item text:start-value reads as that item's startAt; text:continue-numbering is ignored (numbering is derived by adjacency)",
};

/** A hyperlink to another slide of the deck. */
export const ODP_SLIDE_JUMP_LINK: PendingOdpGolden = {
  status: "pending",
  models: "SlideRun.linkSlide / SlideFrameBase.linkSlide",
  candidates: ['xlink:href="#{draw:name}" on the text:a / draw:a', 'a presentation:event-listener presentation:action="jump" child'],
  verification: "Impress",
  writer: "omitted — a linkSlide (and any in-deck #… href) writes no text:a / draw:a at all",
  reader: "a #… text:a keeps the plain-text reading; a #… draw:a stays a verbatim foreign frame",
};

/** Every golden-pending SPEC-097 ODF detail, for the pinning spec. */
export const ODP_TEXT_DEPTH_PENDING: readonly PendingOdpGolden[] = [
  ODP_SHRINK_ON_OVERFLOW,
  ODP_NUMBERING_RESTART,
  ODP_SLIDE_JUMP_LINK,
];
