// The Page setup dialog's Borders tab state (SPEC-074 AC-19..AC-23), kept
// out of the `.vue` for the same two reasons the paragraph draft is
// (lib/paragraphBorders.ts): espree-parsed `.vue` scripts cannot carry the
// annotations, and a pure draft is testable without mounting a dialog.
import { borderSideVisible, type BorderLineStyle, type BorderSide } from "./core/borders";
import type { PageBorderAppliesTo, PageBorderMeasureFrom, PageBorders } from "./documentModel";

/** The four edges a page border can draw on — art borders are out of scope
 * (SPEC-074), so there is no fifth. */
export const PAGE_BORDER_EDGES = ["top", "bottom", "left", "right"] as const;

export type PageBorderEdge = (typeof PAGE_BORDER_EDGES)[number];

/** Word's own default distance from the page edge for a new page border. */
export const DEFAULT_PAGE_BORDER_OFFSET_PT = 24;

const DEFAULT_LINE: BorderSide = { style: "single", widthPt: 1, color: "#000000" };

/** The Borders tab's editable state: one line shared by every edge (AC-19 —
 * the tab offers a single style/width/colour, and which edges take it,
 * AC-20), the offset and what it measures from (AC-21), and the page
 * background (AC-23). */
export interface PageBorderDraft {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
  style: BorderLineStyle;
  widthPt: number;
  color: string;
  offsetPt: number;
  measureFrom: PageBorderMeasureFrom;
  appliesTo: PageBorderAppliesTo;
  backgroundOn: boolean;
  background: string;
}

/** The dialog's opening state for a section. The shared line comes from the
 * first edge that has one — every edge a document written by this product
 * shares a line, and a foreign file with differing edges is shown (and
 * re-saved) with the first edge's line, which is what a single-line tab can
 * express. */
export function pageBorderDraftFrom(
  borders: PageBorders | undefined,
  background: string | undefined,
): PageBorderDraft {
  const line = PAGE_BORDER_EDGES.map((edge) => borders?.[edge]).find((side) => borderSideVisible(side));
  return {
    top: borderSideVisible(borders?.top),
    bottom: borderSideVisible(borders?.bottom),
    left: borderSideVisible(borders?.left),
    right: borderSideVisible(borders?.right),
    style: line ? line.style : DEFAULT_LINE.style,
    widthPt: line ? line.widthPt : DEFAULT_LINE.widthPt,
    color: line && line.color !== "auto" ? line.color : DEFAULT_LINE.color,
    offsetPt: borders?.offsetPt ?? DEFAULT_PAGE_BORDER_OFFSET_PT,
    measureFrom: borders?.measureFrom ?? "page",
    appliesTo: borders?.appliesTo ?? "allPages",
    backgroundOn: typeof background === "string",
    background: typeof background === "string" ? background : "#fdf6e3",
  };
}

/** The draft as a `PageBorders`, or null when no edge is chosen — the
 * `null`-removes value setPageBorders takes. */
export function pageBordersFromDraft(draft: PageBorderDraft): PageBorders | null {
  const line: BorderSide = { style: draft.style, widthPt: draft.widthPt, color: draft.color };
  const borders: PageBorders = {};
  let any = false;
  for (const edge of PAGE_BORDER_EDGES) {
    if (draft[edge]) {
      borders[edge] = { ...line };
      any = true;
    }
  }
  if (!any) {
    return null;
  }
  borders.offsetPt = draft.offsetPt;
  borders.measureFrom = draft.measureFrom;
  if (draft.appliesTo !== "allPages") {
    borders.appliesTo = draft.appliesTo;
  }
  return borders;
}

/** The draft's page background, or null for paper white (AC-23). */
export function pageBackgroundFromDraft(draft: PageBorderDraft): string | null {
  return draft.backgroundOn ? draft.background : null;
}

/** Typed empty seed for a `.vue` ref holding the draft. */
export function emptyPageBorderDraft(): PageBorderDraft {
  return pageBorderDraftFrom(undefined, undefined);
}

/** Typed pass-through for an edge key coming from a `.vue` template's plain
 * string (espree bars the annotation — see documentModel.asTabStopAlignment). */
export function asPageBorderEdge(value: string): PageBorderEdge {
  return value as PageBorderEdge;
}

/** The checked state of an event target, for a `.vue` template that cannot
 * narrow `$event.target` itself (espree — see AppShell.vue). */
export function eventTargetChecked(target: EventTarget | null): boolean {
  return target instanceof HTMLInputElement && target.checked;
}

/** The value of a `<select>`/`<input>` event target, likewise. */
export function eventTargetValue(target: EventTarget | null): string {
  return target instanceof HTMLSelectElement || target instanceof HTMLInputElement ? target.value : "";
}

/** Typed `null` seed for a `PageBorders` field in an espree-parsed `.vue`
 * script or a composable's default payload (see commands.noSelection). */
export function noPageBorders(): PageBorders | null {
  return null;
}

/** Typed `null` seed for the page background, likewise. */
export function noPageBackground(): string | null {
  return null;
}
