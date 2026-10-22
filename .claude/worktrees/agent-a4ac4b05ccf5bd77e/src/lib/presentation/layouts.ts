// The layout catalogue's rules (SPEC-089), framework-free like every other
// presentation/ module (presentation siblings + lib/core only): the standard
// layout templates, classification of a source layout into a `SlideLayoutKind`,
// resolving a slide's layout, and the apply/reset placeholder re-mapping. The
// TYPES (`SlideLayout`, `LayoutSlot`, `SlideLayoutKind`) live in model.ts beside
// `SlideTransition`/`SlideAnimation`; the RULES live here.
//
// "Apply" re-maps a slide's placeholders onto a layout's slots by role and idx,
// keeping their text — the word processor's paragraph-style apply one level up.
// "Reset" snaps every matched placeholder back to its slot geometry. Neither
// deletes anything: unmatched placeholders and non-placeholder frames stay
// exactly where they are and are reported as kept in place.
import { markSlideEdited } from "./text";
import {
  phKey,
  type LayoutSlot,
  type PlaceholderRole,
  type PresentationModel,
  type Slide,
  type SlideLayout,
  type SlideLayoutKind,
  type TextFrame,
} from "./model";

/** cm rounded to 3 decimals — matches `slides.ts`'s `r3` so a synthesised
 * layout's geometry is byte-identical to the from-scratch placeholders. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** The order role-fallback matching walks placeholders in (Reference). */
const ROLE_ORDER: PlaceholderRole[] = ["title", "subtitle", "body", "other"];

/** One standard-layout template: fractional slots the synthesiser instantiates
 * at a deck's slide size, plus the pptx `p:sldLayout/@type` the writer authors
 * it under. Equal to `slides.ts`'s existing placeholder fractions where the
 * kinds overlap. */
export interface StandardLayoutDef {
  kind: SlideLayoutKind;
  /** The name the gallery card shows and the writer stamps as `p:cSld/@name`. */
  name: string;
  /** The `p:sldLayout/@type` the pptx writer authors this layout under. */
  pptxType: string;
  slots: Array<{
    role: PlaceholderRole;
    phType?: string;
    phIdx?: number;
    xFrac: number;
    yFrac: number;
    wFrac: number;
    hFrac: number;
  }>;
}

/** The minimal standard set (Reference table). Order is the from-scratch
 * `slideLayout1..6` order and the gallery's fallback order. */
export const STANDARD_LAYOUTS: StandardLayoutDef[] = [
  {
    kind: "title",
    name: "Title Slide",
    pptxType: "title",
    slots: [
      { role: "title", phType: "ctrTitle", xFrac: 0.08, yFrac: 0.32, wFrac: 0.84, hFrac: 0.18 },
      { role: "subtitle", phType: "subTitle", phIdx: 1, xFrac: 0.08, yFrac: 0.52, wFrac: 0.84, hFrac: 0.12 },
    ],
  },
  {
    kind: "titleAndBody",
    name: "Title and Content",
    pptxType: "obj",
    slots: [
      { role: "title", phType: "title", xFrac: 0.08, yFrac: 0.08, wFrac: 0.84, hFrac: 0.16 },
      { role: "body", phType: "body", phIdx: 1, xFrac: 0.08, yFrac: 0.3, wFrac: 0.84, hFrac: 0.6 },
    ],
  },
  {
    kind: "sectionHeader",
    name: "Section Header",
    pptxType: "secHead",
    slots: [
      { role: "title", phType: "title", xFrac: 0.08, yFrac: 0.42, wFrac: 0.84, hFrac: 0.2 },
      { role: "body", phType: "body", phIdx: 1, xFrac: 0.08, yFrac: 0.64, wFrac: 0.84, hFrac: 0.12 },
    ],
  },
  {
    kind: "twoContent",
    name: "Two Content",
    pptxType: "twoObj",
    slots: [
      { role: "title", phType: "title", xFrac: 0.08, yFrac: 0.08, wFrac: 0.84, hFrac: 0.16 },
      { role: "body", phType: "body", phIdx: 1, xFrac: 0.08, yFrac: 0.3, wFrac: 0.41, hFrac: 0.6 },
      { role: "body", phType: "body", phIdx: 2, xFrac: 0.51, yFrac: 0.3, wFrac: 0.41, hFrac: 0.6 },
    ],
  },
  {
    kind: "titleOnly",
    name: "Title Only",
    pptxType: "titleOnly",
    slots: [{ role: "title", phType: "title", xFrac: 0.08, yFrac: 0.08, wFrac: 0.84, hFrac: 0.16 }],
  },
  { kind: "blank", name: "Blank", pptxType: "blank", slots: [] },
];

/** The stable key a synthesised standard layout carries. */
export function standardLayoutKey(kind: SlideLayoutKind): string {
  return `standard:${kind}`;
}

/** Instantiates a standard-layout template at a deck's slide size into a real
 * `SlideLayout` keyed `standard:{kind}` — the from-scratch parts and the
 * gallery's synthesised fallbacks (AC-5/17). */
export function instantiateStandardLayout(def: StandardLayoutDef, widthCm: number, heightCm: number): SlideLayout {
  return {
    key: standardLayoutKey(def.kind),
    kind: def.kind,
    name: def.name,
    slots: def.slots.map((s) => {
      const slot: LayoutSlot = {
        role: s.role,
        xCm: r3(s.xFrac * widthCm),
        yCm: r3(s.yFrac * heightCm),
        widthCm: r3(s.wFrac * widthCm),
        heightCm: r3(s.hFrac * heightCm),
      };
      if (s.phType !== undefined) slot.phType = s.phType;
      if (s.phIdx !== undefined) slot.phIdx = s.phIdx;
      return slot;
    }),
  };
}

/** The full synthesised standard set at a deck's slide size. */
export function synthesiseStandardLayouts(widthCm: number, heightCm: number): SlideLayout[] {
  return STANDARD_LAYOUTS.map((def) => instantiateStandardLayout(def, widthCm, heightCm));
}

/** The pptx `p:sldLayout/@type` → kind mapping (Reference). An absent type is
 * classified from the slot roles instead (`classifyFromSlots`). */
export function classifyLayout(type: string | undefined, slots: LayoutSlot[]): SlideLayoutKind {
  switch (type) {
    case "title":
      return "title";
    case "obj":
    case "tx":
      return "titleAndBody";
    case "secHead":
      return "sectionHeader";
    case "twoObj":
    case "twoTxTwoObj":
      return "twoContent";
    case "titleOnly":
      return "titleOnly";
    case "blank":
      return "blank";
    case undefined:
      return classifyFromSlots(slots);
    default:
      return "other";
  }
}

/** Classifies a layout with no declared type (odp, and pptx layouts missing
 * `@type`) from the roles its slots carry. */
export function classifyFromSlots(slots: LayoutSlot[]): SlideLayoutKind {
  if (slots.length === 0) {
    return "blank";
  }
  const titles = slots.filter((s) => s.role === "title").length;
  const subtitles = slots.filter((s) => s.role === "subtitle").length;
  const bodies = slots.filter((s) => s.role === "body").length;
  const others = slots.filter((s) => s.role === "other").length;
  if (others > 0) {
    return "other";
  }
  if (titles === 1 && subtitles >= 1 && bodies === 0) {
    return "title";
  }
  if (titles === 1 && bodies >= 2) {
    return "twoContent";
  }
  if (titles === 1 && bodies === 1) {
    return "titleAndBody";
  }
  if (titles === 1 && bodies === 0 && subtitles === 0) {
    return "titleOnly";
  }
  return "other";
}

/** The catalogue entry a slide uses, or undefined when its `layoutKey` names no
 * entry (or is unset). Never repairs the key. */
export function layoutFor(model: PresentationModel, slide: Slide): SlideLayout | undefined {
  if (!slide.layoutKey) {
    return undefined;
  }
  return (model.layouts ?? []).find((l) => l.key === slide.layoutKey);
}

/** The result of an apply/reset: whether it ran, how many placeholders it
 * placed onto slots, how many empty placeholders it minted, and the ids of the
 * frames left exactly where they were (for the "N items kept in place" report).
 */
export interface LayoutApplyReport {
  applied: boolean;
  placed: number;
  created: number;
  keptInPlace: string[];
}

/** The (type#idx) key a text frame matches a slot by. */
function frameKey(frame: TextFrame): string {
  return phKey(frame.phType, frame.phIdx === undefined ? undefined : String(frame.phIdx));
}

function slotKey(slot: LayoutSlot): string {
  return phKey(slot.phType, slot.phIdx === undefined ? undefined : String(slot.phIdx));
}

/** Moves a matched frame onto a slot: geometry adopted, `phType`/`phIdx` from
 * the slot (verbatim write-back), role set. Text, fill, stroke and rotation are
 * left untouched. */
function adoptSlot(frame: TextFrame, slot: LayoutSlot): void {
  frame.xCm = slot.xCm;
  frame.yCm = slot.yCm;
  frame.widthCm = slot.widthCm;
  frame.heightCm = slot.heightCm;
  frame.role = slot.role;
  if (slot.phType !== undefined) {
    frame.phType = slot.phType;
  } else {
    delete frame.phType;
  }
  if (slot.phIdx !== undefined) {
    frame.phIdx = slot.phIdx;
  } else {
    delete frame.phIdx;
  }
}

/** An empty placeholder `TextFrame` minted for a slot with no matching frame. */
function mintSlotFrame(slot: LayoutSlot): TextFrame {
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    xCm: slot.xCm,
    yCm: slot.yCm,
    widthCm: slot.widthCm,
    heightCm: slot.heightCm,
    kind: "text",
    role: slot.role,
    paragraphs: [],
  };
  if (slot.phType !== undefined) frame.phType = slot.phType;
  if (slot.phIdx !== undefined) frame.phIdx = slot.phIdx;
  return frame;
}

/** One empty placeholder `TextFrame` per slot of a layout, in slot order —
 * the frames a `New slide with layout…` slide is born with (AC-8). */
export function newSlideFramesForLayout(layout: SlideLayout): TextFrame[] {
  return layout.slots.map((slot) => mintSlotFrame(slot));
}

/** The candidate placeholder frames of a slide, in document order. */
function placeholderFrames(slide: Slide): TextFrame[] {
  return slide.frames.filter((f): f is TextFrame => f.kind === "text" && f.role !== undefined);
}

/** Three-pass match of a layout's slots against a slide's placeholder frames:
 *  (1) exact phKey; (2) same role in document order; (3) any leftover in role
 *  order title→subtitle→body→other. Each phase runs across all still-unmatched
 *  slots so an exact match is never stolen by an earlier slot's role fallback.
 *  Returns the frame matched per slot index (undefined = mint one). */
function matchSlots(slots: LayoutSlot[], frames: TextFrame[]): (TextFrame | undefined)[] {
  const matched: (TextFrame | undefined)[] = slots.map(() => undefined);
  const consumed = new Set<string>();

  // Pass 1: exact phKey.
  slots.forEach((slot, i) => {
    const key = slotKey(slot);
    const hit = frames.find((f) => !consumed.has(f.id) && frameKey(f) === key);
    if (hit) {
      matched[i] = hit;
      consumed.add(hit.id);
    }
  });
  // Pass 2: same role, document order.
  slots.forEach((slot, i) => {
    if (matched[i]) return;
    const hit = frames.find((f) => !consumed.has(f.id) && f.role === slot.role);
    if (hit) {
      matched[i] = hit;
      consumed.add(hit.id);
    }
  });
  // Pass 3: any leftover, role priority order.
  const leftover = () =>
    ROLE_ORDER.flatMap((role) => frames.filter((f) => !consumed.has(f.id) && f.role === role));
  slots.forEach((slot, i) => {
    if (matched[i]) return;
    void slot;
    const hit = leftover()[0];
    if (hit) {
      matched[i] = hit;
      consumed.add(hit.id);
    }
  });
  return matched;
}

/** The index new placeholder frames splice in at: just before the first
 * non-placeholder frame, so a picture is never covered. Appends when the slide
 * has only placeholders. */
function insertionIndex(slide: Slide): number {
  const i = slide.frames.findIndex((f) => !(f.kind === "text" && f.role !== undefined));
  return i === -1 ? slide.frames.length : i;
}

/** Applies `layoutKey`'s layout to slide `slideIndex`: re-maps placeholders onto
 * its slots (three-pass), mints an empty placeholder for every unmatched slot,
 * sets `slide.layoutKey`, and marks the slide edited. Non-placeholder frames and
 * unmatched placeholders are kept exactly in place. (AC-6) */
export function applySlideLayout(model: PresentationModel, slideIndex: number, layoutKey: string): LayoutApplyReport {
  const slide = model.slides[slideIndex];
  const layout = (model.layouts ?? []).find((l) => l.key === layoutKey);
  if (!slide || !layout) {
    return { applied: false, placed: 0, created: 0, keptInPlace: [] };
  }
  const frames = placeholderFrames(slide);
  const matched = matchSlots(layout.slots, frames);
  const consumed = new Set<string>();

  layout.slots.forEach((slot, i) => {
    const frame = matched[i];
    if (frame) {
      adoptSlot(frame, slot);
      consumed.add(frame.id);
    }
  });
  // Mint an empty placeholder for each unmatched slot, spliced behind
  // non-placeholder frames (slot order preserved).
  const created: TextFrame[] = [];
  layout.slots.forEach((slot, i) => {
    if (!matched[i]) {
      created.push(mintSlotFrame(slot));
    }
  });
  if (created.length > 0) {
    slide.frames.splice(insertionIndex(slide), 0, ...created);
  }

  // Track created frames by id, not object identity: a reactive slide wraps the
  // spliced frames in proxies, so `.includes(rawFrame)` would miss them.
  const createdIds = new Set(created.map((f) => f.id));
  const keptInPlace = slide.frames
    .filter((f) => !consumed.has(f.id) && !createdIds.has(f.id))
    .map((f) => f.id);
  slide.layoutKey = layoutKey;
  markSlideEdited(slide, model.format);
  return { applied: true, placed: consumed.size, created: created.length, keptInPlace };
}

/** Snaps slide `slideIndex`'s placeholders back to its resolved layout's slot
 * geometry (exact-phKey pass only, nothing minted), leaving all else untouched.
 * (AC-7) */
export function resetSlideLayout(model: PresentationModel, slideIndex: number): LayoutApplyReport {
  const slide = model.slides[slideIndex];
  const layout = slide ? layoutFor(model, slide) : undefined;
  if (!slide || !layout) {
    return { applied: false, placed: 0, created: 0, keptInPlace: [] };
  }
  const frames = placeholderFrames(slide);
  const consumed = new Set<string>();
  layout.slots.forEach((slot) => {
    const key = slotKey(slot);
    const hit = frames.find((f) => !consumed.has(f.id) && frameKey(f) === key);
    if (hit) {
      hit.xCm = slot.xCm;
      hit.yCm = slot.yCm;
      hit.widthCm = slot.widthCm;
      hit.heightCm = slot.heightCm;
      consumed.add(hit.id);
    }
  });
  const keptInPlace = slide.frames.filter((f) => !consumed.has(f.id)).map((f) => f.id);
  markSlideEdited(slide, model.format);
  return { applied: true, placed: consumed.size, created: 0, keptInPlace };
}
