// The slide-transition wire mappings (SPEC-085) — the NORMATIVE source for how
// the modelled `SlideTransition` subset (six effects + explicit none) reads
// from and writes to the two formats' markup. Both serialisers and SPEC-088's
// HTML export import these tables so the mapping lives in exactly one place.
//
// Framework-free (imports only the model types); the DOM-reading helpers take a
// parsed element/attribute bag the readers already hold, so this module stays
// unit-testable without mounting anything. Anything the tables do not recognise
// is "out of subset": the reader leaves `slide.transition` unset and keeps the
// source markup preserved verbatim (SPEC-085 AC-3/AC-12).
import type { SlideTransition, TransitionDirection, TransitionType } from "./model";

export type { SlideTransition, TransitionDirection, TransitionType } from "./model";

export const DEFAULT_TRANSITION_DURATION_MS = 600;
export const MIN_TRANSITION_DURATION_MS = 100;
export const MAX_TRANSITION_DURATION_MS = 5000;

/** The six effects, in panel order (None is prepended by the UI). */
export const TRANSITION_TYPES: TransitionType[] = ["fade", "push", "wipe", "cover", "uncover", "split"];

/** Effects whose direction the user can choose. fade/split/none ignore it
 * (split is always horizontal-out). `type` is widened to `string` so untyped
 * .vue reactive values can call it. */
export function usesDirection(type: string): boolean {
  return type === "push" || type === "wipe" || type === "cover" || type === "uncover";
}

/** The default direction for a directional effect (all default to left);
 * undefined for the effects that ignore direction. */
export function defaultDirectionFor(type: string): TransitionDirection | undefined {
  return usesDirection(type) ? "left" : undefined;
}

/** The direction a transition effectively animates with — the model's own, or
 * the per-type default. */
export function effectiveDirection(t: SlideTransition): TransitionDirection | undefined {
  return t.direction ?? defaultDirectionFor(t.type);
}

/** The transition's duration in ms, defaulting and clamping to the model bounds. */
export function effectiveDurationMs(t: SlideTransition): number {
  const ms = t.durationMs ?? DEFAULT_TRANSITION_DURATION_MS;
  return Math.min(MAX_TRANSITION_DURATION_MS, Math.max(MIN_TRANSITION_DURATION_MS, ms));
}

// ---------------------------------------------------------------------------
// pptx (PresentationML)
// ---------------------------------------------------------------------------

/** `p:transition` effect child element → model type. `p:pull` is uncover;
 * `p:cut` (and an absent effect child) is the explicit hard cut. */
const PPTX_EFFECT_TO_TYPE: Record<string, TransitionType | "none"> = {
  fade: "fade",
  push: "push",
  wipe: "wipe",
  cover: "cover",
  pull: "uncover",
  split: "split",
  cut: "none",
};
const TYPE_TO_PPTX_EFFECT: Record<TransitionType, string> = {
  fade: "fade",
  push: "push",
  wipe: "wipe",
  cover: "cover",
  uncover: "pull",
  split: "split",
};
const DIR_TO_PPTX: Record<TransitionDirection, string> = { left: "l", right: "r", up: "u", down: "d" };
const PPTX_TO_DIR: Record<string, TransitionDirection> = { l: "left", r: "right", u: "up", d: "down" };
const PPTX_SPD_TO_MS: Record<string, number> = { slow: 1000, med: 600, fast: 300 };

/** The elements inside `p:transition` that are NOT the effect (a sound cue or
 * the extension list) — skipped when locating the effect child. */
const PPTX_NON_EFFECT = new Set(["sndAc", "extLst"]);

/** Parses a `p:transition` element to a model transition, or `undefined` when
 * the effect is outside the subset (→ keep it preserved). A `p:transition`
 * with no recognised effect child maps to explicit none. */
export function readPptxTransition(el: Element): SlideTransition | undefined {
  let effect: Element | undefined;
  for (const child of Array.from(el.children)) {
    if (PPTX_NON_EFFECT.has(child.localName)) {
      continue;
    }
    effect = child;
    break;
  }
  let type: TransitionType | "none";
  let direction: TransitionDirection | undefined;
  if (!effect) {
    type = "none";
  } else {
    const mapped = PPTX_EFFECT_TO_TYPE[effect.localName];
    if (mapped === undefined) {
      return undefined; // e.g. p:morph, p:zoom — out of subset
    }
    type = mapped;
    if (mapped !== "none" && usesDirection(mapped)) {
      const dir = effect.getAttribute("dir");
      if (dir && PPTX_TO_DIR[dir]) {
        direction = PPTX_TO_DIR[dir];
      }
    }
  }
  const transition: SlideTransition = { type };
  if (direction) {
    transition.direction = direction;
  }
  const durationMs = readPptxDurationMs(el);
  if (durationMs !== undefined) {
    transition.durationMs = durationMs;
  }
  const advClick = el.getAttribute("advClick");
  if (advClick === "0" || advClick === "false") {
    transition.advanceOnClick = false;
  }
  const advTm = el.getAttribute("advTm");
  if (advTm) {
    const n = Number(advTm);
    if (Number.isFinite(n) && n > 0) {
      transition.advanceAfterMs = n;
    }
  }
  return transition;
}

/** Duration in ms from a `p:transition`: a `p14:dur` extension (exact ms) wins,
 * else the `spd` bucket (slow/med/fast). Absent → undefined (model default). */
function readPptxDurationMs(el: Element): number | undefined {
  for (const desc of Array.from(el.getElementsByTagName("*"))) {
    if (desc.localName === "dur") {
      const val = Number(desc.getAttribute("val"));
      if (Number.isFinite(val) && val > 0) {
        return val;
      }
    }
  }
  const spd = el.getAttribute("spd");
  if (spd && PPTX_SPD_TO_MS[spd] !== undefined) {
    return PPTX_SPD_TO_MS[spd];
  }
  return undefined;
}

/** The `spd` bucket a duration quantises to on write: nearest of fast (≤450),
 * med (≤800), slow. The exact ms survives on the model until reopened — the
 * single sanctioned round-trip delta (AC-20). */
export function pptxSpeedFor(durationMs: number): "fast" | "med" | "slow" {
  if (durationMs <= 450) {
    return "fast";
  }
  if (durationMs <= 800) {
    return "med";
  }
  return "slow";
}

/** Serialises a model transition to its `p:transition` element (schema-ordered
 * to sit between `p:clrMapOvr` and `p:timing`). Writes `spd` only — never the
 * p14 extension (out of scope). */
export function writePptxTransition(t: SlideTransition): string {
  const attrs: string[] = [`spd="${pptxSpeedFor(effectiveDurationMs(t))}"`];
  if (t.advanceOnClick === false) {
    attrs.push(`advClick="0"`);
  }
  if (t.advanceAfterMs !== undefined && t.advanceAfterMs > 0) {
    attrs.push(`advTm="${Math.round(t.advanceAfterMs)}"`);
  }
  const inner = pptxEffectXml(t);
  return `<p:transition ${attrs.join(" ")}>${inner}</p:transition>`;
}

function pptxEffectXml(t: SlideTransition): string {
  if (t.type === "none") {
    return "";
  }
  if (t.type === "fade") {
    return `<p:fade/>`;
  }
  if (t.type === "split") {
    return `<p:split orient="horz" dir="out"/>`;
  }
  const element = TYPE_TO_PPTX_EFFECT[t.type];
  const dir = DIR_TO_PPTX[effectiveDirection(t) ?? "left"];
  return `<p:${element} dir="${dir}"/>`;
}

// ---------------------------------------------------------------------------
// odp (ODF drawing-page style)
// ---------------------------------------------------------------------------

/** The transition-relevant attributes read off a page's
 * `style:drawing-page-properties`. */
export interface OdpTransitionAttrs {
  smilType?: string;
  smilSubtype?: string;
  smilDirection?: string;
  smilDur?: string;
  transitionSpeed?: string;
  transitionType?: string;
  presentationDuration?: string;
}

const ODP_SPEED_TO_MS: Record<string, number> = { slow: 1000, medium: 600, fast: 300 };

/** Incoming-travels-`dir` enters from the opposite edge. */
const DIR_TO_FROM_EDGE: Record<TransitionDirection, string> = {
  left: "fromRight",
  right: "fromLeft",
  up: "fromBottom",
  down: "fromTop",
};
const FROM_EDGE_TO_DIR: Record<string, TransitionDirection> = {
  fromRight: "left",
  fromLeft: "right",
  fromBottom: "up",
  fromTop: "down",
};

/** Parses a page's drawing-page transition attributes to a model transition, or
 * `undefined` when the smil combination is outside the subset. Callers only
 * invoke this when `smilType` is present; absent smil = a plain slide. */
export function readOdpTransition(a: OdpTransitionAttrs): SlideTransition | undefined {
  const type = a.smilType;
  const subtype = a.smilSubtype ?? "";
  const reverse = a.smilDirection === "reverse";
  let mapped: { type: TransitionType | "none"; direction?: TransitionDirection } | undefined;

  if (type === "fade") {
    mapped = { type: "fade" };
  } else if (type === "pushWipe") {
    const direction = FROM_EDGE_TO_DIR[subtype];
    if (direction) {
      mapped = { type: "push", direction };
    }
  } else if (type === "barWipe") {
    if (subtype === "leftToRight") {
      mapped = { type: "wipe", direction: reverse ? "left" : "right" };
    } else if (subtype === "topToBottom") {
      mapped = { type: "wipe", direction: reverse ? "up" : "down" };
    }
  } else if (type === "slideWipe") {
    const direction = FROM_EDGE_TO_DIR[subtype];
    if (direction) {
      mapped = { type: reverse ? "uncover" : "cover", direction };
    }
  } else if (type === "barnDoorWipe") {
    mapped = { type: "split" };
  }
  if (!mapped) {
    return undefined;
  }

  const transition: SlideTransition = { type: mapped.type };
  if (mapped.direction) {
    transition.direction = mapped.direction;
  }
  const durationMs = readOdpDurationMs(a);
  if (durationMs !== undefined) {
    transition.durationMs = durationMs;
  }
  if (a.transitionType === "automatic") {
    const auto = parseOdfDurationToMs(a.presentationDuration);
    if (auto !== undefined && auto > 0) {
      transition.advanceAfterMs = auto;
    }
  }
  return transition;
}

function readOdpDurationMs(a: OdpTransitionAttrs): number | undefined {
  const dur = parseOdfDurationToMs(a.smilDur);
  if (dur !== undefined) {
    return dur;
  }
  if (a.transitionSpeed && ODP_SPEED_TO_MS[a.transitionSpeed] !== undefined) {
    return ODP_SPEED_TO_MS[a.transitionSpeed];
  }
  return undefined;
}

/** Parses an ODF duration (`PT0.6S`, `PT5S`) to milliseconds. */
export function parseOdfDurationToMs(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = /^PT([\d.]+)S$/.exec(value.trim());
  if (!match) {
    return undefined;
  }
  const seconds = Number(match[1]);
  return Number.isFinite(seconds) ? Math.round(seconds * 1000) : undefined;
}

/** Milliseconds as an ODF duration (`PT0.6S`). Trailing zeros trimmed. */
export function msToOdfDuration(ms: number): string {
  const seconds = ms / 1000;
  const text = Number.isInteger(seconds) ? String(seconds) : String(Number(seconds.toFixed(3)));
  return `PT${text}S`;
}

/** The smil/presentation attributes a minted `de-DP{n}` drawing-page style
 * carries for a model transition — spliced alongside the copied fill/background
 * properties. Returns "" for explicit none (no visual transition; the caller
 * still mints a style so the replaced original's smil attrs are dropped). */
export function odpTransitionProps(t: SlideTransition): string {
  const parts: string[] = [];
  if (t.type !== "none") {
    const { smilType, smilSubtype, reverse } = odpSmilFor(t);
    parts.push(`smil:type="${smilType}"`);
    if (smilSubtype) {
      parts.push(`smil:subtype="${smilSubtype}"`);
    }
    if (reverse) {
      parts.push(`smil:direction="reverse"`);
    }
    parts.push(`smil:dur="${msToOdfDuration(effectiveDurationMs(t))}"`);
  }
  if (t.advanceAfterMs !== undefined && t.advanceAfterMs > 0) {
    parts.push(`presentation:transition-type="automatic"`);
    parts.push(`presentation:duration="${msToOdfDuration(Math.round(t.advanceAfterMs))}"`);
  } else {
    parts.push(`presentation:transition-type="manual"`);
  }
  return parts.join(" ");
}

function odpSmilFor(t: SlideTransition): { smilType: string; smilSubtype?: string; reverse: boolean } {
  const direction = effectiveDirection(t) ?? "left";
  switch (t.type) {
    case "fade":
      return { smilType: "fade", smilSubtype: "crossfade", reverse: false };
    case "push":
      return { smilType: "pushWipe", smilSubtype: DIR_TO_FROM_EDGE[direction], reverse: false };
    case "wipe":
      return {
        smilType: "barWipe",
        smilSubtype: direction === "left" || direction === "right" ? "leftToRight" : "topToBottom",
        reverse: direction === "left" || direction === "up",
      };
    case "cover":
      return { smilType: "slideWipe", smilSubtype: DIR_TO_FROM_EDGE[direction], reverse: false };
    case "uncover":
      return { smilType: "slideWipe", smilSubtype: DIR_TO_FROM_EDGE[direction], reverse: true };
    case "split":
      return { smilType: "barnDoorWipe", smilSubtype: "horizontal", reverse: false };
    default:
      return { smilType: "fade", reverse: false };
  }
}
