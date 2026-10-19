// The pptx animation wire mapping (SPEC-086) — the NORMATIVE source for how the
// modelled `SlideAnimation[]` subset reads from and writes to a slide's
// `p:timing` SMIL tree. The canonical serialised bytes are pinned by
// `timingGolden.ts`; the reader accepts those plus PowerPoint-authored
// equivalents (reordered attributes, extra behaviour internals it ignores).
//
// Framework-free (imports only the model types, the format-neutral animation
// helpers and core's XML scanner). Anything the preset table does not
// recognise — an unknown preset, a second sequence, a `p:audio` cue, a target
// the slide doesn't contain or only part of a shape (a by-paragraph build) —
// makes the whole timing "out of subset": the reader returns `undefined` and
// the caller keeps the source markup preserved verbatim (AC-4) — its shape
// and paragraph references re-pointed on a rebuild (the "Preserved timing"
// section below).
import type { AnimationEffect, SlideAnimation, SlideParagraph, TransitionDirection } from "../model";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import {
  buildGroups,
  DEFAULT_ANIMATION_DURATION_MS,
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
  usesAnimationDirection,
} from "../animations";

/** (presetID, presetClass) per effect. `entr`/`exit` distinguish the two id-10
 * fades and the two id-2 flies. */
const EFFECT_PRESET: Record<AnimationEffect, { id: number; cls: "entr" | "exit" }> = {
  appear: { id: 1, cls: "entr" },
  fadeIn: { id: 10, cls: "entr" },
  flyIn: { id: 2, cls: "entr" },
  wipeIn: { id: 22, cls: "entr" },
  zoomIn: { id: 23, cls: "entr" },
  fadeOut: { id: 10, cls: "exit" },
  flyOut: { id: 2, cls: "exit" },
};

/** (presetClass:presetID) → effect, the reader's inverse of the table. */
const PRESET_TO_EFFECT: Record<string, AnimationEffect> = {
  "entr:1": "appear",
  "entr:10": "fadeIn",
  "entr:2": "flyIn",
  "entr:22": "wipeIn",
  "entr:23": "zoomIn",
  "exit:10": "fadeOut",
  "exit:2": "flyOut",
};

const DIR_TO_SUBTYPE: Record<TransitionDirection, number> = { left: 8, right: 2, up: 1, down: 4 };
const SUBTYPE_TO_DIR: Record<string, TransitionDirection> = { "8": "left", "2": "right", "1": "up", "4": "down" };

const NODE_TYPE: Record<SlideAnimation["trigger"], string> = {
  click: "clickEffect",
  withPrevious: "withEffect",
  afterPrevious: "afterEffect",
};
const TRIGGER_FOR_NODE: Record<string, SlideAnimation["trigger"]> = {
  clickEffect: "click",
  withEffect: "withPrevious",
  afterEffect: "afterPrevious",
};

/** The wipe filter token per reveal-origin direction (behaviour internal — the
 * reader ignores it). */
const WIPE_FILTER: Record<TransitionDirection, string> = { left: "right", right: "left", up: "down", down: "up" };

// ---------------------------------------------------------------------------
// Write
// ---------------------------------------------------------------------------

function subtypeFor(a: SlideAnimation): number {
  if (a.effect === "zoomIn") {
    return 16;
  }
  if (usesAnimationDirection(a.effect)) {
    return DIR_TO_SUBTYPE[effectiveAnimationDirection(a) ?? "left"];
  }
  return 0;
}

/** A behaviour's `p:cBhvr` (its `p:cTn dur` carries the effect duration when it
 * is the effect's first behaviour — the reader reads that). */
function cBhvr(spid: number, dur: number, attrName?: string): string {
  const attr = attrName ? `<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst>` : "";
  return `<p:cBhvr><p:cTn dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>${attr}</p:cBhvr>`;
}

function tav(tm: string, val: string): string {
  return `<p:tav tm="${tm}"><p:val><p:fltVal val="${val}"/></p:val></p:tav>`;
}

function flyBehaviour(spid: number, dur: number, direction: TransitionDirection, out: boolean): string {
  const axis = direction === "left" || direction === "right" ? "ppt_x" : "ppt_y";
  const off = direction === "left" || direction === "up" ? "-0.5" : "1.5";
  const from = out ? "0" : off;
  const to = out ? off : "0";
  return (
    `<p:anim calcmode="lin" valueType="num">${cBhvr(spid, dur, axis)}` +
    `<p:tavLst>${tav("0", from)}${tav("100000", to)}</p:tavLst></p:anim>`
  );
}

function setVisibility(spid: number, value: "visible" | "hidden"): string {
  return `<p:set>${cBhvr(spid, 1, "style.visibility")}<p:to><p:strVal val="${value}"/></p:to></p:set>`;
}

/** The effect's behaviour tree. The FIRST behaviour always carries the effect
 * duration on its `p:cTn dur` (the reader takes duration from there). */
function behavioursFor(a: SlideAnimation, spid: number): string {
  const dur = effectiveAnimationDurationMs(a);
  const direction = effectiveAnimationDirection(a) ?? "left";
  switch (a.effect) {
    case "appear":
      return setVisibility(spid, "visible");
    case "fadeIn":
      return `<p:animEffect transition="in" filter="fade">${cBhvr(spid, dur)}</p:animEffect>${setVisibility(spid, "visible")}`;
    case "fadeOut":
      return `<p:animEffect transition="out" filter="fade">${cBhvr(spid, dur)}</p:animEffect>${setVisibility(spid, "hidden")}`;
    case "flyIn":
      return flyBehaviour(spid, dur, direction, false);
    case "flyOut":
      return flyBehaviour(spid, dur, direction, true);
    case "wipeIn":
      return `<p:animEffect filter="wipe(${WIPE_FILTER[direction]})">${cBhvr(spid, dur)}</p:animEffect>`;
    case "zoomIn":
      return (
        `<p:animEffect filter="fade">${cBhvr(spid, dur)}</p:animEffect>` +
        `<p:anim calcmode="lin" valueType="num">${cBhvr(spid, dur, "ppt_w")}<p:tavLst>${tav("0", "0")}${tav("100000", "1")}</p:tavLst></p:anim>` +
        `<p:anim calcmode="lin" valueType="num">${cBhvr(spid, dur, "ppt_h")}<p:tavLst>${tav("0", "0")}${tav("100000", "1")}</p:tavLst></p:anim>`
      );
    default:
      return "";
  }
}

function effectPar(a: SlideAnimation, spid: number, nodeType: string): string {
  const preset = EFFECT_PRESET[a.effect];
  return (
    `<p:par><p:cTn presetID="${preset.id}" presetClass="${preset.cls}" presetSubtype="${subtypeFor(a)}" fill="hold" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${effectiveAnimationDelayMs(a)}"/></p:stCondLst>` +
    `<p:childTnLst>${behavioursFor(a, spid)}</p:childTnLst>` +
    `</p:cTn></p:par>`
  );
}

/** One click group: an outer `delay="indefinite"` (the click) wrapping an inner
 * `delay="0"` par that holds the group's effect pars. */
function clickGroupPar(effectsXml: string): string {
  return (
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    effectsXml +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>`
  );
}

/** Serialises a slide's animations to its `p:timing` (schema-ordered to sit
 * after `p:transition`). `spidFor` resolves a frame id to its `p:cNvPr id`
 * (SPEC-081 mints these 2,3,… in z-order). Returns "" for no animations. */
export function writePptxTiming(
  animations: SlideAnimation[] | undefined,
  spidFor: (frameId: string) => number | undefined,
): string {
  if (!animations || animations.length === 0) {
    return "";
  }
  const groups = buildGroups(animations);
  const groupXml = groups
    .map((group) => {
      const effects = group
        .map((a, i) => {
          const spid = spidFor(a.frameId);
          if (spid === undefined) {
            return "";
          }
          const nodeType = i === 0 ? "clickEffect" : NODE_TYPE[a.trigger];
          return effectPar(a, spid, nodeType);
        })
        .join("");
      return clickGroupPar(effects);
    })
    .join("");
  return (
    `<p:timing><p:tnLst><p:par>` +
    `<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    groupXml +
    `</p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
    `</p:seq>` +
    `</p:childTnLst></p:cTn>` +
    `</p:par></p:tnLst></p:timing>`
  );
}

// ---------------------------------------------------------------------------
// Read
// ---------------------------------------------------------------------------

function descendants(el: Element, localName: string): Element[] {
  return Array.from(el.getElementsByTagName("*")).filter((e) => e.localName === localName);
}

const ABORT = Symbol("out-of-subset");

/** Parses a `p:timing` element to modelled animations, or `undefined` when the
 * tree is outside the subset (→ keep it preserved). `resolveSpid` maps a
 * `p:spTgt spid` back to a frame id; an unresolvable target is out of subset. */
export function readPptxTiming(
  timingEl: Element,
  resolveSpid: (spid: string) => string | undefined,
): SlideAnimation[] | undefined {
  try {
    return parseTiming(timingEl, resolveSpid);
  } catch (e) {
    if (e === ABORT) {
      return undefined;
    }
    throw e;
  }
}

function parseTiming(
  timingEl: Element,
  resolveSpid: (spid: string) => string | undefined,
): SlideAnimation[] | undefined {
  // A sound/video cue, or more than one sequence, is out of subset.
  if (descendants(timingEl, "audio").length > 0 || descendants(timingEl, "video").length > 0) {
    return undefined;
  }
  const seqs = descendants(timingEl, "seq");
  if (seqs.length !== 1) {
    return undefined;
  }
  const mainSeq = Array.from(seqs[0].children).find((c) => c.localName === "cTn" && c.getAttribute("nodeType") === "mainSeq");
  if (!mainSeq) {
    return undefined;
  }
  const childTnLst = Array.from(mainSeq.children).find((c) => c.localName === "childTnLst");
  if (!childTnLst) {
    return undefined;
  }
  const groupPars = Array.from(childTnLst.children).filter((c) => c.localName === "par");
  const animations: SlideAnimation[] = [];
  for (const groupPar of groupPars) {
    const effectCtns = descendants(groupPar, "cTn").filter((c) => c.hasAttribute("presetClass"));
    effectCtns.forEach((cTn, indexInGroup) => {
      // Marker pushed and checked below; the loop can't early-return, so a
      // failed parse throws to abandon the whole timing.
      const parsed = parseEffect(cTn, indexInGroup === 0, resolveSpid);
      if (!parsed) {
        throw ABORT;
      }
      animations.push(parsed);
    });
  }
  return animations.length > 0 ? animations : undefined;
}

/** Whether every behaviour of an effect animates whole shapes (the model's
 * frame-level unit): each `p:tgtEl` holds a bare `p:spTgt` and nothing iterates.
 * A sub-target under `p:spTgt` — `p:txEl` (a "by paragraph" text build's
 * paragraph range), `p:graphicEl`/`p:oleChartEl` (chart or diagram element),
 * `p:subSp`, `p:bg` — an ink/sound/slide target, or a `p:iterate` (text by
 * word or letter) animates part of the frame; reading it as the whole frame
 * would replay the frame per part and the writer would save the parts away. */
function isFrameLevel(cTn: Element): boolean {
  if (descendants(cTn, "iterate").length > 0) {
    return false;
  }
  return descendants(cTn, "tgtEl").every((tgtEl) =>
    Array.from(tgtEl.children).every((t) => t.localName === "spTgt" && t.children.length === 0),
  );
}

function parseEffect(
  cTn: Element,
  firstInGroup: boolean,
  resolveSpid: (spid: string) => string | undefined,
): SlideAnimation | undefined {
  const presetClass = cTn.getAttribute("presetClass");
  const presetID = cTn.getAttribute("presetID");
  const effect = PRESET_TO_EFFECT[`${presetClass}:${presetID}`];
  if (!effect || !isFrameLevel(cTn)) {
    return undefined;
  }
  const spTgt = descendants(cTn, "spTgt")[0];
  const spid = spTgt?.getAttribute("spid");
  const frameId = spid ? resolveSpid(spid) : undefined;
  if (!frameId) {
    return undefined;
  }
  const trigger = firstInGroup ? "click" : TRIGGER_FOR_NODE[cTn.getAttribute("nodeType") ?? ""] ?? "click";
  const animation: SlideAnimation = { frameId, effect, trigger };

  if (usesAnimationDirection(effect)) {
    const dir = SUBTYPE_TO_DIR[cTn.getAttribute("presetSubtype") ?? ""];
    if (dir && dir !== "left") {
      animation.direction = dir;
    }
  }
  const delay = Number(firstCond(cTn)?.getAttribute("delay"));
  if (Number.isFinite(delay) && delay > 0) {
    animation.delayMs = delay;
  }
  // `appear` is instant — its `p:set dur="1"` is a marker, not a duration.
  if (effect !== "appear") {
    const dur = firstBehaviourDur(cTn);
    if (dur !== undefined && dur !== DEFAULT_ANIMATION_DURATION_MS) {
      animation.durationMs = dur;
    }
  }
  return animation;
}

/** The `p:spTgt spid`s of the clips a slide's timing starts with the slide —
 * PowerPoint's "Start: Automatically": a `p:cmd type="call"` `playFrom(…)` in
 * the main sequence's first group when that group begins on the slide's own
 * `onBegin` rather than on a click. SPEC-095 reads the cue only this far, to
 * set `autoStart`; the timing itself stays preserved whole (a media cue is out
 * of SPEC-086's subset). */
export function readPptxAutoStartSpids(timingEl: Element): string[] {
  const mainSeq = descendants(timingEl, "cTn").find((c) => c.getAttribute("nodeType") === "mainSeq");
  const childTnLst = mainSeq ? Array.from(mainSeq.children).find((c) => c.localName === "childTnLst") : undefined;
  const firstGroup = childTnLst ? Array.from(childTnLst.children).find((c) => c.localName === "par") : undefined;
  const groupCtn = firstGroup ? Array.from(firstGroup.children).find((c) => c.localName === "cTn") : undefined;
  const stCondLst = groupCtn ? Array.from(groupCtn.children).find((c) => c.localName === "stCondLst") : undefined;
  const withSlide = !!stCondLst && Array.from(stCondLst.children).some((c) => c.getAttribute("evt") === "onBegin");
  if (!groupCtn || !withSlide) {
    return [];
  }
  return descendants(groupCtn, "cmd")
    .filter((cmd) => cmd.getAttribute("type") === "call" && (cmd.getAttribute("cmd") ?? "").startsWith("playFrom"))
    .map((cmd) => descendants(cmd, "spTgt")[0]?.getAttribute("spid") ?? "")
    .filter((spid) => spid !== "");
}

/** The effect's own start condition delay: the first `p:cond` under the
 * effect cTn's direct `p:stCondLst`. */
function firstCond(cTn: Element): Element | undefined {
  const stCondLst = Array.from(cTn.children).find((c) => c.localName === "stCondLst");
  return stCondLst ? Array.from(stCondLst.children).find((c) => c.localName === "cond") : undefined;
}

/** The first nested `p:cTn` (inside the effect's behaviour tree) that carries a
 * numeric `dur` — the effect duration by the canonical form's convention. A
 * `p:set` behaviour is skipped: its `dur="1"` is the instant visibility flip,
 * which PowerPoint writes BEFORE the timed behaviour (the canonical form writes
 * it after), so taking it would read every PowerPoint entrance as 1ms. */
function firstBehaviourDur(effectCtn: Element): number | undefined {
  for (const cTn of descendants(effectCtn, "cTn")) {
    if (cTn === effectCtn || cTn.parentElement?.parentElement?.localName === "set") {
      continue;
    }
    const dur = cTn.getAttribute("dur");
    if (dur && dur !== "indefinite") {
      const n = Number(dur);
      if (Number.isFinite(n)) {
        return n;
      }
    }
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Preserved (out-of-subset) timing
// ---------------------------------------------------------------------------
// A preserved `p:timing` re-emits verbatim, but it names its shapes by their
// SOURCE `p:cNvPr id`, and a rebuilt slide re-mints those ids (2, 3, … in
// z-order). The reader records each targeted frame's source id
// (`sourceWireId`); the serialiser re-points the references through the
// functions below. A drawing-element reference is any NUMERIC `spid`
// attribute — `p:spTgt`, `p:bldP`/`p:bldGraphic`/`p:bldOleChart`/`p:bldDgm`,
// `p:inkTgt`, `p14:bmkTgt`; `p:subSp spid` names a diagram sub-shape by GUID,
// never a slide shape, and is left alone.
//
// A target can also name text INSIDE a shape by position: `p:spTgt/p:txEl`
// holds a `p:pRg` (0-based `a:p` indices, inclusive — a by-paragraph build)
// or a `p:charRg` (character offsets). The model reads every `a:br` as a new
// paragraph and a rebuild writes each as its own `a:p`, so on a rebuilt shape
// those indices shift — the reader records each line's source `a:p` index
// (`pptxSourceParagraph`) for the frames such ranges target, and the
// serialiser re-points each range alongside its shape.
const SPID_ATTR = /(\s(?:[\w.-]+:)?spid=")(\d+)(")/g;

/** The shape ids a preserved `p:timing` string references. */
export function preservedTimingSpids(timingXml: string): Set<string> {
  return new Set(Array.from(timingXml.matchAll(SPID_ATTR), (m) => m[2]));
}

/** A text range a preserved target names inside its shape (`p:txEl`), in the
 * SOURCE's numbering: `pRg` by `a:p` index, `charRg` by character offset;
 * both inclusive of `st` and `end`. */
export interface PreservedTextRange {
  kind: "pRg" | "charRg";
  st: number;
  end: number;
}

/** One `p:txEl` range in a preserved timing string: the `p:spTgt spid` it
 * sits in and the source offsets of its (empty) element's open tag. */
interface TextRangeSite extends PreservedTextRange {
  spid: string;
  start: number;
  stop: number;
}

/** Every `p:txEl` range in a preserved timing string, in document order —
 * `undefined` when the string does not scan (never the case for what the
 * reader stores, which is serialised DOM) or a range is not a pair of
 * non-negative integers. */
function textRangeSites(timingXml: string): TextRangeSite[] | undefined {
  const sites: TextRangeSite[] = [];
  let spid: string | undefined;
  let inTxEl = false;
  let malformed = false;
  try {
    scanXml(timingXml, (event) => {
      if (event.type === "open") {
        if (event.local === "spTgt" && !event.selfClosing) {
          spid = event.attrs.find((a) => a.local === "spid")?.value;
        } else if (event.local === "txEl" && !event.selfClosing) {
          inTxEl = true;
        } else if (inTxEl && spid !== undefined && (event.local === "pRg" || event.local === "charRg")) {
          const st = Number(event.attrs.find((a) => a.local === "st")?.value);
          const end = Number(event.attrs.find((a) => a.local === "end")?.value);
          if (!Number.isInteger(st) || !Number.isInteger(end) || st < 0 || end < 0) {
            malformed = true;
            return "stop";
          }
          sites.push({ spid, kind: event.local, st, end, start: event.start, stop: event.end });
        }
      } else if (event.type === "close") {
        if (event.local === "spTgt") {
          spid = undefined;
        } else if (event.local === "txEl") {
          inTxEl = false;
        }
      }
      return undefined;
    });
  } catch (e) {
    if (e instanceof XmlScanError) {
      return undefined;
    }
    throw e;
  }
  return malformed ? undefined : sites;
}

/** The shape ids a preserved `p:timing` string addresses by text POSITION
 * (a `p:txEl` range inside their `p:spTgt`). */
export function preservedTimingTextTargets(timingXml: string): Set<string> {
  return new Set((textRangeSites(timingXml) ?? []).map((site) => site.spid));
}

/** A source paragraph range re-pointed at a REBUILT body, which writes each
 * model paragraph as its own `a:p`: from the first to the last line read from
 * a source paragraph inside the range (`pptxSourceParagraph`) — so a source
 * paragraph an `a:br` split, or an edit, spans several `a:p` and the build
 * still reveals all of its text. `undefined` when no line from the range is
 * left (the user deleted or merged away that paragraph), and for a character
 * range, whose offsets any text edit moves and nothing records. */
export function rebuiltParagraphRange(
  paragraphs: readonly SlideParagraph[],
  range: PreservedTextRange,
): { st: number; end: number } | undefined {
  if (range.kind !== "pRg") {
    return undefined;
  }
  let st = -1;
  let end = -1;
  paragraphs.forEach((para, index) => {
    const source = para.pptxSourceParagraph;
    if (source !== undefined && source >= range.st && source <= range.end) {
      if (st < 0) {
        st = index;
      }
      end = index;
    }
  });
  return st < 0 ? undefined : { st, end };
}

/** `st`/`end` on a range element's open tag. */
const RANGE_ATTR = /(\s(st|end)\s*=\s*["'])\d+(["'])/g;

/** A preserved `p:timing` with every shape reference re-pointed through
 * `resolve` (source id → the id the rebuilt slide gave that shape) and every
 * `p:txEl` range through `remapText` (source id + source range → the range on
 * the rebuilt shape), or `undefined` when any reference has no shape — or
 * range no text — on the rebuilt slide: the caller then drops the timing
 * whole rather than emit a dangling target (PowerPoint repairs that as
 * corruption), a range that animates the wrong paragraph, or one re-minting
 * handed to a different shape — preserved timing is never half-edited
 * (AC-4). */
export function retargetPreservedTiming(
  timingXml: string,
  resolve: (spid: string) => number | undefined,
  remapText: (spid: string, range: PreservedTextRange) => { st: number; end: number } | undefined,
): string | undefined {
  // Ranges first, while each still sits under its SOURCE spid.
  const sites = textRangeSites(timingXml);
  if (!sites) {
    return undefined;
  }
  let ranged = "";
  let cursor = 0;
  for (const site of sites) {
    const range = remapText(site.spid, site);
    if (!range) {
      return undefined;
    }
    const tag = timingXml
      .slice(site.start, site.stop)
      .replace(RANGE_ATTR, (_m, before: string, name: string, after: string) => `${before}${name === "st" ? range.st : range.end}${after}`);
    ranged += timingXml.slice(cursor, site.start) + tag;
    cursor = site.stop;
  }
  ranged += timingXml.slice(cursor);
  let dangling = false;
  const out = ranged.replace(SPID_ATTR, (match, before: string, spid: string, after: string) => {
    const id = resolve(spid);
    if (id === undefined) {
      dangling = true;
      return match;
    }
    return `${before}${id}${after}`;
  });
  return dangling ? undefined : out;
}
