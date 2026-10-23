// The pptx animation wire mapping (SPEC-086) — the NORMATIVE source for how the
// modelled `SlideAnimation[]` subset reads from and writes to a slide's
// `p:timing` SMIL tree. The canonical serialised bytes are pinned by
// `timingGolden.ts`; the reader accepts those plus PowerPoint-authored
// equivalents (reordered attributes, extra behaviour internals it ignores).
//
// SPEC-105 widens the subset — and this module's preset tables are its
// normative artefact, pinned by `TIMING_DEPTH_GOLDEN`:
// - builds: each expanded step's `p:spTgt` narrows to `p:txEl/p:pRg` (its
//   first and last paragraph), and `p:timing` gains a `p:bldLst` of
//   `p:bldP build="p"` (`bldLvl="1"` by first-level paragraph) after
//   `p:tnLst`, linked to the step effects by `grpId`;
// - effects: `presetClass="emph"` (`p:animRot`, `p:anim` on `ppt_w`/`ppt_h`,
//   `p:animClr`) and `presetClass="path"` (`p:animMotion`, slide-fraction
//   `path`, `pathEditMode="relative"`);
// - repetition: `repeatCount` (×1000, or "indefinite" with an `onNext` end
//   for until-click), `autoRev="1"` and `fill="remove"` (rewind) on the
//   effect's own `p:cTn`;
// - triggers: one `nodeType="interactiveSeq"` `p:seq` per trigger frame,
//   sibling to the mainSeq, begun by an `onClick` on the trigger's `p:spTgt`.
//
// Framework-free (imports only the model types, the format-neutral animation
// helpers and core's XML scanner). Anything the tables do not recognise — an
// unknown preset, a sequence that is neither the mainSeq nor an interactive
// one, an unresolvable trigger, a `p:audio` cue, a target the slide doesn't
// contain, a paragraph range that does not cover the frame exactly as a
// modelled build would, a `p:bldP` disagreeing with the effects, a path no
// preset reproduces — makes the whole timing "out of subset": the reader
// returns `undefined` and the caller keeps the source markup preserved
// verbatim (AC-4) — its shape and paragraph references re-pointed on a
// rebuild (the "Preserved timing" section below).
import type {
  AnimationEffect,
  MotionPath,
  SlideAnimation,
  SlideFrame,
  SlideParagraph,
  TextBuildMode,
  TransitionDirection,
} from "../model";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import {
  DEFAULT_ANIMATION_DURATION_MS,
  DEFAULT_EMPHASIS_COLOR,
  DEFAULT_SCALE_PERCENT,
  DEFAULT_SPIN_DEGREES,
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
  effectiveEmphasisColor,
  effectiveRepeatCount,
  effectiveScalePercent,
  effectiveSpinDegrees,
  MAX_REPEAT_COUNT,
  MAX_SCALE_PERCENT,
  MAX_SPIN_DEGREES,
  MIN_SCALE_PERCENT,
  MIN_SPIN_DEGREES,
  usesAnimationDirection,
} from "../animations";
import {
  buildableParagraphs,
  expandAnimation,
  interactiveBuildSteps,
  mainBuildSteps,
  type BuildStep,
} from "../buildPlayback";
import { formatPathD, inferMotionPath, parseMotionPathD, resolveMotionPath } from "../motionPaths";

type PresetClass = "entr" | "exit" | "emph" | "path";

/** (presetID, presetClass) per effect. `entr`/`exit` distinguish the two id-10
 * fades and the two id-2 flies. The emphasis ids are this editor's table
 * (SPEC-105): colour change 1, grow/shrink 6, pulse 26 and spin 11 — never
 * `emph:8`, which SPEC-086's exotic fixture uses and which must stay out of
 * subset. A player plays an effect by its behaviours; the id only names it.
 * `motionPath` takes its id from `PATH_PRESET_ID`. */
const EFFECT_PRESET: Record<AnimationEffect, { id: number; cls: PresetClass }> = {
  appear: { id: 1, cls: "entr" },
  fadeIn: { id: 10, cls: "entr" },
  flyIn: { id: 2, cls: "entr" },
  wipeIn: { id: 22, cls: "entr" },
  zoomIn: { id: 23, cls: "entr" },
  fadeOut: { id: 10, cls: "exit" },
  flyOut: { id: 2, cls: "exit" },
  colorChange: { id: 1, cls: "emph" },
  growShrink: { id: 6, cls: "emph" },
  spin: { id: 11, cls: "emph" },
  pulse: { id: 26, cls: "emph" },
  motionPath: { id: 0, cls: "path" },
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
  "emph:1": "colorChange",
  "emph:6": "growShrink",
  "emph:11": "spin",
  "emph:26": "pulse",
};

/** `presetClass="path"` presetID per path preset (SPEC-105). The path string
 * itself is normative: the reader re-derives the preset's direction and
 * distance from it, and refuses a path the preset does not reproduce. */
const PATH_PRESET_ID: Record<MotionPath["preset"], number> = { custom: 0, line: 42, arc: 37, turn: 39, loop: 40 };
const PATH_ID_TO_PRESET: Record<string, MotionPath["preset"]> = Object.fromEntries(
  (Object.entries(PATH_PRESET_ID) as [MotionPath["preset"], number][]).map(([preset, id]) => [String(id), preset]),
);

/** The scale a pulse throbs to and back. */
const PULSE_SCALE = 1.1;

const DIR_TO_SUBTYPE: Record<TransitionDirection, number> = { left: 8, right: 2, up: 1, down: 4 };
const SUBTYPE_TO_DIR: Record<string, TransitionDirection> = { "8": "left", "2": "right", "1": "up", "4": "down" };

const TRIGGER_FOR_NODE: Record<string, SlideAnimation["trigger"]> = {
  clickEffect: "click",
  withEffect: "withPrevious",
  afterEffect: "afterPrevious",
};

/** The wipe filter token per reveal-origin direction (behaviour internal — the
 * reader ignores it). */
const WIPE_FILTER: Record<TransitionDirection, string> = { left: "right", right: "left", up: "down", down: "up" };

const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

/** A step's `p:spTgt`: the whole shape, or (a paragraph step) narrowed to its
 * `p:txEl/p:pRg`. */
function spTgt(step: BuildStep, spid: number): string {
  if (step.paragraphIndex === undefined) {
    return `<p:spTgt spid="${spid}"/>`;
  }
  const end = step.paragraphEnd ?? step.paragraphIndex;
  return `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${step.paragraphIndex}" end="${end}"/></p:txEl></p:spTgt>`;
}

/** A behaviour's `p:cBhvr` (its `p:cTn dur` carries the effect duration when it
 * is the effect's first behaviour — the reader reads that). `cTnExtra` lands
 * on the behaviour's `p:cTn` before `fill`. */
function cBhvr(tgt: string, dur: number, attributes: string[] = [], cTnExtra = ""): string {
  const attr =
    attributes.length > 0
      ? `<p:attrNameLst>${attributes.map((n) => `<p:attrName>${n}</p:attrName>`).join("")}</p:attrNameLst>`
      : "";
  return `<p:cBhvr><p:cTn dur="${dur}"${cTnExtra} fill="hold"/><p:tgtEl>${tgt}</p:tgtEl>${attr}</p:cBhvr>`;
}

function tav(tm: string, val: string): string {
  return `<p:tav tm="${tm}"><p:val><p:fltVal val="${val}"/></p:val></p:tav>`;
}

function strTav(tm: string, val: string): string {
  return `<p:tav tm="${tm}"><p:val><p:strVal val="${val}"/></p:val></p:tav>`;
}

function flyBehaviour(tgt: string, dur: number, direction: TransitionDirection, out: boolean): string {
  const axis = direction === "left" || direction === "right" ? "ppt_x" : "ppt_y";
  const off = direction === "left" || direction === "up" ? "-0.5" : "1.5";
  const from = out ? "0" : off;
  const to = out ? off : "0";
  return (
    `<p:anim calcmode="lin" valueType="num">${cBhvr(tgt, dur, [axis])}` +
    `<p:tavLst>${tav("0", from)}${tav("100000", to)}</p:tavLst></p:anim>`
  );
}

function setVisibility(tgt: string, value: "visible" | "hidden"): string {
  return `<p:set>${cBhvr(tgt, 1, ["style.visibility"])}<p:to><p:strVal val="${value}"/></p:to></p:set>`;
}

/** A width-and-height throb to `factor` and back (pulse, grow/shrink): the
 * two `p:anim` keyframe lists `#ppt_w` → `#ppt_w*factor` → `#ppt_w`. */
function scaleThrob(tgt: string, dur: number, factor: number): string {
  return (["ppt_w", "ppt_h"] as const)
    .map(
      (attr) =>
        `<p:anim calcmode="lin" valueType="num">${cBhvr(tgt, dur, [attr])}` +
        `<p:tavLst>${strTav("0", `#${attr}`)}${strTav("50000", `#${attr}*${factor}`)}${strTav("100000", `#${attr}`)}</p:tavLst></p:anim>`,
    )
    .join("");
}

/** A colour change to `color` and back: the text and fill colours animate
 * there over half the duration and auto-reverse over the other half. */
function colorThrob(tgt: string, dur: number, color: string): string {
  const half = Math.round(dur / 2);
  const hex = color.slice(1).toUpperCase();
  return (["style.color", "fillcolor"] as const)
    .map(
      (attr) =>
        `<p:animClr clrSpc="rgb" dir="cw">${cBhvr(tgt, half, [attr], ` autoRev="1"`)}` +
        `<p:to><a:srgbClr val="${hex}"/></p:to></p:animClr>`,
    )
    .join("");
}

/** The effect's behaviour tree for one step. The FIRST behaviour always
 * carries the effect duration on its `p:cTn dur` (the reader takes duration
 * from there; a colour change's is half, auto-reversed). */
function behavioursFor(step: BuildStep, spid: number): string {
  const a = step.animation;
  const tgt = spTgt(step, spid);
  const dur = effectiveAnimationDurationMs(a);
  const direction = effectiveAnimationDirection(a) ?? "left";
  switch (a.effect) {
    case "appear":
      return setVisibility(tgt, "visible");
    case "fadeIn":
      return `<p:animEffect transition="in" filter="fade">${cBhvr(tgt, dur)}</p:animEffect>${setVisibility(tgt, "visible")}`;
    case "fadeOut":
      return `<p:animEffect transition="out" filter="fade">${cBhvr(tgt, dur)}</p:animEffect>${setVisibility(tgt, "hidden")}`;
    case "flyIn":
      return flyBehaviour(tgt, dur, direction, false);
    case "flyOut":
      return flyBehaviour(tgt, dur, direction, true);
    case "wipeIn":
      return `<p:animEffect filter="wipe(${WIPE_FILTER[direction]})">${cBhvr(tgt, dur)}</p:animEffect>`;
    case "zoomIn":
      return (
        `<p:animEffect filter="fade">${cBhvr(tgt, dur)}</p:animEffect>` +
        `<p:anim calcmode="lin" valueType="num">${cBhvr(tgt, dur, ["ppt_w"])}<p:tavLst>${tav("0", "0")}${tav("100000", "1")}</p:tavLst></p:anim>` +
        `<p:anim calcmode="lin" valueType="num">${cBhvr(tgt, dur, ["ppt_h"])}<p:tavLst>${tav("0", "0")}${tav("100000", "1")}</p:tavLst></p:anim>`
      );
    case "pulse":
      return scaleThrob(tgt, dur, PULSE_SCALE);
    case "growShrink":
      return scaleThrob(tgt, dur, effectiveScalePercent(a) / 100);
    case "spin":
      return `<p:animRot by="${Math.round(effectiveSpinDegrees(a) * 60000)}">${cBhvr(tgt, dur, ["r"])}</p:animRot>`;
    case "colorChange":
      return colorThrob(tgt, dur, effectiveEmphasisColor(a));
    case "motionPath": {
      const commands = a.path ? resolveMotionPath(a.path) : undefined;
      if (!commands) {
        return "";
      }
      return (
        `<p:animMotion origin="layout" path="${formatPathD(commands)}" pathEditMode="relative">` +
        `${cBhvr(tgt, dur, ["ppt_x", "ppt_y"])}</p:animMotion>`
      );
    }
    default:
      return "";
  }
}

const UNTIL_NEXT_CLICK = `<p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>`;

/** One step's effect `p:par`. `id` (the step's model order) is written only
 * on a slide with an interactive sequence; `grpId` only on a paragraph step
 * (it links the step to its `p:bldP`). */
function effectPar(step: BuildStep, spid: number, nodeType: string, grpId?: number, id?: number): string {
  const a = step.animation;
  const preset =
    a.effect === "motionPath" ? { id: PATH_PRESET_ID[a.path?.preset ?? "line"], cls: "path" } : EFFECT_PRESET[a.effect];
  const idAttr = id !== undefined ? ` id="${id}"` : "";
  const repeat = a.repeatUntilClick
    ? ` repeatCount="indefinite"`
    : a.repeatCount !== undefined
      ? ` repeatCount="${effectiveRepeatCount(a) * 1000}"`
      : "";
  const autoRev = a.autoReverse ? ` autoRev="1"` : "";
  const fill = a.rewindOnComplete ? "remove" : "hold";
  const grp = grpId !== undefined ? ` grpId="${grpId}"` : "";
  return (
    `<p:par><p:cTn${idAttr} presetID="${preset.id}" presetClass="${preset.cls}" presetSubtype="${subtypeFor(a)}"${repeat}${autoRev} fill="${fill}"${grp} nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${effectiveAnimationDelayMs(a)}"/></p:stCondLst>` +
    (a.repeatUntilClick ? UNTIL_NEXT_CLICK : "") +
    `<p:childTnLst>${behavioursFor(step, spid)}</p:childTnLst>` +
    `</p:cTn></p:par>`
  );
}

/** One click group: an outer par (begun by the click — `delay="indefinite"`,
 * or `"0"` for an interactive sequence's first group, which its trigger's
 * click itself begins) wrapping an inner `delay="0"` par that holds the
 * group's effect pars. */
function clickGroupPar(effectsXml: string, begin = "indefinite"): string {
  return (
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="${begin}"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    effectsXml +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>`
  );
}

/** A step's `nodeType` within its group: the first is the click; a row's
 * first step otherwise joins on its trigger (an interactive row joins its
 * predecessor), a later paragraph step on the row's `buildTrigger`. */
function nodeTypeAt(group: BuildStep[], i: number, interactive: boolean): string {
  if (i === 0) {
    return "clickEffect";
  }
  const a = group[i].animation;
  if (group[i - 1].animation !== a) {
    return !interactive && a.trigger === "afterPrevious" ? "afterEffect" : "withEffect";
  }
  return a.buildTrigger === "withPrevious" ? "withEffect" : "afterEffect";
}

/** Serialises a slide's animations to its `p:timing` (schema-ordered to sit
 * after `p:transition`). `spidFor` resolves a frame id to its `p:cNvPr id`
 * (SPEC-081 mints these 2,3,… in z-order); `frameFor` resolves it to the
 * frame, whose CURRENT paragraphs a build expands over (absent: every row
 * plays whole). Returns "" for no animations. */
export function writePptxTiming(
  animations: SlideAnimation[] | undefined,
  spidFor: (frameId: string) => number | undefined,
  frameFor: (frameId: string) => SlideFrame | undefined = () => undefined,
): string {
  if (!animations || animations.length === 0) {
    return "";
  }
  const main = mainBuildSteps(animations, frameFor);
  const interactive = interactiveBuildSteps(animations, frameFor);
  // Effect node ids in MODEL order, only where an interactive sequence makes
  // document order (main sequence first) lose where an onClickOf row sits.
  const ids = new Map<BuildStep, number>();
  if (interactive.length > 0) {
    const stepsOf = new Map<SlideAnimation, BuildStep[]>();
    for (const step of [...main.flat(), ...interactive.flatMap((s) => s.groups.flat())]) {
      stepsOf.set(step.animation, [...(stepsOf.get(step.animation) ?? []), step]);
    }
    let next = 3;
    for (const animation of animations) {
      for (const step of stepsOf.get(animation) ?? []) {
        ids.set(step, next++);
      }
    }
  }
  // A paragraph build's steps link to their `p:bldP` by grpId: 0 for the first
  // such row on a shape, 1 for the next, … in model order.
  const grpIds = new Map<SlideAnimation, number>();
  const bldPs: string[] = [];
  const perSpid = new Map<number, number>();
  for (const animation of animations) {
    const spid = spidFor(animation.frameId);
    const steps = expandAnimation(animation, frameFor(animation.frameId));
    if (spid === undefined || steps[0].paragraphIndex === undefined) {
      continue;
    }
    const grpId = perSpid.get(spid) ?? 0;
    perSpid.set(spid, grpId + 1);
    grpIds.set(animation, grpId);
    const level = animation.buildBy === "byFirstLevelParagraph" ? ` bldLvl="1"` : "";
    bldPs.push(`<p:bldP spid="${spid}" grpId="${grpId}" build="p"${level}/>`);
  }
  const groupXml = (groups: BuildStep[][], interactiveSeq: boolean): string =>
    groups
      .map((group, gi) => {
        const effects = group
          .map((step, i) => {
            const spid = spidFor(step.frameId);
            if (spid === undefined) {
              return "";
            }
            const grpId = step.paragraphIndex === undefined ? undefined : grpIds.get(step.animation);
            return effectPar(step, spid, nodeTypeAt(group, i, interactiveSeq), grpId, ids.get(step));
          })
          .join("");
        return clickGroupPar(effects, interactiveSeq && gi === 0 ? "0" : "indefinite");
      })
      .join("");
  const mainSeq =
    main.length > 0
      ? `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
        groupXml(main, false) +
        `</p:childTnLst></p:cTn>` +
        `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
        `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
        `</p:seq>`
      : "";
  const interactiveSeqs = interactive
    .map((sequence) => {
      const trigger = spidFor(sequence.triggerFrameId);
      if (trigger === undefined) {
        return "";
      }
      const onClick = `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond>`;
      return (
        `<p:seq concurrent="1" nextAc="seek">` +
        `<p:cTn restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
        `<p:stCondLst>${onClick}</p:stCondLst>` +
        `<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>` +
        `<p:childTnLst>${groupXml(sequence.groups, true)}</p:childTnLst></p:cTn>` +
        `<p:nextCondLst>${onClick}</p:nextCondLst>` +
        `</p:seq>`
      );
    })
    .join("");
  const bldLst = bldPs.length > 0 ? `<p:bldLst>${bldPs.join("")}</p:bldLst>` : "";
  return (
    `<p:timing><p:tnLst><p:par>` +
    `<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    mainSeq +
    interactiveSeqs +
    `</p:childTnLst></p:cTn>` +
    `</p:par></p:tnLst>${bldLst}</p:timing>`
  );
}

// ---------------------------------------------------------------------------
// Read
// ---------------------------------------------------------------------------

function descendants(el: Element, localName: string): Element[] {
  return Array.from(el.getElementsByTagName("*")).filter((e) => e.localName === localName);
}

function child(el: Element, localName: string): Element | undefined {
  return Array.from(el.children).find((c) => c.localName === localName);
}

const ABORT = Symbol("out-of-subset");

/** Parses a `p:timing` element to modelled animations, or `undefined` when the
 * tree is outside the subset (→ keep it preserved). `resolveSpid` maps a
 * `p:spTgt spid` back to a frame id; an unresolvable target is out of subset.
 * `frameFor` resolves a frame id to its frame: a paragraph-range target is
 * modelled as a build only against the frame's paragraphs (absent, any
 * paragraph range is out of subset). */
export function readPptxTiming(
  timingEl: Element,
  resolveSpid: (spid: string) => string | undefined,
  frameFor?: (frameId: string) => SlideFrame | undefined,
): SlideAnimation[] | undefined {
  try {
    return parseTiming(timingEl, resolveSpid, frameFor);
  } catch (e) {
    if (e === ABORT) {
      return undefined;
    }
    throw e;
  }
}

/** One effect `p:cTn` as read: the animation it encodes (trigger still to be
 * settled from its place), its target and where it sits. */
interface WireStep {
  animation: SlideAnimation;
  spid: string;
  range?: { st: number; end: number };
  nodeType: string;
  grpId: string;
  /** The cTn's numeric `id`, when it has one — the model order of a slide
   * with interactive sequences. */
  order?: number;
  opensGroup: boolean;
  groupIndex: number;
}

interface BldP {
  build: string;
  bldLvl: string | null;
  used: boolean;
}

function parseTiming(
  timingEl: Element,
  resolveSpid: (spid: string) => string | undefined,
  frameFor: ((frameId: string) => SlideFrame | undefined) | undefined,
): SlideAnimation[] | undefined {
  // A sound/video cue is out of subset.
  if (descendants(timingEl, "audio").length > 0 || descendants(timingEl, "video").length > 0) {
    return undefined;
  }
  // More than one sequence only as (at most) one mainSeq plus interactive
  // sequences, each begun by a click on a shape this slide has.
  const seqs = descendants(timingEl, "seq");
  if (seqs.length === 0) {
    return undefined;
  }
  let main: Element | undefined;
  const interactive: { cTn: Element; triggerFrameId: string }[] = [];
  for (const seq of seqs) {
    const cTn = child(seq, "cTn");
    const nodeType = cTn?.getAttribute("nodeType");
    if (cTn && nodeType === "mainSeq" && !main) {
      main = cTn;
    } else if (cTn && nodeType === "interactiveSeq") {
      const triggerFrameId = interactiveTrigger(cTn, resolveSpid);
      if (!triggerFrameId) {
        return undefined;
      }
      interactive.push({ cTn, triggerFrameId });
    } else {
      return undefined;
    }
  }
  const bldPs = new Map<string, BldP>();
  for (const bldP of descendants(timingEl, "bldP")) {
    bldPs.set(`${bldP.getAttribute("spid")}:${bldP.getAttribute("grpId") ?? "0"}`, {
      build: bldP.getAttribute("build") ?? "whole",
      bldLvl: bldP.getAttribute("bldLvl"),
      used: false,
    });
  }
  const rows: { animation: SlideAnimation; order?: number }[] = [];
  if (main) {
    rows.push(...collapse(sequenceSteps(main, resolveSpid), undefined, bldPs, frameFor));
  }
  for (const sequence of interactive) {
    rows.push(...collapse(sequenceSteps(sequence.cTn, resolveSpid), sequence.triggerFrameId, bldPs, frameFor));
  }
  // A paragraph build declared but never played disagrees with the effects.
  for (const bldP of bldPs.values()) {
    if (bldP.build === "p" && !bldP.used) {
      return undefined;
    }
  }
  const ordered = rows.length > 0 && rows.every((r) => r.order !== undefined)
    ? rows.toSorted((a, b) => (a.order ?? 0) - (b.order ?? 0))
    : rows;
  return ordered.length > 0 ? ordered.map((r) => r.animation) : undefined;
}

/** An interactive sequence's trigger frame: the bare `p:spTgt` of its
 * `onClick` start condition, resolved — undefined when it is anything else. */
function interactiveTrigger(cTn: Element, resolveSpid: (spid: string) => string | undefined): string | undefined {
  const stCondLst = child(cTn, "stCondLst");
  const cond = stCondLst ? Array.from(stCondLst.children).find((c) => c.getAttribute("evt") === "onClick") : undefined;
  const tgtEl = cond ? child(cond, "tgtEl") : undefined;
  const target = tgtEl ? child(tgtEl, "spTgt") : undefined;
  const spid = target && target.children.length === 0 ? target.getAttribute("spid") : null;
  return spid ? resolveSpid(spid) : undefined;
}

/** Every effect of a sequence's click groups, in document order. */
function sequenceSteps(seqCtn: Element, resolveSpid: (spid: string) => string | undefined): WireStep[] {
  const childTnLst = child(seqCtn, "childTnLst");
  if (!childTnLst) {
    throw ABORT;
  }
  const steps: WireStep[] = [];
  Array.from(childTnLst.children)
    .filter((c) => c.localName === "par")
    .forEach((groupPar, groupIndex) => {
      descendants(groupPar, "cTn")
        .filter((c) => c.hasAttribute("presetClass"))
        .forEach((cTn, indexInGroup) => {
          steps.push(parseEffect(cTn, indexInGroup === 0, groupIndex, resolveSpid));
        });
    });
  return steps;
}

/** The comparable identity of a step's effect (everything but its target and
 * place): a build's steps must all share it. */
function effectKey(a: SlideAnimation): string {
  const { frameId: _frameId, trigger: _trigger, ...rest } = a;
  return JSON.stringify(rest, Object.keys(rest).toSorted());
}

/** How a build's later step begins, from its place: opening a group is a
 * click; otherwise its nodeType. */
function stepTrigger(step: WireStep): SlideAnimation["trigger"] | undefined {
  return step.opensGroup ? "click" : TRIGGER_FOR_NODE[step.nodeType];
}

/** The build modes a `p:bldP` admits for paragraph steps: `bldLvl="1"` only
 * by first-level paragraph; no `bldLvl` (this editor's by-paragraph form, and
 * PowerPoint's default first-level build) either; a deeper level builds every
 * paragraph separately when the text is no deeper than that. */
function buildModes(bldP: BldP | undefined): TextBuildMode[] {
  if (!bldP || bldP.build !== "p") {
    return [];
  }
  if (bldP.bldLvl === null) {
    return ["byParagraph", "byFirstLevelParagraph"];
  }
  if (bldP.bldLvl === "1") {
    return ["byFirstLevelParagraph"];
  }
  return Number(bldP.bldLvl) > 1 ? ["byParagraph"] : [];
}

/** Collapses a sequence's steps into rows: a whole-shape step is one row; a
 * run of paragraph steps on one shape is one build row only when its ranges
 * are exactly the expansion of a build mode over the frame's paragraphs (they
 * partition it, in order, no gap or overlap), every step shares the effect,
 * the later steps agree on one build trigger and the `p:bldP` agrees.
 * `triggerFrameId` marks an interactive sequence. */
function collapse(
  steps: WireStep[],
  triggerFrameId: string | undefined,
  bldPs: Map<string, BldP>,
  frameFor: ((frameId: string) => SlideFrame | undefined) | undefined,
): { animation: SlideAnimation; order?: number }[] {
  const rows: { animation: SlideAnimation; order?: number }[] = [];
  let i = 0;
  while (i < steps.length) {
    const first = steps[i];
    const row: SlideAnimation = { ...first.animation };
    if (triggerFrameId !== undefined) {
      // Every interactive row starts on its trigger's (first) click.
      if ((first.opensGroup && first.groupIndex !== 0) || (!first.opensGroup && first.nodeType !== "withEffect")) {
        throw ABORT;
      }
      row.trigger = "onClickOf";
      row.triggerFrameId = triggerFrameId;
      if (triggerFrameId === row.frameId) {
        throw ABORT;
      }
    } else {
      row.trigger = first.opensGroup ? "click" : TRIGGER_FOR_NODE[first.nodeType] ?? "click";
    }
    const bldP = bldPs.get(`${first.spid}:${first.grpId}`);
    if (!first.range) {
      if (bldP?.build === "p") {
        throw ABORT;
      }
      rows.push({ animation: row, order: first.order });
      i += 1;
      continue;
    }
    const frame = frameFor?.(row.frameId);
    const paragraphs = buildableParagraphs(frame);
    if (!paragraphs || !sourceParagraphsIntact(paragraphs)) {
      throw ABORT;
    }
    const key = effectKey(first.animation);
    let matched = false;
    for (const mode of buildModes(bldP)) {
      const expected = expandAnimation({ ...row, buildBy: mode }, frame);
      const run = steps.slice(i, i + expected.length);
      if (
        run.length !== expected.length ||
        expected.some(
          (e, k) =>
            e.paragraphIndex === undefined ||
            run[k].spid !== first.spid ||
            run[k].grpId !== first.grpId ||
            run[k].range?.st !== e.paragraphIndex ||
            run[k].range?.end !== e.paragraphEnd ||
            effectKey(run[k].animation) !== key,
        )
      ) {
        continue;
      }
      const later = new Set(run.slice(1).map(stepTrigger));
      if (later.size > 1 || later.has(undefined)) {
        continue;
      }
      row.buildBy = mode;
      const buildTrigger = run.length > 1 ? run[1] && stepTrigger(run[1]) : "click";
      if (buildTrigger && buildTrigger !== "click" && buildTrigger !== "onClickOf") {
        row.buildTrigger = buildTrigger;
      }
      bldP!.used = true;
      rows.push({ animation: row, order: first.order });
      i += run.length;
      matched = true;
      break;
    }
    if (!matched) {
      throw ABORT;
    }
  }
  return rows;
}

/** Whether a frame's lines are its source `a:p`s one-for-one (no `a:br`
 * split, no dropped empty paragraph), so `p:pRg` indices address them. */
function sourceParagraphsIntact(paragraphs: readonly SlideParagraph[]): boolean {
  return paragraphs.every((p, i) => p.pptxSourceParagraph === undefined || p.pptxSourceParagraph === i);
}

/** The single target every behaviour of an effect names: one bare
 * `p:spTgt`, or one narrowed to a single `p:txEl/p:pRg`. Anything else — a
 * character range, a chart/diagram/sub-shape element, the background, an
 * ink/sound/slide target, an iterate (text by word or letter), or
 * behaviours that disagree — animates part of the frame in a way the model
 * cannot express, and reading it as the whole frame (or a build) would
 * replay it wrongly. */
function effectTarget(cTn: Element): { spid: string; range?: { st: number; end: number } } | undefined {
  if (descendants(cTn, "iterate").length > 0) {
    return undefined;
  }
  let found: { spid: string; range?: { st: number; end: number } } | undefined;
  for (const behaviour of descendants(cTn, "cBhvr")) {
    const tgtEl = child(behaviour, "tgtEl");
    const targets = tgtEl ? Array.from(tgtEl.children) : [];
    const sp = targets.length === 1 && targets[0].localName === "spTgt" ? targets[0] : undefined;
    const spid = sp?.getAttribute("spid");
    if (!sp || !spid) {
      return undefined;
    }
    let range: { st: number; end: number } | undefined;
    if (sp.children.length > 0) {
      const txEl = sp.children.length === 1 && sp.children[0].localName === "txEl" ? sp.children[0] : undefined;
      const pRg = txEl && txEl.children.length === 1 && txEl.children[0].localName === "pRg" ? txEl.children[0] : undefined;
      const st = Number(pRg?.getAttribute("st"));
      const end = Number(pRg?.getAttribute("end"));
      if (!pRg || !Number.isInteger(st) || !Number.isInteger(end) || st < 0 || end < st) {
        return undefined;
      }
      range = { st, end };
    }
    if (found && (found.spid !== spid || found.range?.st !== range?.st || found.range?.end !== range?.end)) {
      return undefined;
    }
    found = { spid, range };
  }
  return found;
}

/** The behaviour elements (direct children of the effect's `p:childTnLst`). */
function behaviours(cTn: Element, localName: string): Element[] {
  const childTnLst = child(cTn, "childTnLst");
  return childTnLst ? Array.from(childTnLst.children).filter((c) => c.localName === localName) : [];
}

function behaviourCtn(behaviour: Element): Element | undefined {
  const bhvr = child(behaviour, "cBhvr");
  return bhvr ? child(bhvr, "cTn") : undefined;
}

function attrNames(behaviour: Element): string[] {
  return descendants(behaviour, "attrName").map((n) => (n.textContent ?? "").trim());
}

/** The factor of a `#ppt_w*f` keyframe (a pulse / grow-shrink peak). */
function throbFactor(cTn: Element): number | undefined {
  for (const anim of behaviours(cTn, "anim")) {
    if (!attrNames(anim).includes("ppt_w")) {
      continue;
    }
    for (const str of descendants(anim, "strVal")) {
      const m = /^#ppt_w\*([\d.]+)$/.exec((str.getAttribute("val") ?? "").replace(/\s+/g, ""));
      if (m) {
        return Number(m[1]);
      }
    }
  }
  return undefined;
}

/** The emphasis/path parameters an effect's behaviours carry — undefined when
 * they are not the shape this table writes (a PowerPoint preset under the
 * same id that animates differently stays out of subset). Fills `animation`. */
function readEffectOptions(cTn: Element, animation: SlideAnimation, pathPreset: MotionPath["preset"] | undefined): boolean {
  switch (animation.effect) {
    case "spin": {
      const rot = behaviours(cTn, "animRot")[0];
      const deg = Number(rot?.getAttribute("by")) / 60000;
      if (!rot || !Number.isFinite(deg) || deg < MIN_SPIN_DEGREES || deg > MAX_SPIN_DEGREES) {
        return false;
      }
      if (deg !== DEFAULT_SPIN_DEGREES) animation.spinDegrees = deg;
      return true;
    }
    case "pulse":
      return throbFactor(cTn) !== undefined;
    case "growShrink": {
      const factor = throbFactor(cTn);
      const pct = factor === undefined ? NaN : Math.round(factor * 100);
      if (!Number.isFinite(pct) || pct < MIN_SCALE_PERCENT || pct > MAX_SCALE_PERCENT) {
        return false;
      }
      if (pct !== DEFAULT_SCALE_PERCENT) animation.scalePercent = pct;
      return true;
    }
    case "colorChange": {
      const clr = behaviours(cTn, "animClr")[0];
      const to = clr ? child(clr, "to") : undefined;
      const srgb = to ? Array.from(to.children).find((c) => c.localName === "srgbClr" && c.namespaceURI === A_NAMESPACE) : undefined;
      const val = srgb?.getAttribute("val") ?? "";
      if (!clr || behaviourCtn(clr)?.getAttribute("autoRev") !== "1" || !/^[0-9a-fA-F]{6}$/.test(val)) {
        return false;
      }
      const color = `#${val.toLowerCase()}`;
      if (color !== DEFAULT_EMPHASIS_COLOR) animation.emphasisColor = color;
      return true;
    }
    case "motionPath": {
      const motion = behaviours(cTn, "animMotion")[0];
      const commands = motion ? parseMotionPathD(motion.getAttribute("path") ?? "") : undefined;
      const path = commands && pathPreset ? inferMotionPath(pathPreset, commands) : undefined;
      if (!path) {
        return false;
      }
      animation.path = path;
      return true;
    }
    default:
      return true;
  }
}

function parseEffect(
  cTn: Element,
  opensGroup: boolean,
  groupIndex: number,
  resolveSpid: (spid: string) => string | undefined,
): WireStep {
  const presetClass = cTn.getAttribute("presetClass");
  const presetID = cTn.getAttribute("presetID") ?? "";
  const pathPreset = presetClass === "path" ? PATH_ID_TO_PRESET[presetID] : undefined;
  const effect = presetClass === "path" ? (pathPreset ? "motionPath" : undefined) : PRESET_TO_EFFECT[`${presetClass}:${presetID}`];
  const target = effectTarget(cTn);
  const frameId = target ? resolveSpid(target.spid) : undefined;
  if (!effect || !target || !frameId) {
    throw ABORT;
  }
  const animation: SlideAnimation = { frameId, effect, trigger: "click" };

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
  // `appear` is instant — its `p:set dur="1"` is a marker, not a duration. A
  // colour change's behaviours run half the duration and auto-reverse.
  if (effect !== "appear") {
    const dur = firstBehaviourDur(cTn);
    const ms = dur === undefined ? undefined : effect === "colorChange" ? dur * 2 : dur;
    if (ms !== undefined && ms !== DEFAULT_ANIMATION_DURATION_MS) {
      animation.durationMs = ms;
    }
  }
  if (!readEffectOptions(cTn, animation, pathPreset)) {
    throw ABORT;
  }
  const repeat = cTn.getAttribute("repeatCount");
  if (repeat === "indefinite") {
    // Repeat until the next click: indefinitely, ended by the next advance.
    const endCond = child(cTn, "endCondLst");
    if (!endCond || !Array.from(endCond.children).some((c) => c.getAttribute("evt") === "onNext")) {
      throw ABORT;
    }
    animation.repeatUntilClick = true;
  } else if (repeat !== null) {
    const n = Number(repeat) / 1000;
    if (!Number.isInteger(n) || n < 1 || n > MAX_REPEAT_COUNT) {
      throw ABORT;
    }
    if (n > 1) animation.repeatCount = n;
  }
  if (cTn.getAttribute("autoRev") === "1" || cTn.getAttribute("autoRev") === "true") {
    animation.autoReverse = true;
  }
  if (cTn.getAttribute("fill") === "remove") {
    animation.rewindOnComplete = true;
  }
  const id = Number(cTn.getAttribute("id"));
  return {
    animation,
    spid: target.spid,
    range: target.range,
    nodeType: cTn.getAttribute("nodeType") ?? "",
    grpId: cTn.getAttribute("grpId") ?? "0",
    order: cTn.hasAttribute("id") && Number.isInteger(id) ? id : undefined,
    opensGroup,
    groupIndex,
  };
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
