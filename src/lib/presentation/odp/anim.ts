// The odp animation wire mapping (SPEC-086) — the NORMATIVE source for how the
// modelled `SlideAnimation[]` subset reads from and writes to a page's `anim:*`
// SMIL tree. The canonical serialised bytes are pinned by `animGolden.ts`; the
// reader accepts those plus Impress-authored equivalents (reordered attributes,
// extra behaviour internals it ignores).
//
// SPEC-105 widens the subset — this module's preset tables are its normative
// artefact, pinned by `ANIM_DEPTH_GOLDEN`:
// - builds: each expanded step's behaviours target its paragraphs' `text:p`
//   ids (`de-f{n}-p{k}`, minted by odp/serialize.ts on a built frame's
//   paragraphs), which `resolveTarget` maps back to (frame, paragraph); the
//   reader also tolerates Impress's `<anim:iterate anim:iterate-type=
//   "by-paragraph">` as a by-paragraph build, after previous;
// - presets: `ooo-emphasis-{pulse|spin|grow-and-shrink|fill-color}` and
//   `ooo-motionpath-{line|arc|turn|loop|custom}`, over
//   `anim:animateTransform` (rotate / scale), `anim:animateColor` and
//   `anim:animateMotion svg:path` (slide fractions, `smil:path` tolerated);
// - repetition: `smil:repeatCount` (or "indefinite" ended by `smil:end=
//   "next"`), `smil:autoReverse` and `smil:fill="remove"` (rewind);
// - triggers: a sibling `<anim:seq presentation:node-type=
//   "interactive-sequence" smil:begin="de-f{n}.click">` per trigger frame.
//
// Anything the tables do not recognise, a target narrower than a whole
// frame or a build's paragraphs (the frame's text only, text by word or
// letter), an extra sequence that is not interactive with a resolvable
// trigger, or a path no preset reproduces makes the whole tree out of subset:
// the reader returns `undefined` and the caller keeps the source markup
// preserved verbatim (AC-4) — kept pointing at its targets across a rebuild
// (the "Preserved anim" section).
//
// Framework-free; imports the model types, the format-neutral animation
// helpers, and the odf attribute readers.
import { presentationAttr, smilAttr } from "../../core/odfXml";
import type { AnimationEffect, MotionPath, SlideAnimation, SlideFrame, TransitionDirection } from "../model";
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
  isTriggerableKind,
  MAX_ANIMATION_DELAY_MS,
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
import { formatPathD, formatPathNumber, inferMotionPath, parseMotionPathD, resolveMotionPath } from "../motionPaths";

type PresetClass = "entrance" | "exit" | "emphasis" | "motion-path";

/** ooo preset id + class per effect (a motion path's id is per preset —
 * `PATH_PRESET_ID`). */
const EFFECT_PRESET: Record<AnimationEffect, { id: string; cls: PresetClass }> = {
  appear: { id: "ooo-entrance-appear", cls: "entrance" },
  fadeIn: { id: "ooo-entrance-fade-in", cls: "entrance" },
  flyIn: { id: "ooo-entrance-fly-in", cls: "entrance" },
  wipeIn: { id: "ooo-entrance-wipe", cls: "entrance" },
  zoomIn: { id: "ooo-entrance-zoom", cls: "entrance" },
  fadeOut: { id: "ooo-exit-fade-out", cls: "exit" },
  flyOut: { id: "ooo-exit-fly-out", cls: "exit" },
  pulse: { id: "ooo-emphasis-pulse", cls: "emphasis" },
  spin: { id: "ooo-emphasis-spin", cls: "emphasis" },
  growShrink: { id: "ooo-emphasis-grow-and-shrink", cls: "emphasis" },
  colorChange: { id: "ooo-emphasis-fill-color", cls: "emphasis" },
  motionPath: { id: "ooo-motionpath-custom", cls: "motion-path" },
};
const PRESET_TO_EFFECT: Record<string, AnimationEffect> = Object.fromEntries(
  (Object.entries(EFFECT_PRESET) as [AnimationEffect, { id: string }][])
    .filter(([effect]) => effect !== "motionPath")
    .map(([effect, { id }]) => [id, effect]),
);

/** The motion-path preset ids (SPEC-105). The `svg:path` is normative: the
 * reader re-derives the preset's direction and distance from it. */
const PATH_PRESET_ID: Record<MotionPath["preset"], string> = {
  line: "ooo-motionpath-line",
  arc: "ooo-motionpath-arc",
  turn: "ooo-motionpath-turn",
  loop: "ooo-motionpath-loop",
  custom: "ooo-motionpath-custom",
};
const PATH_ID_TO_PRESET: Record<string, MotionPath["preset"]> = Object.fromEntries(
  (Object.entries(PATH_PRESET_ID) as [MotionPath["preset"], string][]).map(([preset, id]) => [id, preset]),
);

/** The scale a pulse throbs to and back. */
const PULSE_SCALE = 1.1;

const DIR_TO_SUB: Record<TransitionDirection, string> = {
  left: "from-left",
  right: "from-right",
  up: "from-top",
  down: "from-bottom",
};
const SUB_TO_DIR: Record<string, TransitionDirection> = {
  "from-left": "left",
  "from-right": "right",
  "from-top": "up",
  "from-bottom": "down",
};

const TRIGGER_FOR_NODE: Record<string, SlideAnimation["trigger"]> = {
  "on-click": "click",
  "with-previous": "withPrevious",
  "after-previous": "afterPrevious",
};

const BARWIPE_SUB: Record<TransitionDirection, string> = {
  left: "leftToRight",
  right: "leftToRight",
  up: "topToBottom",
  down: "topToBottom",
};

/** The `text:p` wire id of a built frame's paragraph: `de-f{n}-p{k}` (k is
 * the 0-based paragraph index). odp/serialize.ts mints it; the writer and
 * reader address paragraphs by it. */
export function paragraphWireId(frameWireId: string, paragraphIndex: number): string {
  return `${frameWireId}-p${paragraphIndex}`;
}

/** What a `smil:targetElement` resolves to: a frame id, or one of a frame's
 * paragraphs. */
export type OdpTarget = string | { frameId: string; paragraphIndex: number };

// ---------------------------------------------------------------------------
// Write
// ---------------------------------------------------------------------------

/** Milliseconds as an ODF SMIL duration in Impress's `0.5s` form. */
function secDur(ms: number): string {
  const seconds = ms / 1000;
  return `${Number.isInteger(seconds) ? seconds : Number(seconds.toFixed(3))}s`;
}

/** One behaviour set per target: the frame's, or each covered paragraph's. */
function behavioursFor(a: SlideAnimation, target: string): string {
  const durMs = effectiveAnimationDurationMs(a);
  const dur = secDur(durMs);
  const direction = effectiveAnimationDirection(a) ?? "left";
  switch (a.effect) {
    case "appear":
      return `<anim:set smil:attributeName="visibility" smil:to="visible" smil:targetElement="${target}"/>`;
    case "fadeIn":
      return `<anim:transitionFilter smil:type="fade" smil:direction="forward" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    case "fadeOut":
      return (
        `<anim:transitionFilter smil:type="fade" smil:direction="reverse" smil:dur="${dur}" smil:targetElement="${target}"/>` +
        `<anim:set smil:attributeName="visibility" smil:to="hidden" smil:targetElement="${target}"/>`
      );
    case "flyIn":
    case "flyOut": {
      const axis = direction === "left" || direction === "right" ? "x" : "y";
      const off = direction === "left" || direction === "up" ? "-100%" : "100%";
      const from = a.effect === "flyOut" ? "0%" : off;
      const to = a.effect === "flyOut" ? off : "0%";
      return `<anim:animate smil:attributeName="${axis}" smil:from="${from}" smil:to="${to}" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    }
    case "wipeIn":
      return `<anim:transitionFilter smil:type="barWipe" smil:subtype="${BARWIPE_SUB[direction]}" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    case "zoomIn":
      return `<anim:animateTransform smil:attributeName="transform" smil:type="scale" smil:from="0,0" smil:to="1,1" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    case "pulse":
    case "growShrink": {
      const f = formatPathNumber(a.effect === "pulse" ? PULSE_SCALE : effectiveScalePercent(a) / 100);
      return `<anim:animateTransform smil:attributeName="transform" smil:type="scale" smil:values="1,1;${f},${f};1,1" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    }
    case "spin":
      return `<anim:animateTransform smil:attributeName="transform" smil:type="rotate" smil:by="${effectiveSpinDegrees(a)}" smil:dur="${dur}" smil:targetElement="${target}"/>`;
    case "colorChange": {
      // To the colour over half the duration, auto-reversed over the rest.
      const half = secDur(Math.round(durMs / 2));
      const color = effectiveEmphasisColor(a);
      return (["color", "fill-color"] as const)
        .map(
          (attr) =>
            `<anim:animateColor smil:attributeName="${attr}" smil:to="${color}" smil:dur="${half}" smil:autoReverse="true" smil:targetElement="${target}"/>`,
        )
        .join("");
    }
    case "motionPath": {
      const commands = a.path ? resolveMotionPath(a.path) : undefined;
      return commands
        ? `<anim:animateMotion svg:path="${formatPathD(commands)}" smil:dur="${dur}" smil:targetElement="${target}"/>`
        : "";
    }
    default:
      return "";
  }
}

/** One step's effect `anim:par`. `xmlId` (the step's model order) is written
 * only on a page with an interactive sequence. */
function effectPar(step: BuildStep, frameTarget: string, nodeType: string, xmlId?: string): string {
  const a = step.animation;
  const preset =
    a.effect === "motionPath" ? { id: PATH_PRESET_ID[a.path?.preset ?? "line"], cls: "motion-path" } : EFFECT_PRESET[a.effect];
  const sub = a.effect === "zoomIn"
    ? ` presentation:preset-sub-type="in"`
    : usesAnimationDirection(a.effect)
      ? ` presentation:preset-sub-type="${DIR_TO_SUB[effectiveAnimationDirection(a) ?? "left"]}"`
      : "";
  const begin = secDur(effectiveAnimationDelayMs(a));
  const repeat = a.repeatUntilClick
    ? ` smil:repeatCount="indefinite" smil:end="next"`
    : a.repeatCount !== undefined
      ? ` smil:repeatCount="${effectiveRepeatCount(a)}"`
      : "";
  const autoReverse = a.autoReverse ? ` smil:autoReverse="true"` : "";
  const fill = a.rewindOnComplete ? ` smil:fill="remove"` : "";
  const id = xmlId !== undefined ? ` xml:id="${xmlId}"` : "";
  const targets =
    step.paragraphIndex === undefined
      ? [frameTarget]
      : Array.from({ length: (step.paragraphEnd ?? step.paragraphIndex) - step.paragraphIndex + 1 }, (_, k) =>
          paragraphWireId(frameTarget, step.paragraphIndex! + k),
        );
  return (
    `<anim:par presentation:node-type="${nodeType}" presentation:preset-class="${preset.cls}"` +
    ` presentation:preset-id="${preset.id}" smil:begin="${begin}"${sub}${repeat}${autoReverse}${fill}${id}>` +
    targets.map((t) => behavioursFor(a, t)).join("") +
    `</anim:par>`
  );
}

/** A step's `presentation:node-type` within its group: the first is the
 * click; a row's first step otherwise joins on its trigger (an interactive
 * row joins its predecessor), a later paragraph step on the row's
 * `buildTrigger`. */
function nodeTypeAt(group: BuildStep[], i: number, interactive: boolean): string {
  if (i === 0) {
    return "on-click";
  }
  const a = group[i].animation;
  if (group[i - 1].animation !== a) {
    return !interactive && a.trigger === "afterPrevious" ? "after-previous" : "with-previous";
  }
  return a.buildTrigger === "withPrevious" ? "with-previous" : "after-previous";
}

/** Serialises a slide's animations to its `anim:*` tree (the slot after the
 * page's content/notes). `targetFor` resolves a frame id to its wire id
 * (`de-f{n}`, minted on the frame's `xml:id`/`draw:id`; a built frame's
 * paragraphs carry `paragraphWireId`s); `frameFor` resolves it to the frame,
 * whose CURRENT paragraphs a build expands over (absent: every row plays
 * whole). Returns "" for none. */
export function writeOdpAnim(
  animations: SlideAnimation[] | undefined,
  targetFor: (frameId: string) => string | undefined,
  frameFor: (frameId: string) => SlideFrame | undefined = () => undefined,
): string {
  if (!animations || animations.length === 0) {
    return "";
  }
  const main = mainBuildSteps(animations, frameFor);
  const interactive = interactiveBuildSteps(animations, frameFor);
  // Effect node ids in MODEL order, only where an interactive sequence makes
  // document order (main sequence first) lose where an onClickOf row sits.
  const ids = new Map<BuildStep, string>();
  if (interactive.length > 0) {
    const stepsOf = new Map<SlideAnimation, BuildStep[]>();
    for (const step of [...main.flat(), ...interactive.flatMap((s) => s.groups.flat())]) {
      stepsOf.set(step.animation, [...(stepsOf.get(step.animation) ?? []), step]);
    }
    let next = 1;
    for (const animation of animations) {
      for (const step of stepsOf.get(animation) ?? []) {
        ids.set(step, `de-n${next++}`);
      }
    }
  }
  const groupXml = (groups: BuildStep[][], interactiveSeq: boolean): string =>
    groups
      .map((group, gi) => {
        const effects = group
          .map((step, i) => {
            const target = targetFor(step.frameId);
            return target ? effectPar(step, target, nodeTypeAt(group, i, interactiveSeq), ids.get(step)) : "";
          })
          .join("");
        const begin = interactiveSeq && gi === 0 ? "0s" : "next";
        return `<anim:par smil:begin="${begin}"><anim:par smil:begin="0s">${effects}</anim:par></anim:par>`;
      })
      .join("");
  const interactiveXml = interactive
    .map((sequence) => {
      const trigger = targetFor(sequence.triggerFrameId);
      return trigger
        ? `<anim:seq presentation:node-type="interactive-sequence" smil:begin="${trigger}.click">${groupXml(sequence.groups, true)}</anim:seq>`
        : "";
    })
    .join("");
  // Self-declare the anim/smil namespaces on the root so the tree is valid
  // wherever it is spliced — a re-read deck's content.xml root may not bind
  // them (SPEC-081's envelope only carries the prefixes its own bytes used).
  return (
    `<anim:par xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"` +
    ` xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"` +
    ` presentation:node-type="timing-root">` +
    `<anim:seq presentation:node-type="main-sequence">${groupXml(main, false)}</anim:seq>` +
    interactiveXml +
    `</anim:par>`
  );
}

// ---------------------------------------------------------------------------
// Read
// ---------------------------------------------------------------------------

function descendants(el: Element, localName: string): Element[] {
  return Array.from(el.getElementsByTagName("*")).filter((e) => e.localName === localName);
}

const ABORT = Symbol("out-of-subset");
const ANIM_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const SVG_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

/** Parses a page's collected `anim:*` children to modelled animations, or
 * `undefined` when the tree is out of subset (→ keep it preserved).
 * `resolveTarget` maps a `smil:targetElement` back to a frame id, or (a
 * built frame's `text:p` id) to that frame's paragraph; `frameFor` resolves a
 * frame id to its frame — paragraph targets are modelled as a build only
 * against its paragraphs (absent, they are out of subset). */
export function readOdpAnim(
  elements: Element[],
  resolveTarget: (wireId: string) => OdpTarget | undefined,
  frameFor?: (frameId: string) => SlideFrame | undefined,
): SlideAnimation[] | undefined {
  try {
    return parseAnim(elements, resolveTarget, frameFor);
  } catch (e) {
    if (e === ABORT) {
      return undefined;
    }
    throw e;
  }
}

/** One effect `anim:par` as read. `range` (paragraph targets) or `iterate`
 * (Impress's by-paragraph iterate over the frame) narrow it. */
interface WireStep {
  animation: SlideAnimation;
  range?: { st: number; end: number };
  iterateIntervalMs?: number;
  nodeType: string;
  order?: number;
  opensGroup: boolean;
  groupIndex: number;
}

function parseAnim(
  elements: Element[],
  resolveTarget: (wireId: string) => OdpTarget | undefined,
  frameFor: ((frameId: string) => SlideFrame | undefined) | undefined,
): SlideAnimation[] | undefined {
  if (elements.length !== 1) {
    return undefined;
  }
  const root = elements[0];
  if (root.localName !== "par" || presentationAttr(root, "node-type") !== "timing-root") {
    return undefined;
  }
  // The single main sequence stands; any other sequence must be interactive,
  // begun by a click on a frame of this page.
  const seqs = descendants(root, "seq");
  const mains = seqs.filter((s) => presentationAttr(s, "node-type") === "main-sequence");
  if (mains.length !== 1) {
    return undefined;
  }
  const interactive: { seq: Element; triggerFrameId: string }[] = [];
  for (const seq of seqs) {
    if (seq === mains[0]) {
      continue;
    }
    const begin = /^\s*([A-Za-z_][\w.-]*)\.click\s*$/.exec(smilAttr(seq, "begin") ?? "");
    const trigger = begin ? resolveTarget(begin[1]) : undefined;
    const triggerKind = typeof trigger === "string" ? frameFor?.(trigger)?.kind : undefined;
    if (
      presentationAttr(seq, "node-type") !== "interactive-sequence" ||
      typeof trigger !== "string" ||
      (triggerKind !== undefined && !isTriggerableKind(triggerKind))
    ) {
      return undefined;
    }
    interactive.push({ seq, triggerFrameId: trigger });
  }
  const rows: { animation: SlideAnimation; order?: number }[] = [];
  rows.push(...collapse(sequenceSteps(mains[0], resolveTarget), undefined, frameFor));
  for (const sequence of interactive) {
    rows.push(...collapse(sequenceSteps(sequence.seq, resolveTarget), sequence.triggerFrameId, frameFor));
  }
  const ordered = rows.length > 0 && rows.every((r) => r.order !== undefined)
    ? rows.toSorted((a, b) => (a.order ?? 0) - (b.order ?? 0))
    : rows;
  return ordered.length > 0 ? ordered.map((r) => r.animation) : undefined;
}

/** Every effect of a sequence's click groups, in document order. */
function sequenceSteps(seq: Element, resolveTarget: (wireId: string) => OdpTarget | undefined): WireStep[] {
  const steps: WireStep[] = [];
  Array.from(seq.children)
    .filter((c) => c.localName === "par")
    .forEach((groupPar, groupIndex) => {
      descendants(groupPar, "par")
        .filter((p) => presentationAttr(p, "preset-class"))
        .forEach((par, indexInGroup) => {
          steps.push(parseEffect(par, indexInGroup === 0, groupIndex, resolveTarget));
        });
    });
  return steps;
}

/** The comparable identity of a step's effect (everything but its place). */
function effectKey(a: SlideAnimation): string {
  const { trigger: _trigger, ...rest } = a;
  return JSON.stringify(rest, Object.keys(rest).toSorted());
}

function stepTrigger(step: WireStep): SlideAnimation["trigger"] | undefined {
  return step.opensGroup ? "click" : TRIGGER_FOR_NODE[step.nodeType];
}

/** Collapses a sequence's steps into rows: a whole-frame step is one row; an
 * Impress by-paragraph iterate is one by-paragraph row, after previous, its
 * interval the delay; a run of paragraph steps on one frame is one build row
 * only when their ranges are exactly a build mode's expansion over the
 * frame's paragraphs (in order, no gap or overlap), every step shares the
 * effect and the later steps agree on one build trigger. */
function collapse(
  steps: WireStep[],
  triggerFrameId: string | undefined,
  frameFor: ((frameId: string) => SlideFrame | undefined) | undefined,
): { animation: SlideAnimation; order?: number }[] {
  const rows: { animation: SlideAnimation; order?: number }[] = [];
  let i = 0;
  while (i < steps.length) {
    const first = steps[i];
    const row: SlideAnimation = { ...first.animation };
    if (triggerFrameId !== undefined) {
      if ((first.opensGroup && first.groupIndex !== 0) || (!first.opensGroup && first.nodeType !== "with-previous")) {
        throw ABORT;
      }
      if (triggerFrameId === row.frameId) {
        throw ABORT;
      }
      row.trigger = "onClickOf";
      row.triggerFrameId = triggerFrameId;
    } else {
      row.trigger = first.opensGroup ? "click" : TRIGGER_FOR_NODE[first.nodeType] ?? "click";
    }
    if (first.iterateIntervalMs !== undefined) {
      row.buildBy = "byParagraph";
      row.buildTrigger = "afterPrevious";
      if (first.iterateIntervalMs > 0) {
        row.delayMs = Math.min(MAX_ANIMATION_DELAY_MS, first.iterateIntervalMs);
      } else {
        delete row.delayMs;
      }
      rows.push({ animation: row, order: first.order });
      i += 1;
      continue;
    }
    if (!first.range) {
      rows.push({ animation: row, order: first.order });
      i += 1;
      continue;
    }
    const frame = frameFor?.(row.frameId);
    if (!buildableParagraphs(frame)) {
      throw ABORT;
    }
    const key = effectKey(first.animation);
    let matched = false;
    for (const mode of ["byParagraph", "byFirstLevelParagraph"] as const) {
      const expected = expandAnimation({ ...row, buildBy: mode }, frame);
      const run = steps.slice(i, i + expected.length);
      if (
        run.length !== expected.length ||
        expected.some(
          (e, k) =>
            e.paragraphIndex === undefined ||
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
      const buildTrigger = run.length > 1 ? stepTrigger(run[1]) : "click";
      if (buildTrigger && buildTrigger !== "click" && buildTrigger !== "onClickOf") {
        row.buildTrigger = buildTrigger;
      }
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

/** Whether a behaviour animates its whole target (no `anim:sub-item`, or
 * "whole"): narrowed to a frame's text or background it animates part of the
 * frame the model cannot express. */
function wholeSubItem(el: Element, allowText = false): boolean {
  const subItem = el.getAttributeNS(ANIM_NAMESPACE, "sub-item");
  return !subItem || subItem === "whole" || (allowText && subItem === "text");
}

/** The behaviour elements of an effect par (its `anim:*` children that are
 * not nested pars/seqs). */
function behaviourElements(par: Element): Element[] {
  return Array.from(par.children).filter((c) => c.localName !== "par" && c.localName !== "seq");
}

/** An `anim:animateTransform`'s transform type, either spelling. */
function transformType(el: Element): string | null {
  return smilAttr(el, "type") ?? el.getAttributeNS(SVG_NAMESPACE, "type");
}

/** The scale of a `smil:values="1,1;f,f;1,1"` throb (pulse, grow/shrink). */
function throbScale(behaviours: Element[]): number | undefined {
  const scale = behaviours.find((b) => b.localName === "animateTransform" && transformType(b) === "scale");
  const m = /^\s*1,1;\s*([\d.]+),\s*([\d.]+);\s*1,1\s*$/.exec(scale ? (smilAttr(scale, "values") ?? "") : "");
  return m && m[1] === m[2] ? Number(m[1]) : undefined;
}

/** Fills the emphasis/path parameters from the behaviours — false when they
 * are not the shape this table writes (an Impress preset that animates
 * differently stays out of subset). */
function readEffectOptions(behaviours: Element[], animation: SlideAnimation, pathPreset: MotionPath["preset"] | undefined): boolean {
  switch (animation.effect) {
    case "spin": {
      const rot = behaviours.find((b) => b.localName === "animateTransform" && transformType(b) === "rotate");
      const deg = Number(rot ? smilAttr(rot, "by") : NaN);
      if (!Number.isFinite(deg) || deg < MIN_SPIN_DEGREES || deg > MAX_SPIN_DEGREES) {
        return false;
      }
      if (deg !== DEFAULT_SPIN_DEGREES) animation.spinDegrees = deg;
      return true;
    }
    case "pulse":
      return throbScale(behaviours) !== undefined;
    case "growShrink": {
      const factor = throbScale(behaviours);
      const pct = factor === undefined ? NaN : Math.round(factor * 100);
      if (!Number.isFinite(pct) || pct < MIN_SCALE_PERCENT || pct > MAX_SCALE_PERCENT) {
        return false;
      }
      if (pct !== DEFAULT_SCALE_PERCENT) animation.scalePercent = pct;
      return true;
    }
    case "colorChange": {
      const clr = behaviours.find((b) => b.localName === "animateColor");
      const to = (clr ? smilAttr(clr, "to") : null) ?? "";
      if (!clr || smilAttr(clr, "autoReverse") !== "true" || !/^#[0-9a-fA-F]{6}$/.test(to)) {
        return false;
      }
      const color = to.toLowerCase();
      if (color !== DEFAULT_EMPHASIS_COLOR) animation.emphasisColor = color;
      return true;
    }
    case "motionPath": {
      const motion = behaviours.find((b) => b.localName === "animateMotion");
      const d = motion ? (motion.getAttributeNS(SVG_NAMESPACE, "path") ?? smilAttr(motion, "path")) : null;
      const commands = d ? parseMotionPathD(d) : undefined;
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
  par: Element,
  opensGroup: boolean,
  groupIndex: number,
  resolveTarget: (wireId: string) => OdpTarget | undefined,
): WireStep {
  const presetId = presentationAttr(par, "preset-id") ?? "";
  const pathPreset = PATH_ID_TO_PRESET[presetId];
  const effect: AnimationEffect | undefined = pathPreset ? "motionPath" : PRESET_TO_EFFECT[presetId];
  if (!effect) {
    throw ABORT;
  }
  let behaviours = behaviourElements(par);
  // Impress's by-paragraph iterate: the behaviours inside, the target on it.
  let iterateIntervalMs: number | undefined;
  let targets: OdpTarget[];
  if (behaviours.length === 1 && behaviours[0].localName === "iterate") {
    const iterate = behaviours[0];
    const target = resolveTarget(smilAttr(iterate, "targetElement") ?? "");
    if (iterate.getAttributeNS(ANIM_NAMESPACE, "iterate-type") !== "by-paragraph" || typeof target !== "string" || !wholeSubItem(iterate, true)) {
      throw ABORT;
    }
    iterateIntervalMs = parseSeconds(iterate.getAttributeNS(ANIM_NAMESPACE, "iterate-interval")) ?? 0;
    behaviours = Array.from(iterate.children);
    targets = [target];
  } else {
    targets = behaviours.map((b) => {
      const target = resolveTarget(smilAttr(b, "targetElement") ?? "");
      if (target === undefined || !wholeSubItem(b)) {
        throw ABORT;
      }
      return target;
    });
  }
  if (behaviours.length === 0 || behaviours.some((b) => b.localName === "iterate" || !wholeSubItem(b, iterateIntervalMs !== undefined))) {
    throw ABORT;
  }
  // Every target on one frame: the frame itself, or a contiguous paragraph run.
  const frameIds = new Set(targets.map((t) => (typeof t === "string" ? t : t.frameId)));
  if (frameIds.size !== 1) {
    throw ABORT;
  }
  const frameId = [...frameIds][0];
  let range: { st: number; end: number } | undefined;
  const paragraphs = targets.filter((t): t is { frameId: string; paragraphIndex: number } => typeof t !== "string");
  if (paragraphs.length > 0) {
    if (paragraphs.length !== targets.length) {
      throw ABORT;
    }
    const indices = [...new Set(paragraphs.map((p) => p.paragraphIndex))].toSorted((a, b) => a - b);
    if (indices.some((v, k) => v !== indices[0] + k)) {
      throw ABORT;
    }
    range = { st: indices[0], end: indices[indices.length - 1] };
  }
  const animation: SlideAnimation = { frameId, effect, trigger: "click" };
  if (usesAnimationDirection(effect)) {
    const dir = SUB_TO_DIR[presentationAttr(par, "preset-sub-type") ?? ""];
    if (dir && dir !== "left") {
      animation.direction = dir;
    }
  }
  const delay = parseSeconds(smilAttr(par, "begin"));
  if (delay !== undefined && delay > 0) {
    animation.delayMs = delay;
  }
  if (effect !== "appear") {
    const dur = parseSeconds(smilAttr(behaviours[0], "dur"));
    const ms = dur === undefined ? undefined : effect === "colorChange" ? dur * 2 : dur;
    if (ms !== undefined && ms !== DEFAULT_ANIMATION_DURATION_MS) {
      animation.durationMs = ms;
    }
  }
  if (!readEffectOptions(behaviours, animation, pathPreset)) {
    throw ABORT;
  }
  const repeat = smilAttr(par, "repeatCount");
  if (repeat === "indefinite") {
    if ((smilAttr(par, "end") ?? "").trim() !== "next") {
      throw ABORT;
    }
    animation.repeatUntilClick = true;
  } else if (repeat !== null) {
    const n = Number(repeat);
    if (!Number.isInteger(n) || n < 1 || n > MAX_REPEAT_COUNT) {
      throw ABORT;
    }
    if (n > 1) animation.repeatCount = n;
  }
  if (smilAttr(par, "autoReverse") === "true") {
    animation.autoReverse = true;
  }
  if (smilAttr(par, "fill") === "remove") {
    animation.rewindOnComplete = true;
  }
  const order = /^de-n(\d+)$/.exec(par.getAttributeNS("http://www.w3.org/XML/1998/namespace", "id") ?? "");
  return {
    animation,
    range,
    iterateIntervalMs,
    nodeType: presentationAttr(par, "node-type") ?? "",
    order: order ? Number(order[1]) : undefined,
    opensGroup,
    groupIndex,
  };
}

/** Parses a SMIL duration/offset (`0.5s`, `250ms`, `2s`) to milliseconds. */
function parseSeconds(value: string | null): number | undefined {
  if (!value) {
    return undefined;
  }
  const ms = /^([\d.]+)ms$/.exec(value.trim());
  if (ms) {
    const n = Number(ms[1]);
    return Number.isFinite(n) ? Math.round(n) : undefined;
  }
  const s = /^([\d.]+)s$/.exec(value.trim());
  if (s) {
    const n = Number(s[1]);
    return Number.isFinite(n) ? Math.round(n * 1000) : undefined;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Preserved (out-of-subset) anim
// ---------------------------------------------------------------------------
// A preserved `anim:*` tree re-emits verbatim and names its targets by SOURCE
// element id. ODF ids are document-wide and the source's are already unique,
// so a rebuilt page re-emits each target's own id (`sourceWireId`, on frames
// and on `text:p` — an Impress by-paragraph build) rather than minting new
// ones; the tree is kept only while every id it references is still defined
// on the page. These scanners work on the preserved string, prefix-agnostic
// (a model stored before the anim was kept byte-verbatim spells the SMIL
// attributes with generated `ns1:`-style prefixes).

/** `name="value"` (either quote) for an attribute whose LOCAL name is one of
 * `names`, under any prefix or none. */
function attrValues(xml: string, names: string): string[] {
  const re = new RegExp(`\\s(?:[\\w.-]+:)?(?:${names})\\s*=\\s*(["'])(.*?)\\1`, "g");
  return Array.from(xml.matchAll(re), (m) => m[2]);
}

/** The element ids a preserved anim string references: every
 * `smil:targetElement` (a shape, or a `text:p` for a by-paragraph build) and
 * `presentation:master-element`, plus the id half of an id-qualified
 * `smil:begin`/`smil:end` value (`id3.click` — an interactive trigger shape;
 * `id7.end` — another node). */
export function preservedAnimReferences(animXml: string): Set<string> {
  const refs = new Set(attrValues(animXml, "targetElement|master-element"));
  for (const value of attrValues(animXml, "begin|end")) {
    for (const part of value.split(";")) {
      const id = /^\s*([A-Za-z_][\w-]*)\.[A-Za-z]/.exec(part)?.[1];
      if (id) {
        refs.add(id);
      }
    }
  }
  return refs;
}

/** Every element id an XML fragment defines (`xml:id`, `draw:id`, `anim:id`,
 * …) — what a preserved anim's references may name on a rebuilt page. */
export function xmlIdsDefinedIn(xml: string): Set<string> {
  return new Set(attrValues(xml, "id"));
}
