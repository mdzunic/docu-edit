// The odp animation wire mapping (SPEC-086) — the NORMATIVE source for how the
// modelled `SlideAnimation[]` subset reads from and writes to a page's `anim:*`
// SMIL tree. The canonical serialised bytes are pinned by `animGolden.ts`; the
// reader accepts those plus Impress-authored equivalents (reordered attributes,
// extra behaviour internals it ignores). Anything the preset table does not
// recognise makes the whole tree out of subset: the reader returns `undefined`
// and the caller keeps the source markup preserved verbatim (AC-4).
//
// Framework-free; imports the model types, the format-neutral animation
// helpers, and the odf attribute readers.
import { presentationAttr, smilAttr } from "../../core/odfXml";
import type { AnimationEffect, SlideAnimation, TransitionDirection } from "../model";
import {
  buildGroups,
  DEFAULT_ANIMATION_DURATION_MS,
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
  usesAnimationDirection,
} from "../animations";

/** ooo preset id + class per effect. */
const EFFECT_PRESET: Record<AnimationEffect, { id: string; cls: "entrance" | "exit" }> = {
  appear: { id: "ooo-entrance-appear", cls: "entrance" },
  fadeIn: { id: "ooo-entrance-fade-in", cls: "entrance" },
  flyIn: { id: "ooo-entrance-fly-in", cls: "entrance" },
  wipeIn: { id: "ooo-entrance-wipe", cls: "entrance" },
  zoomIn: { id: "ooo-entrance-zoom", cls: "entrance" },
  fadeOut: { id: "ooo-exit-fade-out", cls: "exit" },
  flyOut: { id: "ooo-exit-fly-out", cls: "exit" },
};
const PRESET_TO_EFFECT: Record<string, AnimationEffect> = Object.fromEntries(
  (Object.entries(EFFECT_PRESET) as [AnimationEffect, { id: string }][]).map(([effect, { id }]) => [id, effect]),
);

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

const NODE_TYPE: Record<SlideAnimation["trigger"], string> = {
  click: "on-click",
  withPrevious: "with-previous",
  afterPrevious: "after-previous",
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

// ---------------------------------------------------------------------------
// Write
// ---------------------------------------------------------------------------

/** Milliseconds as an ODF SMIL duration in Impress's `0.5s` form. */
function secDur(ms: number): string {
  const seconds = ms / 1000;
  return `${Number.isInteger(seconds) ? seconds : Number(seconds.toFixed(3))}s`;
}

function behavioursFor(a: SlideAnimation, target: string): string {
  const dur = secDur(effectiveAnimationDurationMs(a));
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
    default:
      return "";
  }
}

function effectPar(a: SlideAnimation, target: string, nodeType: string): string {
  const preset = EFFECT_PRESET[a.effect];
  const sub = a.effect === "zoomIn"
    ? ` presentation:preset-sub-type="in"`
    : usesAnimationDirection(a.effect)
      ? ` presentation:preset-sub-type="${DIR_TO_SUB[effectiveAnimationDirection(a) ?? "left"]}"`
      : "";
  const begin = secDur(effectiveAnimationDelayMs(a));
  return (
    `<anim:par presentation:node-type="${nodeType}" presentation:preset-class="${preset.cls}"` +
    ` presentation:preset-id="${preset.id}" smil:begin="${begin}"${sub}>` +
    behavioursFor(a, target) +
    `</anim:par>`
  );
}

/** Serialises a slide's animations to its `anim:*` tree (the slot after the
 * page's content/notes). `targetFor` resolves a frame id to its wire id
 * (`de-f{n}`, minted on the frame's `xml:id`/`draw:id`). Returns "" for none. */
export function writeOdpAnim(
  animations: SlideAnimation[] | undefined,
  targetFor: (frameId: string) => string | undefined,
): string {
  if (!animations || animations.length === 0) {
    return "";
  }
  const groups = buildGroups(animations);
  const groupXml = groups
    .map((group) => {
      const effects = group
        .map((a, i) => {
          const target = targetFor(a.frameId);
          if (!target) {
            return "";
          }
          const nodeType = i === 0 ? "on-click" : NODE_TYPE[a.trigger];
          return effectPar(a, target, nodeType);
        })
        .join("");
      return `<anim:par smil:begin="next"><anim:par smil:begin="0s">${effects}</anim:par></anim:par>`;
    })
    .join("");
  // Self-declare the anim/smil namespaces on the root so the tree is valid
  // wherever it is spliced — a re-read deck's content.xml root may not bind
  // them (SPEC-081's envelope only carries the prefixes its own bytes used).
  return (
    `<anim:par xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"` +
    ` xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"` +
    ` presentation:node-type="timing-root">` +
    `<anim:seq presentation:node-type="main-sequence">${groupXml}</anim:seq>` +
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

/** Parses a page's collected `anim:*` children to modelled animations, or
 * `undefined` when the tree is out of subset (→ keep it preserved).
 * `resolveTarget` maps a `smil:targetElement` back to a frame id. */
export function readOdpAnim(
  elements: Element[],
  resolveTarget: (wireId: string) => string | undefined,
): SlideAnimation[] | undefined {
  try {
    return parseAnim(elements, resolveTarget);
  } catch (e) {
    if (e === ABORT) {
      return undefined;
    }
    throw e;
  }
}

function parseAnim(
  elements: Element[],
  resolveTarget: (wireId: string) => string | undefined,
): SlideAnimation[] | undefined {
  if (elements.length !== 1) {
    return undefined;
  }
  const root = elements[0];
  if (root.localName !== "par" || presentationAttr(root, "node-type") !== "timing-root") {
    return undefined;
  }
  const seqs = descendants(root, "seq").filter((s) => presentationAttr(s, "node-type") === "main-sequence");
  if (seqs.length !== 1) {
    return undefined;
  }
  const groupPars = Array.from(seqs[0].children).filter((c) => c.localName === "par");
  const animations: SlideAnimation[] = [];
  for (const groupPar of groupPars) {
    const effectPars = descendants(groupPar, "par").filter((p) => presentationAttr(p, "preset-class"));
    effectPars.forEach((par, indexInGroup) => {
      const parsed = parseEffect(par, indexInGroup === 0, resolveTarget);
      if (!parsed) {
        throw ABORT;
      }
      animations.push(parsed);
    });
  }
  return animations.length > 0 ? animations : undefined;
}

function parseEffect(
  par: Element,
  firstInGroup: boolean,
  resolveTarget: (wireId: string) => string | undefined,
): SlideAnimation | undefined {
  const effect = PRESET_TO_EFFECT[presentationAttr(par, "preset-id") ?? ""];
  if (!effect) {
    return undefined;
  }
  const behaviour = firstBehaviour(par);
  const wireId = behaviour ? smilAttr(behaviour, "targetElement") : null;
  const frameId = wireId ? resolveTarget(wireId) : undefined;
  if (!frameId) {
    return undefined;
  }
  const trigger = firstInGroup ? "click" : TRIGGER_FOR_NODE[presentationAttr(par, "node-type") ?? ""] ?? "click";
  const animation: SlideAnimation = { frameId, effect, trigger };

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
    const dur = parseSeconds(behaviour ? smilAttr(behaviour, "dur") : null);
    if (dur !== undefined && dur !== DEFAULT_ANIMATION_DURATION_MS) {
      animation.durationMs = dur;
    }
  }
  return animation;
}

/** The first behaviour element inside an effect par (an `anim:*` child that is
 * not a nested par/seq). */
function firstBehaviour(par: Element): Element | undefined {
  return Array.from(par.children).find((c) => c.localName !== "par" && c.localName !== "seq");
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
