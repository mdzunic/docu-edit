// The deck designer's plan (SPEC-115): the only shapes a model's reply may
// take — a `DeckPlan` for a new deck, a `DeckRedesign` for the open one — and
// the deterministic validators that turn untrusted JSON into one of them.
//
// The design vocabulary is closed and comes from the product, never from the
// model: layout kinds are SPEC-089's standard set (`STANDARD_LAYOUTS`, i.e.
// the `SlideLayoutKind` union minus "other"), theme names are SPEC-094's
// `BUILT_IN_THEMES`, slide sizes are SPEC-090's presets as SPEC-106 names them
// (`DECK_SIZE_NAMES`). A value outside it is replaced by the nearest default
// and reported as a `DeckDropRecord`; only a reply that is not a plan at all
// is rejected (a `DocumentGenerationError` naming the path), which is what
// sends it back for its one repair.
//
// Product-neutral in shape and framework-free: imports lib/core,
// lib/presentation and lib/generate's deck modules only (boundary.spec.ts).
import { BUILT_IN_THEMES, DEFAULT_THEME } from "../core/theme";
import { STANDARD_LAYOUTS } from "../presentation/layouts";
import { DEFAULT_PRESENTATION_TITLE, type SlideLayoutKind } from "../presentation/model";
import { MAX_INDENT_LEVEL } from "../presentation/text";
import { slideLocation, type DeckDropRecord } from "../generate/deckDrops";
import { DECK_SIZE_NAMES, type DeckSizeName } from "../generate/deckSlides";
import { DocumentGenerationError } from "../generate/errors";

export { planToDescription } from "./planToDescription";

/** A layout kind a plan may name: SPEC-089's union, never "other". */
export type PlanLayoutKind = Exclude<SlideLayoutKind, "other">;

/** A bullet: a plain string is level 0; the object form carries SPEC-106's
 * indent level (0..MAX_INDENT_LEVEL). */
export type PlannedBullet = string | { text: string; level: number };

export interface PlannedSlide {
  layout: PlanLayoutKind;
  title: string;
  /** 0..MAX_PLANNED_BULLETS, each at most MAX_BULLET_CHARS characters. */
  bullets: PlannedBullet[];
  note?: string;
}

export interface DeckPlan {
  title: string;
  /** A BUILT_IN_THEMES name, in the registry's spelling. */
  theme: string;
  size: DeckSizeName;
  /** 1..MAX_PLANNED_SLIDES. */
  slides: PlannedSlide[];
}

/** Redesign: the open deck's slides are fixed and the model only arranges
 * them. */
export interface RedesignedSlide {
  /** 0-based index of an existing slide; each index exactly once. */
  source: number;
  /** null = keep the slide's current layout. */
  layout: PlanLayoutKind | null;
}

export interface DeckRedesign {
  /** A BUILT_IN_THEMES name; null = keep the deck's theme. */
  theme: string | null;
  /** Exactly slideCount entries after validation. */
  slides: RedesignedSlide[];
}

export const MAX_PLANNED_SLIDES = 60;
export const MAX_PLANNED_BULLETS = 8;
export const MAX_BULLET_CHARS = 200;
export const DEFAULT_PLAN_SIZE: DeckSizeName = "16:9";

/** The layout kinds, in the registry's order — read from STANDARD_LAYOUTS,
 * never restated, so the vocabulary cannot drift from the product. */
export const PLAN_LAYOUT_KINDS: readonly PlanLayoutKind[] = Object.freeze(
  STANDARD_LAYOUTS.map((def) => def.kind).filter((kind): kind is PlanLayoutKind => kind !== "other"),
);

/** The theme names, in the registry's order. */
export const PLAN_THEME_NAMES: readonly string[] = Object.freeze(BUILT_IN_THEMES.map((theme) => theme.name));

/** The slide sizes a plan may name. */
export const PLAN_SIZES: readonly DeckSizeName[] = DECK_SIZE_NAMES;

export function isPlanLayoutKind(value: unknown): value is PlanLayoutKind {
  return typeof value === "string" && (PLAN_LAYOUT_KINDS as readonly string[]).includes(value);
}

/** The registry spelling of a built-in theme name, matched ignoring letter
 * case and surrounding whitespace; undefined for any other value. */
export function normaliseThemeName(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const wanted = value.trim().toLowerCase();
  return PLAN_THEME_NAMES.find((name) => name.toLowerCase() === wanted);
}

/** The "nearest default" layout for a slide whose layout is unknown. */
export function defaultLayoutFor(bulletCount: number): PlanLayoutKind {
  return bulletCount > 0 ? "titleAndBody" : "titleOnly";
}

export function bulletText(bullet: PlannedBullet): string {
  return typeof bullet === "string" ? bullet : bullet.text;
}

export function bulletLevel(bullet: PlannedBullet): number {
  return typeof bullet === "string" ? 0 : bullet.level;
}

/** A bullet in its canonical form: a plain string at level 0. */
export function makeBullet(text: string, level: number): PlannedBullet {
  return level > 0 ? { text, level } : text;
}

/** A short quotation of `text` for a drop record's reason. */
function quote(text: string): string {
  const chars = Array.from(text);
  return `“${chars.length > 60 ? `${chars.slice(0, 60).join("")}…` : text}”`;
}

// --- validatePlan -------------------------------------------------------------

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** A field that defaults when absent: JSON has no `undefined`, so `null`
 * counts as absent too. Present and not a string is a rejection. */
function optionalString(value: unknown, path: string): string | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new DocumentGenerationError(`${path}: expected a string`);
  }
  return value;
}

/** Cuts `text` to MAX_BULLET_CHARS characters (code points, so a surrogate
 * pair is never split). */
function capBullet(text: string): { text: string; cut: boolean } {
  const chars = Array.from(text);
  return chars.length > MAX_BULLET_CHARS
    ? { text: chars.slice(0, MAX_BULLET_CHARS).join(""), cut: true }
    : { text, cut: false };
}

function validateBullets(raw: unknown, path: string, location: string, dropped: DeckDropRecord[]): PlannedBullet[] {
  if (raw === undefined || raw === null) {
    return [];
  }
  if (!Array.isArray(raw)) {
    dropped.push({ kind: "bullet", location, reason: "the bullets were not a list, so none were kept" });
    return [];
  }
  const bullets: PlannedBullet[] = [];
  raw.forEach((entry, j) => {
    const bulletPath = `${path}[${j}]`;
    let text: string;
    if (typeof entry === "string") {
      text = entry;
    } else if (isPlainObject(entry)) {
      if (typeof entry.text !== "string") {
        throw new DocumentGenerationError(`${bulletPath}.text: expected a string`);
      }
      text = entry.text;
    } else {
      throw new DocumentGenerationError(`${bulletPath}: expected a string or { text, level }`);
    }
    if (j >= MAX_PLANNED_BULLETS) {
      dropped.push({
        kind: "bullet",
        location,
        reason: `a slide holds at most ${MAX_PLANNED_BULLETS} bullets; ${quote(text)} dropped`,
      });
      return;
    }
    let level = 0;
    const rawLevel = isPlainObject(entry) ? entry.level : undefined;
    if (rawLevel !== undefined && rawLevel !== null) {
      const requested = typeof rawLevel === "number" && Number.isFinite(rawLevel) ? Math.trunc(rawLevel) : 0;
      level = Math.min(Math.max(requested, 0), MAX_INDENT_LEVEL);
      if (level !== rawLevel) {
        dropped.push({
          kind: "bullet",
          location,
          reason: `bullet ${j + 1}'s level ${JSON.stringify(rawLevel)} is outside 0–${MAX_INDENT_LEVEL}; used ${level}`,
        });
      }
    }
    const capped = capBullet(text);
    if (capped.cut) {
      dropped.push({
        kind: "bullet",
        location,
        reason: `bullet ${j + 1} cut at ${MAX_BULLET_CHARS} characters`,
      });
    }
    bullets.push(makeBullet(capped.text, level));
  });
  return bullets;
}

function validatePlannedSlide(raw: unknown, index: number, dropped: DeckDropRecord[]): PlannedSlide {
  const path = `slides[${index}]`;
  const location = slideLocation(index);
  if (!isPlainObject(raw)) {
    throw new DocumentGenerationError(`${path}: expected an object`);
  }
  const title = optionalString(raw.title, `${path}.title`) ?? "";
  const note = optionalString(raw.note, `${path}.note`);
  let bullets = validateBullets(raw.bullets, `${path}.bullets`, location, dropped);

  let layout: PlanLayoutKind;
  const requested = raw.layout;
  const generatedKind =
    requested === "chart" || raw.chart !== undefined ? "chart" : requested === "table" || raw.table !== undefined ? "table" : null;
  if (generatedKind !== null) {
    // SPEC-091/092 own charts and tables; a plan asking for one gets a
    // bullet slide and a drop record.
    layout = "titleAndBody";
    dropped.push({
      kind: generatedKind,
      location,
      reason: `a ${generatedKind} cannot be generated; the slide is a titleAndBody slide instead`,
    });
  } else if (isPlanLayoutKind(requested)) {
    layout = requested;
  } else {
    layout = defaultLayoutFor(bullets.length);
    dropped.push({
      kind: "layout",
      location,
      reason:
        requested === undefined || requested === null
          ? `no layout named; used ${layout}`
          : `unknown layout ${JSON.stringify(requested)}; used ${layout}`,
    });
  }

  if (layout === "title" && bullets.length > 1) {
    // A title slide shows one line under its title — its subtitle.
    for (const extra of bullets.slice(1)) {
      dropped.push({
        kind: "bullet",
        location,
        reason: `a title slide keeps one line under its title; ${quote(bulletText(extra))} dropped`,
      });
    }
    bullets = bullets.slice(0, 1);
  }

  const slide: PlannedSlide = { layout, title, bullets };
  if (note !== undefined && note.trim() !== "") {
    slide.note = note;
  }
  return slide;
}

/** Validates a model's (or an edited) plan. Defaults rather than rejects
 * wherever a default exists — every closed-vocabulary substitution and every
 * cut recorded — and rejects (throwing `DocumentGenerationError`, which takes
 * the one repair) only a value that is not an object; `slides` missing, not
 * an array or empty; a slide that is not an object; or a title, bullet text or
 * note present but not a string. Unknown keys are ignored. */
export function validatePlan(value: unknown): { plan: DeckPlan; dropped: DeckDropRecord[] } {
  if (!isPlainObject(value)) {
    throw new DocumentGenerationError("plan: expected a JSON object");
  }
  if (!Array.isArray(value.slides) || value.slides.length === 0) {
    throw new DocumentGenerationError(
      value.slides === undefined ? "slides: missing; expected a non-empty array of slides" : "slides: expected a non-empty array of slides",
    );
  }
  const dropped: DeckDropRecord[] = [];

  const title = (optionalString(value.title, "title") ?? "").trim() || DEFAULT_PRESENTATION_TITLE;

  let theme = normaliseThemeName(value.theme);
  if (theme === undefined) {
    theme = DEFAULT_THEME.name;
    dropped.push({
      kind: "theme",
      location: "deck",
      reason:
        value.theme === undefined || value.theme === null
          ? `no theme named; used ${theme}`
          : `unknown theme ${JSON.stringify(value.theme)}; used ${theme}`,
    });
  }

  let size: DeckSizeName;
  if (typeof value.size === "string" && (PLAN_SIZES as readonly string[]).includes(value.size.trim())) {
    size = value.size.trim() as DeckSizeName;
  } else {
    size = DEFAULT_PLAN_SIZE;
    dropped.push({
      kind: "slideSize",
      location: "deck",
      reason:
        value.size === undefined || value.size === null
          ? `no slide size named; used ${size}`
          : `unknown slide size ${JSON.stringify(value.size)}; used ${size}`,
    });
  }

  const rawSlides = value.slides as unknown[];
  if (rawSlides.length > MAX_PLANNED_SLIDES) {
    dropped.push({
      kind: "slideCount",
      location: "deck",
      reason: `${rawSlides.length} slides planned; a plan holds at most ${MAX_PLANNED_SLIDES}, so slides ${MAX_PLANNED_SLIDES + 1}–${rawSlides.length} were cut`,
    });
  }
  const slides = rawSlides.slice(0, MAX_PLANNED_SLIDES).map((raw, index) => validatePlannedSlide(raw, index, dropped));
  return { plan: { title, theme, size, slides }, dropped };
}

// --- validateRedesign ---------------------------------------------------------

/** Validates a redesign reply against the open deck's `slideCount`. Never
 * rejects for the count: an entry whose `source` is not an integer in
 * 0..slideCount-1, or repeats an earlier entry's `source`, is discarded; an
 * unknown (or "other") layout and an unknown theme become "keep current";
 * every index the reply omits is appended in ascending order with its layout
 * kept — each recorded. The result always has exactly `slideCount` entries,
 * every index once. Rejects only a value that is not an object, or whose
 * `slides` is not an array. */
export function validateRedesign(value: unknown, slideCount: number): { redesign: DeckRedesign; dropped: DeckDropRecord[] } {
  if (!isPlainObject(value)) {
    throw new DocumentGenerationError("redesign: expected a JSON object");
  }
  if (!Array.isArray(value.slides)) {
    throw new DocumentGenerationError(
      value.slides === undefined ? "slides: missing; expected an array of { source, layout }" : "slides: expected an array of { source, layout }",
    );
  }
  const dropped: DeckDropRecord[] = [];

  let theme: string | null = null;
  if (value.theme !== undefined && value.theme !== null) {
    theme = normaliseThemeName(value.theme) ?? null;
    if (theme === null) {
      dropped.push({ kind: "theme", location: "deck", reason: `unknown theme ${JSON.stringify(value.theme)}; kept the current theme` });
    }
  }

  const seen = new Set<number>();
  const slides: RedesignedSlide[] = [];
  (value.slides as unknown[]).forEach((entry, i) => {
    const source = isPlainObject(entry) ? entry.source : undefined;
    if (typeof source !== "number" || !Number.isInteger(source) || source < 0 || source >= slideCount) {
      dropped.push({
        kind: "slideOrder",
        location: "deck",
        reason: `entry ${i + 1} names ${source === undefined ? "no slide" : `slide index ${JSON.stringify(source)}`}, which is not a slide of this deck; ignored`,
      });
      return;
    }
    if (seen.has(source)) {
      dropped.push({ kind: "slideOrder", location: slideLocation(source), reason: `named more than once; entry ${i + 1} ignored` });
      return;
    }
    seen.add(source);
    const requested = (entry as Record<string, unknown>).layout;
    let layout: PlanLayoutKind | null = null;
    if (isPlanLayoutKind(requested)) {
      layout = requested;
    } else if (requested !== undefined && requested !== null) {
      dropped.push({
        kind: "layout",
        location: slideLocation(source),
        reason: `unknown layout ${JSON.stringify(requested)}; kept the current layout`,
      });
    }
    slides.push({ source, layout });
  });
  for (let source = 0; source < slideCount; source++) {
    if (!seen.has(source)) {
      slides.push({ source, layout: null });
      dropped.push({
        kind: "slideOrder",
        location: slideLocation(source),
        reason: "not placed by the model; kept, after the placed slides, with its current layout",
      });
    }
  }
  return { redesign: { theme, slides }, dropped };
}
