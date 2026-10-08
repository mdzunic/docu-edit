// Deck sections (SPEC-099, the normative section model): an ordered list of
// named sections on the deck plus one `sectionKey` per slide. Sections are
// contiguous runs in deck order — what pptx's `p14:sectionLst` encodes and
// what odp carries in a `meta.xml` custom property — so membership needs
// nothing but the per-slide key. Framework-free like every presentation/
// module: the commands mutate the model in place, mark it `structureDirty`
// (the writers rebuild presentation.xml / meta.xml) and return what the view
// should select; `normaliseSections` is the idempotent repair every
// structural command (slides.ts), paste and import ends with.
//
// Layering rule (presentation/boundary.spec.ts): presentation siblings and
// src/lib/core only.
import type { DeckSection, PresentationModel, SectionRun } from "./model";

/** The longest a section name may be, after trimming. */
export const MAX_SECTION_NAME_LENGTH = 120;

/** The name of the section `addSection` mints to cover a leading run, and the
 * one a section gets when it is added with a blank name. */
export const DEFAULT_SECTION_NAME = "Default Section";

/** The deck's sections as runs of slides, in deck order. Assumes a normalised
 * deck (every slide keyed, runs contiguous); `[]` for a deck with none. */
export function sectionsOf(model: PresentationModel): SectionRun[] {
  const runs: SectionRun[] = [];
  let cursor = 0;
  for (const section of model.sections ?? []) {
    const start = cursor;
    while (cursor < model.slides.length && model.slides[cursor].sectionKey === section.key) {
      cursor += 1;
    }
    runs.push({ section, startIndex: start, endIndex: cursor });
  }
  return runs;
}

/** The run holding slide `index` (the last run when `index` is past the end),
 * or undefined for a deck with no sections. */
export function runAt(model: PresentationModel, index: number): SectionRun | undefined {
  const runs = sectionsOf(model);
  const key = model.slides[index]?.sectionKey;
  if (key !== undefined) {
    return runs.find((run) => run.section.key === key);
  }
  return runs[runs.length - 1];
}

/** A section name trimmed and clamped to `MAX_SECTION_NAME_LENGTH`; `null`
 * when nothing is left (an empty name is rejected, never stored). */
export function cleanSectionName(name: string): string | null {
  const trimmed = name.trim().slice(0, MAX_SECTION_NAME_LENGTH).trim();
  return trimmed.length > 0 ? trimmed : null;
}

/** A fresh deck-stable key: `de-sec{n}`, n one past the highest minted so far. */
export function mintSectionKey(model: PresentationModel): string {
  let max = 0;
  for (const section of model.sections ?? []) {
    const match = /^de-sec(\d+)$/.exec(section.key);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `de-sec${max + 1}`;
}

/** Adds a section starting at slide `atIndex`, taking the slides from there to
 * the end of the run that held it (the whole deck when there were no sections)
 * and sitting right after that run's section. With no sections yet and
 * `atIndex > 0`, a leading "Default Section" is minted first so the invariant
 * (no leading unsectioned run) holds. Returns the new section's key. */
export function addSection(model: PresentationModel, atIndex: number, name: string): string {
  const at = Math.min(Math.max(atIndex, 0), model.slides.length);
  if (!model.sections || model.sections.length === 0) {
    model.sections = [];
    if (at > 0) {
      const lead: DeckSection = { key: mintSectionKey(model), name: DEFAULT_SECTION_NAME };
      model.sections.push(lead);
      for (const slide of model.slides) {
        slide.sectionKey = lead.key;
      }
    }
  }
  const section: DeckSection = { key: mintSectionKey(model), name: cleanSectionName(name) ?? DEFAULT_SECTION_NAME };
  const host = at < model.slides.length ? runAt(model, at) : undefined;
  if (host) {
    for (let i = at; i < host.endIndex; i++) {
      model.slides[i].sectionKey = section.key;
    }
    const hostSlot = model.sections.findIndex((entry) => entry.key === host.section.key);
    model.sections.splice(hostSlot + 1, 0, section);
  } else {
    // No slide at the index (past the end, or a first section over every slide).
    if (model.sections.length === 0) {
      for (const slide of model.slides) {
        slide.sectionKey = section.key;
      }
    }
    model.sections.push(section);
  }
  normaliseSections(model);
  model.structureDirty = true;
  return section.key;
}

/** Renames a section to a trimmed 1..120-character name. False (and nothing
 * changed) when the key names no section or the name is empty after trim. */
export function renameSection(model: PresentationModel, key: string, name: string): boolean {
  const section = model.sections?.find((entry) => entry.key === key);
  const clean = cleanSectionName(name);
  if (!section || clean === null) {
    return false;
  }
  section.name = clean;
  model.structureDirty = true;
  return true;
}

/** Moves a section — and its whole run of slides — `delta` places among the
 * sections (clamped). Returns the moved run's new first slide index (the slide
 * to select), or -1 when the key names no section. */
export function moveSection(model: PresentationModel, key: string, delta: number): number {
  const runs = sectionsOf(model);
  const from = runs.findIndex((run) => run.section.key === key);
  if (from < 0) {
    return -1;
  }
  const to = Math.min(Math.max(from + delta, 0), runs.length - 1);
  if (to !== from) {
    const [moved] = runs.splice(from, 1);
    runs.splice(to, 0, moved);
    const slides = model.slides;
    model.slides = runs.flatMap((run) => slides.slice(run.startIndex, run.endIndex));
    model.sections = runs.map((run) => run.section);
    model.structureDirty = true;
  }
  let start = 0;
  for (const run of runs) {
    if (run.section.key === key) {
      break;
    }
    start += run.endIndex - run.startIndex;
  }
  return Math.min(start, Math.max(model.slides.length - 1, 0));
}

/** Removes a section. `"sectionOnly"` keeps its slides, which join the
 * preceding section (the following one when it was first); removing the only
 * section leaves a deck with none. `"withSlides"` deletes its slides too.
 * Returns the slide index to select afterwards, or -1 when the key names no
 * section. */
export function removeSection(
  model: PresentationModel,
  key: string,
  mode: "sectionOnly" | "withSlides",
): number {
  const runs = sectionsOf(model);
  const index = runs.findIndex((run) => run.section.key === key);
  if (index < 0 || !model.sections) {
    return -1;
  }
  const run = runs[index];
  if (mode === "withSlides") {
    model.slides.splice(run.startIndex, run.endIndex - run.startIndex);
  } else {
    const heir = runs[index - 1] ?? runs[index + 1];
    for (let i = run.startIndex; i < run.endIndex; i++) {
      if (heir) {
        model.slides[i].sectionKey = heir.section.key;
      } else {
        delete model.slides[i].sectionKey;
      }
    }
  }
  model.sections = model.sections.filter((entry) => entry.key !== key);
  normaliseSections(model);
  model.structureDirty = true;
  return model.slides.length === 0 ? 0 : Math.min(run.startIndex, model.slides.length - 1);
}

/** Moves slide `slideIndex` into section `key` — the sorter's "a slide dragged
 * past a header adopts the section it lands in". A key naming no section is
 * ignored. */
export function setSlideSection(model: PresentationModel, slideIndex: number, key: string): void {
  const slide = model.slides[slideIndex];
  if (!slide || !model.sections?.some((entry) => entry.key === key)) {
    return;
  }
  slide.sectionKey = key;
  normaliseSections(model);
  model.structureDirty = true;
}

/** Restores the section invariants after any structural change; idempotent,
 * and never itself marks the deck dirty (the command it ends does that):
 *  - no sections (absent or `[]`) → `sections` absent and no slide keyed;
 *  - a slide whose `sectionKey` names no entry adopts its predecessor's (the
 *    first entry for `slides[0]`), as does one whose section's run already
 *    closed, so runs stay contiguous;
 *  - `sections` re-orders to first-appearance order, an emptied section
 *    keeping its slot;
 *  - duplicate entries are dropped, so a key held by no slide and no entry
 *    cannot survive. */
export function normaliseSections(model: PresentationModel): void {
  if (!model.sections || model.sections.length === 0) {
    delete model.sections;
    for (const slide of model.slides) {
      delete slide.sectionKey;
    }
    return;
  }
  const seen = new Set<string>();
  const entries = model.sections.filter((entry) => {
    if (seen.has(entry.key)) {
      return false;
    }
    seen.add(entry.key);
    return true;
  });
  const known = new Set(entries.map((entry) => entry.key));
  const closed = new Set<string>();
  const order: string[] = [];
  let previous: string | undefined;
  for (const slide of model.slides) {
    let key = slide.sectionKey;
    if (key === undefined || !known.has(key) || (key !== previous && closed.has(key))) {
      key = previous ?? entries[0].key;
    }
    if (key !== previous) {
      if (previous !== undefined) {
        closed.add(previous);
      }
      order.push(key);
      previous = key;
    }
    if (slide.sectionKey !== key) {
      slide.sectionKey = key;
    }
  }
  // Empty sections keep their slot; the others fill the remaining slots in
  // first-appearance order.
  const byKey = new Map(entries.map((entry) => [entry.key, entry]));
  const appearing = order.map((key) => byKey.get(key)!);
  let next = 0;
  const reordered = entries.map((entry) => (order.includes(entry.key) ? appearing[next++] : entry));
  const changed = reordered.length !== model.sections.length || reordered.some((entry, i) => entry !== model.sections![i]);
  if (changed) {
    model.sections = reordered;
  }
}

/** The section key a slide inserted at `index` (before the slide now there)
 * joins: its predecessor's, else the first section's; undefined for a deck
 * with no sections. Paste and import stamp it before normalising. */
export function sectionKeyForInsertion(model: PresentationModel, index: number): string | undefined {
  if (!model.sections || model.sections.length === 0) {
    return undefined;
  }
  return model.slides[index - 1]?.sectionKey ?? model.slides[index]?.sectionKey ?? model.sections[0].key;
}
