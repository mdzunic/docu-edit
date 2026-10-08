// The odp mapping for deck sections (SPEC-099) — a stated loss, not an
// invention: ODF 1.3 has no deck-section construct (`presentation:show` is a
// custom slide show, a different feature), so sections ride in `meta.xml` as
// ONE custom property, `meta:name="docu-edit:sections"`
// `meta:value-type="string"`, whose value is the compact JSON
// `{"v":1,"sections":[{"key":…,"name":…,"n":…}]}` — `n` the run length, so
// membership needs no per-slide attribute. LibreOffice keeps it as a custom
// property and displays sections nowhere (the sorter says so once). Read rides
// `parseOdfCustomProperties` (the odp reader already calls it); write rides
// `buildMetaXml`.
import type { CustomProperty } from "../../core/documentProperties";
import type { DeckSection, PresentationModel, Slide } from "../model";
import { cleanSectionName, DEFAULT_SECTION_NAME, normaliseSections, sectionsOf } from "../sections";

/** The custom-property name the sections ride under. */
export const SECTIONS_PROPERTY_NAME = "docu-edit:sections";

/** A key the odp mapping accepts back from the wire: a pptx brace GUID or a
 * minted `de-sec{n}` (anything else is re-minted). */
const SAFE_KEY = /^(?:de-sec\d{1,9}|\{[0-9A-Fa-f-]{36}\})$/;

/** The property value for the deck's sections, or undefined for a deck with
 * none: `{"v":1,"sections":[{key,name,n}]}`, compact, in deck order. */
export function encodeSectionsValue(model: PresentationModel): string | undefined {
  if (!model.sections || model.sections.length === 0) {
    return undefined;
  }
  const sections = sectionsOf(model).map((run) => ({
    key: run.section.key,
    name: run.section.name,
    n: run.endIndex - run.startIndex,
  }));
  return JSON.stringify({ v: 1, sections });
}

/** The deck's custom properties as the writer emits them: every unrelated
 * property in order, then the sections property when the deck has sections. */
export function customPropertiesWithSections(
  properties: CustomProperty[] | undefined,
  model: PresentationModel,
): CustomProperty[] | undefined {
  const others = (properties ?? []).filter((property) => property.name !== SECTIONS_PROPERTY_NAME);
  const value = encodeSectionsValue(model);
  const all = value === undefined ? others : [...others, { name: SECTIONS_PROPERTY_NAME, type: "text" as const, value }];
  return all.length > 0 ? all : undefined;
}

/** Decodes the property value onto the read slides: runs assigned in order, a
 * total overrunning the deck clamped (the overflow sections read empty) and a
 * shortfall extending the last run. Returns the sections (each slide's
 * `sectionKey` set), or undefined when the value is not a v1 payload with at
 * least one section — the property then stays an ordinary custom property. */
export function decodeSectionsValue(value: string, slides: Slide[]): DeckSection[] | undefined {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return undefined;
  }
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    return undefined;
  }
  const payload = parsed as Record<string, unknown>;
  if (payload.v !== 1 || !Array.isArray(payload.sections)) {
    return undefined;
  }
  const sections: DeckSection[] = [];
  const lengths: number[] = [];
  const declared = new Set(
    payload.sections.map((entry) => (entry as Record<string, unknown> | null)?.key).filter((key) => typeof key === "string"),
  );
  const keys = new Set<string>();
  let minted = 0;
  for (const entry of payload.sections) {
    if (typeof entry !== "object" || entry === null) {
      return undefined;
    }
    const raw = entry as Record<string, unknown>;
    if (typeof raw.name !== "string" || typeof raw.n !== "number" || !Number.isInteger(raw.n) || raw.n < 0) {
      return undefined;
    }
    let key = typeof raw.key === "string" && SAFE_KEY.test(raw.key) ? raw.key : undefined;
    if (key === undefined || keys.has(key)) {
      do {
        minted += 1;
        key = `de-sec${minted}`;
      } while (keys.has(key) || declared.has(key));
    }
    keys.add(key);
    sections.push({ key, name: cleanSectionName(raw.name) ?? DEFAULT_SECTION_NAME });
    lengths.push(raw.n);
  }
  if (sections.length === 0) {
    return undefined;
  }
  let cursor = 0;
  sections.forEach((section, i) => {
    const end = Math.min(cursor + lengths[i], slides.length);
    for (; cursor < end; cursor++) {
      slides[cursor].sectionKey = section.key;
    }
  });
  // A shortfall extends the last run.
  const last = sections[sections.length - 1].key;
  for (; cursor < slides.length; cursor++) {
    slides[cursor].sectionKey = last;
  }
  const holder: Pick<PresentationModel, "slides" | "sections"> = { slides, sections };
  normaliseSections(holder);
  return holder.sections;
}

/** The sections property's raw value in a parsed custom-property list, if any. */
export function sectionsValueOf(properties: CustomProperty[] | undefined): string | undefined {
  return properties?.find((property) => property.name === SECTIONS_PROPERTY_NAME && property.type === "text")?.value;
}
