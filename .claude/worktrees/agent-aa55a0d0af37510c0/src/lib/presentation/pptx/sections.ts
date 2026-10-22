// The pptx wire mapping for deck sections (SPEC-099): PowerPoint stores them
// in `p:presentation/p:extLst` as one `p:ext` holding a `p14:sectionLst`, each
// `p14:section` naming its slides by `p:sldId/@id`. The reader models any list
// it finds — so a section list is never left preserved and the model is the
// single source — and the writer recomputes membership from the model's slide
// order against the `256 + index` ids `buildSubstitution` re-mints, splicing
// the list back in place (see `spliceSectionLst`). `sectionsGolden.ts` pins
// the wire detail.
import { childElements, escapeXmlAttribute, MARKUP_COMPATIBILITY_NAMESPACE } from "../../core/xml";
import type { DeckSection, PresentationModel, Slide } from "../model";
import { cleanSectionName, DEFAULT_SECTION_NAME, normaliseSections } from "../sections";

/** The `p:ext` uri PowerPoint files its section list under. */
export const SECTIONS_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
/** The PowerPoint 2010 (`p14`) namespace the section elements live in. */
export const P14_NAMESPACE = "http://schemas.microsoft.com/office/powerpoint/2010/main";

const P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** A brace GUID, the only id form PowerPoint accepts on a `p14:section`. */
const BRACE_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;
/** The GUIDs a minted `de-sec{n}` key is written as — deterministic both ways,
 * so a section added here keeps its key across save → reopen. */
const MINTED_GUID = /^\{00000000-DE5E-4C00-8000-([0-9A-F]{12})\}$/i;

/** The `p14:section@id` a section key is written as: a source GUID verbatim,
 * a minted `de-sec{n}` as its reserved GUID. */
export function sectionWireId(key: string): string {
  if (BRACE_GUID.test(key)) {
    return key;
  }
  const minted = /^de-sec(\d+)$/.exec(key);
  const n = minted ? Number(minted[1]) : hashKey(key);
  return `{00000000-DE5E-4C00-8000-${n.toString(16).toUpperCase().padStart(12, "0")}}`;
}

/** The model key a `p14:section@id` reads back as (the inverse of
 * `sectionWireId`); undefined when the id is no brace GUID. */
export function sectionKeyFromWireId(id: string): string | undefined {
  const minted = MINTED_GUID.exec(id);
  if (minted) {
    return `de-sec${parseInt(minted[1], 16)}`;
  }
  return BRACE_GUID.test(id) ? id : undefined;
}

/** A stable 32-bit number for a key that is neither a GUID nor minted (never
 * produced by this editor; a guard so the writer always emits a valid id). */
function hashKey(key: string): number {
  let hash = 0;
  for (const char of key) {
    hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  }
  return hash + 0x1_0000_0000;
}

/** Models the deck's `p14:sectionLst` (when there is one) onto the read
 * slides: names in document order, membership by `p14:sldId/@id` resolved
 * through the `p:sldIdLst` walk (`slideIds[i]` is model slide i's
 * `p:sldId/@id`), unknown ids ignored and an unassigned slide joining the
 * preceding run. Returns the sections, or undefined when the deck declares
 * none. Each slide's `sectionKey` is set. */
export function readSections(presentation: Element, slides: Slide[], slideIds: (string | null)[]): DeckSection[] | undefined {
  const list = findSectionLst(presentation);
  if (!list) {
    return undefined;
  }
  const sections: DeckSection[] = [];
  const keys = new Set<string>();
  const indexById = new Map<string, number>();
  slideIds.forEach((id, index) => {
    if (id !== null && !indexById.has(id)) {
      indexById.set(id, index);
    }
  });
  const assigned = new Set<number>();
  let minted = 0;
  for (const element of childElements(list, P14_NAMESPACE, "section")) {
    let key = sectionKeyFromWireId(element.getAttribute("id") ?? "");
    if (key === undefined || keys.has(key)) {
      do {
        minted += 1;
        key = `de-sec${minted}`;
      } while (keys.has(key));
    }
    keys.add(key);
    sections.push({ key, name: cleanSectionName(element.getAttribute("name") ?? "") ?? DEFAULT_SECTION_NAME });
    for (const idList of childElements(element, P14_NAMESPACE, "sldIdLst")) {
      for (const sldId of childElements(idList, P14_NAMESPACE, "sldId")) {
        const index = indexById.get(sldId.getAttribute("id") ?? "");
        if (index !== undefined && !assigned.has(index)) {
          assigned.add(index);
          slides[index].sectionKey = key;
        }
      }
    }
  }
  if (sections.length === 0) {
    return undefined;
  }
  // Unassigned slides join the preceding run (the first section at the head).
  const holder: Pick<PresentationModel, "slides" | "sections"> = { slides, sections };
  normaliseSections(holder);
  return holder.sections;
}

/** The `p14:sectionLst` under `p:extLst/p:ext[@uri]`, by namespace. */
function findSectionLst(presentation: Element): Element | undefined {
  for (const extLst of childElements(presentation, P_NAMESPACE, "extLst")) {
    for (const ext of childElements(extLst, P_NAMESPACE, "ext")) {
      const list = childElements(ext, P14_NAMESPACE, "sectionLst")[0];
      if (list) {
        return list;
      }
    }
  }
  return undefined;
}

/** The deck's `sectionLst` element under `prefix`, membership recomputed from
 * the model's slide order against the freshly minted `256 + index` ids. The
 * namespace is declared on the element itself, so it is well-formed wherever
 * it lands. */
export function serializeSectionLst(model: PresentationModel, prefix = "p14"): string {
  const slideIdsByKey = new Map<string, number[]>();
  model.slides.forEach((slide, index) => {
    if (slide.sectionKey !== undefined) {
      const ids = slideIdsByKey.get(slide.sectionKey) ?? [];
      ids.push(256 + index);
      slideIdsByKey.set(slide.sectionKey, ids);
    }
  });
  const sections = (model.sections ?? [])
    .map((section) => {
      const ids = (slideIdsByKey.get(section.key) ?? []).map((id) => `<${prefix}:sldId id="${id}"/>`).join("");
      return (
        `<${prefix}:section name="${escapeXmlAttribute(section.name)}" id="${sectionWireId(section.key)}">` +
        `<${prefix}:sldIdLst>${ids}</${prefix}:sldIdLst></${prefix}:section>`
      );
    })
    .join("");
  return `<${prefix}:sectionLst xmlns:${prefix}="${P14_NAMESPACE}">${sections}</${prefix}:sectionLst>`;
}

/** The whole `p:ext` block carrying the deck's sections (what a fresh splice
 * inserts; `SECTIONS_GOLDEN` pins it). */
export function sectionsExtXml(model: PresentationModel): string {
  return `<p:ext uri="${SECTIONS_EXT_URI}">${serializeSectionLst(model)}</p:ext>`;
}

const SECTION_LST = /<([A-Za-z0-9]+):sectionLst[\s\S]*?<\/\1:sectionLst>/;
const SECTIONS_EXT = /<p:ext\b[^>]*\buri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"[^>]*>[\s\S]*?<\/p:ext>/i;

/** Splices the model's sections into presentation.xml, mirroring
 * `spliceSldIdLst`: an existing `sectionLst` is replaced in place — reusing
 * the prefix it was written under and leaving the enclosing and every sibling
 * `p:ext` byte-verbatim — else a `p:ext` is spliced before `</p:extLst>`, else
 * a whole `p:extLst` goes last in `p:presentation`; `xmlns:p14` and
 * `mc:Ignorable="p14"` join the root only when absent. A deck with no
 * sections loses its section `p:ext` (and an `p:extLst` it leaves empty). */
export function spliceSectionLst(presXml: string, model: PresentationModel): string {
  const hasSections = model.sections !== undefined && model.sections.length > 0;
  const existing = SECTION_LST.exec(presXml);
  if (!hasSections) {
    if (!existing) {
      return presXml;
    }
    const withoutExt = SECTIONS_EXT.test(presXml) ? presXml.replace(SECTIONS_EXT, "") : presXml.replace(SECTION_LST, "");
    return withoutExt.replace(/<p:extLst>\s*<\/p:extLst>/, "");
  }
  if (existing) {
    const prefix = existing[1];
    return presXml.replace(SECTION_LST, () => serializeSectionLst(model, prefix));
  }
  const ext = sectionsExtXml(model);
  let spliced: string;
  if (/<\/p:extLst>/.test(presXml)) {
    spliced = presXml.replace(/<\/p:extLst>(?![\s\S]*<\/p:extLst>)/, `${ext}</p:extLst>`);
  } else if (/<p:extLst\s*\/>/.test(presXml)) {
    spliced = presXml.replace(/<p:extLst\s*\/>/, `<p:extLst>${ext}</p:extLst>`);
  } else {
    spliced = presXml.replace(/<\/p:presentation>\s*$/, `<p:extLst>${ext}</p:extLst></p:presentation>`);
  }
  return declareP14(spliced);
}

/** Adds `xmlns:p14` (and `xmlns:mc` + `mc:Ignorable="p14"`) to the
 * `p:presentation` start tag, each only when absent; an existing
 * `mc:Ignorable` list gains `p14` rather than being replaced. */
function declareP14(presXml: string): string {
  return presXml.replace(/<p:presentation\b[^>]*>/, (tag) => {
    let out = tag;
    const add = (attr: string): void => {
      out = out.replace(/\s*(\/?)>$/, ` ${attr}$1>`);
    };
    if (!/\sxmlns:p14="/.test(out)) {
      add(`xmlns:p14="${P14_NAMESPACE}"`);
    }
    const mcPrefix = /\sxmlns:([A-Za-z0-9]+)="http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006"/.exec(out)?.[1];
    if (!mcPrefix) {
      add(`xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}"`);
      add(`mc:Ignorable="p14"`);
      return out;
    }
    const ignorable = new RegExp(`\\s${mcPrefix}:Ignorable="([^"]*)"`).exec(out);
    if (!ignorable) {
      add(`${mcPrefix}:Ignorable="p14"`);
    } else if (!ignorable[1].split(/\s+/).includes("p14")) {
      out = out.replace(ignorable[0], ` ${mcPrefix}:Ignorable="${`${ignorable[1]} p14`.trim()}"`);
    }
    return out;
  });
}
