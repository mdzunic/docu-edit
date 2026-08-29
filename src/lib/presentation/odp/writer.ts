// The `.odp` writer (SPEC-080). Like the pptx writer, three paths chosen by
// what the model holds:
//
//  1. No retained package (from-scratch / `stripRetention`ed deck): build the
//     minimal ODF package — `mimetype` (first, STOREd), META-INF/manifest.xml,
//     meta.xml, styles.xml (one `Default` master page carrying the page size),
//     content.xml with one `draw:page` per model slide (empty placeholder
//     frames; SPEC-081 owns real frame serialisation).
//  2. Retained package, nothing structural changed: re-emit every part
//     verbatim, `mimetype` first + STOREd — the preparation-milestone
//     behaviour, so a no-edit download stays byte-identical (SPEC-008 / AC-51).
//  3. Retained package + structural changes: re-emit every part verbatim
//     except content.xml, rebuilt as `envelope.prefix + pages + envelope.suffix`
//     — each untouched page contributes its `sourceXml` verbatim
//     (splice-equality), duplicated pages the source `sourceXml` with
//     `draw:name` uniqued, new pages generated markup. META-INF/manifest.xml
//     is spliced only when parts were added/removed (never, this spec — pages
//     live inside content.xml).
import { base64ToBytes, decodedPartBytes } from "../../core/packageParts";
import { buildMetaXml, type OdfMetaInput } from "../../core/odfMeta";
import { orientationOf } from "../../core/pageSetup";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import type { MediaPart } from "../../core/media";
import { escapeXmlAttribute } from "../../core/xml";
import {
  DEFAULT_PRESENTATION_TITLE,
  type LayoutSlot,
  type PlaceholderRole,
  type PresentationModel,
  type SlideFrame,
  type SlideLayout,
} from "../model";
import { classifyFromSlots, synthesiseStandardLayouts } from "../layouts";
import { ODP_MIMETYPE } from "./reader";
import { drawAttr, serializePage, spliceAutomaticStyles, StyleRegistry } from "./serialize";
import { CHART_OBJECT_CONTENT_TYPE, odfChartContentXml } from "../../core/chartPartOdf";

/** Every media part referenced by an ImageFrame anywhere in the deck (groups
 * descended), from `model.media` — the writer's source of truth for inserted
 * `Pictures/…` parts (SPEC-082). */
function referencedMedia(model: PresentationModel): MediaPart[] {
  const used = new Set<string>();
  const walk = (frames: SlideFrame[]): void => {
    for (const frame of frames) {
      if (frame.kind === "image") {
        used.add(frame.mediaId);
      } else if (frame.kind === "group") {
        walk(frame.children);
      }
    }
  };
  for (const slide of model.slides) {
    walk(slide.frames);
  }
  return (model.media ?? []).filter((part) => used.has(part.mediaId));
}

/** Splices `<manifest:file-entry>` rows for any media part not already
 * declared, before `</manifest:manifest>` (AC-49). */
function spliceManifest(manifestXml: string, media: MediaPart[]): string {
  const additions = media
    .filter((part) => !manifestXml.includes(`manifest:full-path="${part.mediaId}"`))
    .map(
      (part) =>
        `<manifest:file-entry manifest:full-path="${escapeXmlAttribute(part.mediaId)}" manifest:media-type="${escapeXmlAttribute(part.contentType)}"/>`,
    )
    .join("");
  if (additions.length === 0) {
    return manifestXml;
  }
  return manifestXml.replace(/<\/manifest:manifest>/, `${additions}</manifest:manifest>`);
}

export { ODP_MIMETYPE };

interface OdpChartPlan {
  /** Regenerated `Object N/content.xml` bytes by part name. */
  partBytes: Map<string, Uint8Array>;
  /** Fresh object directories (e.g. "Object 3") — appended, and given
   *  `META-INF/manifest.xml` entries for the object dir and its content.xml. */
  newObjects: string[];
}

/** Plans the embedded chart objects a rebuild writes (SPEC-092 AC-21). Every
 * touched chart frame is regenerated: an `edited` chart's `content.xml` is
 * rewritten in its `Object N/` directory; a chart created here — or a duplicate
 * sharing a `sourcePart` (D8) — mints a fresh `Object {n}` past the highest
 * existing, recorded on the frame so the page serialiser resolves its href.
 * Untouched and unsupported charts are left alone (verbatim pass-through). */
function planOdpCharts(model: PresentationModel, parts: Map<string, Uint8Array>): OdpChartPlan {
  const partBytes = new Map<string, Uint8Array>();
  const newObjects: string[] = [];
  const claimed = new Set<string>();
  let maxIndex = 0;
  for (const name of parts.keys()) {
    const m = /^Object (\d+)\//.exec(name);
    if (m) {
      maxIndex = Math.max(maxIndex, Number(m[1]));
    }
  }
  for (const slide of model.slides) {
    for (const frame of slide.frames) {
      if (frame.kind !== "chart" || frame.xml !== undefined || frame.unsupportedChartXml !== undefined) {
        continue;
      }
      let target = frame.sourcePart;
      if (!target || claimed.has(target) || !parts.has(target)) {
        maxIndex += 1;
        const dir = `Object ${maxIndex}`;
        target = `${dir}/content.xml`;
        frame.sourcePart = target;
        newObjects.push(dir);
      }
      claimed.add(target);
      partBytes.set(target, encode(odfChartContentXml(frame, frame.widthCm, frame.heightCm)));
    }
  }
  return { partBytes, newObjects };
}

/** `<manifest:file-entry>` rows for a chart plan's fresh object directories:
 * the `Object N/` dir (media type `application/vnd.oasis.opendocument.chart`)
 * and its `Object N/content.xml` (`text/xml`). */
function odpChartManifestEntries(plan: OdpChartPlan): string {
  return plan.newObjects
    .map(
      (dir) =>
        `<manifest:file-entry manifest:full-path="${escapeXmlAttribute(dir)}/" manifest:media-type="${CHART_OBJECT_CONTENT_TYPE}"/>` +
        `<manifest:file-entry manifest:full-path="${escapeXmlAttribute(dir)}/content.xml" manifest:media-type="text/xml"/>`,
    )
    .join("");
}

const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL_NS = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

const encoder = new TextEncoder();
const encode = (text: string): Uint8Array => encoder.encode(text);

// ── Layout catalogue writing (SPEC-089) ────────────────────────────────────

/** The `presentation:object` value an ODF page-layout placeholder carries for a
 * role — the inverse of the reader's `roleOf`. */
function objectForRole(role: PlaceholderRole): string {
  switch (role) {
    case "title":
      return "title";
    case "subtitle":
      return "subtitle";
    case "body":
      return "outline";
    default:
      return "object";
  }
}

/** One authored `<style:presentation-page-layout>` for a catalogue entry with
 * no counterpart in the deck's styles.xml (SPEC-089 AC-16). Exported so the
 * layout golden asserts the splice reproduces LAYOUT_GOLDEN's markup. */
export function authorPageLayoutStyle(name: string, layout: SlideLayout): string {
  const placeholders = layout.slots
    .map(
      (slot: LayoutSlot) =>
        `<presentation:placeholder presentation:object="${objectForRole(slot.role)}"` +
        ` svg:x="${slot.xCm}cm" svg:y="${slot.yCm}cm" svg:width="${slot.widthCm}cm" svg:height="${slot.heightCm}cm"/>`,
    )
    .join("");
  return `<style:presentation-page-layout style:name="${escapeXmlAttribute(name)}">${placeholders}</style:presentation-page-layout>`;
}

interface OdpLayoutPlan {
  /** layoutKey → the `presentation:presentation-page-layout-name` to stamp. */
  nameForKey: Map<string, string>;
  /** Authored `<style:presentation-page-layout>` blocks for styles.xml. */
  styleBlocks: string[];
}

/** Resolves each slide's `layoutKey` to a page-layout style name: an existing
 * deck style (the `${master}::${pageLayout}` key's second half) is stamped
 * as-is; a synthesised layout is authored once as `de-AL{n}` (SPEC-089). A
 * page whose key names no layout (`${master}::`) is left unstamped. */
function planOdpLayouts(model: PresentationModel): OdpLayoutPlan {
  const nameForKey = new Map<string, string>();
  const styleBlocks: string[] = [];
  let authored = 0;
  const keys = new Set<string>();
  for (const slide of model.slides) {
    if (slide.layoutKey) {
      keys.add(slide.layoutKey);
    }
  }
  for (const key of keys) {
    if (key.startsWith("standard:")) {
      const entry = (model.layouts ?? []).find((l) => l.key === key);
      if (!entry) {
        continue;
      }
      authored += 1;
      const name = `de-AL${authored}`;
      styleBlocks.push(authorPageLayoutStyle(name, entry));
      nameForKey.set(key, name);
      continue;
    }
    const sep = key.indexOf("::");
    const pageLayout = sep >= 0 ? key.slice(sep + 2) : "";
    if (pageLayout) {
      nameForKey.set(key, pageLayout);
    }
  }
  return { nameForKey, styleBlocks };
}

/** Splices authored page-layout styles before `</office:styles>` (SPEC-089
 * AC-16); creates the `office:styles` container when the file lacks one. */
function spliceOfficeStyles(stylesXml: string, blocks: string[]): string {
  if (blocks.length === 0) {
    return stylesXml;
  }
  const joined = blocks.join("");
  if (/<\/office:styles>/.test(stylesXml)) {
    return stylesXml.replace(/<\/office:styles>/, `${joined}</office:styles>`);
  }
  if (/<office:automatic-styles[ >]/.test(stylesXml)) {
    return stylesXml.replace(/<office:automatic-styles[ >]/, (m) => `<office:styles>${joined}</office:styles>${m}`);
  }
  return stylesXml.replace(/<\/office:document-styles>/, `<office:styles>${joined}</office:styles></office:document-styles>`);
}

// ── Deck-size write-back (SPEC-090) ─────────────────────────────────────────

/** The meta.xml input for the deck's metadata — the deck's `properties` with
 * `title` falling back to the app-bar title, passing custom/unsupported/extra
 * fields straight back so nothing this dialog does not edit is lost. Shared by
 * the from-scratch and substitution paths. */
function metaInputFor(model: PresentationModel): OdfMetaInput {
  const props = model.properties ?? {};
  return {
    title: props.title ?? model.title,
    subject: props.subject,
    author: props.author,
    lastModifiedBy: props.lastModifiedBy,
    keywords: props.keywords,
    comments: props.comments,
    created: props.created,
    extraFields: props.extraFields,
    customProperties: props.customProperties,
    unsupportedCustomProperties: props.unsupportedCustomProperties,
  };
}

/** The `style:page-layout-name` on each `style:master-page` START tag (a notes
 * page layout on a `presentation:notes` CHILD is a separate element, so this
 * naturally excludes it — SPEC-090). */
function masterPageLayoutNames(xml: string): string[] {
  const names = new Set<string>();
  for (const tag of xml.match(/<style:master-page\b[^>]*>/g) ?? []) {
    const match = /\bstyle:page-layout-name="([^"]*)"/.exec(tag);
    if (match) {
      names.add(match[1]);
    }
  }
  return [...names];
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Sets an ODF attribute on a start/empty tag, replacing its value in place or
 * inserting it before the tag's close when absent (SPEC-090 "inserting an
 * attribute when absent, leaving every other byte alone"). */
function setOdfAttr(tag: string, name: string, value: string): string {
  const re = new RegExp(`\\b${escapeRegExp(name)}="[^"]*"`);
  if (re.test(tag)) {
    return tag.replace(re, `${name}="${value}"`);
  }
  return tag.replace(/\s*\/?>$/, (end) => ` ${name}="${value}"${end}`);
}

/** Rewrites `fo:page-width`, `fo:page-height` and `style:print-orientation` on
 * every named `style:page-layout`'s `style:page-layout-properties`. */
function rewritePageLayoutSize(
  xml: string,
  names: string[],
  widthCm: number,
  heightCm: number,
  orientation: string,
): string {
  let out = xml;
  for (const name of names) {
    const re = new RegExp(
      `(<style:page-layout\\b[^>]*\\bstyle:name="${escapeRegExp(name)}"[^>]*>)([\\s\\S]*?)(</style:page-layout>)`,
    );
    out = out.replace(re, (_full, open: string, inner: string, close: string) => {
      let body = inner;
      if (/<style:page-layout-properties\b/.test(body)) {
        body = body.replace(/<style:page-layout-properties\b[^>]*\/?>/, (props) => {
          let updated = setOdfAttr(props, "fo:page-width", `${widthCm}cm`);
          updated = setOdfAttr(updated, "fo:page-height", `${heightCm}cm`);
          updated = setOdfAttr(updated, "style:print-orientation", orientation);
          return updated;
        });
      } else {
        body =
          `<style:page-layout-properties fo:page-width="${widthCm}cm" fo:page-height="${heightCm}cm"` +
          ` style:print-orientation="${orientation}"/>` +
          body;
      }
      return open + body + close;
    });
  }
  return out;
}

/** Applies the deck size to whichever part carries the master pages' page
 * layouts: styles.xml first (the common case), else the rebuilt content.xml's
 * automatic styles (the reader's own fallback), else a fresh `PMde1` layout
 * spliced into styles.xml and referenced from each bare master page. */
function rewriteOdpDeckSize(
  stylesXml: string | undefined,
  contentXml: string | undefined,
  model: PresentationModel,
): { stylesXml?: string; contentXml?: string } {
  const widthCm = model.slideWidthCm;
  const heightCm = model.slideHeightCm;
  const orientation = orientationOf({ widthCm, heightCm });
  const stylesNames = stylesXml ? masterPageLayoutNames(stylesXml) : [];
  if (stylesXml && stylesNames.length > 0) {
    return { stylesXml: rewritePageLayoutSize(stylesXml, stylesNames, widthCm, heightCm, orientation) };
  }
  const contentNames = contentXml ? masterPageLayoutNames(contentXml) : [];
  if (contentXml && contentNames.length > 0) {
    return { contentXml: rewritePageLayoutSize(contentXml, contentNames, widthCm, heightCm, orientation) };
  }
  if (stylesXml) {
    const layout =
      `<style:page-layout style:name="PMde1"><style:page-layout-properties fo:page-width="${widthCm}cm"` +
      ` fo:page-height="${heightCm}cm" style:print-orientation="${orientation}"/></style:page-layout>`;
    let out = spliceAutoStyleLayout(stylesXml, layout);
    out = out.replace(/<style:master-page\b[^>]*>/g, (tag) =>
      /\bstyle:page-layout-name="/.test(tag) ? tag : setOdfAttr(tag, "style:page-layout-name", "PMde1"),
    );
    return { stylesXml: out };
  }
  return {};
}

/** Splices a page-layout style into styles.xml's `office:automatic-styles`
 * (creating the container just before `</office:document-styles>` if absent). */
function spliceAutoStyleLayout(stylesXml: string, layout: string): string {
  if (/<\/office:automatic-styles>/.test(stylesXml)) {
    return stylesXml.replace(/<\/office:automatic-styles>/, `${layout}</office:automatic-styles>`);
  }
  return stylesXml.replace(
    /<\/office:document-styles>/,
    `<office:automatic-styles>${layout}</office:automatic-styles></office:document-styles>`,
  );
}

export function odpFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.odp`;
}

export interface WriteOdpOptions {
  /** Injectable so tests assert an exact meta.xml (the clock-coupled-test
   * rule): the timestamp stamped into a from-scratch meta.xml. */
  modified?: Date;
}

/** Serialises a presentation to `.odp` bytes (see the module header). */
export async function writeOdp(model: PresentationModel, options: WriteOdpOptions = {}): Promise<{ bytes: Uint8Array }> {
  const hasPackage = model.opaque !== undefined && model.opaque.length > 0;
  if (!hasPackage) {
    return { bytes: await zip(buildFromScratch(model, options.modified ?? new Date())) };
  }
  const parts = decodedPartBytes(model.opaque);
  const deckDirty = model.deckDirty === true;
  // A deck-level change (SPEC-090) joins the structural condition: it rewrites
  // styles.xml's page-layout size + meta.xml (and, in scale mode, dirties the
  // slides so content.xml re-emits their geometry too).
  const structural = model.structureDirty === true || deckDirty || model.slides.some((slide) => slide.dirty === true);
  const mimetype = parts.get("mimetype") ?? encode(ODP_MIMETYPE);
  const entries: ZipEntryInput[] = [{ name: "mimetype", data: mimetype, store: true }];

  // Layout changes (SPEC-089): a layout change lives inside preserved["@page"];
  // a synthesised layout also authors a styles.xml page-layout style. A deck
  // whose layouts were untouched authors nothing, so styles.xml stays verbatim.
  const layoutPlan = planOdpLayouts(model);
  let contentXml =
    structural && model.odpContentEnvelope ? rebuildContentXml(model, model.odpContentEnvelope, layoutPlan) : undefined;

  // styles.xml: SPEC-089's authored layout blocks and/or SPEC-090's deck-size
  // rewrite (fo:page-width/height + style:print-orientation on the layouts the
  // master pages reference).
  let stylesXml: string | undefined;
  if (structural && parts.has("styles.xml") && (layoutPlan.styleBlocks.length > 0 || deckDirty)) {
    stylesXml = decoder.decode(parts.get("styles.xml")!);
    if (layoutPlan.styleBlocks.length > 0) {
      stylesXml = spliceOfficeStyles(stylesXml, layoutPlan.styleBlocks);
    }
  }
  if (deckDirty) {
    const rewritten = rewriteOdpDeckSize(stylesXml, contentXml, model);
    stylesXml = rewritten.stylesXml ?? stylesXml;
    contentXml = rewritten.contentXml ?? contentXml;
  }

  // meta.xml is rebuilt from the deck metadata on a deck change (SPEC-090),
  // passing custom/unsupported/extra fields straight back so nothing is lost;
  // firstSlideNumber is deliberately NOT written to odp (Decisions).
  const metaXml =
    deckDirty && parts.has("meta.xml") ? buildMetaXml(metaInputFor(model), options.modified ?? new Date()) : undefined;

  // Inserted images: model.media is the source of truth (overrides an opaque
  // twin), new `Pictures/…` parts are appended, and the manifest gains their
  // entries (AC-49). Only touched when an image was actually inserted.
  const media = structural ? referencedMedia(model) : [];
  const mediaBytes = new Map(media.map((part) => [part.mediaId, base64ToBytes(part.data)]));
  const appended = media.filter((part) => !parts.has(part.mediaId));
  // Embedded charts (SPEC-092): regenerate a touched chart's Object N/content.xml
  // and mint fresh objects, splicing their manifest entries.
  const charts = structural ? planOdpCharts(model, parts) : { partBytes: new Map<string, Uint8Array>(), newObjects: [] };
  const chartManifest = odpChartManifestEntries(charts);
  const manifestChanged = appended.length > 0 || chartManifest.length > 0;

  for (const [name, data] of parts) {
    if (name === "mimetype") {
      continue;
    }
    if (name === "content.xml" && contentXml !== undefined) {
      entries.push({ name, data: encode(contentXml) });
    } else if (name === "styles.xml" && stylesXml !== undefined) {
      entries.push({ name, data: encode(stylesXml) });
    } else if (name === "meta.xml" && metaXml !== undefined) {
      entries.push({ name, data: encode(metaXml) });
    } else if (name === "META-INF/manifest.xml" && manifestChanged) {
      const withMedia = appended.length > 0 ? spliceManifest(decoder.decode(data), appended) : decoder.decode(data);
      const withCharts = chartManifest.length > 0 ? withMedia.replace(/<\/manifest:manifest>/, `${chartManifest}</manifest:manifest>`) : withMedia;
      entries.push({ name, data: encode(withCharts) });
    } else if (charts.partBytes.has(name)) {
      entries.push({ name, data: charts.partBytes.get(name)! });
    } else if (mediaBytes.has(name)) {
      entries.push({ name, data: mediaBytes.get(name)! });
    } else {
      entries.push({ name, data });
    }
  }
  for (const part of appended) {
    entries.push({ name: part.mediaId, data: mediaBytes.get(part.mediaId)! });
  }
  for (const dir of charts.newObjects) {
    const name = `${dir}/content.xml`;
    entries.push({ name, data: charts.partBytes.get(name)! });
  }
  return { bytes: await zip(entries) };
}

const decoder = new TextDecoder("utf-8");

// --- substitution: rebuild content.xml from the page envelope --------------

function rebuildContentXml(
  model: PresentationModel,
  envelope: { prefix: string; suffix: string },
  layoutPlan: OdpLayoutPlan = { nameForKey: new Map(), styleBlocks: [] },
): string {
  const usedNames = new Set<string>();
  const registry = new StyleRegistry();
  // A source that (pathologically) already uses `de-`-prefixed names cannot
  // collide with a generated one — seed the counters past every such name.
  // The same string doubles as the lookup source for an original drawing-page
  // style's fill/background when a transition mints its replacement (AC-11).
  const styleSource = envelope.prefix + envelope.suffix + model.slides.map((s) => s.sourceXml ?? "").join("");
  registry.seedFrom(styleSource);
  let lastMaster = "Default";
  const pages = model.slides.map((slide) => {
    if (slide.sourceXml) {
      // Untouched or duplicated page: re-emit its source markup verbatim (its
      // draw:name uniqued for a duplicate) — splice-equality (AC-33).
      const master = drawNameAttr(slide.sourceXml, "draw:master-page-name");
      if (master) {
        lastMaster = master;
      }
      return uniqueDrawName(slide.sourceXml, usedNames);
    }
    // Dirty or new page: rebuild from the model, keeping its original
    // draw:page attributes (AC-37) and generating automatic styles.
    const preservedMaster = slide.preserved?.["@page"]
      ? drawAttr(slide.preserved["@page"], "draw:master-page-name")
      : undefined;
    const master = preservedMaster ?? lastMaster;
    lastMaster = master;
    const layoutName = slide.layoutKey ? layoutPlan.nameForKey.get(slide.layoutKey) : undefined;
    return serializePage(slide, registry, usedNames, master, styleSource, layoutName);
  });
  const content = envelope.prefix + pages.join("") + envelope.suffix;
  return registry.hasStyles() ? spliceAutomaticStyles(content, registry.xml()) : content;
}

/** The value of an attribute like `draw:name="…"` in a page's markup. */
function drawNameAttr(xml: string, attr: string): string | undefined {
  const match = new RegExp(`${attr}="([^"]*)"`).exec(xml);
  return match ? match[1] : undefined;
}

/** A page's `sourceXml` with its `draw:name` made unique against `used` (and
 * recorded there). A duplicated page shares its source's name until here. */
function uniqueDrawName(sourceXml: string, used: Set<string>): string {
  const current = drawNameAttr(sourceXml, "draw:name");
  if (current === undefined) {
    return sourceXml;
  }
  if (!used.has(current)) {
    used.add(current);
    return sourceXml;
  }
  const fresh = uniqueName(current, used);
  return sourceXml.replace(/draw:name="[^"]*"/, `draw:name="${fresh}"`);
}

function uniqueName(base: string, used: Set<string>): string {
  let candidate = base;
  let n = 2;
  while (used.has(candidate)) {
    candidate = `${base}_${n++}`;
  }
  used.add(candidate);
  return candidate;
}

// --- from-scratch package --------------------------------------------------

/** The standard-set index a from-scratch page references: its `layoutKey` when
 * it names a standard layout, else the kind inferred from its placeholder
 * roles. */
function standardIndexForSlide(slide: PresentationModel["slides"][number], standard: SlideLayout[]): number {
  if (slide.layoutKey) {
    const named = standard.findIndex((l) => l.key === slide.layoutKey);
    if (named >= 0) {
      return named;
    }
  }
  const slots: LayoutSlot[] = slide.frames
    .filter((f): f is SlideFrame & { kind: "text"; role: PlaceholderRole } => f.kind === "text" && f.role !== undefined)
    .map((f) => ({ role: f.role, xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 }));
  const kind = classifyFromSlots(slots);
  const inferred = standard.findIndex((l) => l.kind === kind);
  return inferred >= 0 ? inferred : 0;
}

function buildFromScratch(model: PresentationModel, modified: Date): ZipEntryInput[] {
  const used = new Set<string>();
  const registry = new StyleRegistry();
  // The minimal standard set (SPEC-089 AC-17): six page-layout styles in
  // styles.xml, each page stamped with the one its content implies.
  const standardLayouts = synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm);
  const layoutStyleNames = standardLayouts.map((_, i) => `de-AL${i + 1}`);
  const pageLayoutStyles = standardLayouts
    .map((layout, i) => authorPageLayoutStyle(layoutStyleNames[i], layout))
    .join("");
  const pages = model.slides
    .map((slide) => {
      const idx = standardIndexForSlide(slide, standardLayouts);
      return serializePage(slide, registry, used, "Default", "", layoutStyleNames[idx]);
    })
    .join("");

  const base =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:text="${TEXT_NS}"` +
    ` xmlns:draw="${DRAW_NS}" xmlns:fo="${FO_NS}" xmlns:svg="${SVG_NS}" xmlns:xlink="http://www.w3.org/1999/xlink"` +
    ` xmlns:presentation="${PRESENTATION_NS}" xmlns:smil="${SMIL_NS}" xmlns:table="${TABLE_NS}" office:version="1.2">` +
    `<office:automatic-styles></office:automatic-styles>` +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    `</office:document-content>`;
  const content = registry.hasStyles() ? spliceAutomaticStyles(base, registry.xml()) : base;

  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:draw="${DRAW_NS}"` +
    ` xmlns:fo="${FO_NS}" xmlns:svg="${SVG_NS}" xmlns:presentation="${PRESENTATION_NS}" office:version="1.2">` +
    `<office:styles>${pageLayoutStyles}</office:styles>` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="${model.slideWidthCm}cm"` +
    ` fo:page-height="${model.slideHeightCm}cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles>` +
    `<style:master-page style:name="Default" style:page-layout-name="PL1"/>` +
    `</office:master-styles>` +
    `</office:document-styles>`;

  const meta = buildMetaXml(metaInputFor(model), modified);

  const media = referencedMedia(model);
  const mediaEntries = media
    .map(
      (part) =>
        `<manifest:file-entry manifest:full-path="${escapeXmlAttribute(part.mediaId)}" manifest:media-type="${escapeXmlAttribute(part.contentType)}"/>`,
    )
    .join("");
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="${ODP_MIMETYPE}"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    mediaEntries +
    `</manifest:manifest>`;

  const entries: ZipEntryInput[] = [
    { name: "mimetype", data: encode(ODP_MIMETYPE), store: true },
    { name: "content.xml", data: encode(content) },
    { name: "styles.xml", data: encode(styles) },
    { name: "meta.xml", data: encode(meta) },
    { name: "META-INF/manifest.xml", data: encode(manifest) },
  ];
  for (const part of media) {
    entries.push({ name: part.mediaId, data: base64ToBytes(part.data) });
  }
  return entries;
}
