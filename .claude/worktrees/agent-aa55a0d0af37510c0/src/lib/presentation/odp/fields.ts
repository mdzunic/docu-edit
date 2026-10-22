// The odp wire model for deck field placeholders (SPEC-093): the
// `presentation:footer-decl`/`date-time-decl` block regenerated into the
// content envelope's prefix, the `number:date-style` an auto date needs, the
// master-page placeholders a from-scratch deck emits, and the per-page plan
// (`use-*-name` + `display-*`) `serializePage` stamps. One source of truth so
// the substitution path, the from-scratch path and the golden check agree.
//
// Field placeholders are never serialised as `draw:frame`s (AC-80) — a page
// shows a field purely through its declaration reference + display flag.
import { escapeXmlAttribute, escapeXmlText } from "../../core/xml";
import { cmToOdfLength } from "../../core/odfXml";
import { dateTimeFormatEntry, odtDateStyleName } from "../../core/dateTimeFormats";
import { DEFAULT_DATE_FORMAT, defaultFieldRect, isTitleSlide } from "../headerFooter";
import type { DeckHeaderFooter, PresentationModel, Slide, SlideFieldKind } from "../model";

export const NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/** The `presentation:class` an odp master placeholder / display flag uses. */
const ODP_CLASS: Record<SlideFieldKind, "footer" | "page-number" | "date-time"> = {
  footer: "footer",
  slideNumber: "page-number",
  date: "date-time",
};

/** A page's resolved field references + visibility (SPEC-093). */
export interface OdpPageFieldPlan {
  footerName?: string;
  dateName?: string;
  showFooter: boolean;
  showPageNumber: boolean;
  showDateTime: boolean;
}

/** One date-time declaration's identity. */
interface DateDecl {
  source: "fixed" | "current-date";
  dataStyleName?: string;
  text: string;
}

/** The deck-wide declaration plan (SPEC-093): the regenerated decl block, the
 * `number:date-style` styles auto dates need, and each page's plan. */
export interface OdpDeclarationPlan {
  declXml: string;
  dateStyleXml: string;
  pagePlans: Map<Slide, OdpPageFieldPlan>;
}

/** Whether a slide's fields are suppressed (a title slide with the deck's
 * "don't show on title slide" set). */
function suppressed(model: PresentationModel, index: number): boolean {
  return model.headerFooter?.showOnTitleSlide === false && isTitleSlide(model, index);
}

/** The date declaration a slide's date field references: a per-slide fixed
 * literal, a preserved out-of-subset style, or the deck default. */
function slideDateDecl(hf: DeckHeaderFooter | undefined, field: { text?: string; wireStyleName?: string }): DateDecl {
  if (field.text !== undefined) {
    return { source: "fixed", text: field.text };
  }
  if (field.wireStyleName !== undefined) {
    return { source: "current-date", dataStyleName: field.wireStyleName, text: "" };
  }
  return deckDateDecl(hf);
}

/** The deck's default date declaration (auto from the catalogue format, or a
 * deck-wide fixed date). */
function deckDateDecl(hf: DeckHeaderFooter | undefined): DateDecl {
  if (hf?.dateSource === "fixed") {
    return { source: "fixed", text: hf.dateFixedText ?? "" };
  }
  const entry = dateTimeFormatEntry(hf?.dateFormat ?? DEFAULT_DATE_FORMAT);
  return { source: "current-date", dataStyleName: entry ? odtDateStyleName(entry) : undefined, text: "" };
}

const dateKey = (decl: DateDecl): string =>
  decl.source === "fixed" ? `fixed:${decl.text}` : `auto:${decl.dataStyleName ?? ""}`;

/** Parses `value → name` from a source prefix's footer declarations, so an
 * unchanged value reuses its source name (AC-77) and verbatim pages stay
 * valid. */
function sourceFooterNames(prefix: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const m of prefix.matchAll(
    /<presentation:footer-decl\b[^>]*\bpresentation:name="([^"]*)"[^>]*>([\s\S]*?)<\/presentation:footer-decl>/g,
  )) {
    map.set(m[2], m[1]); // escaped value → name
  }
  return map;
}

/** Parses `dateKey → name` from a source prefix's date-time declarations. */
function sourceDateNames(prefix: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const m of prefix.matchAll(/<presentation:date-time-decl\b([^>]*?)(?:\/>|>([\s\S]*?)<\/presentation:date-time-decl>)/g)) {
    const attrs = m[1];
    const name = /\bpresentation:name="([^"]*)"/.exec(attrs)?.[1];
    if (!name) continue;
    const source = /\bpresentation:source="fixed"/.test(attrs) ? "fixed" : "current-date";
    const dataStyle = /\bstyle:data-style-name="([^"]*)"/.exec(attrs)?.[1];
    const key = source === "fixed" ? `fixed:${m[2] ?? ""}` : `auto:${dataStyle ?? ""}`;
    map.set(key, name);
  }
  return map;
}

/** The highest `${kind}${n}` declaration index already present, so fresh names
 * never collide with a source name we did not reuse. */
function maxDeclIndex(prefix: string, kind: string): number {
  let max = 0;
  for (const m of prefix.matchAll(new RegExp(`presentation:name="${kind}(\\d+)"`, "g"))) {
    max = Math.max(max, Number(m[1]));
  }
  return max;
}

/** Plans the deck's declarations + each page's field references (SPEC-093).
 * `sourcePrefix` lets an unchanged value reuse its source declaration name so
 * verbatim pages keep resolving. */
export function planOdpDeclarations(model: PresentationModel, sourcePrefix = ""): OdpDeclarationPlan {
  const hf = model.headerFooter;
  const srcFooters = sourceFooterNames(sourcePrefix);
  const srcDates = sourceDateNames(sourcePrefix);
  let footerMax = maxDeclIndex(sourcePrefix, "ftr");
  let dateMax = maxDeclIndex(sourcePrefix, "dtd");

  const footerNames = new Map<string, string>(); // plain value → name
  const footerDecls: { name: string; value: string }[] = [];
  const footerName = (value: string): string => {
    const cached = footerNames.get(value);
    if (cached) return cached;
    const name = srcFooters.get(escapeXmlText(value)) ?? `ftr${++footerMax}`;
    footerNames.set(value, name);
    footerDecls.push({ name, value });
    return name;
  };

  const dateNames = new Map<string, string>(); // dateKey → name
  const dateDecls: { name: string; decl: DateDecl }[] = [];
  const dateName = (decl: DateDecl): string => {
    const key = dateKey(decl);
    const cached = dateNames.get(key);
    if (cached) return cached;
    const name = srcDates.get(key) ?? `dtd${++dateMax}`;
    dateNames.set(key, name);
    dateDecls.push({ name, decl });
    return name;
  };

  const pagePlans = new Map<Slide, OdpPageFieldPlan>();
  model.slides.forEach((slide, index) => {
    const hidden = suppressed(model, index);
    const fieldOf = (kind: SlideFieldKind) => slide.fields?.find((f) => f.kind === kind);
    const plan: OdpPageFieldPlan = {
      showFooter: !hidden && fieldOf("footer") !== undefined,
      showPageNumber: !hidden && fieldOf("slideNumber") !== undefined,
      showDateTime: !hidden && fieldOf("date") !== undefined,
    };
    if (plan.showFooter) {
      const entry = fieldOf("footer")!;
      plan.footerName = footerName(entry.text ?? hf?.footerText ?? "");
    }
    if (plan.showDateTime) {
      plan.dateName = dateName(slideDateDecl(hf, fieldOf("date")!));
    }
    pagePlans.set(slide, plan);
  });

  const declXml =
    footerDecls
      .map((d) => `<presentation:footer-decl presentation:name="${d.name}">${escapeXmlText(d.value)}</presentation:footer-decl>`)
      .join("") + dateDecls.map((d) => dateDeclXml(d.name, d.decl)).join("");

  // A number:date-style per distinct in-subset auto date (an out-of-subset
  // style is preserved in the source, never re-emitted).
  const seenStyle = new Set<string>();
  let dateStyleXml = "";
  for (const { decl } of dateDecls) {
    if (decl.source !== "current-date" || !decl.dataStyleName) continue;
    const entry = dateTimeFormatEntry(hf?.dateFormat ?? DEFAULT_DATE_FORMAT);
    if (!entry || odtDateStyleName(entry) !== decl.dataStyleName || seenStyle.has(decl.dataStyleName)) continue;
    seenStyle.add(decl.dataStyleName);
    dateStyleXml += `<number:date-style style:name="${decl.dataStyleName}">${entry.odtStyleBody}</number:date-style>`;
  }

  return { declXml, dateStyleXml, pagePlans };
}

function dateDeclXml(name: string, decl: DateDecl): string {
  if (decl.source === "fixed") {
    return `<presentation:date-time-decl presentation:name="${name}" presentation:source="fixed">${escapeXmlText(decl.text)}</presentation:date-time-decl>`;
  }
  const style = decl.dataStyleName ? ` style:data-style-name="${escapeXmlAttribute(decl.dataStyleName)}"` : "";
  return `<presentation:date-time-decl presentation:name="${name}" presentation:source="current-date"${style}/>`;
}

/** The `presentation:display-*` property string a page's drawing-page style
 * carries (SPEC-093 AC-79). */
export function odpDisplayProps(plan: OdpPageFieldPlan): string {
  return (
    `presentation:display-footer="${plan.showFooter}" ` +
    `presentation:display-page-number="${plan.showPageNumber}" ` +
    `presentation:display-date-time="${plan.showDateTime}"`
  );
}

/** Whether the deck configures any field (the trigger for the from-scratch
 * master placeholders + the declaration block). */
export function deckHasFields(model: PresentationModel): boolean {
  return model.headerFooter !== undefined || model.slides.some((slide) => (slide.fields?.length ?? 0) > 0);
}

/** The master-page placeholder frames a from-scratch deck emits (SPEC-093
 * AC-90): one `presentation:class` frame per field kind, geometry from
 * `DEFAULT_FIELD_BAND`, with the ODF body element the class expects. */
export function odpMasterPlaceholders(model: PresentationModel): string {
  const kinds: SlideFieldKind[] = ["date", "footer", "slideNumber"];
  const bodyFor: Record<SlideFieldKind, string> = {
    footer: "<presentation:footer/>",
    slideNumber: "<text:page-number>1</text:page-number>",
    date: "<presentation:date-time/>",
  };
  return kinds
    .map((kind) => {
      const rect = defaultFieldRect(kind, model.slideWidthCm, model.slideHeightCm);
      return (
        `<draw:frame presentation:class="${ODP_CLASS[kind]}" ` +
        `svg:x="${cmToOdfLength(rect.xCm)}" svg:y="${cmToOdfLength(rect.yCm)}" ` +
        `svg:width="${cmToOdfLength(rect.widthCm)}" svg:height="${cmToOdfLength(rect.heightCm)}">` +
        `<draw:text-box>${bodyFor[kind]}</draw:text-box></draw:frame>`
      );
    })
    .join("");
}
