// Serialises a DocumentModel into an OpenDocument Text package — the
// inverse of reader.ts and the ODF mirror of docx/writer.ts. Every field
// the reader maps into the model is written back out (text, run formatting,
// paragraph formatting, style references, page layout), each as the exact
// inverse of the reader's mapping, so open -> save -> re-open reproduces
// the model.
//
// Emission choices that keep the round trip exact:
//  - ODF has no direct formatting: every run and paragraph property is
//    expressed through an automatic style in content.xml, referenced by
//    text:style-name. A paragraph's automatic style is named after its
//    styleId so the model's style reference survives re-opening; run styles
//    get generated "T<n>" names, which is lossless because the reader never
//    records span style names in the model.
//  - Styles carry the model's *effective* formatting and no parents. The
//    reader resolves style chains into the model, so re-emitting the merged
//    result as a flat style reproduces it without reconstructing the
//    source's style hierarchy.
//  - styles.xml's default-style carries the model's default font family and
//    size — the one place the reader looks for defaultStyles.fontFamily/
//    fontSizePt. The reader merges those defaults into every run it reads;
//    that cannot disturb a round trip because any model that came from an
//    ODF file already carries the merged values on each run.
//  - Lengths are formatted by the exact-round-trip helpers in units.ts, so
//    re-parsing yields bit-identical doubles.
//  - Literal spaces that ODF whitespace collapsing would eat (leading,
//    consecutive, after a tab or line break) are emitted as text:s.
import type {
  DocumentModel,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunFormatting,
} from "../documentModel";
import { zip } from "../docx/zipWriter";
import { ODT_MIMETYPE } from "./reader";
import { HIGHLIGHT_NAMES } from "./properties";
import { FO_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TEXT_NAMESPACE } from "./xml";
import { cmToOdfLength, multiplierToOdfPercentage, ptToOdfLength } from "./units";

export { ODT_MIMETYPE };

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

const MANIFEST_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const DC_NAMESPACE = "http://purl.org/dc/elements/1.1/";

// Declared on every document root; styles.xml needs all four, the others a
// subset, and surplus declarations are harmless.
const DOCUMENT_XMLNS =
  ` xmlns:office="${OFFICE_NAMESPACE}"` +
  ` xmlns:style="${STYLE_NAMESPACE}"` +
  ` xmlns:text="${TEXT_NAMESPACE}"` +
  ` xmlns:fo="${FO_NAMESPACE}"`;

/** The fallback title for a blank/whitespace-only model title — "Document 1"
 * per SPEC-007's decisions (unlike the .docx writer, which falls back to the
 * product's default title). */
export const UNTITLED_ODT_TITLE = "Document 1";

/** The suggested download file name: `<title>.odt` (AC-2). */
export function odtFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : UNTITLED_ODT_TITLE}.odt`;
}

/** Escapes text for use in XML text content or attribute values. */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// The OOXML highlight-name -> ODF background colour mapping, inverted from
// the reader's palette.
const HIGHLIGHT_COLORS: Record<string, string> = Object.fromEntries(
  Object.entries(HIGHLIGHT_NAMES).map(([hex, name]) => [name, hex]),
);

/** A run's highlight as an fo:background-color value: palette names map to
 * their hex colour, raw CSS colours pass through, anything else (an OOXML
 * name outside the standard palette) has no ODF representation. */
function highlightToOdfColor(highlight: string | undefined): string | undefined {
  if (highlight === undefined) {
    return undefined;
  }
  return highlight.startsWith("#") ? highlight : HIGHLIGHT_COLORS[highlight];
}

/** A style:text-properties element for the run formatting, or "" when there
 * is nothing to say. Explicit false values are written as "normal"/"none" —
 * the model only holds false where the source explicitly switched the
 * property off, and dropping that would change the re-opened model. */
function textPropertiesXml(formatting: RunFormatting): string {
  const attrs: string[] = [];
  if (formatting.fontFamily !== undefined) {
    attrs.push(`fo:font-family="${escapeXml(formatting.fontFamily)}"`);
  }
  if (formatting.fontSizePt !== undefined) {
    attrs.push(`fo:font-size="${ptToOdfLength(formatting.fontSizePt)}"`);
  }
  if (formatting.bold !== undefined) {
    attrs.push(`fo:font-weight="${formatting.bold ? "bold" : "normal"}"`);
  }
  if (formatting.italic !== undefined) {
    attrs.push(`fo:font-style="${formatting.italic ? "italic" : "normal"}"`);
  }
  if (formatting.underline !== undefined) {
    attrs.push(`style:text-underline-style="${formatting.underline ? "solid" : "none"}"`);
  }
  if (formatting.strikethrough !== undefined) {
    attrs.push(`style:text-line-through-style="${formatting.strikethrough ? "solid" : "none"}"`);
  }
  if (formatting.color !== undefined) {
    attrs.push(`fo:color="${escapeXml(formatting.color)}"`);
  }
  const background = highlightToOdfColor(formatting.highlight);
  if (background !== undefined) {
    attrs.push(`fo:background-color="${escapeXml(background)}"`);
  }
  return attrs.length > 0 ? `<style:text-properties ${attrs.join(" ")}/>` : "";
}

/** A style:paragraph-properties element for the paragraph formatting, or ""
 * when there is nothing to say. */
function paragraphPropertiesXml(formatting: ParagraphFormatting): string {
  const attrs: string[] = [];
  if (formatting.alignment !== undefined) {
    // The model's alignment values are all legal fo:text-align keywords.
    attrs.push(`fo:text-align="${formatting.alignment}"`);
  }
  if (formatting.indentLeftCm !== undefined) {
    attrs.push(`fo:margin-left="${cmToOdfLength(formatting.indentLeftCm)}"`);
  }
  if (formatting.indentRightCm !== undefined) {
    attrs.push(`fo:margin-right="${cmToOdfLength(formatting.indentRightCm)}"`);
  }
  if (formatting.indentFirstLineCm !== undefined) {
    attrs.push(`fo:text-indent="${cmToOdfLength(formatting.indentFirstLineCm)}"`);
  }
  if (formatting.spaceBeforePt !== undefined) {
    attrs.push(`fo:margin-top="${ptToOdfLength(formatting.spaceBeforePt)}"`);
  }
  if (formatting.spaceAfterPt !== undefined) {
    attrs.push(`fo:margin-bottom="${ptToOdfLength(formatting.spaceAfterPt)}"`);
  }
  // A single fo:line-height carries either rule; lineSpacingPt wins when
  // both are set, matching the model's documented precedence and the .docx
  // writer.
  if (formatting.lineSpacingPt !== undefined) {
    attrs.push(`fo:line-height="${ptToOdfLength(formatting.lineSpacingPt)}"`);
  } else if (formatting.lineSpacing !== undefined) {
    attrs.push(`fo:line-height="${multiplierToOdfPercentage(formatting.lineSpacing)}"`);
  }
  return attrs.length > 0 ? `<style:paragraph-properties ${attrs.join(" ")}/>` : "";
}

/** Assigns and serialises the automatic styles of content.xml. Paragraph
 * styles are named after the paragraph's styleId (all styleIds are reserved
 * up front, so a generated name can never shadow one); styleless formatted
 * paragraphs get "P<n>" and runs share "T<n>" styles per distinct
 * formatting. */
interface AutomaticStyles {
  /** Serialised style:style elements, in first-use order. */
  defs: string[];
  paragraphNameFor(paragraph: Paragraph): string | undefined;
  textNameFor(formatting: RunFormatting): string;
}

function createAutomaticStyles(model: DocumentModel): AutomaticStyles {
  const takenNames = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.styleId) {
        takenNames.add(block.styleId);
      }
    }
  }

  // Paragraph and text styles share one defs list, serialised in first-use
  // order — deterministic, so a re-saved model reproduces identical bytes.
  const defs: string[] = [];
  // The serialised properties element doubles as the formatting signature:
  // emission order is fixed, so equal formatting yields equal XML.
  const paragraphNames = new Map<string, string>();
  const textNames = new Map<string, string>();
  const boundStyleIds = new Set<string>();

  const claimName = (base: string, separator: string): string => {
    for (let n = 1; ; n++) {
      const candidate = `${base}${separator}${n}`;
      if (!takenNames.has(candidate)) {
        takenNames.add(candidate);
        return candidate;
      }
    }
  };

  return {
    defs,

    paragraphNameFor(paragraph: Paragraph): string | undefined {
      const propertiesXml = paragraphPropertiesXml(paragraph.formatting ?? {});
      // Nothing to reference and nothing to say: emit no style at all, so
      // the paragraph re-opens without a styleId, matching the model.
      if (!paragraph.styleId && propertiesXml === "") {
        return undefined;
      }
      const key = `${paragraph.styleId ?? ""} ${propertiesXml}`;
      const existing = paragraphNames.get(key);
      if (existing !== undefined) {
        return existing;
      }
      let name: string;
      if (paragraph.styleId && !boundStyleIds.has(paragraph.styleId)) {
        // The first formatting seen under a styleId gets the id itself, so
        // the reference round-trips.
        boundStyleIds.add(paragraph.styleId);
        name = paragraph.styleId;
      } else {
        // A styleless paragraph, or (possible only for models that came
        // from .docx, where direct formatting exists) a styleId already
        // bound to different formatting.
        name = paragraph.styleId ? claimName(paragraph.styleId, "-") : claimName("P", "");
      }
      paragraphNames.set(key, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${escapeXml(name)}" style:family="paragraph"/>`
          : `<style:style style:name="${escapeXml(name)}" style:family="paragraph">${propertiesXml}</style:style>`,
      );
      return name;
    },

    textNameFor(formatting: RunFormatting): string {
      const propertiesXml = textPropertiesXml(formatting);
      const existing = textNames.get(propertiesXml);
      if (existing !== undefined) {
        return existing;
      }
      const name = `T${textNames.size + 1}`;
      textNames.set(propertiesXml, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${name}" style:family="text"/>`
          : `<style:style style:name="${name}" style:family="text">${propertiesXml}</style:style>`,
      );
      return name;
    },
  };
}

/** Whitespace context threaded through a paragraph's runs: the last
 * character emitted so far — the writer-side mirror of the reader's
 * InlineAccumulator.lastChar, deciding which spaces need text:s. */
interface WhitespaceState {
  lastChar: string;
}

/** A run's text as literal text, text:s, text:tab and text:line-break —
 * the inverse of the reader's whitespace handling. A literal space survives
 * re-opening only after a non-whitespace character; every other space
 * (leading, consecutive, after a tab or break) must be an explicit text:s. */
function runContentXml(text: string, state: WhitespaceState): string {
  let xml = "";
  let buffer = "";
  const flush = () => {
    if (buffer.length > 0) {
      xml += escapeXml(buffer);
      buffer = "";
    }
  };
  for (let i = 0; i < text.length; ) {
    const char = text[i];
    if (char === "\t") {
      flush();
      xml += "<text:tab/>";
      state.lastChar = "\t";
      i++;
    } else if (char === "\n") {
      flush();
      xml += "<text:line-break/>";
      state.lastChar = "\n";
      i++;
    } else if (char === " ") {
      let count = 1;
      while (i + count < text.length && text[i + count] === " ") {
        count++;
      }
      let explicit = count;
      if (state.lastChar !== "" && !" \t\n".includes(state.lastChar)) {
        buffer += " ";
        explicit--;
      }
      flush();
      if (explicit > 0) {
        xml += explicit > 1 ? `<text:s text:c="${explicit}"/>` : "<text:s/>";
      }
      state.lastChar = " ";
      i += count;
    } else {
      buffer += char;
      state.lastChar = char;
      i++;
    }
  }
  flush();
  return xml;
}

function runXml(run: Run, styles: AutomaticStyles, state: WhitespaceState): string {
  const styleName = styles.textNameFor(run.formatting);
  return `<text:span text:style-name="${styleName}">${runContentXml(run.text, state)}</text:span>`;
}

/** A paragraph as text:p — or text:h with text:outline-level when the model
 * carries a heading level, so the level survives re-opening regardless of
 * the style's definition. */
function paragraphXml(paragraph: Paragraph, styles: AutomaticStyles): string {
  const styleName = styles.paragraphNameFor(paragraph);
  const tag = paragraph.headingLevel !== undefined ? "text:h" : "text:p";
  let attrs = styleName !== undefined ? ` text:style-name="${escapeXml(styleName)}"` : "";
  if (paragraph.headingLevel !== undefined) {
    attrs += ` text:outline-level="${paragraph.headingLevel}"`;
  }
  const state: WhitespaceState = { lastChar: "" };
  const content = paragraph.runs.map((run) => runXml(run, styles, state)).join("");
  return content.length > 0 ? `<${tag}${attrs}>${content}</${tag}>` : `<${tag}${attrs}/>`;
}

/** The content part: automatic styles plus the document body. All sections'
 * blocks are written in order (page setup lives on defaultStyles, not per
 * section); a model with no blocks still gets one empty text:p, mirroring
 * the .docx writer. */
export function buildContentXml(model: DocumentModel): string {
  const styles = createAutomaticStyles(model);
  const blocks = model.sections.flatMap((section) => section.blocks);
  // Serialised before the styles are read: paragraphXml registers each
  // automatic style on first use.
  const body =
    blocks.length > 0 ? blocks.map((block) => paragraphXml(block, styles)).join("") : "<text:p/>";
  return (
    `${XML_DECLARATION}<office:document-content${DOCUMENT_XMLNS} office:version="1.2">` +
    `<office:automatic-styles>${styles.defs.join("")}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body>` +
    "</office:document-content>"
  );
}

/** The styles part: a paragraph default-style carrying the model's default
 * font (the reader's source for defaultStyles.fontFamily/fontSizePt), and
 * the page layout applied through the first master page (AC-15). */
export function buildStylesXml(model: DocumentModel): string {
  const { fontFamily, fontSizePt, pageSize, margins } = model.defaultStyles;
  // ODF stores landscape pages with width/height already swapped; the
  // orientation attribute just marks them, mirroring the .docx writer's
  // w:orient.
  const orientation =
    pageSize.widthCm > pageSize.heightCm ? ' style:print-orientation="landscape"' : "";
  return (
    `${XML_DECLARATION}<office:document-styles${DOCUMENT_XMLNS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph">' +
    `<style:text-properties fo:font-family="${escapeXml(fontFamily)}" fo:font-size="${ptToOdfLength(fontSizePt)}"/>` +
    "</style:default-style>" +
    "</office:styles>" +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1">' +
    `<style:page-layout-properties fo:page-width="${cmToOdfLength(pageSize.widthCm)}"` +
    ` fo:page-height="${cmToOdfLength(pageSize.heightCm)}"${orientation}` +
    ` fo:margin-top="${cmToOdfLength(margins.topCm)}" fo:margin-right="${cmToOdfLength(margins.rightCm)}"` +
    ` fo:margin-bottom="${cmToOdfLength(margins.bottomCm)}" fo:margin-left="${cmToOdfLength(margins.leftCm)}"/>` +
    "</style:page-layout>" +
    "</office:automatic-styles>" +
    "<office:master-styles>" +
    '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
    "</office:master-styles>" +
    "</office:document-styles>"
  );
}

/** A timestamp in the W3CDTF profile ODF metadata uses (whole seconds,
 * UTC), e.g. "2026-08-03T12:00:00Z". */
function toW3CDTF(date: Date): string {
  return date.toISOString().replace(/\.\d+Z$/, "Z");
}

/** meta.xml with the document title (AC-19) and dc:date, ODF's modification
 * timestamp (AC-20). */
export function buildMetaXml(title: string, modified: Date): string {
  return (
    `${XML_DECLARATION}<office:document-meta xmlns:office="${OFFICE_NAMESPACE}"` +
    ` xmlns:dc="${DC_NAMESPACE}" office:version="1.2">` +
    `<office:meta><dc:title>${escapeXml(title)}</dc:title>` +
    `<dc:date>${toW3CDTF(modified)}</dc:date></office:meta>` +
    "</office:document-meta>"
  );
}

/** settings.xml (AC-7). The app has no settings to persist; office:settings
 * is optional in the schema, and omitting it entirely is valid while an
 * empty office:settings element is not. */
export function buildSettingsXml(): string {
  return `${XML_DECLARATION}<office:document-settings xmlns:office="${OFFICE_NAMESPACE}" office:version="1.2"/>`;
}

/** META-INF/manifest.xml listing every part (AC-8). The mimetype entry is
 * not listed — per ODF its media type belongs to the "/" entry. */
export function buildManifestXml(): string {
  const part = (path: string) =>
    `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml"/>`;
  return (
    `${XML_DECLARATION}<manifest:manifest xmlns:manifest="${MANIFEST_NAMESPACE}" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIMETYPE}"/>` +
    part("content.xml") +
    part("styles.xml") +
    part("meta.xml") +
    part("settings.xml") +
    "</manifest:manifest>"
  );
}

export interface WriteOdtOptions {
  /** The dc:date modification timestamp; defaults to now. Injectable so
   * tests can assert an exact meta.xml. */
  modified?: Date;
}

/** Serialises the model into `.odt` bytes: a valid ODF package whose
 * uncompressed mimetype entry comes first (AC-3, AC-9, AC-10), followed by
 * content.xml, styles.xml, meta.xml, settings.xml and META-INF/manifest.xml
 * (AC-4..AC-8). */
export async function writeOdt(
  model: DocumentModel,
  options: WriteOdtOptions = {},
): Promise<Uint8Array> {
  const modified = options.modified ?? new Date();
  const title = model.title.trim().length > 0 ? model.title : UNTITLED_ODT_TITLE;
  const encoder = new TextEncoder();
  const content = buildContentXml(model);
  return zip([
    // ODF requires the mimetype entry first and STOREd so format sniffers
    // can read it from the raw bytes at a fixed offset.
    { name: "mimetype", data: encoder.encode(ODT_MIMETYPE), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(buildStylesXml(model)) },
    { name: "meta.xml", data: encoder.encode(buildMetaXml(title, modified)) },
    { name: "settings.xml", data: encoder.encode(buildSettingsXml()) },
    { name: "META-INF/manifest.xml", data: encoder.encode(buildManifestXml()) },
  ]);
}
