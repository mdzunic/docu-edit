#!/usr/bin/env node
// One-off generator for the SPEC-071 freeform fixtures (run manually:
// `node scripts/build-spec071-fixtures.mjs`). Like build-spec051-fixtures.mjs
// (preset shapes), these are hand-built packages reproducing what Word /
// LibreOffice Writer emit for hand-drawn geometry — NOT round-tripped through
// this project's own writer — so opening them is a genuine test of the custom
// geometry readers (parseCustGeom in src/lib/docx/shape.ts, parseEnhancedPath
// in src/lib/core/enhancedPath.ts), not "the reader understands its own
// writer's output".
//
// Neither fixture carries the `de:smooth` extension attribute this project
// writes for its own path shapes: a foreign file has no per-vertex smoothness
// to declare, and the reader must cope without it. Smoothness round-tripping
// is covered by src/lib/shapePath.roundTrip.spec.ts instead.
//
// Shapes are written in a fixed document order so the e2e spec can address
// each one by its 0-based `data-testid="drawing-<n>"` index:
//   0  Word freeform         open polyline (moveTo + lnTo x4)      outline only
//   1  Closed freeform       polygon (moveTo + lnTo x4 + close)    filled
//   2  Closed curve          moveTo + cubicBezTo x4 + close        filled
//   3  Open scribble         moveTo + cubicBezTo x2                outline only
//   4  Smooth and corner mix cubic + line + close                  filled
//   5  Reduced primitives    quadratic + arc, converted on read    filled
//   6  Formula-driven        guide-referencing coordinates         PRESERVED
//   7  Unsupported geometry  outside what the model expresses      PRESERVED
//
// 0..5 are path shapes: they parse into a `ShapePath` and paint their real
// outline (AC-4, AC-5). 6 and 7 are what the narrowed placeholder scope keeps
// on SPEC-051's preservation path (AC-31, AC-32) — 6 because its coordinates
// are `draw:equation`/`a:gd` results a flattened path could not reproduce, 7
// because the geometry itself is outside what the model expresses. The two
// formats express "7" differently, since each has a different natural way to
// be unmodellable: .docx uses a `prstGeom` outside the catalogue, .odt an
// enhanced-path using the `W` (elliptical-quadrant) command the parser does
// not model.
import { deflateRawSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function u32le(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32LE(value >>> 0, 0);
  return buf;
}
function u16le(value) {
  const buf = Buffer.alloc(2);
  buf.writeUInt16LE(value, 0);
  return buf;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

// --- minimal ZIP writer (shared shape with build-spec051-fixtures.mjs) -----

function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.isBuffer(entry.content) ? entry.content : Buffer.from(entry.content, "utf-8");
    const method = entry.store ? 0 : 8;
    const dataBytes = entry.store ? rawBytes : deflateRawSync(rawBytes);
    const crc = crc32(rawBytes);
    const version = entry.store ? 10 : 20;

    const localHeader = Buffer.concat([
      u32le(0x04034b50),
      u16le(version),
      u16le(0),
      u16le(method),
      u16le(0),
      u16le(0),
      u32le(crc),
      u32le(dataBytes.length),
      u32le(rawBytes.length),
      u16le(nameBytes.length),
      u16le(0),
      nameBytes,
    ]);
    localParts.push(localHeader, dataBytes);

    const centralHeader = Buffer.concat([
      u32le(0x02014b50),
      u16le(20),
      u16le(version),
      u16le(0),
      u16le(method),
      u16le(0),
      u16le(0),
      u32le(crc),
      u32le(dataBytes.length),
      u32le(rawBytes.length),
      u16le(nameBytes.length),
      u16le(0),
      u16le(0),
      u16le(0),
      u16le(0),
      u32le(0),
      u32le(offset),
      nameBytes,
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  }

  const centralStart = offset;
  const centralBytes = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32le(0x06054b50),
    u16le(0),
    u16le(0),
    u16le(entries.length),
    u16le(entries.length),
    u32le(centralBytes.length),
    u32le(centralStart),
    u16le(0),
  ]);
  return Buffer.concat([...localParts, centralBytes, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function escapeXml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Every path is authored once, on the 21600 grid both formats use, and each
// builder renders it into its own command vocabulary — so the two fixtures
// hold the *same* eight shapes and the e2e spec can assert one set of
// outlines against both.
const GRID = 21600;

// The shared entry list. `steps` is the authored path (absent for the two
// preserved entries, which each builder expresses in its own way).
const SHAPES = [
  {
    key: "freeform",
    label: "Word freeform",
    fill: undefined, // open: outline only, per AC-9
    steps: [
      { op: "move", x: 0, y: 21600 },
      { op: "line", x: 5400, y: 0 },
      { op: "line", x: 10800, y: 16200 },
      { op: "line", x: 16200, y: 2700 },
      { op: "line", x: 21600, y: 21600 },
    ],
  },
  {
    key: "closedPolygon",
    label: "Closed freeform",
    fill: true,
    steps: [
      { op: "move", x: 10800, y: 0 },
      { op: "line", x: 21600, y: 8100 },
      { op: "line", x: 17280, y: 21600 },
      { op: "line", x: 4320, y: 21600 },
      { op: "line", x: 0, y: 8100 },
      { op: "close" },
    ],
  },
  {
    // A circle drawn as four cubics (the k = 0.5523 control-point rule) —
    // what a Curve tool produces, and geometry a bounding box cannot fake:
    // its corners are outside the outline.
    key: "closedCurve",
    label: "Closed curve",
    fill: true,
    steps: [
      { op: "move", x: 0, y: 10800 },
      { op: "cubic", x1: 0, y1: 4835, x2: 4835, y2: 0, x: 10800, y: 0 },
      { op: "cubic", x1: 16765, y1: 0, x2: 21600, y2: 4835, x: 21600, y: 10800 },
      { op: "cubic", x1: 21600, y1: 16765, x2: 16765, y2: 21600, x: 10800, y: 21600 },
      { op: "cubic", x1: 4835, y1: 21600, x2: 0, y2: 16765, x: 0, y: 10800 },
      { op: "close" },
    ],
  },
  {
    key: "openScribble",
    label: "Open scribble",
    fill: undefined,
    steps: [
      { op: "move", x: 0, y: 10800 },
      { op: "cubic", x1: 3600, y1: 2160, x2: 7200, y2: 19440, x: 10800, y: 10800 },
      { op: "cubic", x1: 14400, y1: 2160, x2: 18000, y2: 19440, x: 21600, y: 10800 },
    ],
  },
  {
    // One curved side meeting two straight ones: the vertices where the
    // cubic meets a line are corners, the ones inside the curve are smooth.
    key: "mixed",
    label: "Smooth and corner mix",
    fill: true,
    steps: [
      { op: "move", x: 0, y: 21600 },
      { op: "cubic", x1: 5400, y1: 0, x2: 16200, y2: 0, x: 21600, y: 21600 },
      { op: "line", x: 10800, y: 16200 },
      { op: "close" },
    ],
  },
];

// Shape 5 is authored per-format below (a quadratic + an arc), because the
// two formats spell those primitives differently; both reduce to the same
// cubics on read. Shapes 6 and 7 carry no `steps` at all.
const REDUCED_LABEL = "Reduced primitives";
const FORMULA_LABEL = "Formula-driven geometry";
const UNSUPPORTED_LABEL = "Unsupported geometry";

const OUTLINE = { color: "cc0000", widthPt: 1.5 };
const FILL = "3366cc";
const WIDTH_CM = 4;
const HEIGHT_CM = 3;

// --- .docx fixture -----------------------------------------------------------
{
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

  const cmToEmu = (cm) => Math.round(cm * 360000);
  const ptToEmu = (pt) => Math.round(pt * 12700);

  const pt = (x, y) => `<a:pt x="${x}" y="${y}"/>`;

  /** Renders an authored step list as `a:pathLst` command elements. */
  function custGeomXml(steps) {
    let body = "";
    for (const step of steps) {
      if (step.op === "move") {
        body += `<a:moveTo>${pt(step.x, step.y)}</a:moveTo>`;
      } else if (step.op === "line") {
        body += `<a:lnTo>${pt(step.x, step.y)}</a:lnTo>`;
      } else if (step.op === "cubic") {
        body += `<a:cubicBezTo>${pt(step.x1, step.y1)}${pt(step.x2, step.y2)}${pt(step.x, step.y)}</a:cubicBezTo>`;
      } else {
        body += "<a:close/>";
      }
    }
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="${GRID}" b="${GRID}"/>` +
      `<a:pathLst><a:path w="${GRID}" h="${GRID}">${body}</a:path></a:pathLst></a:custGeom>`;
  }

  // Shape 5: a quadratic and a 90-degree arc, the two primitives the model
  // reduces to cubics at the reader boundary (AC-29). swAng is in 1/60000
  // of a degree, so 5400000 is a quarter turn.
  const reducedGeomXml =
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="${GRID}" b="${GRID}"/>` +
    `<a:pathLst><a:path w="${GRID}" h="${GRID}">` +
    `<a:moveTo>${pt(0, 21600)}</a:moveTo>` +
    `<a:quadBezTo>${pt(10800, 0)}${pt(21600, 10800)}</a:quadBezTo>` +
    '<a:arcTo wR="10800" hR="10800" stAng="0" swAng="5400000"/>' +
    "<a:close/></a:path></a:pathLst></a:custGeom>";

  // Shape 6: coordinates that are guide names, not numbers — the shape of a
  // parametric custGeom whose outline follows a modifier drag. Flattening it
  // would freeze that behaviour, so it stays preserved verbatim (AC-31).
  const formulaGeomXml =
    "<a:custGeom>" +
    '<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>' +
    '<a:gdLst><a:gd name="x1" fmla="*/ w adj 100000"/></a:gdLst>' +
    `<a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="${GRID}" b="${GRID}"/>` +
    `<a:pathLst><a:path w="${GRID}" h="${GRID}">` +
    '<a:moveTo><a:pt x="x1" y="0"/></a:moveTo>' +
    `<a:lnTo>${pt(21600, 21600)}</a:lnTo>` +
    `<a:lnTo>${pt(0, 21600)}</a:lnTo>` +
    "<a:close/></a:path></a:pathLst></a:custGeom>";

  // Shape 7: a preset outside SPEC-069's catalogue — the placeholder case the
  // spec names explicitly. `gear9` is a real OOXML preset with no catalogue
  // entry (and no plausible one: it is nine parametric teeth, not an outline
  // the model could carry).
  const unsupportedGeomXml = '<a:prstGeom prst="gear9"><a:avLst/></a:prstGeom>';

  function shapeDrawing({ id, geomXml, fill, altText }) {
    const cx = cmToEmu(WIDTH_CM);
    const cy = cmToEmu(HEIGHT_CM);
    const spPr =
      "<wps:spPr>" +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      geomXml +
      (fill ? `<a:solidFill><a:srgbClr val="${FILL}"/></a:solidFill>` : "<a:noFill/>") +
      `<a:ln w="${ptToEmu(OUTLINE.widthPt)}"><a:solidFill><a:srgbClr val="${OUTLINE.color}"/></a:solidFill></a:ln>` +
      "</wps:spPr>";
    const wsp = `<wps:wsp xmlns:wps="${WPS}"><wps:cNvSpPr/>${spPr}</wps:wsp>`;
    const graphic = `<a:graphic xmlns:a="${A}"><a:graphicData uri="${WPS}">${wsp}</a:graphicData></a:graphic>`;
    return (
      `<w:drawing xmlns:wp="${WP}">` +
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      `<wp:docPr id="${id}" name="Shape ${id}" descr="${escapeXml(altText)}"/>` +
      "<wp:cNvGraphicFramePr/>" +
      graphic +
      "</wp:inline></w:drawing>"
    );
  }

  let nextId = 1;
  const drawings = SHAPES.map((entry) =>
    shapeDrawing({ id: nextId++, geomXml: custGeomXml(entry.steps), fill: entry.fill, altText: entry.label }),
  );
  drawings.push(shapeDrawing({ id: nextId++, geomXml: reducedGeomXml, fill: true, altText: REDUCED_LABEL }));
  drawings.push(shapeDrawing({ id: nextId++, geomXml: formulaGeomXml, fill: true, altText: FORMULA_LABEL }));
  drawings.push(shapeDrawing({ id: nextId++, geomXml: unsupportedGeomXml, fill: true, altText: UNSUPPORTED_LABEL }));

  const documentXml =
    `${XML}<w:document xmlns:w="${W}"><w:body>` +
    drawings.map((d) => `<w:p><w:r>${d}</w:r></w:p>`).join("") +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";

  const stylesXml =
    `${XML}<w:styles xmlns:w="${W}">` +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>' +
    "</w:rPr></w:rPrDefault></w:docDefaults>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    "</w:styles>";

  const contentTypesXml =
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>";

  const packageRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    "</Relationships>";

  const documentRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    "</Relationships>";

  const corePropsXml =
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:title>Freeform fixture</dc:title>" +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>' +
    "</cp:coreProperties>";

  const bytes = buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
    { name: "docProps/core.xml", content: corePropsXml },
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "docx", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "word-freeform.docx"), bytes);
    console.log("Wrote", path.join(outDir, "word-freeform.docx"));
  }
}

// --- .odt fixture --------------------------------------------------------
{
  const NS =
    ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
    ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
    ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
    ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
    ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
    ' xmlns:xlink="http://www.w3.org/1999/xlink"';

  const VIEW_BOX = `0 0 ${GRID} ${GRID}`;

  /** Renders an authored step list as a `draw:enhanced-path` command string.
   * ODF requires the terminating `N`. */
  function enhancedPath(steps) {
    const parts = [];
    for (const step of steps) {
      if (step.op === "move") {
        parts.push(`M ${step.x} ${step.y}`);
      } else if (step.op === "line") {
        parts.push(`L ${step.x} ${step.y}`);
      } else if (step.op === "cubic") {
        parts.push(`C ${step.x1} ${step.y1} ${step.x2} ${step.y2} ${step.x} ${step.y}`);
      } else {
        parts.push("Z");
      }
    }
    parts.push("N");
    return parts.join(" ");
  }

  function graphicStyleXml(name, fill) {
    const fillAttrs = fill ? ` draw:fill="solid" draw:fill-color="#${FILL}"` : ' draw:fill="none"';
    const strokeAttrs =
      ' draw:stroke="solid"' +
      ` svg:stroke-color="#${OUTLINE.color}"` +
      ` svg:stroke-width="${((OUTLINE.widthPt / 72) * 2.54).toFixed(4)}cm"`;
    return (
      `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="Graphics">` +
      `<style:graphic-properties${fillAttrs}${strokeAttrs}` +
      ' style:vertical-pos="from-top" style:vertical-rel="paragraph"' +
      ' style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>' +
      "</style:style>"
    );
  }

  // Shape 5: the ODF spellings of the two reduced primitives — `Q`
  // (quadratic) and `A` (arcto, given as a bounding box then a start and end
  // point). The arc runs from (21600 10800) to (10800 21600) around the box
  // centre, matching the .docx fixture's quarter-turn `a:arcTo` (AC-29).
  const reducedPath = "M 0 21600 Q 10800 0 21600 10800 A 0 0 21600 21600 21600 10800 10800 21600 Z N";

  // Shape 6: an enhanced-path whose first coordinate is a `draw:equation`
  // result rather than a number — the parametric case that cannot be
  // flattened without losing the modifier behaviour (AC-31).
  const formulaGeometryXml =
    `<draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="${VIEW_BOX}"` +
    ' draw:modifiers="5400"' +
    ' draw:enhanced-path="M ?f0 0 L 21600 21600 L 0 21600 Z N">' +
    '<draw:equation draw:name="f0" draw:formula="$0 *2"/>' +
    "</draw:enhanced-geometry>";

  // Shape 7: `W` (elliptical quadrant) — a real enhanced-path command the
  // model does not carry, so the shape stays on the preservation path and
  // paints as the dashed placeholder (AC-32).
  const unsupportedGeometryXml =
    `<draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="${VIEW_BOX}"` +
    ' draw:enhanced-path="M 0 0 W 0 0 21600 21600 0 0 21600 21600 N"/>';

  function shapeElement({ styleName, geometryXml, altText }) {
    return (
      `<draw:custom-shape text:anchor-type="as-char" draw:style-name="${styleName}"` +
      ` svg:width="${WIDTH_CM}cm" svg:height="${HEIGHT_CM}cm">` +
      `<svg:desc>${escapeXml(altText)}</svg:desc>` +
      geometryXml +
      "</draw:custom-shape>"
    );
  }

  let styleCounter = 1;
  const styles = [];
  const paragraphs = [];

  function addShape({ geometryXml, fill, altText }) {
    const styleName = `gr${styleCounter++}`;
    styles.push(graphicStyleXml(styleName, fill));
    paragraphs.push(`<text:p>${shapeElement({ styleName, geometryXml, altText })}</text:p>`);
  }

  for (const entry of SHAPES) {
    addShape({
      geometryXml:
        `<draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="${VIEW_BOX}"` +
        ` draw:enhanced-path="${escapeXml(enhancedPath(entry.steps))}"/>`,
      fill: entry.fill,
      altText: entry.label,
    });
  }
  addShape({
    geometryXml:
      `<draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="${VIEW_BOX}"` +
      ` draw:enhanced-path="${escapeXml(reducedPath)}"/>`,
    fill: true,
    altText: REDUCED_LABEL,
  });
  addShape({ geometryXml: formulaGeometryXml, fill: true, altText: FORMULA_LABEL });
  addShape({ geometryXml: unsupportedGeometryXml, fill: true, altText: UNSUPPORTED_LABEL });

  const contentXml =
    `${XML}<office:document-content${NS} office:version="1.2">` +
    `<office:automatic-styles>${styles.join("")}</office:automatic-styles>` +
    `<office:body><office:text>${paragraphs.join("")}</office:text></office:body>` +
    "</office:document-content>";

  const stylesXml =
    `${XML}<office:document-styles${NS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph">' +
    '<style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt"/>' +
    "</style:default-style>" +
    '<style:style style:name="Graphics" style:family="graphic"/>' +
    "</office:styles>" +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1">' +
    '<style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"' +
    ' fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/>' +
    "</style:page-layout>" +
    "</office:automatic-styles>" +
    "<office:master-styles>" +
    '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
    "</office:master-styles>" +
    "</office:document-styles>";

  const metaXml =
    `${XML}<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">' +
    "<office:meta><dc:title>Freeform fixture</dc:title><dc:date>2024-01-01T00:00:00Z</dc:date></office:meta>" +
    "</office:document-meta>";

  const settingsXml =
    '<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"/>';

  const manifestXml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">' +
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.text"/>' +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>' +
    "</manifest:manifest>";

  const bytes = buildZip([
    { name: "mimetype", content: "application/vnd.oasis.opendocument.text", store: true },
    { name: "content.xml", content: contentXml },
    { name: "styles.xml", content: stylesXml },
    { name: "meta.xml", content: metaXml },
    { name: "settings.xml", content: settingsXml },
    { name: "META-INF/manifest.xml", content: manifestXml },
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "odf", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "lo-freeform.odt"), bytes);
    console.log("Wrote", path.join(outDir, "lo-freeform.odt"));
  }
}
