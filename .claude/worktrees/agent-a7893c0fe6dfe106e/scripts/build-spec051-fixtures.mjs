#!/usr/bin/env node
// One-off generator for the SPEC-051 shape fixtures (run manually:
// `node scripts/build-spec051-fixtures.mjs`). Like build-spec024-fixtures.mjs
// (images), these are hand-built packages reproducing what Word/LibreOffice
// would emit for drawn shapes — NOT round-tripped through this project's own
// writer — so opening them is a genuine test of the reader (src/lib/docx/shape.ts,
// src/lib/odf/shape.ts), not just "the reader understands its own writer's
// output".
//
// That claim was only half true until 2026-08-29: the .docx side emitted bare
// `<w:drawing>` elements, and Word emits none. Word wraps every shape in
// `mc:AlternateContent` (an `mc:Choice Requires="wps"` beside a legacy VML
// `mc:Fallback`), which the reader could not see through — so SPEC-051's
// shape import passed this fixture while failing on every real Word document.
// Keep the wrapper when editing this script; a bare drawing here tests markup
// no producer writes. Each fixture contains, per AC-63/AC-64: every catalogue geometry
// (SHAPE_CATALOGUE in src/lib/documentModel.ts), a grouped pair of shapes
// sharing a `de:group` ordinal, a text box (a fill-less/outline-less rect
// with a wps:txbx/draw:text-box), and one geometry outside the catalogue
// (rendered as a dashed placeholder on open, AC-60).
//
// Re-running this must leave every part's *content* identical to the
// committed fixtures — that is the check that this script has not drifted
// from them. The packaged bytes may still differ by a few: deflate output is
// not stable across zlib/Node versions, so `git checkout` the four fixture
// files when only their compressed sizes moved.
//
// Shapes are written in a fixed document order so the e2e spec can address
// each one by its 0-based `data-testid="drawing-<n>"` index (DocShapeBlock.vue
// assigns that index by walking Shape blocks in document order):
//   0  line              7  star5              14 group member A (rect)
//   1  arrow             8  rightArrow         15 group member B (roundRect)
//   2  rect              9  leftArrow          16 text box (rect, no fill/outline)
//   3  roundRect        10  upArrow            17 unsupported geometry ("gear9"/"gear")
//   4  ellipse          11  downArrow          18 group container member A
//   5  triangle         12  wedgeRectCallout   19 group container member B
//   6  diamond          13  wedgeRoundRectCallout
//
// 14/15 are the FLAT grouping this editor writes (`de:group` on two sibling
// drawings); 18/19 are one real container — a Word `wpg:wgp` in the .docx,
// a LibreOffice `draw:g` in the .odt — that group.ts flattens into two
// blocks. Both shapes of grouping exist in the wild and the fixtures have to
// carry both.
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

// --- minimal ZIP writer (shared shape with build-spec024-fixtures.mjs) -----

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

// The catalogue this fixture must exercise (mirrors SHAPE_CATALOGUE order in
// src/lib/documentModel.ts) plus the grouped pair / text box / unsupported
// entries appended after it, all sharing one flat ordered list so both
// builders (.docx/.odt) walk the identical sequence.
const CATALOGUE = [
  { geometry: "line", label: "Line" },
  { geometry: "arrow", label: "Arrow" },
  { geometry: "rect", label: "Rectangle" },
  { geometry: "roundRect", label: "Rounded rectangle" },
  { geometry: "ellipse", label: "Ellipse" },
  { geometry: "triangle", label: "Triangle" },
  { geometry: "diamond", label: "Diamond" },
  { geometry: "star5", label: "Star (5-point)" },
  { geometry: "rightArrow", label: "Right arrow" },
  { geometry: "leftArrow", label: "Left arrow" },
  { geometry: "upArrow", label: "Up arrow" },
  { geometry: "downArrow", label: "Down arrow" },
  { geometry: "wedgeRectCallout", label: "Rectangular callout" },
  { geometry: "wedgeRoundRectCallout", label: "Rounded callout" },
];

const LINE_GEOMETRIES = new Set(["line", "arrow"]);

// --- .docx fixture -----------------------------------------------------------
{
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
  const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
  const V = "urn:schemas-microsoft-com:vml";
  const O = "urn:schemas-microsoft-com:office:office";
  const W10 = "urn:schemas-microsoft-com:office:word";
  const DE = "urn:docu-edit:shape";

  // Every prefix an `mc:Choice/@Requires` below names, declared where Word
  // declares them — on `w:document`. MCE resolves Requires through the
  // document's own declarations, so a fixture that declared `wps` only on the
  // inner `wps:wsp` would leave the condition unevaluable and send every
  // reader to the VML fallback.
  const DOCUMENT_NAMESPACES =
    ` xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wp="${WP}" xmlns:a="${A}"` +
    ` xmlns:wps="${WPS}" xmlns:wpg="${WPG}" xmlns:wp14="${WP14}"` +
    ` xmlns:v="${V}" xmlns:o="${O}" xmlns:w10="${W10}" xmlns:de="${DE}"` +
    ' mc:Ignorable="wp14"';

  // Word writes NOTHING that uses a post-2006 namespace as a bare
  // `w:drawing`: a `wps:wsp` shape, a `wpg:wgp` group, a text box — each is
  // an `mc:Choice` gated on the namespace it needs, beside an `mc:Fallback`
  // holding the legacy VML that Word 2007 reads instead. Reproducing that
  // here is the whole point of this fixture: with bare drawings the suite
  // passed for years against markup Word does not produce, and no Word shape
  // ever reached the reader (see docx/reader.ts's runDrawingElement).
  //
  // The fallback is deliberately real markup, not a stub: if a regression
  // ever makes the reader take the Fallback branch instead of the Choice,
  // these shapes vanish and the e2e count assertion fails loudly.
  let vmlShapeId = 1026;
  function alternateContent(requires, drawing, { fill, outline, widthCm, heightCm }) {
    const spid = `_x0000_s${vmlShapeId++}`;
    const style = `position:absolute;width:${widthCm}cm;height:${heightCm}cm;z-index:251658240`;
    const fillAttr = fill ? ` fillcolor="#${fill.toLowerCase()}"` : ' filled="f"';
    const strokeAttr = outline
      ? ` strokecolor="#${outline.color.toLowerCase()}" strokeweight="${outline.widthPt}pt"`
      : ' stroked="f"';
    return (
      `<mc:AlternateContent><mc:Choice Requires="${requires}">${drawing}</mc:Choice>` +
      `<mc:Fallback><w:pict><v:rect id="${spid}" o:spid="${spid}" style="${style}"` +
      `${fillAttr}${strokeAttr}><w10:wrap type="square"/></v:rect></w:pict></mc:Fallback>` +
      "</mc:AlternateContent>"
    );
  }

  const GEOMETRY_TO_PRST = {
    line: "line",
    arrow: "straightConnector1",
    rect: "rect",
    roundRect: "roundRect",
    ellipse: "ellipse",
    triangle: "triangle",
    diamond: "diamond",
    star5: "star5",
    rightArrow: "rightArrow",
    leftArrow: "leftArrow",
    upArrow: "upArrow",
    downArrow: "downArrow",
    wedgeRectCallout: "wedgeRectCallout",
    wedgeRoundRectCallout: "wedgeRoundRectCallout",
  };

  const cmToEmu = (cm) => Math.round(cm * 360000);
  const ptToEmu = (pt) => Math.round(pt * 12700);

  function dashXml(dash) {
    if (dash === "dash") {
      return '<a:prstDash val="dash"/>';
    }
    if (dash === "dot") {
      return '<a:prstDash val="sysDot"/>';
    }
    return "";
  }

  function outlineXml(outline, extra = "") {
    if (!outline) {
      return "<a:ln><a:noFill/></a:ln>";
    }
    return (
      `<a:ln w="${ptToEmu(outline.widthPt)}">` +
      `<a:solidFill><a:srgbClr val="${outline.color}"/></a:solidFill>` +
      dashXml(outline.dash) +
      extra +
      "</a:ln>"
    );
  }

  function fillXml(fill) {
    return fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
  }

  function txbxXml(paragraphs, vAlign = "t") {
    const pXml = paragraphs
      .map((text) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`)
      .join("");
    return (
      `<wps:txbx><w:txbxContent>${pXml}</w:txbxContent></wps:txbx>` +
      `<wps:bodyPr wrap="square" anchor="${vAlign}"/>`
    );
  }

  // Builds one `<w:drawing>` for a shape. `opts.prst` overrides the geometry
  // lookup (used for the unsupported "gear9" entry, which is not in
  // GEOMETRY_TO_PRST).
  function shapeDrawing(opts) {
    const {
      id,
      geometry,
      prst = GEOMETRY_TO_PRST[geometry],
      widthCm,
      heightCm,
      fill,
      outline,
      rotationDeg,
      altText,
      groupOrdinal,
      wrap = "inline",
      offsetXCm = 0,
      offsetYCm = 0,
      zOrder,
      txbx,
    } = opts;
    const cx = cmToEmu(widthCm);
    const cy = cmToEmu(heightCm);
    const isLine = LINE_GEOMETRIES.has(geometry);

    const descrAttr = altText ? ` descr="${escapeXml(altText)}"` : "";
    const groupAttrs = groupOrdinal !== undefined ? ` de:group="${groupOrdinal}"` : "";
    const docPr = `<wp:docPr id="${id}" name="Shape ${id}"${descrAttr}${groupAttrs}/>`;

    const rotAttr = rotationDeg ? ` rot="${Math.round(rotationDeg * 60000)}"` : "";
    const tailEnd = geometry === "arrow" ? '<a:tailEnd type="triangle" w="med" len="med"/>' : "";
    const spPr =
      "<wps:spPr>" +
      `<a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
      (isLine ? "" : fillXml(fill)) +
      outlineXml(outline, tailEnd) +
      "</wps:spPr>";

    const wsp = `<wps:wsp><wps:cNvSpPr/>${spPr}${txbx ?? ""}</wps:wsp>`;
    const graphic = `<a:graphic><a:graphicData uri="${WPS}">${wsp}</a:graphicData></a:graphic>`;

    const drawing =
      wrap === "inline"
        ? "<w:drawing>" +
          '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
          `<wp:extent cx="${cx}" cy="${cy}"/>` +
          '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
          docPr +
          "<wp:cNvGraphicFramePr/>" +
          graphic +
          "</wp:inline></w:drawing>"
        : "<w:drawing>" +
          '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"' +
          ` relativeHeight="${251658240 + (zOrder ?? 0)}" behindDoc="0" locked="0"` +
          ' layoutInCell="1" allowOverlap="1">' +
          '<wp:simplePos x="0" y="0"/>' +
          `<wp:positionH relativeFrom="column"><wp:posOffset>${cmToEmu(offsetXCm)}</wp:posOffset></wp:positionH>` +
          `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${cmToEmu(offsetYCm)}</wp:posOffset></wp:positionV>` +
          `<wp:extent cx="${cx}" cy="${cy}"/>` +
          '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
          '<wp:wrapSquare wrapText="bothSides"/>' +
          docPr +
          "<wp:cNvGraphicFramePr/>" +
          graphic +
          "</wp:anchor></w:drawing>";
    // `wps` is the 2010 shape namespace, so Word gates every one of these —
    // inline and floating alike — on it.
    return alternateContent("wps", drawing, { fill, outline, widthCm, heightCm });
  }

  let nextId = 1;
  const drawings = [];

  for (const entry of CATALOGUE) {
    const isLine = LINE_GEOMETRIES.has(entry.geometry);
    drawings.push(
      shapeDrawing({
        id: nextId++,
        geometry: entry.geometry,
        widthCm: isLine ? 5 : 4,
        heightCm: isLine ? 0 : 3,
        fill: isLine ? undefined : "3366CC",
        outline: { color: "000000", widthPt: 1.5, dash: "solid" },
        altText: `${entry.label} shape`,
      }),
    );
  }

  // Grouped pair — both floating (square wrap) and sharing de:group="1", per
  // shape.ts's groupOrdinal round-trip (an ordinal, not an id).
  drawings.push(
    shapeDrawing({
      id: nextId++,
      geometry: "rect",
      widthCm: 2,
      heightCm: 2,
      fill: "1155CC",
      outline: { color: "000000", widthPt: 1, dash: "solid" },
      altText: "Group shape A",
      groupOrdinal: 1,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      zOrder: 0,
    }),
  );
  drawings.push(
    shapeDrawing({
      id: nextId++,
      geometry: "roundRect",
      widthCm: 2,
      heightCm: 2,
      fill: "22AA55",
      outline: { color: "000000", widthPt: 1, dash: "solid" },
      altText: "Group shape B",
      groupOrdinal: 1,
      wrap: "square",
      offsetXCm: 3,
      offsetYCm: 0,
      zOrder: 1,
    }),
  );

  // Text box: geometry "rect" with no fill, no outline, and a wps:txbx.
  drawings.push(
    shapeDrawing({
      id: nextId++,
      geometry: "rect",
      widthCm: 6,
      heightCm: 3,
      fill: undefined,
      outline: undefined,
      altText: "Text box",
      txbx: txbxXml(["First paragraph in the text box.", "Second paragraph in the text box."], "t"),
    }),
  );

  // Unsupported geometry: a prst outside GEOMETRY_TO_PRST — the reader falls
  // back to geometry "rect" with unsupportedGeometryXml set (AC-60). This
  // entry has to stay outside SPEC-069's catalogue to mean anything: it was
  // "pentagon" until SPEC-069 made that a real preset, and the committed
  // fixture bytes moved to "gear9" while this script did not — a drift that
  // would have broken AC-60 the next time anyone regenerated. Nine
  // parametric teeth is not an outline the model could plausibly gain.
  drawings.push(
    shapeDrawing({
      id: nextId++,
      geometry: "rect",
      prst: "gear9",
      widthCm: 4,
      heightCm: 3,
      fill: "CC6600",
      outline: { color: "000000", widthPt: 1, dash: "solid" },
      altText: "Unsupported shape",
    }),
  );

  // A REAL Word group (`wpg:wgp`), as distinct from the flat `de:group` pair
  // at 14/15 above: Word's own Group command writes a container whose members
  // carry no `wp:anchor` of their own, positioned in the group's `chOff`/
  // `chExt` child space. docx/group.ts flattens it into one block per member
  // sharing a minted ordinal, so this one drawing yields TWO entries (18, 19).
  //
  // The child space is deliberately 2x the group box on both axes, so a
  // reader that ignores `chExt` places and sizes both members wrongly rather
  // than coincidentally right.
  function groupMemberXml({ id, prst, fill, x, y, cx, cy, descr }) {
    return (
      `<wps:wsp><wps:cNvPr id="${id}" name="Shape ${id}" descr="${escapeXml(descr)}"/><wps:cNvSpPr/>` +
      `<wps:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
      fillXml(fill) +
      outlineXml({ color: "000000", widthPt: 1, dash: "solid" }) +
      "</wps:spPr><wps:bodyPr/></wps:wsp>"
    );
  }

  {
    const groupId = nextId++;
    const boxCx = cmToEmu(8);
    const boxCy = cmToEmu(3);
    const members =
      groupMemberXml({
        id: nextId++,
        prst: "rect",
        fill: "1155CC",
        x: 0,
        y: 0,
        cx: cmToEmu(4),
        cy: cmToEmu(4),
        descr: "Word group member A",
      }) +
      groupMemberXml({
        id: nextId++,
        prst: "ellipse",
        fill: "109618",
        x: cmToEmu(8),
        y: cmToEmu(2),
        cx: cmToEmu(4),
        cy: cmToEmu(4),
        descr: "Word group member B",
      });
    const groupDrawing =
      "<w:drawing>" +
      '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"' +
      ' relativeHeight="251658242" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${boxCx}" cy="${boxCy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
      `<wp:docPr id="${groupId}" name="Group ${groupId}" descr="Word group"/>` +
      "<wp:cNvGraphicFramePr/>" +
      `<a:graphic><a:graphicData uri="${WPG}">` +
      "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>" +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${boxCx}" cy="${boxCy}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${boxCx * 2}" cy="${boxCy * 2}"/></a:xfrm>` +
      "</wpg:grpSpPr>" +
      members +
      "</wpg:wgp></a:graphicData></a:graphic>" +
      "</wp:anchor></w:drawing>";
    drawings.push(
      alternateContent("wpg", groupDrawing, {
        fill: "1155CC",
        outline: { color: "000000", widthPt: 1, dash: "solid" },
        widthCm: 8,
        heightCm: 3,
      }),
    );
  }

  const documentXml =
    `${XML}<w:document${DOCUMENT_NAMESPACES}><w:body>` +
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
    "<dc:title>Shapes fixture</dc:title>" +
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
    writeFileSync(path.join(outDir, "word-shapes.docx"), bytes);
    console.log("Wrote", path.join(outDir, "word-shapes.docx"));
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

  const GEOMETRY_TO_DRAW_TYPE = {
    rect: "rectangle",
    roundRect: "round-rectangle",
    ellipse: "ellipse",
    triangle: "isosceles-triangle",
    diamond: "diamond",
    star5: "star5",
    rightArrow: "right-arrow",
    leftArrow: "left-arrow",
    upArrow: "up-arrow",
    downArrow: "down-arrow",
    wedgeRectCallout: "rectangular-callout",
    wedgeRoundRectCallout: "round-rectangular-callout",
  };

  const cm = (value) => `${value}cm`;

  function graphicStyleXml(name, { fill, outline, wrap }) {
    const fillAttrs = fill ? ` draw:fill="solid" draw:fill-color="${fill}"` : ' draw:fill="none"';
    let strokeAttrs = ' draw:stroke="none"';
    if (outline) {
      const widthCm = (outline.widthPt / 72) * 2.54;
      strokeAttrs =
        ` draw:stroke="${outline.dash === "solid" ? "solid" : "dash"}"` +
        ` svg:stroke-color="${outline.color}"` +
        ` svg:stroke-width="${cm(widthCm.toFixed(4))}"`;
    }
    let wrapAttrs = "";
    if (wrap === "square") {
      wrapAttrs = ' style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"';
    }
    return (
      `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="Graphics">` +
      `<style:graphic-properties${fillAttrs}${strokeAttrs}${wrapAttrs}` +
      ' style:vertical-pos="from-top" style:vertical-rel="paragraph"' +
      ' style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>' +
      "</style:style>"
    );
  }

  function textBoxXml(paragraphs, vAlign = "top") {
    const pXml = paragraphs.map((text) => `<text:p>${escapeXml(text)}</text:p>`).join("");
    return `<draw:text-box draw:text-vertical-align="${vAlign}">${pXml}</draw:text-box>`;
  }

  // Builds one shape element (`draw:custom-shape` for area geometries,
  // `draw:line` for Line/Arrow) plus its automatic graphic style.
  function shapeElement(opts) {
    const {
      styleName,
      geometry,
      drawType = GEOMETRY_TO_DRAW_TYPE[geometry],
      widthCm,
      heightCm,
      fill,
      outline,
      altText,
      groupOrdinal,
      rotationDeg,
      wrap = "inline",
      offsetXCm = 0,
      offsetYCm = 0,
      textParagraphs,
    } = opts;
    const isLine = LINE_GEOMETRIES.has(geometry);
    const styleXml = graphicStyleXml(styleName, { fill, outline, wrap });

    const anchorType = wrap === "inline" ? "as-char" : "paragraph";
    const descXml = altText ? `<svg:desc>${escapeXml(altText)}</svg:desc>` : "";
    const deAttrs =
      groupOrdinal !== undefined || rotationDeg
        ? ` xmlns:de="urn:docu-edit:shape"${groupOrdinal !== undefined ? ` de:group="${groupOrdinal}"` : ""}${
            rotationDeg ? ` de:rotation="${rotationDeg}"` : ""
          }`
        : "";

    if (isLine) {
      const x1 = offsetXCm;
      const y1 = offsetYCm;
      const x2 = offsetXCm + widthCm;
      const y2 = offsetYCm + heightCm;
      const markerAttr = geometry === "arrow" ? ' draw:marker-end="Arrow" draw:marker-end-width="0.3cm"' : "";
      const shapeXml =
        `<draw:line text:anchor-type="${anchorType}" draw:style-name="${styleName}"${deAttrs}` +
        ` svg:x1="${cm(x1)}" svg:y1="${cm(y1)}" svg:x2="${cm(x2)}" svg:y2="${cm(y2)}"${markerAttr}>` +
        descXml +
        "</draw:line>";
      return { styleXml, shapeXml };
    }

    const posAttrs = wrap !== "inline" ? ` svg:x="${cm(offsetXCm)}" svg:y="${cm(offsetYCm)}"` : "";
    const geometryXml = `<draw:enhanced-geometry draw:type="${drawType}"/>`;
    const textXml = textParagraphs ? textBoxXml(textParagraphs) : "";
    const shapeXml =
      `<draw:custom-shape text:anchor-type="${anchorType}" draw:style-name="${styleName}"${deAttrs}${posAttrs}` +
      ` svg:width="${cm(widthCm)}" svg:height="${cm(heightCm)}">` +
      descXml +
      textXml +
      geometryXml +
      "</draw:custom-shape>";
    return { styleXml, shapeXml };
  }

  let styleCounter = 1;
  const styles = [];
  const paragraphs = [];

  function addShape(opts) {
    const styleName = `gr${styleCounter++}`;
    const { styleXml, shapeXml } = shapeElement({ ...opts, styleName });
    styles.push(styleXml);
    paragraphs.push(`<text:p>${shapeXml}</text:p>`);
  }

  for (const entry of CATALOGUE) {
    const isLine = LINE_GEOMETRIES.has(entry.geometry);
    addShape({
      geometry: entry.geometry,
      widthCm: isLine ? 5 : 4,
      heightCm: isLine ? 0 : 3,
      fill: isLine ? undefined : "#3366cc",
      outline: { color: "#000000", widthPt: 1.5, dash: "solid" },
      altText: `${entry.label} shape`,
    });
  }

  // Grouped pair — both floating (wrap "square") and sharing de:group="1".
  addShape({
    geometry: "rect",
    widthCm: 2,
    heightCm: 2,
    fill: "#1155cc",
    outline: { color: "#000000", widthPt: 1, dash: "solid" },
    altText: "Group shape A",
    groupOrdinal: 1,
    wrap: "square",
    offsetXCm: 0,
    offsetYCm: 0,
  });
  addShape({
    geometry: "roundRect",
    widthCm: 2,
    heightCm: 2,
    fill: "#22aa55",
    outline: { color: "#000000", widthPt: 1, dash: "solid" },
    altText: "Group shape B",
    groupOrdinal: 1,
    wrap: "square",
    offsetXCm: 3,
    offsetYCm: 0,
  });

  // Text box: geometry "rect" with no fill, no outline, and a draw:text-box.
  addShape({
    geometry: "rect",
    widthCm: 6,
    heightCm: 3,
    fill: undefined,
    outline: undefined,
    altText: "Text box",
    textParagraphs: ["First paragraph in the text box.", "Second paragraph in the text box."],
  });

  // Unsupported geometry: a draw:type outside GEOMETRY_TO_DRAW_TYPE — the
  // reader falls back to geometry "rect" with unsupportedGeometryXml set.
  // "gear" for the same reason the .docx side uses "gear9" (see there):
  // ODF's "pentagon" is now the catalogue's homePlate.
  addShape({
    geometry: "rect",
    drawType: "gear",
    widthCm: 4,
    heightCm: 3,
    fill: "#cc6600",
    outline: { color: "#000000", widthPt: 1, dash: "solid" },
    altText: "Unsupported shape",
  });

  // A REAL LibreOffice group (`draw:g`), the .odt twin of the `wpg:wgp` the
  // .docx side carries: a container whose members have no anchor of their
  // own, positioned by their own `svg:x`/`svg:y` in the group's frame (ODF
  // has no `chOff`/`chExt` equivalent). odf/group.ts flattens it into one
  // block per member sharing a minted ordinal, so this one element yields
  // TWO entries (18, 19) — the same indices the .docx group lands on.
  {
    const memberStyles = [];
    const memberXml = [
      {
        geometry: "rect",
        fill: "#1155cc",
        offsetXCm: 0,
        offsetYCm: 0,
        altText: "ODF group member A",
      },
      {
        geometry: "ellipse",
        fill: "#109618",
        offsetXCm: 4,
        offsetYCm: 1,
        altText: "ODF group member B",
      },
    ].map((opts) => {
      const styleName = `gr${styleCounter++}`;
      const { styleXml, shapeXml } = shapeElement({
        ...opts,
        styleName,
        widthCm: 2,
        heightCm: 2,
        outline: { color: "#000000", widthPt: 1, dash: "solid" },
        // "square" only so the member carries svg:x/svg:y; the group owns the
        // real anchoring, and odf/group.ts overwrites the wrap from it.
        wrap: "square",
      });
      memberStyles.push(styleXml);
      return shapeXml;
    });
    styles.push(...memberStyles);
    paragraphs.push(
      '<text:p><draw:g draw:name="Group 1" text:anchor-type="paragraph" draw:z-index="10">' +
        '<svg:desc>ODF group</svg:desc>' +
        memberXml.join("") +
        "</draw:g></text:p>",
    );
  }

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
    "<office:meta><dc:title>Shapes fixture</dc:title><dc:date>2024-01-01T00:00:00Z</dc:date></office:meta>" +
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
    writeFileSync(path.join(outDir, "lo-shapes.odt"), bytes);
    console.log("Wrote", path.join(outDir, "lo-shapes.odt"));
  }
}
