#!/usr/bin/env node
// One-off generator for the SPEC-057 native drawing-group fixtures (run
// manually: `node scripts/build-spec057-fixtures.mjs`). Sibling of
// build-spec051-fixtures.mjs — hand-built packages reproducing what
// Word/LibreOffice emit for native groups (`wpg:wgp` / `draw:g`), NOT
// round-tripped through this project's own writer, so opening them genuinely
// tests the group readers (src/lib/docx/groupReader.ts, src/lib/odf/groupReader.ts).
//
// Adds word-groups.docx (native wpg:wgp; the plain-pair group uses a
// non-identity child coordinate space) and lo-groups.odt (native draw:g), each
// holding, in a fixed document order so e2e/SPEC-057.spec.ts can address each
// group by its 0-based drawing-<n> index:
//   0  plain two-shape group
//   1  nested group (a shape + a nested pair)
//   2  mixed group: a shape, an image and a text box (two paragraphs)
//   3  rotated group (container @rot in .docx; the baked form in .odt)
//   4  preserved-member group (a member outside the shape catalogue → sealed)
//   5  chart-member group (a shape + a chart; AC-81) — .docx wpg:graphicFrame
//      referencing word/charts/chart1.xml, .odt draw:object referencing an
//      embedded chart subpackage (Object 1). The chart XML below is exactly
//      what this project's own chart writers emit, so the readers parse it.
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
      u32le(0x04034b50), u16le(version), u16le(0), u16le(method),
      u16le(0), u16le(0), u32le(crc), u32le(dataBytes.length),
      u32le(rawBytes.length), u16le(nameBytes.length), u16le(0),
    ]);
    localParts.push(localHeader, nameBytes, dataBytes);
    const central = Buffer.concat([
      u32le(0x02014b50), u16le(version), u16le(version), u16le(0), u16le(method),
      u16le(0), u16le(0), u32le(crc), u32le(dataBytes.length), u32le(rawBytes.length),
      u16le(nameBytes.length), u16le(0), u16le(0), u16le(0), u16le(0), u32le(0), u32le(offset),
    ]);
    centralParts.push(central, nameBytes);
    offset += localHeader.length + nameBytes.length + dataBytes.length;
  }
  const centralBytes = Buffer.concat(centralParts);
  const centralStart = offset;
  const end = Buffer.concat([
    u32le(0x06054b50), u16le(0), u16le(0), u16le(entries.length), u16le(entries.length),
    u32le(centralBytes.length), u32le(centralStart), u16le(0),
  ]);
  return Buffer.concat([...localParts, centralBytes, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function escapeXml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// A real 1×1 transparent PNG so both decodeImageDimensions (reads IHDR) and a
// browser <img>/canvas render it in the e2e.
const PNG_1x1 = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQAY3Y2wAAAAAElFTkSuQmCC",
  "base64",
);

// The chart-member group's chart content — verbatim what this project's own
// chart writers emit for a columnClustered chart (categories A/B, one series),
// so the group readers parse each back to a first-class Chart member.
const DOCX_CHART_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart><c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx><c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111111111"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

const ODT_CHART_XML =
  '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.2"><office:body><office:chart><chart:chart chart:class="chart:bar" chart:vertical="true" chart:stacked="false" svg:width="5cm" svg:height="4cm"><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y"><chart:grid chart:class="major"/></chart:axis><chart:series chart:class="chart:bar" chart:color="#4472c4" chart:values-cell-range-address="local-table.$B$2:.$B$3" chart:label-cell-address="local-table.$B$1"/></chart:plot-area><table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column/></table:table-columns><table:table-header-rows><table:table-row><table:table-cell/><table:table-cell office:value-type="string"><text:p>S</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-rows><table:table-row><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="string"><text:p>B</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell></table:table-row></table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>';

// --- .docx fixture -----------------------------------------------------------
{
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
  const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const emu = (cm) => Math.round(cm * 360000);
  let nextId = 1;

  // A wps:wsp member at child-space off/ext (EMU), optional prst override
  // (for the unsupported geometry), rotation and text box.
  function wspMember({ prst = "rect", x, y, w, h, fill = "3366CC", rot, altText, txbx }) {
    const rotAttr = rot ? ` rot="${Math.round(rot * 60000)}"` : "";
    const descrAttr = altText ? ` descr="${escapeXml(altText)}"` : "";
    const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
    return (
      `<wps:wsp xmlns:wps="${WPS}" xmlns:a="${A}">` +
      `<wps:cNvPr id="${nextId++}" name="Shape"${descrAttr}/><wps:cNvSpPr/>` +
      "<wps:spPr>" +
      `<a:xfrm${rotAttr}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
      fillXml +
      "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>" +
      "</wps:spPr>" +
      (txbx
        ? `<wps:txbx><w:txbxContent xmlns:w="${W}">${txbx.map((t) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(t)}</w:t></w:r></w:p>`).join("")}</w:txbxContent></wps:txbx><wps:bodyPr/>`
        : "") +
      "</wps:wsp>"
    );
  }
  // A pic:pic member referencing rId1.
  function picMember({ x, y, w, h, altText }) {
    const descrAttr = altText ? ` descr="${escapeXml(altText)}"` : "";
    return (
      `<pic:pic xmlns:pic="${PIC}" xmlns:a="${A}">` +
      `<pic:nvPicPr><pic:cNvPr id="${nextId++}" name="Picture"${descrAttr}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      "</pic:pic>"
    );
  }
  // A wpg:graphicFrame chart member referencing the chart part via `relId`.
  function chartMember({ x, y, w, h, relId, altText }) {
    const descrAttr = altText ? ` descr="${escapeXml(altText)}"` : "";
    return (
      `<wpg:graphicFrame xmlns:wpg="${WPG}" xmlns:a="${A}" xmlns:r="${R}">` +
      `<wpg:cNvPr id="${nextId++}" name="Chart"${descrAttr}/><wpg:cNvFrPr/>` +
      `<wpg:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></wpg:xfrm>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/>` +
      "</a:graphicData></a:graphic></wpg:graphicFrame>"
    );
  }
  // A nested wpg:grpSp at parent-space off/ext with an identity child space.
  function nestedGroup({ x, y, w, h, members }) {
    return (
      `<wpg:grpSp xmlns:wpg="${WPG}" xmlns:a="${A}">` +
      `<wpg:cNvPr id="${nextId++}" name="Group"/><wpg:cNvGrpSpPr/>` +
      "<wpg:grpSpPr>" +
      `<a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
      "</wpg:grpSpPr>" +
      members +
      "</wpg:grpSp>"
    );
  }
  // A floating group anchor holding wpg:wgp. `chScale` > 1 makes a non-identity
  // child coordinate space (chExt = chScale × ext); members are then authored
  // in that scaled space.
  function groupAnchor({ w, h, rot, chScale = 1, altText, members, posX = 3, posY = 0 }) {
    const cx = emu(w);
    const cy = emu(h);
    const rotAttr = rot ? ` rot="${Math.round(rot * 60000)}"` : "";
    const descrAttr = altText ? ` descr="${escapeXml(altText)}"` : "";
    const chExtCx = Math.round(cx * chScale);
    const chExtCy = Math.round(cy * chScale);
    return (
      `<w:drawing xmlns:wp="${WP}">` +
      '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(posX)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(posY)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:wrapSquare wrapText="bothSides"/>' +
      `<wp:docPr id="${nextId++}" name="Group"${descrAttr}/>` +
      "<wp:cNvGraphicFramePr/>" +
      `<a:graphic xmlns:a="${A}">` +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      `<wpg:wgp xmlns:wpg="${WPG}">` +
      "<wpg:cNvGrpSpPr/>" +
      `<wpg:grpSpPr><a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${chExtCx}" cy="${chExtCy}"/></a:xfrm></wpg:grpSpPr>` +
      members +
      "</wpg:wgp>" +
      "</a:graphicData></a:graphic></wp:anchor></w:drawing>"
    );
  }

  const groups = [];
  // 0 — plain pair, non-identity child space (chScale 2 → child coords doubled).
  groups.push(
    groupAnchor({
      w: 5, h: 2, chScale: 2, altText: "Plain pair", posY: 0,
      members: wspMember({ x: 0, y: 0, w: 4, h: 4, fill: "1155CC", altText: "A" }) + wspMember({ prst: "roundRect", x: 6, y: 0, w: 4, h: 4, fill: "22AA55", altText: "B" }),
    }),
  );
  // 1 — nested: a shape + a nested pair (identity space).
  groups.push(
    groupAnchor({
      w: 7, h: 2, altText: "Nested", posY: 4,
      members:
        wspMember({ prst: "triangle", x: 0, y: 0, w: 2, h: 2, fill: "CC3366", altText: "Solo" }) +
        nestedGroup({ x: 3, y: 0, w: 4, h: 2, members: wspMember({ x: 0, y: 0, w: 2, h: 2, altText: "Nested A" }) + wspMember({ prst: "ellipse", x: 2, y: 0, w: 2, h: 2, fill: "AA5522", altText: "Nested B" }) }),
    }),
  );
  // 2 — mixed: a shape, an image and a text box (identity space).
  groups.push(
    groupAnchor({
      w: 5, h: 5, altText: "Mixed", posY: 8,
      members:
        wspMember({ x: 0, y: 0, w: 2, h: 2, fill: "3366CC", altText: "Shape member" }) +
        picMember({ x: 3, y: 0, w: 2, h: 2, altText: "Image member" }) +
        wspMember({ x: 0, y: 3, w: 4, h: 2, fill: "", altText: "Text box member", txbx: ["First paragraph.", "Second paragraph."] }),
    }),
  );
  // 3 — rotated container (@rot 30°).
  groups.push(
    groupAnchor({
      w: 5, h: 2, rot: 30, altText: "Rotated", posY: 15,
      members: wspMember({ x: 0, y: 0, w: 2, h: 2, fill: "1155CC", altText: "R-A" }) + wspMember({ prst: "ellipse", x: 3, y: 0, w: 2, h: 2, fill: "22AA55", altText: "R-B" }),
    }),
  );
  // 4 — preserved member (prst outside the catalogue → sealed).
  groups.push(
    groupAnchor({
      w: 5, h: 2, altText: "Preserved", posY: 19,
      members: wspMember({ prst: "pentagon", x: 0, y: 0, w: 2, h: 2, fill: "CC6600", altText: "Preserved member" }) + wspMember({ x: 3, y: 0, w: 2, h: 2, fill: "3366CC", altText: "Normal member" }),
    }),
  );
  // 5 — chart member (a shape + a wpg:graphicFrame chart → first-class Chart).
  groups.push(
    groupAnchor({
      w: 8, h: 4, altText: "Chart group", posY: 23,
      members: wspMember({ x: 0, y: 0, w: 2, h: 2, fill: "3366CC", altText: "Chart sibling" }) + chartMember({ x: 3, y: 0, w: 5, h: 4, relId: "rIdChart1", altText: "Chart member" }),
    }),
  );

  const documentXml =
    `${XML}<w:document xmlns:w="${W}"><w:body>` +
    groups.map((g) => `<w:p><w:r>${g}</w:r></w:p>`).join("") +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";

  const stylesXml =
    `${XML}<w:styles xmlns:w="${W}">` +
    "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    "</w:styles>";

  const contentTypesXml =
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
    "</Types>";

  const packageRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";

  const documentRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>' +
    "</Relationships>";

  const bytes = buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
    { name: "word/charts/chart1.xml", content: DOCX_CHART_XML },
    { name: "word/media/image1.png", content: PNG_1x1, store: true },
  ]);
  for (const outDir of [path.join(__dirname, "..", "e2e", "fixtures"), path.join(__dirname, "..", "src", "lib", "docx", "fixtures")]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "word-groups.docx"), bytes);
    console.log("Wrote", path.join(outDir, "word-groups.docx"));
  }
}

// --- .odt fixture ------------------------------------------------------------
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
  const cmv = (v) => `${v}cm`;
  const styles = [];
  let styleCounter = 1;

  function shapeStyle({ fill, wrap }) {
    const name = `gr${styleCounter++}`;
    const fillAttrs = fill ? ` draw:fill="solid" draw:fill-color="${fill}"` : ' draw:fill="none"';
    const wrapAttrs = wrap === "square" ? ' style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"' : "";
    styles.push(
      `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="Graphics">` +
      `<style:graphic-properties${fillAttrs} draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0.05cm"${wrapAttrs}` +
      ' style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>' +
      "</style:style>",
    );
    return name;
  }
  const groupStyleName = (() => {
    const name = "grpwrap";
    styles.push(
      `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="Graphics">` +
      '<style:graphic-properties style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"' +
      ' style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>' +
      "</style:style>",
    );
    return name;
  })();

  // Absolute-positioned members (svg:x/svg:y in the anchor frame).
  function shapeMember({ drawType = "rectangle", x, y, w, h, fill = "#3366cc", altText, textParagraphs }) {
    const name = shapeStyle({ fill, wrap: "square" });
    const descXml = altText ? `<svg:desc>${escapeXml(altText)}</svg:desc>` : "";
    const textXml = textParagraphs ? `<draw:text-box draw:text-vertical-align="top">${textParagraphs.map((t) => `<text:p>${escapeXml(t)}</text:p>`).join("")}</draw:text-box>` : "";
    return (
      `<draw:custom-shape text:anchor-type="paragraph" draw:style-name="${name}" svg:x="${cmv(x)}" svg:y="${cmv(y)}" svg:width="${cmv(w)}" svg:height="${cmv(h)}">` +
      descXml + textXml + `<draw:enhanced-geometry draw:type="${drawType}"/>` +
      "</draw:custom-shape>"
    );
  }
  function imageMember({ x, y, w, h, altText }) {
    const descXml = altText ? `<svg:desc>${escapeXml(altText)}</svg:desc>` : "";
    return (
      `<draw:frame text:anchor-type="paragraph" svg:x="${cmv(x)}" svg:y="${cmv(y)}" svg:width="${cmv(w)}" svg:height="${cmv(h)}">` +
      '<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
      descXml +
      "</draw:frame>"
    );
  }
  // A chart member: a draw:frame/draw:object referencing the embedded chart
  // subpackage "Object 1".
  function chartMember({ x, y, w, h, altText }) {
    const descXml = altText ? `<svg:desc>${escapeXml(altText)}</svg:desc>` : "";
    return (
      `<draw:frame text:anchor-type="paragraph" svg:x="${cmv(x)}" svg:y="${cmv(y)}" svg:width="${cmv(w)}" svg:height="${cmv(h)}">` +
      '<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
      descXml +
      "</draw:frame>"
    );
  }
  function group({ altText, members }) {
    const titleXml = altText ? `<svg:title>${escapeXml(altText)}</svg:title><svg:desc>${escapeXml(altText)}</svg:desc>` : "";
    return `<text:p><draw:g text:anchor-type="paragraph" draw:style-name="${groupStyleName}">${titleXml}${members}</draw:g></text:p>`;
  }
  function nestedGroup({ members }) {
    return `<draw:g>${members}</draw:g>`;
  }

  const paragraphs = [];
  // 0 — plain pair.
  paragraphs.push(group({ altText: "Plain pair", members: shapeMember({ x: 2, y: 2, w: 2, h: 2, fill: "#1155cc", altText: "A" }) + shapeMember({ drawType: "round-rectangle", x: 5, y: 2, w: 2, h: 2, fill: "#22aa55", altText: "B" }) }));
  // 1 — nested.
  paragraphs.push(group({ altText: "Nested", members: shapeMember({ drawType: "isosceles-triangle", x: 2, y: 8, w: 2, h: 2, fill: "#cc3366", altText: "Solo" }) + nestedGroup({ members: shapeMember({ x: 5, y: 8, w: 2, h: 2, altText: "Nested A" }) + shapeMember({ drawType: "ellipse", x: 7, y: 8, w: 2, h: 2, fill: "#aa5522", altText: "Nested B" }) }) }));
  // 2 — mixed: shape, image, text box.
  paragraphs.push(group({ altText: "Mixed", members: shapeMember({ x: 2, y: 13, w: 2, h: 2, fill: "#3366cc", altText: "Shape member" }) + imageMember({ x: 5, y: 13, w: 2, h: 2, altText: "Image member" }) + shapeMember({ x: 2, y: 16, w: 4, h: 2, fill: undefined, altText: "Text box member", textParagraphs: ["First paragraph.", "Second paragraph."] }) }));
  // 3 — rotated (baked form: two plain members, no group rotation).
  paragraphs.push(group({ altText: "Rotated", members: shapeMember({ x: 10, y: 2, w: 2, h: 2, fill: "#1155cc", altText: "R-A" }) + shapeMember({ drawType: "ellipse", x: 13, y: 2, w: 2, h: 2, fill: "#22aa55", altText: "R-B" }) }));
  // 4 — preserved member (draw:type outside the catalogue → sealed).
  paragraphs.push(group({ altText: "Preserved", members: shapeMember({ drawType: "pentagon", x: 10, y: 8, w: 2, h: 2, fill: "#cc6600", altText: "Preserved member" }) + shapeMember({ x: 13, y: 8, w: 2, h: 2, fill: "#3366cc", altText: "Normal member" }) }));
  // 5 — chart member (a shape + a draw:object chart → first-class Chart).
  paragraphs.push(group({ altText: "Chart group", members: shapeMember({ x: 10, y: 13, w: 2, h: 2, fill: "#3366cc", altText: "Chart sibling" }) + chartMember({ x: 13, y: 13, w: 4, h: 3, altText: "Chart member" }) }));

  const contentXml =
    `${XML}<office:document-content${NS} office:version="1.2">` +
    `<office:automatic-styles>${styles.join("")}</office:automatic-styles>` +
    `<office:body><office:text>${paragraphs.join("")}</office:text></office:body>` +
    "</office:document-content>";

  const stylesXml =
    `${XML}<office:document-styles${NS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph"><style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt"/></style:default-style>' +
    '<style:style style:name="Graphics" style:family="graphic"/>' +
    "</office:styles>" +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout>' +
    "</office:automatic-styles>" +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>' +
    "</office:document-styles>";

  const metaXml =
    `${XML}<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    "<office:meta><dc:title>Groups fixture</dc:title><dc:date>2024-01-01T00:00:00Z</dc:date></office:meta></office:document-meta>";
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
    '<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>' +
    '<manifest:file-entry manifest:full-path="Object 1/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.chart"/>' +
    '<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>' +
    "</manifest:manifest>";

  const bytes = buildZip([
    { name: "mimetype", content: "application/vnd.oasis.opendocument.text", store: true },
    { name: "content.xml", content: contentXml },
    { name: "styles.xml", content: stylesXml },
    { name: "meta.xml", content: metaXml },
    { name: "settings.xml", content: settingsXml },
    { name: "Pictures/image1.png", content: PNG_1x1, store: true },
    { name: "Object 1/content.xml", content: ODT_CHART_XML },
    { name: "META-INF/manifest.xml", content: manifestXml },
  ]);
  for (const outDir of [path.join(__dirname, "..", "e2e", "fixtures"), path.join(__dirname, "..", "src", "lib", "odf", "fixtures")]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "lo-groups.odt"), bytes);
    console.log("Wrote", path.join(outDir, "lo-groups.odt"));
  }
}
