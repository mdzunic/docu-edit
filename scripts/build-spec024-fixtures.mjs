#!/usr/bin/env node
// One-off generator for the SPEC-024 image fixtures (run manually:
// `node scripts/build-spec024-fixtures.mjs`). Like table.docx/lo-embedded.odt,
// these are hand-built packages reproducing what Word/LibreOffice actually
// emit for pictures — not round-tripped through this project's own writer —
// so opening them is a genuine test of the reader, not just "the reader
// understands its own writer's output": a `wp:inline` picture and a
// `wp:anchor`/`wrapSquare` floating picture for .docx (AC-28), a
// `text:anchor-type="as-char"` and a `text:anchor-type="paragraph"` +
// `style:wrap="parallel"` `draw:frame` for .odt (AC-29). Each fixture embeds
// two distinct small PNGs (solid-colour, built here) so size/position are
// independently checkable per image.
import { deflateRawSync, deflateSync } from "node:zlib";
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
function u32be(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32BE(value >>> 0, 0);
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

// --- minimal PNG encoder (24-bit RGB, no filtering) -------------------------

function pngChunk(type, data) {
  const typeBuf = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBuf, data]);
  return Buffer.concat([u32be(data.length), body, u32be(crc32(body))]);
}

function buildPng(width, height, [r, g, b]) {
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (rowBytes + 1);
    raw[rowStart] = 0; // filter type: none
    for (let x = 0; x < width; x++) {
      const px = rowStart + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type: RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

// A landscape photo-ish image (300x200 at 96 DPI = 7.9cm x 5.29cm) for the
// inline picture, and a small square logo-ish image (120x120 = 3.18cm) for
// the floating one — different enough that AC-28/29's "renders at the size
// the source app shows" is actually checking something.
const inlinePng = buildPng(300, 200, [0x2a, 0x5c, 0x8a]);
const floatingPng = buildPng(120, 120, [0xc0, 0x39, 0x2b]);

// --- minimal ZIP writer (shared shape with build-spec022-fixtures.mjs) -----

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

// --- .docx fixture -----------------------------------------------------------
// EMU: 914400 per inch, 96px/inch source images -> cx/cy = px * 9525.
{
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const inlineCx = 300 * 9525;
  const inlineCy = 200 * 9525;
  const floatCx = 120 * 9525;
  const floatCy = 120 * 9525;

  function pictureGraphic(id, name, relId, cx, cy) {
    return (
      `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}">` +
      `<pic:pic xmlns:pic="${PIC}">` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      "</pic:pic></a:graphicData></a:graphic>"
    );
  }

  // A real Word inline picture run, as Word itself emits it.
  const inlineDrawing =
    `<w:drawing xmlns:wp="${WP}">` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${inlineCx}" cy="${inlineCy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:docPr id="1" name="Picture 1" descr="A blue landscape photo"/>' +
    "<wp:cNvGraphicFramePr/>" +
    pictureGraphic(1, "Picture 1", "rId4", inlineCx, inlineCy) +
    "</wp:inline></w:drawing>";

  // A real Word floating (anchored) picture: square wrap, in front of text.
  const anchorDrawing =
    `<w:drawing xmlns:wp="${WP}">` +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"' +
    ' relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="${floatCx}" cy="${floatCy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:wrapSquare wrapText="bothSides"/>' +
    '<wp:docPr id="2" name="Picture 2" descr="A red square logo"/>' +
    "<wp:cNvGraphicFramePr/>" +
    pictureGraphic(2, "Picture 2", "rId5", floatCx, floatCy) +
    "</wp:anchor></w:drawing>";

  const documentXml =
    `${XML}<w:document xmlns:w="${W}"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">Inline picture: </w:t></w:r><w:r>${inlineDrawing}</w:r><w:r><w:t xml:space="preserve"> end of paragraph.</w:t></w:r></w:p>` +
    `<w:p><w:r>${anchorDrawing}</w:r><w:r><w:t>Text beside the floating picture.</w:t></w:r></w:p>` +
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
    '<Default Extension="png" ContentType="image/png"/>' +
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
    `<Relationship Id="rId4" Type="${R}/image" Target="media/image1.png"/>` +
    `<Relationship Id="rId5" Type="${R}/image" Target="media/image2.png"/>` +
    "</Relationships>";

  const corePropsXml =
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:title>Images fixture</dc:title>" +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>' +
    "</cp:coreProperties>";

  const bytes = buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
    { name: "word/media/image1.png", content: inlinePng },
    { name: "word/media/image2.png", content: floatingPng },
    { name: "docProps/core.xml", content: corePropsXml },
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "docx", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "images.docx"), bytes);
    console.log("Wrote", path.join(outDir, "images.docx"));
  }
}

// --- .odt fixture --------------------------------------------------------
{
  const inlineWidthCm = ((300 / 96) * 2.54).toFixed(4);
  const inlineHeightCm = ((200 / 96) * 2.54).toFixed(4);
  const floatWidthCm = ((120 / 96) * 2.54).toFixed(4);
  const floatHeightCm = ((120 / 96) * 2.54).toFixed(4);

  const NS =
    ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
    ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
    ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
    ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
    ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
    ' xmlns:xlink="http://www.w3.org/1999/xlink"';

  // A real LibreOffice inline (as-char anchored) picture frame.
  const inlineFrame =
    `<draw:frame draw:name="Image1" text:anchor-type="as-char"` +
    ` svg:width="${inlineWidthCm}cm" svg:height="${inlineHeightCm}cm">` +
    '<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
    "<svg:desc>A blue landscape photo</svg:desc>" +
    "</draw:frame>";

  // A real LibreOffice floating (paragraph-anchored) picture frame, wrap
  // "parallel" (square) via its graphic-family automatic style "fr1".
  const floatFrame =
    `<draw:frame draw:name="Image2" draw:style-name="fr1" text:anchor-type="paragraph"` +
    ` svg:width="${floatWidthCm}cm" svg:height="${floatHeightCm}cm">` +
    '<draw:image xlink:href="Pictures/image2.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
    "<svg:desc>A red square logo</svg:desc>" +
    "</draw:frame>";

  const contentXml =
    `${XML}<office:document-content${NS} office:version="1.2">` +
    "<office:automatic-styles>" +
    '<style:style style:name="fr1" style:family="graphic">' +
    '<style:graphic-properties style:wrap="parallel" style:vertical-pos="from-top" style:vertical-rel="paragraph"' +
    ' style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>' +
    "</style:style>" +
    "</office:automatic-styles>" +
    "<office:body><office:text>" +
    `<text:p>Inline picture: ${inlineFrame} end of paragraph.</text:p>` +
    `<text:p>${floatFrame}Text beside the floating picture.</text:p>` +
    "</office:text></office:body></office:document-content>";

  const stylesXml =
    `${XML}<office:document-styles${NS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph">' +
    '<style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt"/>' +
    "</style:default-style>" +
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
    "<office:meta><dc:title>Images fixture</dc:title><dc:date>2024-01-01T00:00:00Z</dc:date></office:meta>" +
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
    '<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>' +
    '<manifest:file-entry manifest:full-path="Pictures/image2.png" manifest:media-type="image/png"/>' +
    "</manifest:manifest>";

  const bytes = buildZip([
    { name: "mimetype", content: "application/vnd.oasis.opendocument.text", store: true },
    { name: "content.xml", content: contentXml },
    { name: "styles.xml", content: stylesXml },
    { name: "meta.xml", content: metaXml },
    { name: "settings.xml", content: settingsXml },
    { name: "Pictures/image1.png", content: inlinePng },
    { name: "Pictures/image2.png", content: floatingPng },
    { name: "META-INF/manifest.xml", content: manifestXml },
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "odf", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "images.odt"), bytes);
    console.log("Wrote", path.join(outDir, "images.odt"));
  }
}
