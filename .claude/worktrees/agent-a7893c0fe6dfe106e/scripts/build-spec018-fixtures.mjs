#!/usr/bin/env node
// One-off generator for the SPEC-018 list fixtures under e2e/fixtures/.
// Not part of the app build — run manually
// (`node scripts/build-spec018-fixtures.mjs`) if the fixtures ever need to
// be regenerated or extended.
//
// Like the other fixtures (see build-docx-fixtures.mjs), these are NOT
// files exported from real Word/LibreOffice installs — they are hand-built
// packages reproducing each producer's list markup:
//
// - lists.docx: Word-style numbering.xml — a bullet abstractNum using the
//   Symbol/Courier/Wingdings characters Word writes (F0B7, "o", F0A7), a
//   decimal/lowerLetter abstractNum, and one whose level 0 starts at 5;
//   level 0 of the bullet definition carries a non-default 1080-twip
//   indent. Paragraphs reference the nums via w:numPr.
// - lists.odt: LibreOffice-style content.xml — automatic text:list-style
//   definitions (bullet chars, num-format/start-value, space-before +
//   min-label-width indents), nested text:list structures for sub-levels,
//   and a list starting at 5. Level 1 of the bullet style carries the same
//   non-default 1.905 cm indent.
//
// Both encode the same document, so the AC-20/AC-21 renders can be asserted
// with the same expectations.
import { deflateRawSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(outDir, { recursive: true });

function u32(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32LE(value >>> 0, 0);
  return buf;
}
function u16(value) {
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

// entries: {name, content, store?} — store:true keeps the entry uncompressed
// (required for the mimetype entry, which must also come first).
function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.from(entry.content, "utf-8");
    const method = entry.store ? 0 : 8;
    const dataBytes = entry.store ? rawBytes : deflateRawSync(rawBytes);
    const crc = crc32(rawBytes);
    const version = entry.store ? 10 : 20;

    const localHeader = Buffer.concat([
      u32(0x04034b50),
      u16(version),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(crc),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
    ]);
    localParts.push(localHeader, dataBytes);

    const centralHeader = Buffer.concat([
      u32(0x02014b50),
      u16(20),
      u16(version),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(crc),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      nameBytes,
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  }

  const centralStart = offset;
  const centralBytes = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(centralBytes.length),
    u32(centralStart),
    u16(0),
  ]);
  return Buffer.concat([...localParts, centralBytes, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// --- lists.docx ------------------------------------------------------------

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Word's Symbol bullet (F0B7) and Wingdings square (F0A7); level 1 keeps
// Word's Courier New "o".
const SYMBOL_BULLET = "";
const WINGDINGS_SQUARE = "";

function lvl(ilvl, numFmt, lvlText, start, leftTwips) {
  return (
    `<w:lvl w:ilvl="${ilvl}">` +
    `<w:start w:val="${start}"/>` +
    `<w:numFmt w:val="${numFmt}"/>` +
    `<w:lvlText w:val="${lvlText}"/>` +
    '<w:lvlJc w:val="left"/>' +
    `<w:pPr><w:ind w:left="${leftTwips}" w:hanging="360"/></w:pPr>` +
    "</w:lvl>"
  );
}

const numberingXml =
  `${XML}<w:numbering xmlns:w="${W}">` +
  // Bullets: level 0 has a NON-default 1080-twip (1.905 cm) indent.
  `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
  lvl(0, "bullet", SYMBOL_BULLET, 1, 1080) +
  lvl(1, "bullet", "o", 1, 1440) +
  lvl(2, "bullet", WINGDINGS_SQUARE, 1, 2160) +
  "</w:abstractNum>" +
  // Numbers: decimal then lowerLetter.
  `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>` +
  lvl(0, "decimal", "%1.", 1, 720) +
  lvl(1, "lowerLetter", "%2.", 1, 1440) +
  "</w:abstractNum>" +
  // A decimal list starting at 5 (AC-22).
  `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>` +
  lvl(0, "decimal", "%1.", 5, 720) +
  "</w:abstractNum>" +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>' +
  "</w:numbering>";

function listP(text, numId, ilvl) {
  return (
    "<w:p><w:pPr>" +
    `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` +
    `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  );
}

const documentXml =
  `${XML}<w:document xmlns:w="${W}"><w:body>` +
  "<w:p><w:r><w:t>Groceries</w:t></w:r></w:p>" +
  listP("Apples", 1, 0) +
  listP("Bananas", 1, 0) +
  listP("Green ones", 1, 1) +
  listP("Wrapped ones", 1, 2) +
  "<w:p><w:r><w:t>Steps</w:t></w:r></w:p>" +
  listP("First", 2, 0) +
  listP("Second", 2, 0) +
  listP("Detail", 2, 1) +
  "<w:p><w:r><w:t>Continue from five</w:t></w:r></w:p>" +
  listP("Fifth", 3, 0) +
  listP("Sixth", 3, 0) +
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
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
  "</Types>";

const packageRelsXml =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships>";

const documentRelsXml =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
  "</Relationships>";

writeFileSync(
  path.join(outDir, "lists.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
    { name: "word/numbering.xml", content: numberingXml },
  ]),
);

// --- lists.odt -------------------------------------------------------------

const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";
const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

function bulletLevel(level, char, spaceBefore) {
  return (
    `<text:list-level-style-bullet text:level="${level}" text:bullet-char="${char}">` +
    `<style:list-level-properties text:space-before="${spaceBefore}cm" text:min-label-width="0.635cm"/>` +
    "</text:list-level-style-bullet>"
  );
}

function numberLevel(level, format, start, spaceBefore) {
  const startAttr = start !== 1 ? ` text:start-value="${start}"` : "";
  return (
    `<text:list-level-style-number text:level="${level}" style:num-format="${format}" style:num-suffix="."${startAttr}>` +
    `<style:list-level-properties text:space-before="${spaceBefore}cm" text:min-label-width="0.635cm"/>` +
    "</text:list-level-style-number>"
  );
}

const odtContentXml =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  `<office:document-content xmlns:office="${OFFICE}" xmlns:text="${TEXT}" xmlns:style="${STYLE}" xmlns:fo="${FO}" office:version="1.2">` +
  "<office:automatic-styles>" +
  // Bullets: level 1 has a NON-default 1.905 cm indent (1.27 + 0.635).
  '<text:list-style style:name="WL1">' +
  bulletLevel(1, "•", "1.27") +
  bulletLevel(2, "◦", "1.905") +
  bulletLevel(3, "▪", "3.175") +
  "</text:list-style>" +
  '<text:list-style style:name="WL2">' +
  numberLevel(1, "1", 1, "0.635") +
  numberLevel(2, "a", 1, "1.905") +
  "</text:list-style>" +
  '<text:list-style style:name="WL3">' +
  numberLevel(1, "1", 5, "0.635") +
  "</text:list-style>" +
  "</office:automatic-styles>" +
  "<office:body><office:text>" +
  "<text:p>Groceries</text:p>" +
  '<text:list text:style-name="WL1">' +
  "<text:list-item><text:p>Apples</text:p></text:list-item>" +
  "<text:list-item><text:p>Bananas</text:p>" +
  "<text:list><text:list-item><text:p>Green ones</text:p>" +
  "<text:list><text:list-item><text:p>Wrapped ones</text:p></text:list-item></text:list>" +
  "</text:list-item></text:list>" +
  "</text:list-item>" +
  "</text:list>" +
  "<text:p>Steps</text:p>" +
  '<text:list text:style-name="WL2">' +
  "<text:list-item><text:p>First</text:p></text:list-item>" +
  "<text:list-item><text:p>Second</text:p>" +
  "<text:list><text:list-item><text:p>Detail</text:p></text:list-item></text:list>" +
  "</text:list-item>" +
  "</text:list>" +
  "<text:p>Continue from five</text:p>" +
  '<text:list text:style-name="WL3">' +
  "<text:list-item><text:p>Fifth</text:p></text:list-item>" +
  "<text:list-item><text:p>Sixth</text:p></text:list-item>" +
  "</text:list>" +
  "</office:text></office:body></office:document-content>";

const odtStylesXml =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  `<office:document-styles xmlns:office="${OFFICE}" xmlns:style="${STYLE}" xmlns:fo="${FO}" office:version="1.2">` +
  "<office:styles>" +
  '<style:default-style style:family="paragraph">' +
  '<style:text-properties fo:font-family="Calibri" fo:font-size="11pt"/>' +
  "</style:default-style>" +
  "</office:styles>" +
  "<office:automatic-styles>" +
  '<style:page-layout style:name="pm1">' +
  '<style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"' +
  ' fo:margin-top="2.54cm" fo:margin-right="2.54cm" fo:margin-bottom="2.54cm" fo:margin-left="2.54cm"/>' +
  "</style:page-layout>" +
  "</office:automatic-styles>" +
  "<office:master-styles>" +
  '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
  "</office:master-styles>" +
  "</office:document-styles>";

const odtManifestXml =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  `<manifest:manifest xmlns:manifest="${MANIFEST}" manifest:version="1.2">` +
  `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIMETYPE}"/>` +
  '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
  '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
  "</manifest:manifest>";

writeFileSync(
  path.join(outDir, "lists.odt"),
  buildZip([
    { name: "mimetype", content: ODT_MIMETYPE, store: true },
    { name: "content.xml", content: odtContentXml },
    { name: "styles.xml", content: odtStylesXml },
    { name: "META-INF/manifest.xml", content: odtManifestXml },
  ]),
);

console.log("Wrote lists.docx and lists.odt to", outDir);
