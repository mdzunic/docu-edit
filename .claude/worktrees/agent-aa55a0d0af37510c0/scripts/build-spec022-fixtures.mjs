#!/usr/bin/env node
// One-off generator for the SPEC-022 table fixture under e2e/fixtures/.
// Not part of the app build — run manually
// (`node scripts/build-spec022-fixtures.mjs`) if the fixture ever needs to
// be regenerated or extended.
//
// Like the other fixtures (see build-docx-fixtures.mjs), table.docx is NOT
// a real Word export — it is a hand-built package reproducing Word's table
// markup: w:tbl with w:tblPr (single-line borders), a w:tblGrid carrying
// the column widths, and per-cell w:tcPr/w:tcW. The table covers the three
// e2e-rendered shapes: plain text cells, one empty cell (just a bare
// w:p), and one cell holding two paragraphs.
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
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function p(text) {
  return text.length > 0 ? `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` : "<w:p/>";
}

function tc(content) {
  return `<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
}

const tableXml =
  "<w:tbl>" +
  "<w:tblPr>" +
  '<w:tblW w:w="9026" w:type="dxa"/>' +
  "<w:tblBorders>" +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:start w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:end w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  "</w:tblBorders>" +
  '<w:tblLayout w:type="fixed"/>' +
  "</w:tblPr>" +
  '<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>' +
  // Row 1: a Heading1-styled cell paragraph (cells accept styles) beside a
  // plain text cell.
  "<w:tr>" +
  tc('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>') +
  tc(p("Bravo")) +
  "</w:tr>" +
  // Row 2: an EMPTY first cell (a bare w:p) beside a text cell.
  `<w:tr>${tc(p(""))}${tc(p("Charlie"))}</w:tr>` +
  // Row 3: a multi-paragraph first cell.
  `<w:tr>${tc(p("First line") + p("Second line"))}${tc(p("Delta"))}</w:tr>` +
  "</w:tbl>";

const documentXml =
  `${XML}<w:document xmlns:w="${W}"><w:body>` +
  p("Before the table") +
  tableXml +
  p("After the table") +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr></w:body></w:document>";

const stylesXml =
  `${XML}<w:styles xmlns:w="${W}">` +
  "<w:docDefaults><w:rPrDefault><w:rPr>" +
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>' +
  "</w:rPr></w:rPrDefault></w:docDefaults>" +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1">' +
  '<w:name w:val="heading 1"/>' +
  '<w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>' +
  "</w:style>" +
  "</w:styles>";

const contentTypesXml =
  `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  "</Types>";

const packageRelsXml =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships>";

const documentRelsXml =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  "</Relationships>";

writeFileSync(
  path.join(outDir, "table.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
  ]),
);

console.log("Wrote", path.join(outDir, "table.docx"));
