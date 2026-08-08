#!/usr/bin/env node
// One-off generator for the three SPEC-004 test fixtures under
// src/lib/docx/fixtures/. Not part of the app build — run manually
// (`node scripts/build-docx-fixtures.mjs`) if the fixtures ever need to be
// regenerated or extended.
//
// These are NOT files exported from real Word/LibreOffice/Google Docs (this
// environment has no network access and none of those tools installed) —
// they are hand-built OPC packages that reproduce each producer's known
// structural quirks (style naming, section orientation, proofErr markers,
// direct-vs-styled formatting balance) closely enough to exercise the same
// code paths a genuine export would. See /lane/implement-summary.md for the
// full rationale.
import { deflateRawSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, "..", "src", "lib", "docx", "fixtures");
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

function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.from(entry.content, "utf-8");
    const dataBytes = deflateRawSync(rawBytes);
    const method = 8;

    const localHeader = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(0),
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
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(0),
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

  const centralDirectory = Buffer.concat(centralParts);
  const localData = Buffer.concat(localParts);
  const eocd = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(centralDirectory.length),
    u32(localData.length),
    u16(0),
  ]);

  return Buffer.concat([localData, centralDirectory, eocd]);
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

function documentXml(bodyXml, sectPrXml) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
    ${sectPrXml}
  </w:body>
</w:document>`;
}

function sectPr({ w, h, orient, margin }) {
  return `<w:sectPr>
    <w:pgSz w:w="${w}" w:h="${h}"${orient ? ` w:orient="${orient}"` : ""}/>
    <w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="708" w:footer="708" w:gutter="0"/>
  </w:sectPr>`;
}

// ---------------------------------------------------------------------
// Fixture A: Word-authored
// ---------------------------------------------------------------------
const wordBody = `
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly Report</w:t></w:r></w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">This is a paragraph with </w:t></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
    <w:r><w:t xml:space="preserve"> and </w:t></w:r>
    <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
    <w:r><w:t xml:space="preserve"> Tab</w:t><w:tab/><w:t xml:space="preserve">here. </w:t></w:r>
    <w:r><w:t>Line one</w:t><w:br/><w:t>Line two</w:t></w:r>
  </w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section Two</w:t></w:r></w:p>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:ind w:left="720" w:right="720" w:firstLine="360"/>
      <w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/>
    </w:pPr>
    <w:r><w:t>Centered paragraph with custom spacing and indent.</w:t></w:r>
  </w:p>`;

const wordStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr>
  </w:style>
</w:styles>`;

writeFileSync(
  path.join(outDir, "word-authored.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: CONTENT_TYPES },
    { name: "_rels/.rels", content: ROOT_RELS },
    { name: "word/_rels/document.xml.rels", content: DOCUMENT_RELS },
    {
      name: "word/document.xml",
      content: documentXml(wordBody, sectPr({ w: 11906, h: 16838, margin: 1440 })),
    },
    { name: "word/styles.xml", content: wordStyles },
  ]),
);

// ---------------------------------------------------------------------
// Fixture B: LibreOffice-authored
// ---------------------------------------------------------------------
const libreOfficeBody = `
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Meeting Minutes</w:t></w:r></w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">Normal text </w:t></w:r>
    <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underlined</w:t></w:r>
    <w:r><w:t xml:space="preserve"> and </w:t></w:r>
    <w:r><w:rPr><w:strike/></w:rPr><w:t>struck through</w:t></w:r>
    <w:r><w:t>.</w:t></w:r>
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">   spaced text   </w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Action Items</w:t></w:r></w:p>
  <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Right aligned text.</w:t></w:r></w:p>`;

const libreOfficeStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr>
  </w:style>
</w:styles>`;

writeFileSync(
  path.join(outDir, "libreoffice-authored.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: CONTENT_TYPES },
    { name: "_rels/.rels", content: ROOT_RELS },
    { name: "word/_rels/document.xml.rels", content: DOCUMENT_RELS },
    {
      name: "word/document.xml",
      content: documentXml(
        libreOfficeBody,
        sectPr({ w: 16838, h: 11906, orient: "landscape", margin: 1134 }),
      ),
    },
    { name: "word/styles.xml", content: libreOfficeStyles },
  ]),
);

// ---------------------------------------------------------------------
// Fixture C: Google-Docs-exported
// ---------------------------------------------------------------------
const googleDocsBody = `
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Project Overview</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Goals</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Sub-goal</w:t></w:r></w:p>
  <w:p>
    <w:proofErr w:type="spellStart"/>
    <w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="28"/><w:color w:val="1155CC"/></w:rPr><w:t>Colored</w:t></w:r>
    <w:proofErr w:type="spellEnd"/>
    <w:r><w:t xml:space="preserve"> text in Arial 14pt.</w:t></w:r>
  </w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">Some </w:t></w:r>
    <w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>highlighted text</w:t></w:r>
    <w:r><w:t xml:space="preserve"> here.</w:t></w:r>
  </w:p>
  <w:p><w:r><w:t>Line one</w:t><w:br/><w:t>Line two</w:t></w:r></w:p>`;

const googleDocsStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>`;

writeFileSync(
  path.join(outDir, "google-docs-exported.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: CONTENT_TYPES },
    { name: "_rels/.rels", content: ROOT_RELS },
    { name: "word/_rels/document.xml.rels", content: DOCUMENT_RELS },
    {
      name: "word/document.xml",
      content: documentXml(googleDocsBody, sectPr({ w: 12240, h: 15840, margin: 1440 })),
    },
    { name: "word/styles.xml", content: googleDocsStyles },
  ]),
);

// ---------------------------------------------------------------------
// Fixture D: Word-authored, containing a table of contents (SPEC-033
// AC-19, AC-20) — the shape real Word emits: a `TOC` complex field
// (fldChar begin/separate/end around an instrText instruction) whose
// cached result is TOCn-styled paragraphs, each entry's page number itself
// a NESTED `PAGEREF` complex field inside the entry's own hyperlink. The
// reader must track field nesting depth to find the OUTER field's true end
// rather than stopping at the first entry's nested PAGEREF end.
// ---------------------------------------------------------------------
const wordTocBody = `
  <w:p>
    <w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:hyperlink w:anchor="_Toc1" w:history="1">
      <w:r><w:t>Introduction</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:hyperlink>
  </w:p>
  <w:p>
    <w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>
    <w:hyperlink w:anchor="_Toc2" w:history="1">
      <w:r><w:t>Background</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:hyperlink>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
  <w:p><w:r><w:t>Opening paragraph.</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc2"/><w:r><w:t>Background</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
  <w:p><w:r><w:t>More detail.</w:t></w:r></w:p>`;

writeFileSync(
  path.join(outDir, "word-toc.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: CONTENT_TYPES },
    { name: "_rels/.rels", content: ROOT_RELS },
    { name: "word/_rels/document.xml.rels", content: DOCUMENT_RELS },
    {
      name: "word/document.xml",
      content: documentXml(wordTocBody, sectPr({ w: 11906, h: 16838, margin: 1440 })),
    },
    { name: "word/styles.xml", content: wordStyles },
  ]),
);

console.log("Wrote fixtures to", outDir);
