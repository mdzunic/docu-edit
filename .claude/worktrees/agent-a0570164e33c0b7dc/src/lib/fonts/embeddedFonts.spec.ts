import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import {
  deobfuscateOdttf,
  extractDocxEmbeddedFonts,
  extractOdtEmbeddedFonts,
  fsTypeAllowsEmbedding,
  readFsType,
} from "./embeddedFonts";

function bundled(file: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(__dirname, "../../../public/fonts", file)));
}

const encoder = new TextEncoder();

describe("deobfuscateOdttf", () => {
  const key = "{6ADCF233-6C4E-4B2A-9F11-0011223344556}".slice(0, 38);
  const guid = "6ADCF2336C4E4B2A9F11001122334455";

  it("is symmetric — obfuscating then de-obfuscating restores the bytes", () => {
    const original = bundled("Carlito-Regular.ttf");
    const obfuscated = deobfuscateOdttf(original, guid);
    expect(Array.from(obfuscated.slice(0, 32))).not.toEqual(Array.from(original.slice(0, 32)));
    const restored = deobfuscateOdttf(obfuscated, guid);
    expect(Array.from(restored)).toEqual(Array.from(original));
  });

  it("only touches the first 32 bytes", () => {
    const original = bundled("Carlito-Regular.ttf");
    const obfuscated = deobfuscateOdttf(original, guid);
    expect(Array.from(obfuscated.slice(32))).toEqual(Array.from(original.slice(32)));
  });

  it("leaves bytes untouched when the key is malformed", () => {
    const original = bundled("Carlito-Regular.ttf");
    expect(Array.from(deobfuscateOdttf(original, "not-a-guid"))).toEqual(Array.from(original));
    void key;
  });
});

describe("readFsType / fsTypeAllowsEmbedding", () => {
  it("reads the OS/2 fsType of a real bundled face", () => {
    // Liberation/Carlito ship as freely embeddable (installable) fonts.
    const fsType = readFsType(bundled("LiberationSans-Regular.ttf"));
    expect(fsType).not.toBeNull();
    expect(fsTypeAllowsEmbedding(fsType)).toBe(true);
  });

  it("treats the restricted-licence bit as forbidding embedding", () => {
    expect(fsTypeAllowsEmbedding(0x0002)).toBe(false);
    expect(fsTypeAllowsEmbedding(0x0002 | 0x0100)).toBe(false);
    expect(fsTypeAllowsEmbedding(0x0004)).toBe(true); // preview & print
    expect(fsTypeAllowsEmbedding(0x0008)).toBe(true); // editable
    expect(fsTypeAllowsEmbedding(0)).toBe(true);
    expect(fsTypeAllowsEmbedding(null)).toBe(true);
  });
});

describe("extractDocxEmbeddedFonts", () => {
  const fontKey = "{6ADCF233-6C4E-4B2A-9F11-001122334455}";

  function buildParts(fontBytes: Uint8Array): Map<string, Uint8Array> {
    const obfuscated = deobfuscateOdttf(fontBytes, fontKey);
    const parts = new Map<string, Uint8Array>();
    parts.set(
      "word/fontTable.xml",
      encoder.encode(
        `<?xml version="1.0"?><w:fonts xmlns:w="${"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}" ` +
          `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
          `<w:font w:name="CoolFont"><w:embedRegular r:id="rId1" w:fontKey="${fontKey}"/>` +
          `<w:embedBold r:id="rId2" w:fontKey="${fontKey}"/></w:font></w:fonts>`,
      ),
    );
    parts.set(
      "word/_rels/fontTable.xml.rels",
      encoder.encode(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>` +
          `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>` +
          `</Relationships>`,
      ),
    );
    parts.set("word/fonts/font1.odttf", obfuscated);
    parts.set("word/fonts/font2.odttf", obfuscated);
    return parts;
  }

  it("extracts and de-obfuscates each embedded face, keeping the family and weight (AC-11)", () => {
    const original = bundled("Carlito-Regular.ttf");
    const faces = extractDocxEmbeddedFonts(buildParts(original));
    expect(faces).toHaveLength(2);
    const regular = faces.find((f) => !f.bold);
    const bold = faces.find((f) => f.bold);
    expect(regular?.family).toBe("CoolFont");
    expect(Array.from(regular?.bytes ?? [])).toEqual(Array.from(original));
    expect(regular?.embeddable).toBe(true);
    expect(bold?.bold).toBe(true);
    expect(bold?.italic).toBe(false);
  });

  it("returns nothing when there is no fontTable part", () => {
    expect(extractDocxEmbeddedFonts(new Map())).toEqual([]);
  });
});

describe("extractOdtEmbeddedFonts", () => {
  it("extracts fonts referenced from the font-face-decls, reading the Fonts part (AC-13)", () => {
    const original = bundled("Caladea-Regular.ttf");
    const decls =
      `<office:font-face-decls xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
      `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
      `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
      `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
      `xmlns:loext="urn:oasis:names:tc:openoffice:xmlns:loext:1.0">` +
      `<style:font-face style:name="F1" svg:font-family="MyEmbedded">` +
      `<svg:font-face-src><svg:font-face-uri xlink:href="Fonts/MyEmbedded.ttf" ` +
      `loext:font-weight="bold" loext:font-style="italic">` +
      `<svg:font-face-format svg:string="truetype"/></svg:font-face-uri></svg:font-face-src>` +
      `</style:font-face></office:font-face-decls>`;
    const parts = new Map<string, Uint8Array>([["Fonts/MyEmbedded.ttf", original]]);
    const faces = extractOdtEmbeddedFonts(decls, parts);
    expect(faces).toHaveLength(1);
    expect(faces[0].family).toBe("MyEmbedded");
    expect(faces[0].bold).toBe(true);
    expect(faces[0].italic).toBe(true);
    expect(Array.from(faces[0].bytes)).toEqual(Array.from(original));
    expect(faces[0].embeddable).toBe(true);
  });

  it("returns nothing without font-face declarations", () => {
    expect(extractOdtEmbeddedFonts(undefined, new Map())).toEqual([]);
  });
});
