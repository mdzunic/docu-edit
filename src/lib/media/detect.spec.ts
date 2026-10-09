// detectMedia (SPEC-107 AC-6): every signature, plus the SVG
// false-positive cases the spec calls out — an XHTML file, an XML file whose
// COMMENT mentions svg, and an empty file. The sniff runs on bytes alone;
// no test here involves a file name.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { detectMedia } from "../core/packageDetect";
import {
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildPngFixtureBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
  buildXrefStreamPdfBytes,
} from "./fixtures/build";

const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../../e2e/fixtures/spec107-pages.pdf")));

function utf8(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

describe("detectMedia", () => {
  it("recognises all six formats from their bytes", async () => {
    expect(detectMedia(PAGES_PDF)).toBe("pdf");
    expect(detectMedia(await buildXrefStreamPdfBytes())).toBe("pdf");
    expect(detectMedia(await buildPngFixtureBytes())).toBe("png");
    expect(detectMedia(buildJpegFixtureBytes())).toBe("jpeg");
    expect(detectMedia(buildGifFixtureBytes())).toBe("gif");
    expect(detectMedia(buildWebpFixtureBytes())).toBe("webp");
    expect(detectMedia(buildSvgFixtureBytes())).toBe("svg");
  });

  it("finds a %PDF- header within the first 1024 bytes, not beyond", () => {
    const prefixed = new Uint8Array(100 + 20);
    prefixed.set(utf8("%PDF-1.4"), 100);
    expect(detectMedia(prefixed)).toBe("pdf");
    const buried = new Uint8Array(2000 + 20);
    buried.set(utf8("%PDF-1.4"), 1500);
    expect(detectMedia(buried)).toBeNull();
  });

  it("accepts SVG behind a BOM, XML declaration, doctype and comments", () => {
    const svg =
      "﻿<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
      "<!-- a leading comment -->\n<!DOCTYPE svg>\n<!-- another -->\n" +
      "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
    expect(detectMedia(utf8(svg))).toBe("svg");
  });

  it("accepts an Illustrator-style SVG whose doctype declares entities in an internal subset", () => {
    const svg =
      '<?xml version="1.0" encoding="utf-8"?>\n' +
      "<!-- Generator: Adobe Illustrator 27.0.0, SVG Export Plug-In . SVG Version: 6.00 Build 0)  -->\n" +
      '<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd" [\n' +
      '  <!ENTITY ns_extend "http://ns.adobe.com/Extensibility/1.0/">\n' +
      "  <!-- a comment ] with > inside the subset -->\n" +
      '  <!ENTITY ns_ai "http://ns.adobe.com/AdobeIllustrator/10.0/">\n' +
      "]>\n" +
      '<svg version="1.1" xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>';
    expect(detectMedia(utf8(svg))).toBe("svg");
  });

  it("accepts a UTF-16 SVG behind its byte-order mark", () => {
    const text = '<?xml version="1.0" encoding="UTF-16"?><svg xmlns="http://www.w3.org/2000/svg"/>';
    const le = new Uint8Array(2 + text.length * 2);
    const be = new Uint8Array(2 + text.length * 2);
    le.set([0xff, 0xfe]);
    be.set([0xfe, 0xff]);
    for (let i = 0; i < text.length; i += 1) {
      le[2 + i * 2] = text.charCodeAt(i);
      be[3 + i * 2] = text.charCodeAt(i);
    }
    expect(detectMedia(le)).toBe("svg");
    expect(detectMedia(be)).toBe("svg");
  });

  it("only looks within the first 4 KiB for the root element", () => {
    const padded = `<!-- ${"x".repeat(5000)} --><svg xmlns="http://www.w3.org/2000/svg"/>`;
    expect(detectMedia(utf8(padded))).toBeNull();
  });

  it("accepts a namespace-prefixed svg root", () => {
    expect(detectMedia(utf8('<s:svg xmlns:s="http://www.w3.org/2000/svg"/>'))).toBe("svg");
  });

  it("rejects an XHTML file", () => {
    const xhtml =
      '<?xml version="1.0"?><!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml"><body/></html>';
    expect(detectMedia(utf8(xhtml))).toBeNull();
  });

  it("rejects an XML file whose comment mentions svg", () => {
    const xml = '<?xml version="1.0"?><!-- contains <svg> in a comment --><data/>';
    expect(detectMedia(utf8(xml))).toBeNull();
  });

  it("rejects an element merely starting with the letters svg", () => {
    expect(detectMedia(utf8("<svgfoo/>"))).toBeNull();
  });

  it("rejects an empty file and plain text", () => {
    expect(detectMedia(new Uint8Array(0))).toBeNull();
    expect(detectMedia(utf8("hello world"))).toBeNull();
  });

  it("tells the two GIF versions apart from other GIF-ish bytes, and WebP from other RIFF files", () => {
    expect(detectMedia(utf8("GIF87a\u0001\u0000"))).toBe("gif");
    expect(detectMedia(utf8("GIF89a\u0001\u0000"))).toBe("gif");
    expect(detectMedia(utf8("GIF88a\u0001\u0000"))).toBeNull();
    expect(detectMedia(utf8("RIFF\u0004\u0000\u0000\u0000WAVEfmt "))).toBeNull();
    expect(detectMedia(new Uint8Array([0xff, 0xd8]))).toBeNull();
  });

  it("rejects ZIP-shaped bytes (packages take the package path)", () => {
    expect(detectMedia(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]))).toBeNull();
  });
});
