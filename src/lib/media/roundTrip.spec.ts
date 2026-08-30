// SPEC-107 AC: a media file is stored verbatim — the bytes that came in are
// the bytes that go back out, byte-for-byte, for all six formats. This is
// round-trip fidelity taken to its trivial-but-absolute case: no writer, no
// re-encode, no metadata rewrite.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { assembleMediaModel } from "./open";
import { mediaFileName, mediaSourceBytes } from "./download";
import type { MediaFormat } from "./model";
import {
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildPagesPdfBytes,
  buildPngFixtureBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
} from "./fixtures/build";

const FIXTURES: Array<{ format: MediaFormat; bytes: () => Promise<Uint8Array> | Uint8Array }> = [
  { format: "pdf", bytes: () => buildPagesPdfBytes() },
  { format: "png", bytes: () => buildPngFixtureBytes() },
  { format: "jpeg", bytes: () => buildJpegFixtureBytes() },
  { format: "gif", bytes: () => buildGifFixtureBytes() },
  { format: "svg", bytes: () => buildSvgFixtureBytes() },
  { format: "webp", bytes: () => buildWebpFixtureBytes() },
];

describe("media round trip", () => {
  for (const fixture of FIXTURES) {
    it(`${fixture.format}: download bytes are identical to the source`, async () => {
      const source = await fixture.bytes();
      const model = await assembleMediaModel("fixture", fixture.format, source);
      const out = mediaSourceBytes(model);
      expect(equalBytes(out, source)).toBe(true);
    });
  }

  it("downloads under the title with the format's extension", async () => {
    const model = await assembleMediaModel("holiday photo", "jpeg", buildJpegFixtureBytes());
    expect(mediaFileName(model)).toBe("holiday photo.jpeg");
    model.title = "  ";
    expect(mediaFileName(model)).toBe("Untitled.jpeg");
  });
});
