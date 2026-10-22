// The content-type/extension mapping tests, moved verbatim from
// docx/drawing.spec.ts when the helpers moved to core/media.ts (both format
// packages share them; the AC references are SPEC-024's).
import { describe, it, expect } from "vitest";
import {
  ACCEPTED_MEDIA_MIME_TYPES,
  CONTENT_TYPE_BY_EXTENSION,
  contentTypeForPartName,
  encodePendingMediaInBackground,
  ensureMediaEncoded,
  extensionForContentType,
  flushPendingMedia,
  isAcceptedMediaType,
  MAX_MEDIA_SIZE_BYTES,
  MEDIA_ACCEPT_ATTR,
  MEDIA_TOO_LARGE_MESSAGE,
  mediaDataUri,
  pendingMediaByteLength,
  registerPendingMedia,
  type MediaPart,
} from "./media";
import { bytesToBase64 } from "./packageParts";

describe("contentTypeForPartName / extensionForContentType (AC-1, AC-2)", () => {
  it("maps every supported extension to its content type", () => {
    expect(contentTypeForPartName("word/media/image1.png")).toBe("image/png");
    expect(contentTypeForPartName("word/media/image2.jpg")).toBe("image/jpeg");
    expect(contentTypeForPartName("word/media/image3.jpeg")).toBe("image/jpeg");
    expect(contentTypeForPartName("word/media/image4.gif")).toBe("image/gif");
    expect(contentTypeForPartName("word/media/image5.svg")).toBe("image/svg+xml");
  });

  it("is case-insensitive on the extension", () => {
    expect(contentTypeForPartName("word/media/image1.PNG")).toBe("image/png");
  });

  it("returns undefined for an unknown extension or no extension", () => {
    expect(contentTypeForPartName("word/media/image1.bmp")).toBeUndefined();
    expect(contentTypeForPartName("word/media/image1")).toBeUndefined();
  });

  it("maps every content type CONTENT_TYPE_BY_EXTENSION covers back to its canonical extension", () => {
    expect(extensionForContentType("image/png")).toBe("png");
    expect(extensionForContentType("image/jpeg")).toBe("jpeg");
    expect(extensionForContentType("image/gif")).toBe("gif");
    expect(extensionForContentType("image/svg+xml")).toBe("svg");
  });

  it("falls back to png for an unknown content type", () => {
    expect(extensionForContentType("image/bmp")).toBe("png");
  });

  it("CONTENT_TYPE_BY_EXTENSION covers the image and media (SPEC-095) extensions", () => {
    expect(Object.keys(CONTENT_TYPE_BY_EXTENSION).toSorted()).toEqual([
      "gif",
      "jpeg",
      "jpg",
      "m4a",
      "mp3",
      "mp4",
      "ogg",
      "png",
      "svg",
      "wav",
      "webm",
      "webp",
    ]);
  });
});

// SPEC-095: the audio/video half of the media vocabulary, plus the
// extension/content-type rows the writers need so a clip never rides in a
// `.png`-named part.
describe("media (audio/video) vocabulary (SPEC-095)", () => {
  it("accepts the broad container types and rejects others", () => {
    expect([...ACCEPTED_MEDIA_MIME_TYPES]).toEqual([
      "video/mp4",
      "video/webm",
      "video/ogg",
      "audio/mpeg",
      "audio/mp4",
      "audio/ogg",
      "audio/wav",
    ]);
    expect(MEDIA_ACCEPT_ATTR).toBe("video/mp4,video/webm,video/ogg,audio/mpeg,audio/mp4,audio/ogg,audio/wav");
    expect(isAcceptedMediaType("video/mp4")).toBe(true);
    expect(isAcceptedMediaType("audio/wav")).toBe(true);
    expect(isAcceptedMediaType("video/quicktime")).toBe(false);
    expect(isAcceptedMediaType("image/png")).toBe(false);
  });

  it("caps inserted media at 25 MB with a message", () => {
    expect(MAX_MEDIA_SIZE_BYTES).toBe(25 * 1024 * 1024);
    expect(MEDIA_TOO_LARGE_MESSAGE).toContain("25 MB");
  });

  it("maps media extensions to content types and back", () => {
    expect(contentTypeForPartName("ppt/media/media1.mp4")).toBe("video/mp4");
    expect(contentTypeForPartName("ppt/media/media1.webm")).toBe("video/webm");
    expect(contentTypeForPartName("ppt/media/media1.m4a")).toBe("audio/mp4");
    expect(contentTypeForPartName("ppt/media/media1.mp3")).toBe("audio/mpeg");
    expect(contentTypeForPartName("Pictures/clip.wav")).toBe("audio/wav");
    expect(contentTypeForPartName("Pictures/poster.webp")).toBe("image/webp");

    expect(extensionForContentType("video/mp4")).toBe("mp4");
    expect(extensionForContentType("video/webm")).toBe("webm");
    expect(extensionForContentType("audio/mpeg")).toBe("mp3");
    expect(extensionForContentType("audio/mp4")).toBe("m4a");
    expect(extensionForContentType("audio/wav")).toBe("wav");
    expect(extensionForContentType("image/webp")).toBe("webp");
  });
});

// SPEC-047: the deferred-media encode (a large `.docx`'s image base64 pass
// moved off the open path). A part registered as pending carries no `data`
// until a consumer forces the encode; every force-encode is idempotent.
describe("deferred media encoding (SPEC-047)", () => {
  const bytes = new Uint8Array([1, 2, 3, 4, 5, 250, 128, 0]);

  function pendingPart(): MediaPart {
    const part: MediaPart = { mediaId: "rId1", contentType: "image/png", data: "" };
    registerPendingMedia(part, bytes);
    return part;
  }

  it("reports the pending byte length until the part is encoded", () => {
    const part = pendingPart();
    expect(part.data).toBe("");
    expect(pendingMediaByteLength(part)).toBe(bytes.length);
    ensureMediaEncoded(part);
    expect(part.data).toBe(bytesToBase64(bytes));
    expect(pendingMediaByteLength(part)).toBe(0);
  });

  it("ensureMediaEncoded is idempotent and a no-op for a never-deferred part", () => {
    const part = pendingPart();
    ensureMediaEncoded(part);
    const encoded = part.data;
    ensureMediaEncoded(part);
    expect(part.data).toBe(encoded);
    // A part that never deferred keeps whatever data it already had.
    const eager: MediaPart = { mediaId: "rId2", contentType: "image/png", data: "already" };
    ensureMediaEncoded(eager);
    expect(eager.data).toBe("already");
    ensureMediaEncoded(undefined);
  });

  it("mediaDataUri force-encodes a still-deferred part on demand", () => {
    const part = pendingPart();
    expect(mediaDataUri(part)).toBe(`data:image/png;base64,${bytesToBase64(bytes)}`);
    // and the part is now materialised for every later reader
    expect(pendingMediaByteLength(part)).toBe(0);
  });

  it("flushPendingMedia encodes an entire list (and tolerates undefined)", () => {
    const parts = [pendingPart(), pendingPart(), { mediaId: "x", contentType: "image/png", data: "eager" }];
    flushPendingMedia(parts);
    expect(parts.every((p) => p.data.length > 0)).toBe(true);
    flushPendingMedia(undefined);
  });

  it("encodePendingMediaInBackground encodes every pending part across chunk yields", async () => {
    const parts = Array.from({ length: 6 }, () => pendingPart());
    // A tiny chunk size forces several macrotask yields between encodes.
    await encodePendingMediaInBackground(parts, 4);
    for (const part of parts) {
      expect(part.data).toBe(bytesToBase64(bytes));
      expect(pendingMediaByteLength(part)).toBe(0);
    }
  });
});
