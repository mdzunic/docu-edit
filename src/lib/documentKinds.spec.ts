// The kind registry (SPEC-107 adds the third kind, "media"): every
// kind-dispatched call site derives from these values, so they are pinned
// here once rather than re-asserted wherever a PDF happens to be routed.
import { describe, expect, it } from "vitest";
import {
  editorPathFor,
  kindOfFormat,
  MEDIA_FORMAT_CONTENT_TYPES,
  MEDIA_FORMATS,
  OPEN_FILE_ACCEPT,
  OPENABLE_EXTENSIONS,
} from "./documentKinds";
import { ACCEPTED_IMAGE_MIME_TYPES, LIBRARY_IMAGE_MIME_TYPES } from "./core/media";
import { MEDIA_CONTENT_TYPES } from "./media/model";

describe("documentKinds — the media kind (SPEC-107)", () => {
  it("lists the six media formats in display order", () => {
    expect(MEDIA_FORMATS).toEqual(["pdf", "png", "jpeg", "gif", "svg", "webp"]);
  });

  it("routes every media format, and only those, to the media kind", () => {
    for (const format of MEDIA_FORMATS) {
      expect(kindOfFormat(format)).toBe("media");
    }
    expect(kindOfFormat("pptx")).toBe("presentation");
    expect(kindOfFormat("odp")).toBe("presentation");
    for (const format of ["docx", "odt", "md", "new", "jpg"]) {
      expect(kindOfFormat(format)).toBe("text");
    }
  });

  it("gives each kind its own editor route", () => {
    expect(editorPathFor("media", "m1")).toBe("/media/m1");
    expect(editorPathFor("presentation", "p1")).toBe("/slides/p1");
    expect(editorPathFor("text", "t1")).toBe("/doc/t1");
  });

  it("opens the media extensions — jpg and jpeg both — through every picker and the drop check", () => {
    for (const extension of ["pdf", "png", "jpg", "jpeg", "gif", "svg", "webp"]) {
      expect(OPENABLE_EXTENSIONS).toContain(extension);
      expect(OPEN_FILE_ACCEPT.split(",")).toContain(`.${extension}`);
    }
  });

  it("derives the media content types from the media product's own table", () => {
    expect(MEDIA_FORMAT_CONTENT_TYPES).toEqual([
      "application/pdf",
      "image/png",
      "image/jpeg",
      "image/gif",
      "image/svg+xml",
      "image/webp",
    ]);
    expect(MEDIA_FORMAT_CONTENT_TYPES).toEqual(MEDIA_FORMATS.map((format) => MEDIA_CONTENT_TYPES[format]));
  });

  it("keeps the library's image types (+ WebP) apart from SPEC-024's insert-image list", () => {
    expect(ACCEPTED_IMAGE_MIME_TYPES).toEqual(["image/png", "image/jpeg", "image/gif", "image/svg+xml"]);
    expect(LIBRARY_IMAGE_MIME_TYPES).toEqual([...ACCEPTED_IMAGE_MIME_TYPES, "image/webp"]);
    // Every image format the library holds is served under one of them.
    expect(MEDIA_FORMAT_CONTENT_TYPES.filter((type) => type.startsWith("image/")).toSorted()).toEqual(
      [...LIBRARY_IMAGE_MIME_TYPES].toSorted(),
    );
  });
});
