// SPEC-093 phase-one core move: the date/time catalogue and renderer now live
// here (product-neutral), re-exported by src/lib/fields.ts. This spec pins the
// catalogue's new home; src/lib/fields.spec.ts pins the same behaviour through
// the re-export and must keep passing unchanged.
import { describe, expect, it } from "vitest";
import {
  DATE_TIME_CATALOGUE,
  canonicalDateTimeFormatFromDocx,
  dateTimeFormatEntry,
  formatFromOdtDateStyleName,
  isSupportedDateTimeFormat,
  odtDateStyleName,
  renderDateTime,
} from "./dateTimeFormats";
import { DATE_TIME_CATALOGUE as REEXPORTED } from "../fields";

const INSTANT = new Date(2026, 4, 14, 15, 4, 5); // Thu 14 May 2026, 15:04:05 local

describe("core/dateTimeFormats catalogue", () => {
  it("keeps the nine catalogue entries in spec order", () => {
    expect(DATE_TIME_CATALOGUE.map((e) => e.format)).toEqual([
      "d MMMM yyyy",
      "d MMM yyyy",
      "dd/MM/yyyy",
      "MM/dd/yyyy",
      "yyyy-MM-dd",
      "EEEE, d MMMM yyyy",
      "HH:mm",
      "h:mm a",
      "HH:mm:ss",
    ]);
  });

  it("is the very object src/lib/fields.ts re-exports", () => {
    expect(REEXPORTED).toBe(DATE_TIME_CATALOGUE);
  });

  it("recognises catalogue patterns and rejects foreign ones", () => {
    expect(isSupportedDateTimeFormat("d MMMM yyyy")).toBe(true);
    expect(isSupportedDateTimeFormat("d MMMM, yyyy 'at' HH:mm")).toBe(false);
  });

  it("renders each date pattern against a fixed instant (en-US)", () => {
    expect(renderDateTime("d MMMM yyyy", INSTANT, "en-US")).toBe("14 May 2026");
    expect(renderDateTime("d MMM yyyy", INSTANT, "en-US")).toBe("14 May 2026");
    expect(renderDateTime("dd/MM/yyyy", INSTANT, "en-US")).toBe("14/05/2026");
    expect(renderDateTime("MM/dd/yyyy", INSTANT, "en-US")).toBe("05/14/2026");
    expect(renderDateTime("yyyy-MM-dd", INSTANT, "en-US")).toBe("2026-05-14");
    expect(renderDateTime("EEEE, d MMMM yyyy", INSTANT, "en-US")).toBe("Thursday, 14 May 2026");
  });

  it("returns null for a format outside the catalogue", () => {
    expect(renderDateTime("not a pattern", INSTANT, "en-US")).toBeNull();
  });

  it("round-trips the odt date-style name for catalogue entries", () => {
    for (const entry of DATE_TIME_CATALOGUE) {
      const name = odtDateStyleName(entry);
      expect(formatFromOdtDateStyleName(name)).toBe(entry.format);
    }
    expect(formatFromOdtDateStyleName("N123")).toBeNull();
  });

  it("maps docx picture strings back to canonical patterns", () => {
    expect(canonicalDateTimeFormatFromDocx("dddd, d MMMM yyyy")).toBe("EEEE, d MMMM yyyy");
    expect(canonicalDateTimeFormatFromDocx("d MMMM yyyy")).toBe("d MMMM yyyy");
    expect(canonicalDateTimeFormatFromDocx("something foreign")).toBe("something foreign");
  });

  it("looks up entries by canonical pattern", () => {
    expect(dateTimeFormatEntry("dd/MM/yyyy")?.docxPicture).toBe("dd/MM/yyyy");
    expect(dateTimeFormatEntry("nope")).toBeUndefined();
  });
});
