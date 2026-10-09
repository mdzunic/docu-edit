import { describe, it, expect } from "vitest";
import { formatBadgeLabel, formatRelativeTime, formatTimestamp, recoveryNotice, spreadsheetFormatLabel } from "./format";

describe("formatTimestamp", () => {
  it("formats a date as 'YYYY-MM-DD HH:mm' in local time", () => {
    expect(formatTimestamp(new Date(2026, 7, 2, 14, 35, 59))).toBe(
      "2026-08-02 14:35",
    );
  });

  it("zero-pads month, day, hours and minutes", () => {
    expect(formatTimestamp(new Date(2026, 0, 5, 9, 7, 0))).toBe(
      "2026-01-05 09:07",
    );
  });
});

// SPEC-003 — relative timestamps for the library's "last modified" column.
describe("formatRelativeTime", () => {
  const MINUTE = 60_000;
  const HOUR = 60 * MINUTE;
  const DAY = 24 * HOUR;
  const now = 1_000_000_000_000;

  it("says 'just now' under a minute, and for future timestamps", () => {
    expect(formatRelativeTime(now, now)).toBe("just now");
    expect(formatRelativeTime(now - 59_000, now)).toBe("just now");
    expect(formatRelativeTime(now + 5_000, now)).toBe("just now");
  });

  it("formats minutes with singular/plural", () => {
    expect(formatRelativeTime(now - MINUTE, now)).toBe("1 minute ago");
    expect(formatRelativeTime(now - 59 * MINUTE, now)).toBe("59 minutes ago");
  });

  it("formats hours (e.g. '2 hours ago')", () => {
    expect(formatRelativeTime(now - HOUR, now)).toBe("1 hour ago");
    expect(formatRelativeTime(now - 2 * HOUR, now)).toBe("2 hours ago");
    expect(formatRelativeTime(now - 23 * HOUR, now)).toBe("23 hours ago");
  });

  it("formats days, months and years", () => {
    expect(formatRelativeTime(now - DAY, now)).toBe("1 day ago");
    expect(formatRelativeTime(now - 29 * DAY, now)).toBe("29 days ago");
    expect(formatRelativeTime(now - 31 * DAY, now)).toBe("1 month ago");
    expect(formatRelativeTime(now - 200 * DAY, now)).toBe("6 months ago");
    expect(formatRelativeTime(now - 400 * DAY, now)).toBe("1 year ago");
    expect(formatRelativeTime(now - 800 * DAY, now)).toBe("2 years ago");
  });
});

// SPEC-003 — format badge labels (AC-6).
describe("formatBadgeLabel", () => {
  it("maps each stored format to its badge text", () => {
    expect(formatBadgeLabel("docx")).toBe("DOCX");
    expect(formatBadgeLabel("odt")).toBe("ODT");
    expect(formatBadgeLabel("new")).toBe("New");
  });
});

const at = (label: string) => new Date(`2026-09-05T${label}Z`).getTime();

describe("recoveryNotice (SPEC-010)", () => {
  it("dates the draft relatively, not as a bare clock reading", () => {
    // The bug: `toLocaleTimeString()` rendered "3:36:58 PM" for a draft from a
    // minute ago and for one from last Tuesday alike — precise to the second
    // and useless for the only question being asked, which is whether these
    // changes are recent enough to want back.
    expect(recoveryNotice(at("15:36:58"), at("15:41:00"))).toBe(
      "Recovered unsaved changes from 4 minutes ago",
    );
    expect(recoveryNotice(at("15:36:58"), at("15:37:10"))).toBe(
      "Recovered unsaved changes from just now",
    );
  });

  it("keeps the prefix the recovery bar is identified by", () => {
    expect(recoveryNotice(at("15:00:00"), at("18:00:00"))).toMatch(
      /^Recovered unsaved changes from /,
    );
  });

  it("spans days, so a stale draft cannot read as a recent one", () => {
    const tuesday = new Date("2026-09-01T15:36:58Z").getTime();
    expect(recoveryNotice(tuesday, at("15:36:58"))).toBe(
      "Recovered unsaved changes from 4 days ago",
    );
  });
});

describe("workbook badges and labels (SPEC-116)", () => {
  it("badges a workbook by format + package variant", () => {
    expect(formatBadgeLabel("xlsx")).toBe("XLSX");
    expect(formatBadgeLabel("xlsx", "macroEnabled")).toBe("XLSM");
    expect(formatBadgeLabel("xlsx", "template")).toBe("XLTX");
    expect(formatBadgeLabel("xlsx", "macroEnabledTemplate")).toBe("XLTM");
    expect(formatBadgeLabel("ods")).toBe("ODS");
    expect(formatBadgeLabel("ods", "template")).toBe("OTS");
    // A variant means nothing outside the spreadsheet kind.
    expect(formatBadgeLabel("pptx", "template")).toBe("PPTX");
  });

  it("names each format and variant as the Reference table does", () => {
    expect(spreadsheetFormatLabel("xlsx")).toBe("Excel workbook (.xlsx)");
    expect(spreadsheetFormatLabel("xlsx", "macroEnabled")).toBe("Excel macro-enabled workbook (.xlsm)");
    expect(spreadsheetFormatLabel("xlsx", "template")).toBe("Excel template (.xltx)");
    expect(spreadsheetFormatLabel("xlsx", "macroEnabledTemplate")).toBe("Excel macro-enabled template (.xltm)");
    expect(spreadsheetFormatLabel("ods")).toBe("OpenDocument Spreadsheet (.ods)");
    expect(spreadsheetFormatLabel("ods", "template")).toBe("OpenDocument Spreadsheet template (.ots)");
  });
});
