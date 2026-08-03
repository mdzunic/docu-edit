import { describe, it, expect } from "vitest";
import {
  formatBadgeLabel,
  formatRelativeTime,
  formatTimestamp,
} from "./format";

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
