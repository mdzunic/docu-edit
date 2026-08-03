import type { DocumentFormat } from "./documentStore";

/**
 * Format a date for display as a document title, e.g. "2026-08-02 14:35".
 * Local time in a fixed layout, so output is deterministic across locales.
 */
export function formatTimestamp(date: Date): string {
  const pad = (value: number): string => String(value).padStart(2, "0");
  const day = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  const time = `${pad(date.getHours())}:${pad(date.getMinutes())}`;
  return `${day} ${time}`;
}

/**
 * Format an epoch-ms timestamp relative to `nowMs`, e.g. "2 hours ago".
 * Timestamps in the future (clock skew) are treated as "just now".
 */
export function formatRelativeTime(timestampMs: number, nowMs: number): string {
  const unit = (value: number, name: string): string =>
    value === 1 ? `1 ${name} ago` : `${value} ${name}s ago`;

  const seconds = Math.max(0, Math.floor((nowMs - timestampMs) / 1000));
  if (seconds < 60) {
    return "just now";
  }
  const minutes = Math.floor(seconds / 60);
  if (minutes < 60) {
    return unit(minutes, "minute");
  }
  const hours = Math.floor(minutes / 60);
  if (hours < 24) {
    return unit(hours, "hour");
  }
  const days = Math.floor(hours / 24);
  if (days < 30) {
    return unit(days, "day");
  }
  const months = Math.floor(days / 30);
  if (months < 12) {
    return unit(months, "month");
  }
  return unit(Math.max(1, Math.floor(days / 365)), "year");
}

const BADGE_LABELS: Record<DocumentFormat, string> = {
  docx: "DOCX",
  odt: "ODT",
  new: "New",
};

/** The display label for a library row's format badge. */
export function formatBadgeLabel(format: DocumentFormat): string {
  return BADGE_LABELS[format];
}
