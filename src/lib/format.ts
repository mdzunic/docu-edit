import type { DocumentFormat } from "./documentStore";
import type { DocumentFormat as ModelDocumentFormat } from "./documentModel";

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
  md: "MD",
  new: "New",
};

/** The display label for a library row's format badge. */
export function formatBadgeLabel(format: DocumentFormat): string {
  return BADGE_LABELS[format];
}

const DOCUMENT_FORMAT_LABELS: Record<ModelDocumentFormat, string> = {
  docx: "Word (.docx)",
  odt: "OpenDocument Text (.odt)",
  md: "Markdown (.md)",
};

/** The full format name for the Properties dialog's General tab (SPEC-041
 * AC-2) — "" for a document with no source/saved format yet (AC-18a). */
export function documentFormatLabel(format?: ModelDocumentFormat): string {
  return format ? DOCUMENT_FORMAT_LABELS[format] : "";
}

/** A byte count as a human-readable size (SPEC-041 AC-2's General tab),
 * e.g. 245_000 -> "245 KB". Binary units (1 KB = 1024 B), one decimal place
 * from MB up, none below it — file sizes don't need sub-byte/sub-KB
 * precision. */
export function formatFileSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  const kb = bytes / 1024;
  if (kb < 1024) {
    return `${Math.round(kb)} KB`;
  }
  const mb = kb / 1024;
  if (mb < 1024) {
    return `${mb.toFixed(1)} MB`;
  }
  return `${(mb / 1024).toFixed(1)} GB`;
}
