import type { DocumentFormat } from "./documentStore";
import type { DocumentFormat as ModelDocumentFormat } from "./documentModel";
import type { PresentationFormat } from "./presentation/model";
import type { PackageVariant, SpreadsheetFormat } from "./spreadsheet/model";
import { spreadsheetFileExtension } from "./spreadsheet/fileTypes";
import { kindOfFormat } from "./documentKinds";

const pad = (value: number): string => String(value).padStart(2, "0");

/**
 * Format a date for display as a document title, e.g. "2026-08-02 14:35".
 * Local time in a fixed layout, so output is deterministic across locales.
 */
export function formatTimestamp(date: Date): string {
  const day = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  const time = `${pad(date.getHours())}:${pad(date.getMinutes())}`;
  return `${day} ${time}`;
}

const unit = (value: number, name: string): string =>
  value === 1 ? `1 ${name} ago` : `${value} ${name}s ago`;

/**
 * Format an epoch-ms timestamp relative to `nowMs`, e.g. "2 hours ago".
 * Timestamps in the future (clock skew) are treated as "just now".
 */
export function formatRelativeTime(timestampMs: number, nowMs: number): string {
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
  pptx: "PPTX",
  odp: "ODP",
  pdf: "PDF",
  png: "PNG",
  jpeg: "JPEG",
  gif: "GIF",
  svg: "SVG",
  webp: "WEBP",
  xlsx: "XLSX",
  ods: "ODS",
  csv: "CSV",
  tsv: "TSV",
  new: "New",
};

/** The display label for a library row's format badge. Accepts plain strings
 * (runtime-typed .vue props) and falls back to the upper-case format for any
 * value outside the known formats. A workbook is badged by format + package
 * variant (SPEC-116): XLSX, XLSM, XLTX, XLTM, ODS or OTS — its extension. */
export function formatBadgeLabel(format: string, packageVariant?: string): string {
  if (packageVariant && kindOfFormat(format) === "spreadsheet") {
    return spreadsheetFileExtension(format as SpreadsheetFormat, packageVariant as PackageVariant).toUpperCase();
  }
  return BADGE_LABELS[format as DocumentFormat] ?? format.toUpperCase();
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

const PRESENTATION_FORMAT_LABELS: Record<PresentationFormat, string> = {
  pptx: "PowerPoint (.pptx)",
  odp: "OpenDocument Presentation (.odp)",
};

/** The full format name a presentation's view header shows. */
export function presentationFormatLabel(format: PresentationFormat): string {
  return PRESENTATION_FORMAT_LABELS[format];
}

const SPREADSHEET_FORMAT_LABELS: Record<SpreadsheetFormat, Partial<Record<PackageVariant | "plain", string>>> = {
  xlsx: {
    plain: "Excel workbook (.xlsx)",
    macroEnabled: "Excel macro-enabled workbook (.xlsm)",
    template: "Excel template (.xltx)",
    macroEnabledTemplate: "Excel macro-enabled template (.xltm)",
  },
  ods: {
    plain: "OpenDocument Spreadsheet (.ods)",
    template: "OpenDocument Spreadsheet template (.ots)",
  },
  csv: { plain: "Comma-separated values (.csv)" },
  tsv: { plain: "Tab-separated values (.tsv)" },
};

/** The full format name a workbook's view header shows (SPEC-116), by
 * format + package variant. */
export function spreadsheetFormatLabel(format: SpreadsheetFormat, variant?: PackageVariant): string {
  const labels = SPREADSHEET_FORMAT_LABELS[format];
  return (variant !== undefined ? labels[variant] : undefined) ?? labels.plain ?? "";
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

/** The recovery bar's sentence (SPEC-010).
 *
 * Relative, not a clock reading: this used to render `toLocaleTimeString()`,
 * so a draft from last Tuesday and one from a minute ago both showed as a
 * bare "3:36:58 PM" — precise to the second and useless for the only question
 * being asked, which is whether these changes are recent enough to want. */
export function recoveryNotice(savedAtMs: number, nowMs: number): string {
  return `Recovered unsaved changes from ${formatRelativeTime(savedAtMs, nowMs)}`;
}
