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
