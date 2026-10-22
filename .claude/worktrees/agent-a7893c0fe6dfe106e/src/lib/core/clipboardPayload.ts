// The product-neutral half of the app's private clipboard flavours (SPEC-016,
// hoisted by SPEC-099): every internal payload rides `ClipboardEvent.
// clipboardData` under a private MIME type as JSON labelled `{ mimeType,
// version, … }`, and every reader treats the clipboard as an untrusted
// boundary — anything can write any string under any key — so it checks the
// label and then rebuilds the body field by field from a whitelist. The text
// side's `lib/clipboard.ts` (DOC_MODEL_MIME) and the presentation side's
// `lib/presentation/slideClipboard.ts` (SLIDES_MIME) both build on this
// envelope and its guards.

const SAFE_COLOR = /^#[0-9a-f]{3,8}$/i;
const SAFE_COLOR_NAME = /^[a-zA-Z]+$/;

/** A colour string safe to splice into a model: `#rgb`…`#rrggbbaa` hex or a
 * bare CSS colour keyword. */
export function isSafeColor(value: unknown): value is string {
  return typeof value === "string" && (SAFE_COLOR.test(value) || SAFE_COLOR_NAME.test(value));
}

/** A real, finite number — never NaN/Infinity, never a numeric string. */
export function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** Serialises a payload body under the `{ mimeType, version }` label — the
 * label first, then the body's own fields in their order. */
export function buildClipboardEnvelope(mimeType: string, version: number, body: Record<string, unknown>): string {
  return JSON.stringify({ mimeType, version, ...body });
}

/** Parses a raw clipboard string and checks its label: the parsed object when
 * it is a JSON object carrying exactly this `mimeType` and `version`, else
 * null. The caller then rebuilds the body from a whitelist — the returned
 * record is never spliced into a model as-is. */
export function parseClipboardEnvelope(raw: string, mimeType: string, version: number): Record<string, unknown> | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    return null;
  }
  const payload = parsed as Record<string, unknown>;
  if (payload.mimeType !== mimeType || payload.version !== version) {
    return null;
  }
  return payload;
}
