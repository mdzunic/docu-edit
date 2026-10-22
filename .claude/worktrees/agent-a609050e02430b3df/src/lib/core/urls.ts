// URL safety and bare-URL detection (SPEC-025), moved down into core by
// SPEC-097 so the presentation product can sanitise slide hyperlinks without
// importing the text side. Phase one of the two-phase move documentModel.ts
// used for MediaPart: the definitions live here and src/lib/links.ts
// re-exports them, so every existing import path (LinkDialog.vue, useLinks.ts,
// links.spec.ts) keeps working unchanged. None of this touches a document
// model — it is neutral string logic.

// The schemes that execute script or smuggle content when navigated to.
// Rejected on entry (the link dialog), stripped on every other ingress path
// (file open, paste, autolink) and re-checked before any navigation.
export const DANGEROUS_SCHEMES: ReadonlySet<string> = new Set(["javascript", "data", "vbscript"]);

/** The URL's scheme, lowercased — parsed the forgiving way browsers do:
 * ASCII control characters and whitespace are stripped first, so
 * "  java\tscript:alert(1)" is still recognised as `javascript`. */
export function urlScheme(url: string): string | null {
  let cleaned = "";
  for (const char of url) {
    if (char.charCodeAt(0) > 0x20) {
      cleaned += char;
    }
  }
  const match = /^([a-z][a-z0-9+.-]*):/.exec(cleaned.toLowerCase());
  return match ? match[1] : null;
}

/** Whether the URL's scheme is one of the dangerous set above. */
export function isDangerousUrl(url: string): boolean {
  const scheme = urlScheme(url);
  return scheme !== null && DANGEROUS_SCHEMES.has(scheme);
}

/** The URL a link may actually carry: trimmed, or null when it must be
 * rejected (dialog entry, AC-14..16) or stripped (file open / paste,
 * AC-17). */
export function sanitizeHref(url: string): string | null {
  const trimmed = url.trim();
  if (trimmed.length === 0 || isDangerousUrl(trimmed)) {
    return null;
  }
  return trimmed;
}

// --- autolink detection (AC-5..7) -------------------------------------------

/** A bare URL found in a paragraph's text: the covered character span plus
 * the (possibly scheme-prefixed) href to link it to. */
export interface AutolinkMatch {
  start: number;
  end: number;
  href: string;
}

// Punctuation a sentence hangs on a URL's end ("see http://x.com.") — never
// part of the autolinked target, matching what word processors do.
const TRAILING_PUNCTUATION = /[.,;:!?)\]}>"']+$/;
const LEADING_WRAPPERS = /^[([{<"']+/;

/** Finds a bare URL ending at `endOffset` in `text` — the word between the
 * previous whitespace and `endOffset`, with surrounding punctuation peeled
 * off. Recognises http(s) URLs and `www.` shorthand (prefixed with https://
 * in the href). Dangerous schemes never match (AC-17's paste path). */
export function detectBareUrl(text: string, endOffset: number): AutolinkMatch | null {
  const upto = text.slice(0, Math.max(0, Math.min(endOffset, text.length)));
  let start = upto.length;
  while (start > 0 && !/\s/.test(upto[start - 1])) {
    start--;
  }
  let word = upto.slice(start);
  const lead = LEADING_WRAPPERS.exec(word);
  if (lead) {
    start += lead[0].length;
    word = word.slice(lead[0].length);
  }
  let end = upto.length;
  const trail = TRAILING_PUNCTUATION.exec(word);
  if (trail) {
    word = word.slice(0, word.length - trail[0].length);
    end -= trail[0].length;
  }
  if (word.length === 0) {
    return null;
  }
  let href: string | null = null;
  if (/^https?:\/\/\S+$/i.test(word)) {
    href = word;
  } else if (/^www\.[^\s.]+\.\S+$/i.test(word)) {
    href = `https://${word}`;
  }
  if (href === null || isDangerousUrl(href)) {
    return null;
  }
  return { start, end, href };
}
