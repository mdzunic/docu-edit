// Supported spell-check languages and language-tag helpers (SPEC-035). The
// document model treats language as a per-run override (`RunFormatting.lang`)
// falling back to the document default (`DefaultStyles.language`), itself
// falling back to `DEFAULT_LANGUAGE` — every piece of text always resolves to
// *some* language even though both model fields are optional (the same
// "absent means the product default" convention as measurementUnit/showRuler
// in documentModel.ts).

export interface LanguageOption {
  /** BCP-47-ish tag as stored in the model and written to both file formats
   * (`language-REGION`, e.g. "en-US"). */
  tag: string;
  /** Display name shown in the language selector and status bar. */
  label: string;
}

/** The document/product default when nothing else specifies a language —
 * also what a brand-new blank document starts with. */
export const DEFAULT_LANGUAGE = "en-US";

/** Languages with a bundled dictionary (SPEC-035: "at minimum, dictionaries
 * for English (US), English (UK), German, French, Spanish and Czech").
 * Order is the language selector's display order. */
export const SUPPORTED_LANGUAGES: LanguageOption[] = [
  { tag: "en-US", label: "English (US)" },
  { tag: "en-GB", label: "English (UK)" },
  { tag: "de-DE", label: "German" },
  { tag: "fr-FR", label: "French" },
  { tag: "es-ES", label: "Spanish" },
  { tag: "cs-CZ", label: "Czech" },
];

const SUPPORTED_TAGS = new Set(SUPPORTED_LANGUAGES.map((option) => option.tag));

/** Canonicalises a language tag's casing (language lowercase, region
 * uppercase) and drops any script/variant subtags beyond language-region —
 * the dictionaries this product ships only key on those two. A tag this
 * loses information from (e.g. "zh-Hant-TW") still normalises to *some*
 * usable tag ("zh") rather than being dropped entirely. */
export function normalizeLanguageTag(tag: string): string {
  const parts = tag.trim().split(/[-_]/).filter(Boolean);
  if (parts.length === 0) {
    return tag;
  }
  const language = parts[0].toLowerCase();
  const region = parts.slice(1).find((part) => /^[a-zA-Z]{2}$/.test(part));
  return region ? `${language}-${region.toUpperCase()}` : language;
}

/** Whether a bundled dictionary exists for a language tag (SPEC-035 AC-19):
 * text in an unavailable language is silently skipped rather than checked
 * or erroring. */
export function isDictionaryAvailable(tag: string | undefined): boolean {
  return tag !== undefined && SUPPORTED_TAGS.has(normalizeLanguageTag(tag));
}

/** Splits a tag into ODF's separate `fo:language`/`fo:country` attributes. */
export function splitLanguageTag(tag: string): { language: string; country?: string } {
  const normalized = normalizeLanguageTag(tag);
  const [language, country] = normalized.split("-");
  return country ? { language, country } : { language };
}

/** Composes ODF's `fo:language`/`fo:country` pair back into one model tag. */
export function composeLanguageTag(language: string, country: string | null | undefined): string {
  const lang = language.trim().toLowerCase();
  return country ? `${lang}-${country.trim().toUpperCase()}` : lang;
}

/** A language tag's display label: the bundled dictionary's name, else the
 * raw tag itself (AC: a document declaring a language with no installed
 * dictionary, e.g. "pt-BR", still shows something meaningful in the status
 * bar rather than a blank or generic label). */
export function languageLabel(tag: string | undefined): string {
  if (!tag) {
    return languageLabel(DEFAULT_LANGUAGE);
  }
  const normalized = normalizeLanguageTag(tag);
  return SUPPORTED_LANGUAGES.find((option) => option.tag === normalized)?.label ?? tag;
}

/** The effective language for a run: its own tag, else the document
 * default, else the product default — the one fallback chain every
 * consumer (spell check, the rendered `lang` attribute, the status bar)
 * shares. */
export function effectiveLanguage(
  runLang: string | undefined,
  documentLanguage: string | undefined,
): string {
  return runLang ?? documentLanguage ?? DEFAULT_LANGUAGE;
}

/** Typed pass-through for a LanguageOption[] arriving from a `.vue`
 * component's untyped Array prop (same espree constraint as
 * documentModel.ts's asCharacterFormatKey). */
export function asLanguageOptions(value: unknown): LanguageOption[] {
  return value as LanguageOption[];
}

/** The decimal separator a language writes numbers with (SPEC-054 AC-19): a
 * decimal tab stop aligns text on this character. Derived from the platform's
 * Intl data for the effective language, falling back to "." when Intl is
 * unavailable or the tag does not resolve — the same character English uses,
 * and the one the decimal-stop layout treats as its primary separator before
 * the AC-20 "first separator character" fallback. */
export function decimalSeparatorFor(tag: string | undefined): string {
  const normalized = normalizeLanguageTag(tag ?? DEFAULT_LANGUAGE);
  try {
    const parts = new Intl.NumberFormat(normalized).formatToParts(1.1);
    const decimal = parts.find((part) => part.type === "decimal");
    return decimal?.value ?? ".";
  } catch {
    return ".";
  }
}

/** ODF's "no linguistic content" language code (ISO 639-2's `zxx`) — the
 * convention ODF producers (including LibreOffice's own "None (Do not check
 * spelling)" language entry) use to mark text as not to be spell-checked,
 * since ODF has no separate boolean flag the way `w:noProof` is in OOXML.
 * The .odt reader/writer maps this to/from `RunFormatting.noProof` so the
 * rest of the app only ever deals with one "do not check spelling" concept
 * regardless of source format; a run that is both `noProof` and has a real
 * language set loses that language on an .odt round trip specifically (ODF
 * itself has no room to keep both), which mirrors the limitation real ODF
 * consumers have. */
export const ODF_NO_LINGUISTIC_CONTENT = "zxx";
