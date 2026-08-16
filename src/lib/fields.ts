// The unified field evaluator (SPEC-053). A field is a run carrying a
// `FieldSpec` (see documentModel.ts) whose displayed value is *computed* from
// the document and the clock rather than stored as literal text. This module
// owns the closed date/time format catalogue, the per-format renderer, and
// the evaluator that turns a `FieldSpec` into the string a reader displays.
//
// Framework-free per CLAUDE.md: no Vue imports, unit-testable without mounting
// a component. Determinism comes from `clock.ts` — the evaluator never calls
// `Date.now()`/`new Date()` directly; the caller passes the instant read from
// `now()`.
import type { CaretPosition } from "./textEditing";
import { blockText, compareCaret, paragraphAt } from "./textEditing";
import { formatOrdinal } from "./core/lists";
import { documentBodyStats } from "./wordCount";
import { effectiveLanguage } from "./core/language";
import { bookmarkCaretIncludingHidden } from "./links";
import type { DocumentModel, FieldSpec, PageNumberFormat, RefFormat } from "./documentModel";

/** The text a cross-reference shows when its target bookmark is gone
 * (SPEC-053) — Word's exact wording, so the string round-trips as content
 * both applications recognise. */
export const REF_SOURCE_NOT_FOUND = "Error! Reference source not found.";

/** One entry of the closed date/time format catalogue (SPEC-053 "Date and
 * time formats"): the pattern stored in `FieldSpec.format`, the `.docx`
 * picture string it writes as, and the `.odt` number-style body written into
 * styles.xml. `isTime` splits the catalogue into `date`-field and `time`-field
 * defaults for the Date and time dialog. */
export interface DateTimeFormatEntry {
  /** The canonical pattern stored in the model (also the dialog's value). */
  format: string;
  /** Word's picture string, written as `DATE \@ "<picture>"`. */
  docxPicture: string;
  /** The `<number:date-style>`/`<number:time-style>` element name (`date` or
   * `time`) and its child body, written into the `.odt`'s styles.xml. */
  odtStyleTag: "date" | "time";
  odtStyleBody: string;
  /** True for the clock formats, so the dialog groups them under "Time". */
  isTime: boolean;
}

// The nine supported patterns, in the order the spec's table lists them.
// `odtStyleBody` is the children of `<number:date-style>`/`<number:time-style>`.
export const DATE_TIME_CATALOGUE: readonly DateTimeFormatEntry[] = [
  {
    format: "d MMMM yyyy",
    docxPicture: "d MMMM yyyy",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:day/><number:text> </number:text><number:month number:textual="true" number:style="long"/><number:text> </number:text><number:year number:style="long"/>',
    isTime: false,
  },
  {
    format: "d MMM yyyy",
    docxPicture: "d MMM yyyy",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:day/><number:text> </number:text><number:month number:textual="true"/><number:text> </number:text><number:year number:style="long"/>',
    isTime: false,
  },
  {
    format: "dd/MM/yyyy",
    docxPicture: "dd/MM/yyyy",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/>',
    isTime: false,
  },
  {
    format: "MM/dd/yyyy",
    docxPicture: "MM/dd/yyyy",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:month number:style="long"/><number:text>/</number:text><number:day number:style="long"/><number:text>/</number:text><number:year number:style="long"/>',
    isTime: false,
  },
  {
    format: "yyyy-MM-dd",
    docxPicture: "yyyy-MM-dd",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/>',
    isTime: false,
  },
  {
    format: "EEEE, d MMMM yyyy",
    docxPicture: "dddd, d MMMM yyyy",
    odtStyleTag: "date",
    odtStyleBody:
      '<number:day-of-week number:style="long"/><number:text>, </number:text><number:day/><number:text> </number:text><number:month number:textual="true" number:style="long"/><number:text> </number:text><number:year number:style="long"/>',
    isTime: false,
  },
  {
    format: "HH:mm",
    docxPicture: "HH:mm",
    odtStyleTag: "time",
    odtStyleBody:
      '<number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/>',
    isTime: true,
  },
  {
    format: "h:mm a",
    docxPicture: "h:mm am/pm",
    odtStyleTag: "time",
    odtStyleBody:
      '<number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text> </number:text><number:am-pm/>',
    isTime: true,
  },
  {
    format: "HH:mm:ss",
    docxPicture: "HH:mm:ss",
    odtStyleTag: "time",
    odtStyleBody:
      '<number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/>',
    isTime: true,
  },
];

/** The catalogue as a lookup on the canonical pattern. */
const CATALOGUE_BY_FORMAT = new Map(DATE_TIME_CATALOGUE.map((entry) => [entry.format, entry]));

/** The catalogue keyed by `.docx` picture string, for the reader mapping a
 * `DATE \@ "<picture>"` back to a canonical pattern. */
const CATALOGUE_BY_DOCX_PICTURE = new Map(DATE_TIME_CATALOGUE.map((entry) => [entry.docxPicture, entry]));

/** Whether a picture string is one this editor renders (SPEC-053): a picture
 * outside the catalogue is kept verbatim and the field displays its `cached`
 * value read-only, because implementing Word's picture-string language is a
 * spec of its own. */
export function isSupportedDateTimeFormat(format: string): boolean {
  return CATALOGUE_BY_FORMAT.has(format);
}

/** Maps a `.docx` `DATE`/`TIME` picture string to a canonical catalogue
 * pattern, or returns the picture verbatim when it is outside the catalogue
 * (kept as an unsupported picture). */
export function canonicalDateTimeFormatFromDocx(picture: string): string {
  return CATALOGUE_BY_DOCX_PICTURE.get(picture.trim())?.format ?? picture;
}

/** The catalogue entry for a canonical pattern, or `undefined` for an
 * unsupported picture. */
export function dateTimeFormatEntry(format: string): DateTimeFormatEntry | undefined {
  return CATALOGUE_BY_FORMAT.get(format);
}

/** The deterministic `number:date-style`/`number:time-style` name the `.odt`
 * writer gives a catalogue format, keyed by its catalogue index so the reader
 * recovers the exact pattern from the referenced style name without matching
 * the style body — a foreign producer's own date-style names simply don't
 * match and read as an unsupported picture (rendered from `cached`). */
export function odtDateStyleName(entry: DateTimeFormatEntry): string {
  return `FieldDT${DATE_TIME_CATALOGUE.indexOf(entry)}`;
}

/** Recovers a catalogue format from an `.odt` date/time field's referenced
 * `style:data-style-name`, or null when the name isn't one this editor wrote. */
export function formatFromOdtDateStyleName(name: string): string | null {
  const match = /^FieldDT(\d+)$/.exec(name);
  if (!match) {
    return null;
  }
  return DATE_TIME_CATALOGUE[Number(match[1])]?.format ?? null;
}

// --- date/time rendering -----------------------------------------------------

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** Renders `instant` under a supported catalogue pattern in `language`
 * (SPEC-035). Numeric components come straight off the `Date` (Western digits,
 * matching the spec's en-GB examples and locale-independent for the slash and
 * ISO patterns, whose ordering the pattern — not the locale — dictates); only
 * the month name, weekday name and AM/PM marker are localised through `Intl`.
 * Returns `null` for an unsupported pattern, so the caller falls back to the
 * field's `cached` value. */
export function renderDateTime(format: string, instant: Date, language: string): string | null {
  if (!CATALOGUE_BY_FORMAT.has(format)) {
    return null;
  }
  const d = instant.getDate();
  const monthLong = new Intl.DateTimeFormat(language, { month: "long" }).format(instant);
  const monthShort = new Intl.DateTimeFormat(language, { month: "short" }).format(instant);
  const weekdayLong = new Intl.DateTimeFormat(language, { weekday: "long" }).format(instant);
  const yyyy = String(instant.getFullYear()).padStart(4, "0");
  const MM = pad2(instant.getMonth() + 1);
  const dd = pad2(d);
  const hours24 = instant.getHours();
  const HH = pad2(hours24);
  const mm = pad2(instant.getMinutes());
  const ss = pad2(instant.getSeconds());
  const h = String(((hours24 + 11) % 12) + 1);
  // The AM/PM marker. ICU renders the English pair lowercase for some locales
  // (en-GB → "pm"), but the spec's catalogue shows the uppercase "PM"; so an
  // English-style "am"/"pm" is normalised to the uppercase pair, while a
  // genuinely localised marker (other scripts) is kept as ICU gives it.
  const dayPeriodPart = new Intl.DateTimeFormat(language, { hour: "numeric", hour12: true })
    .formatToParts(instant)
    .find((part) => part.type === "dayPeriod");
  const rawDayPeriod = dayPeriodPart?.value ?? "";
  const a = rawDayPeriod === "" || /^[ap]\.?m\.?$/i.test(rawDayPeriod)
    ? hours24 < 12
      ? "AM"
      : "PM"
    : rawDayPeriod;
  switch (format) {
    case "d MMMM yyyy":
      return `${d} ${monthLong} ${yyyy}`;
    case "d MMM yyyy":
      return `${d} ${monthShort} ${yyyy}`;
    case "dd/MM/yyyy":
      return `${dd}/${MM}/${yyyy}`;
    case "MM/dd/yyyy":
      return `${MM}/${dd}/${yyyy}`;
    case "yyyy-MM-dd":
      return `${yyyy}-${MM}-${dd}`;
    case "EEEE, d MMMM yyyy":
      return `${weekdayLong}, ${d} ${monthLong} ${yyyy}`;
    case "HH:mm":
      return `${HH}:${mm}`;
    case "h:mm a":
      return `${h}:${mm} ${a}`;
    case "HH:mm:ss":
      return `${HH}:${mm}:${ss}`;
    default:
      return null;
  }
}

// --- field evaluation --------------------------------------------------------

/** The context a field evaluation reads from. `model`, `now` and `fileName`
 * are always available; the pagination-dependent members are supplied by the
 * caller that knows layout (rendering, the update commands), and left absent
 * where the value should fall back to `cached`. */
export interface FieldContext {
  model: DocumentModel;
  /** The instant from `clock.now()` — never read the wall clock here. */
  now: Date;
  /** The document's file name (SPEC-053 `fileName`); "" when unsaved. */
  fileName: string;
  /** The field run's own position, for `ref` above/below and self page. */
  here?: CaretPosition | null;
  /** The page the field sits on (`page`, and `ref` page context). */
  pageNumber?: number;
  /** The document's total page count (`numpages`). */
  totalPages?: number;
  /** The page a bookmark target sits on (`pageRef`, `ref` show:page). */
  pageOfBookmark?: (name: string) => number | null;
}

/** A custom document property's value by name (SPEC-041), or "" when absent. */
function customPropertyValue(model: DocumentModel, name: string): string {
  return model.properties?.customProperties?.find((property) => property.name === name)?.value ?? "";
}

/** The paragraph a bookmark names, for a cross-reference's text — or null
 * when the target no longer exists. */
function bookmarkTargetText(model: DocumentModel, name: string): string | null {
  const caret = bookmarkCaretIncludingHidden(model, name);
  if (!caret) {
    return null;
  }
  const paragraph = paragraphAt(model, caret);
  return paragraph ? blockText(paragraph) : null;
}

/** The above/below position of a cross-reference target relative to the
 * reference itself (SPEC-053), using textEditing's document-order comparator. */
function relativePosition(model: DocumentModel, targetName: string, here: CaretPosition | null | undefined): string | null {
  if (!here) {
    return null;
  }
  const target = bookmarkCaretIncludingHidden(model, targetName);
  if (!target) {
    return null;
  }
  return compareCaret(target, here) <= 0 ? "above" : "below";
}

/** Evaluates a cross-reference (`ref`/`pageRef`) to its display string,
 * returning `REF_SOURCE_NOT_FOUND` when the target is gone (SPEC-053). */
function evaluateRef(field: Extract<FieldSpec, { kind: "ref" | "pageRef" }>, ctx: FieldContext): string {
  const exists = bookmarkCaretIncludingHidden(ctx.model, field.target) !== null;
  if (!exists) {
    return REF_SOURCE_NOT_FOUND;
  }
  const show: RefFormat = field.kind === "pageRef" ? "page" : field.show;
  if (show === "page") {
    const page = ctx.pageOfBookmark?.(field.target);
    return page != null ? String(page) : field.cached;
  }
  if (show === "position") {
    return relativePosition(ctx.model, field.target, ctx.here) ?? field.cached;
  }
  return bookmarkTargetText(ctx.model, field.target) ?? REF_SOURCE_NOT_FOUND;
}

/** Computes a field's live display string (SPEC-053). Fields whose value is
 * not recomputable from `ctx` (unsupported types, unsupported date pictures,
 * pagination values with no layout in `ctx`) fall back to their `cached`
 * result — the value the file was saved with, which is the truth for a field
 * this editor does not recalculate. Merge chips and `page`/`numpages` are not
 * handled here: chips resolve via `mergeChips.ts`, and page fields render
 * through `pageNumbering.ts`'s `fieldValueText`, which needs page context. */
export function evaluateField(field: FieldSpec, ctx: FieldContext): string {
  switch (field.kind) {
    case "page":
      return formatOrdinal(field.format, ctx.pageNumber ?? 1);
    case "numpages":
      return formatOrdinal(field.format, ctx.totalPages ?? 1);
    case "date":
    case "time":
      return renderDateTime(field.format, ctx.now, effectiveLanguage(undefined, ctx.model.defaultStyles.language)) ?? field.cached;
    case "author":
      return ctx.model.properties?.author ?? "";
    case "docTitle":
      return ctx.model.properties?.title ?? "";
    case "docSubject":
      return ctx.model.properties?.subject ?? "";
    case "docKeywords":
      return ctx.model.properties?.keywords ?? "";
    case "fileName":
      return field.path ? ctx.fileName : ctx.fileName.replace(/^.*[\\/]/, "");
    case "wordCount":
      return String(documentBodyStats(ctx.model).words);
    case "docProperty":
      return customPropertyValue(ctx.model, field.name);
    case "ref":
    case "pageRef":
      return evaluateRef(field, ctx);
    case "unsupported":
      return field.cached;
    default:
      // page/numpages/merge chips — not evaluated here (see doc comment).
      return "";
  }
}

/** The message shown for a field this editor cannot recompute (SPEC-053):
 * an unsupported field type, or a date/time field whose picture string is
 * outside the catalogue. Both are shown read-only with their saved value. */
export const UNSUPPORTED_FIELD_MESSAGE =
  "This field type isn't supported — its value is shown as it was saved";

/** Whether a field is shown read-only with `UNSUPPORTED_FIELD_MESSAGE`
 * (SPEC-053): an `unsupported` type, or a date/time field carrying a picture
 * string outside the catalogue. */
export function isReadOnlyField(field: FieldSpec): boolean {
  if (field.kind === "unsupported") {
    return true;
  }
  if (field.kind === "date" || field.kind === "time") {
    return !isSupportedDateTimeFormat(field.format);
  }
  return false;
}

/** The `.docx` numbering-switch fragment for a page-number format, shared by
 * the writer and any picture builder. */
export function pageNumberFormatOrdinal(format: PageNumberFormat, value: number): string {
  return formatOrdinal(format, value);
}
