// The product-neutral date/time format catalogue and renderer (SPEC-093's
// two-phase core move of SPEC-053's machinery — the shape of SPEC-079's
// core/drawingMl.ts move). The closed catalogue, the per-format renderer and
// the `.docx`/`.odt` name mappings live here so BOTH products can share one
// date machinery: the word processor's field evaluator (src/lib/fields.ts,
// which re-exports every symbol so its consumers and specs are untouched) and
// the presentation editor's header/footer resolver
// (src/lib/presentation/headerFooter.ts). Framework-free, no document model,
// no Vue — importable by either product (boundary.spec.ts).
//
// Determinism comes from the caller: nothing here reads the wall clock; the
// instant is always passed in (see core/clock.ts).

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
