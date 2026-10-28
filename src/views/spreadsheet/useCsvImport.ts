// SPEC-126's import bar and "Text import options" dialog for a csv/tsv
// workbook (CsvImportBar.vue, CsvImportDialog.vue): the bar says how the file
// was read — "Opened as semicolon-separated, Windows-1252, numbers and dates
// read as German · Change…" (the locale part only when it is not the user's
// own) — and shows on every csv/tsv workbook view until dismissed. Change…
// opens the dialog, preselected from `model.csv`, whose Separator, Character
// set and "Numbers and dates" selects drive a live preview of the first
// CSV_PREVIEW_RECORDS records, re-parsed from the retained source bytes on
// every change.
//
// The dialog re-reads the RETAINED SOURCE BYTES with the chosen options.
// While the workbook is unedited (no dirty sheet, no structureDirty) its
// primary action is "Apply": reimportCsv replaces the first sheet's rows in
// place — same workbook id, same sheet id, still unedited (so the download
// is still the source bytes) — the undo stack is reset and the record
// written through. Once edited the primary action is "Open as new
// workbook": a new workbook titled "<title> (re-imported)" is read from the
// same bytes, stored, and navigated to, leaving the edited one untouched.
import { computed, ref, watch, type ShallowRef } from "vue";
import type { Router } from "vue-router";
import { languageLabel } from "../../lib/core/language";
import { editorPathFor } from "../../lib/documentKinds";
import { storeSpreadsheet, storeSpreadsheetEdit } from "../../lib/documentStore";
import { csvSourceBytes, parseCsvText, readCsvBytes, reimportCsv, type CsvReadOptions } from "../../lib/spreadsheet/csv/parse";
import { decodeCsvBytes, sepLineOf } from "../../lib/spreadsheet/csv/sniff";
import { isCsvFormat, type CsvDelimiter, type CsvDialect, type CsvEncoding, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { isSpreadsheetLocaleTag, SPREADSHEET_LOCALE_TAGS, type SpreadsheetLocaleTag } from "../../lib/spreadsheet/numberFormat/locale";
import { useSpreadsheetLocale } from "./useSpreadsheetLocale";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

/** How many records the dialog's preview shows. */
export const CSV_PREVIEW_RECORDS = 20;

/** The dialog's Separator choices, in order. */
export const CSV_SEPARATOR_OPTIONS: readonly { value: string; label: string; delimiter: CsvDelimiter; word: string }[] = [
  { value: "comma", label: "Comma", delimiter: ",", word: "comma-separated" },
  { value: "semicolon", label: "Semicolon", delimiter: ";", word: "semicolon-separated" },
  { value: "tab", label: "Tab", delimiter: "\t", word: "tab-separated" },
  { value: "pipe", label: "Pipe", delimiter: "|", word: "pipe-separated" },
];

/** The dialog's Character set choices, in order. */
export const CSV_ENCODING_OPTIONS: readonly { value: CsvEncoding; label: string; word: string }[] = [
  { value: "utf-8", label: "Unicode (UTF-8)", word: "UTF-8" },
  { value: "utf-16le", label: "Unicode (UTF-16 LE)", word: "UTF-16" },
  { value: "utf-16be", label: "Unicode (UTF-16 BE)", word: "UTF-16" },
  { value: "windows-1252", label: "Western (Windows-1252)", word: "Windows-1252" },
];

/** The dialog's "Numbers and dates" choices: the six locales by language. */
export const CSV_LOCALE_OPTIONS: readonly { value: SpreadsheetLocaleTag; label: string }[] = SPREADSHEET_LOCALE_TAGS.map((tag) => ({
  value: tag,
  label: languageLabel(tag),
}));

/** The bar's sentence before " · Change…": "Opened as semicolon-separated,
 * Windows-1252, numbers and dates read as German". */
export function csvImportBarText(dialect: Pick<CsvDialect, "delimiter" | "encoding" | "localeTag">, userLocaleTag: SpreadsheetLocaleTag): string {
  const separator = CSV_SEPARATOR_OPTIONS.find((option) => option.delimiter === dialect.delimiter)?.word ?? "comma-separated";
  const encoding = CSV_ENCODING_OPTIONS.find((option) => option.value === dialect.encoding)?.word ?? "UTF-8";
  const locale = dialect.localeTag === userLocaleTag ? "" : `, numbers and dates read as ${languageLabel(dialect.localeTag)}`;
  return `Opened as ${separator}, ${encoding}${locale}`;
}

/** The title "Open as new workbook" stores the re-read under. */
export function reimportedTitle(title: string): string {
  return `${title} (re-imported)`;
}

/** The first CSV_PREVIEW_RECORDS records of `bytes` read with the chosen
 * separator and character set (a `sep=` line is not a record), or the
 * reason they cannot be read. */
export function csvImportPreview(bytes: Uint8Array, delimiter: CsvDelimiter, encoding: CsvEncoding): { records: string[][]; error: string } {
  try {
    const { text } = decodeCsvBytes(bytes, encoding);
    const sep = sepLineOf(text);
    const records: string[][] = [];
    parseCsvText(sep ? text.slice(sep.length) : text, delimiter, (fields) => {
      records.push(fields);
      return records.length >= CSV_PREVIEW_RECORDS ? "stop" : undefined;
    });
    return { records, error: "" };
  } catch (error) {
    return { records: [], error: error instanceof Error ? error.message : String(error) };
  }
}

export interface CsvImportDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    /** Shows a re-read model in place of the one on show. */
    replace(model: SpreadsheetModel): void;
  };
  undo: Pick<SpreadsheetUndo, "reset">;
  router: Pick<Router, "push">;
  /** SPEC-138: true while the view is read-only (another tab holds the
   * workbook, or a version is previewed in its place) — "Apply" then does
   * nothing. Absent: never. */
  locked?: () => boolean;
}

export function useCsvImport(deps: CsvImportDeps) {
  const { tag: userLocaleTag } = useSpreadsheetLocale();
  const dismissed = ref(false);
  const dialogOpen = ref(false);
  /** The dialog's selections: a CSV_SEPARATOR_OPTIONS value, an encoding and
   * a locale tag. */
  const separator = ref("comma");
  const encoding = ref<string>("utf-8");
  const localeTag = ref<string>("en-US");
  const error = ref("");

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;

  // Another workbook on show (a re-import opened as new, a route change):
  // its bar shows again and no dialog carries over.
  watch(
    () => deps.book.workbook.value?.id,
    () => {
      dismissed.value = false;
      dialogOpen.value = false;
    },
  );

  const dialect = computed((): CsvDialect | null => {
    const current = deps.book.workbook.value;
    return current && isCsvFormat(current.format) ? (current.csv ?? null) : null;
  });

  /** The bar's sentence; "" when no bar shows. */
  const barText = computed(() => (dialect.value && !dismissed.value ? csvImportBarText(dialect.value, userLocaleTag.value) : ""));

  /** Whether the workbook was edited (SPEC-121's flags): the dialog's
   * primary action is then "Open as new workbook". */
  const edited = computed(() => {
    const current = deps.book.workbook.value;
    return current !== null && (current.structureDirty === true || current.sheets.some((sheet) => sheet.dirty));
  });

  const chosen = computed(() => ({
    delimiter: CSV_SEPARATOR_OPTIONS.find((option) => option.value === separator.value)?.delimiter ?? ",",
    encoding: CSV_ENCODING_OPTIONS.find((option) => option.value === encoding.value)?.value ?? "utf-8",
    localeTag: isSpreadsheetLocaleTag(localeTag.value) ? localeTag.value : userLocaleTag.value,
  }));

  /** The retained bytes, decoded from the record once per opening. */
  const source = computed(() => {
    const current = deps.book.workbook.value;
    return dialogOpen.value && current ? csvSourceBytes(current) : null;
  });

  const preview = computed(() => (source.value ? csvImportPreview(source.value, chosen.value.delimiter, chosen.value.encoding) : { records: [], error: "" }));

  function dismiss(): void {
    dismissed.value = true;
  }

  function openDialog(): void {
    const current = dialect.value;
    if (!current) {
      return;
    }
    separator.value = CSV_SEPARATOR_OPTIONS.find((option) => option.delimiter === current.delimiter)?.value ?? "comma";
    encoding.value = current.encoding;
    localeTag.value = current.localeTag;
    error.value = "";
    dialogOpen.value = true;
  }

  function closeDialog(): void {
    dialogOpen.value = false;
  }

  function readOptions(): CsvReadOptions {
    return { userLocaleTag: userLocaleTag.value, ...chosen.value };
  }

  /** "Apply" (unedited only): the first sheet re-read in place. */
  function apply(): void {
    const current = model();
    if (!current || !isCsvFormat(current.format) || edited.value || deps.locked?.()) {
      return;
    }
    let next: SpreadsheetModel;
    try {
      next = reimportCsv(current, readOptions());
    } catch (failure) {
      error.value = failure instanceof Error ? failure.message : String(failure);
      return;
    }
    dialogOpen.value = false;
    deps.book.replace(next);
    deps.undo.reset();
    void storeSpreadsheetEdit(next);
  }

  /** "Open as new workbook": the same bytes, re-read into a new workbook. */
  function openAsNew(): void {
    const current = model();
    if (!current || !isCsvFormat(current.format)) {
      return;
    }
    let fresh: SpreadsheetModel;
    try {
      fresh = readCsvBytes(csvSourceBytes(current), current.format, reimportedTitle(current.title), readOptions());
    } catch (failure) {
      error.value = failure instanceof Error ? failure.message : String(failure);
      return;
    }
    dialogOpen.value = false;
    storeSpreadsheet(fresh);
    void deps.router.push(editorPathFor("spreadsheet", fresh.id));
  }

  return {
    barText,
    dismiss,
    dialogOpen,
    openDialog,
    closeDialog,
    separator,
    encoding,
    localeTag,
    preview,
    edited,
    error,
    apply,
    openAsNew,
  };
}

export type CsvImport = ReturnType<typeof useCsvImport>;

/** Narrows a component's runtime-typed `csvImport` prop (the .vue parser
 * bars casts). */
export function asCsvImport(value: unknown): CsvImport {
  return value as CsvImport;
}
