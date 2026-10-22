// The library's open progress for a workbook (SPEC-137 "Opening a large
// workbook shows progress"): from the file choice until navigation the
// library shows `spreadsheet-open-progress` (role="status") reading
// "Opening {title}… {n}%", n the share of the workbook's row chunks that has
// arrived from the parse worker (spreadsheetParseClient's progress, handed
// through the open funnel's onSpreadsheetProgress). A file named like a
// workbook package shows it from the moment it is chosen (0%); any other
// file only once the funnel, which routes by content, reports it is a
// workbook. HomeView wires begin() into the picker's open and end() into its
// settling; LibraryDialogs renders `text`.
import { computed, ref, type ComputedRef } from "vue";
import { titleFromFileName, type OpenFileOptions } from "../../lib/openDocument";
import { SPREADSHEET_PACKAGE_EXTENSIONS } from "../../lib/spreadsheet/fileTypes";

/** The status line for `title` at `fraction` (0..1) of its rows. */
export function spreadsheetOpenProgressText(title: string, fraction: number): string {
  const percent = Math.min(100, Math.max(0, Math.floor(fraction * 100)));
  return `Opening ${title}… ${percent}%`;
}

function namedLikeWorkbook(fileName: string): boolean {
  const dot = fileName.lastIndexOf(".");
  return dot >= 0 && SPREADSHEET_PACKAGE_EXTENSIONS.includes(fileName.slice(dot + 1).toLowerCase());
}

export interface SpreadsheetOpenProgress {
  /** The status text — "" while no workbook is opening. */
  text: ComputedRef<string>;
  /** A file was chosen: the funnel options that report its progress. */
  begin(file: File): OpenFileOptions;
  /** The open settled (navigated away, or failed). */
  end(): void;
}

export function useSpreadsheetOpenProgress(): SpreadsheetOpenProgress {
  const title = ref<string | null>(null);
  const fraction = ref(0);
  /** Which begin() the callbacks belong to: a stale open never moves a newer one. */
  let generation = 0;

  function begin(file: File): OpenFileOptions {
    generation += 1;
    const mine = generation;
    const name = titleFromFileName(file.name);
    title.value = namedLikeWorkbook(file.name) ? name : null;
    fraction.value = 0;
    return {
      onSpreadsheetProgress(value: number) {
        if (mine !== generation) {
          return;
        }
        title.value = name;
        fraction.value = Math.max(fraction.value, value);
      },
    };
  }

  function end(): void {
    generation += 1;
    title.value = null;
    fraction.value = 0;
  }

  return {
    text: computed(() => (title.value === null ? "" : spreadsheetOpenProgressText(title.value, fraction.value))),
    begin,
    end,
  };
}
