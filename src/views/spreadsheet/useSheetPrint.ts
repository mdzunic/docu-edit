// SPEC-136's print preview and browser printing for the workbook view.
// Ctrl/Cmd+P and File ▸ "Print…" open the preview: SheetPrintPages in the
// grid area beside the reused PrintPreviewDialog, whose options slot holds
// the scope ("Active sheet" | "Selection" | "Entire workbook") and whose
// actions slot holds "Page setup…" and "Download PDF". The job is print.ts's
// paginateWorkbook of the scope; the page range parses through
// core/pageRange.ts's parseNumberRange(input, total, "page"); an empty job
// says "This sheet has nothing to print." and cannot print. Print marks
// `body.is-sheet-printing`, injects the job's `@page` rules
// (sheetPrintPageRuleCss), mounts the chosen pages off-screen at 1:1 and
// calls window.print() — App.vue's print CSS shows only those pages, one per
// sheet of paper; a job over PRINT_PAGE_LIMIT pages is refused. Closing the
// preview turns page-break lines on for its sheet (usePageSetup). Nothing
// here writes the model.
import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from "vue";
import { parseNumberRange } from "../../lib/core/pageRange";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import {
  NOTHING_TO_PRINT_MESSAGE,
  paginateWorkbook,
  PRINT_LIMIT_MESSAGE,
  PRINT_PAGE_LIMIT,
  sheetPrintPageRuleCss,
  type PrintScope,
  type SheetPrintJob,
  type SheetPrintPage,
} from "../../lib/spreadsheet/print";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { defaultPrintFont, layoutPrintPage, printSheetContext, type PrintPageLayout, type PrintSheetContext } from "./sheetPrintLayout";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { PageSetup } from "./usePageSetup";
import type { SheetPdfExport } from "./useSheetPdfExport";
import type { SheetSelectionController } from "./useSheetSelection";
import type { asWorkbookBook } from "./useWorkbookLoad";

export const PRINT_SCOPES: readonly { value: PrintScope; label: string }[] = [
  { value: "sheet", label: "Active sheet" },
  { value: "selection", label: "Selection" },
  { value: "workbook", label: "Entire workbook" },
];

/** A page of the job with its place in it ("Page N of M"). */
export interface PlacedPage {
  page: SheetPrintPage;
  index: number;
}

export interface SheetPrintDeps {
  book: ReturnType<typeof asWorkbookBook>;
  selection: Pick<SheetSelectionController, "selection">;
  pageSetup: Pick<PageSetup, "pageBreaks" | "printPreferences" | "dialog">;
  locale: () => SpreadsheetLocale;
}

/** Typed pass-throughs for SheetPrintPages' runtime props (espree). */
export function asPlacedPages(value: unknown): PlacedPage[] {
  return value as PlacedPage[];
}
export function asLayoutFor(value: unknown): (placed: PlacedPage, total: number) => PrintPageLayout {
  return value as (placed: PlacedPage, total: number) => PrintPageLayout;
}

export function useSheetPrint(deps: SheetPrintDeps) {
  // Each page laid out once per model state (layouts are pure functions of it).
  let layoutModel: SpreadsheetModel | null = null;
  let layoutDate = new Date();
  const contexts = new Map<string, PrintSheetContext>();
  const layouts = new Map<string, PrintPageLayout>();
  function layoutFor(placed: PlacedPage, total: number): PrintPageLayout {
    const model = deps.book.workbook.value!;
    if (model !== layoutModel || contexts.size > 64) {
      layoutModel = model;
      layoutDate = new Date();
      contexts.clear();
      layouts.clear();
    }
    const key = `${placed.page.sheetId}|${placed.index}|${total}|${placed.page.number}|${placed.page.rows[0]}|${placed.page.cols[0]}|${placed.page.scale}`;
    let layout = layouts.get(key);
    if (layout === undefined) {
      let ctx = contexts.get(placed.page.sheetId);
      const sheet = model.sheets.find((each) => each.id === placed.page.sheetId);
      if (ctx === undefined || ctx.sheet !== sheet) {
        ctx = printSheetContext(model, sheet!, deps.locale(), layoutDate);
        contexts.set(placed.page.sheetId, ctx);
      }
      layout = layoutPrintPage(ctx, placed.page, total);
      layouts.set(key, layout);
    }
    return layout;
  }
  /** The default font family pages print in. */
  const fontFamily = computed(() => {
    const model = deps.book.workbook.value;
    return model ? (defaultPrintFont(model) ?? "") : "";
  });

  const isOpen = ref(false);
  const scope = ref<PrintScope>("sheet");
  const rangeMode = ref<"all" | "custom">("all");
  const rangeText = ref("");
  const notice = ref("");
  const printing = ref(false);
  /** The pages mounted off-screen for the browser's print, and their `page:` names. */
  const printPages = shallowRef<PlacedPage[]>([]);
  const printPageNames = shallowRef<string[]>([]);
  // The preview's body class (App.vue's print CSS drops its fit-to-area scale).
  watch(isOpen, (now) => {
    if (typeof document !== "undefined") {
      document.body.classList.toggle("is-sheet-print-preview", now);
    }
  });
  onScopeDispose(() => {
    if (typeof document !== "undefined") {
      document.body.classList.remove("is-sheet-print-preview");
    }
  });

  /** The job a scope prints, from the live model (never written). */
  function jobFor(which: PrintScope): SheetPrintJob {
    const model = deps.book.workbook.value;
    if (!model) {
      return { pages: [] };
    }
    return paginateWorkbook(model, which, { selection: deps.selection.selection.value.ranges, preferences: deps.pageSetup.printPreferences.value });
  }

  const job = computed((): SheetPrintJob => (isOpen.value ? jobFor(scope.value) : { pages: [] }));
  const pageCount = computed(() => job.value.pages.length);
  const range = computed(() => (rangeMode.value === "custom" ? parseNumberRange(rangeText.value, pageCount.value, "page") : null));
  const rangeError = computed(() => range.value?.error ?? "");
  /** The pages the range chooses, each with its place in the job. */
  const chosen = computed((): PlacedPage[] => {
    const all = job.value.pages.map((page, index) => ({ page, index }));
    const numbers = range.value?.pages;
    return numbers === undefined ? all : numbers.map((n) => all[n - 1]!).filter((each) => each !== undefined);
  });
  const canPrint = computed(() => pageCount.value > 0 && rangeError.value === "" && !printing.value);

  function open(): void {
    if (deps.book.workbook.value === null) {
      return;
    }
    notice.value = "";
    isOpen.value = true;
  }

  function close(): void {
    if (!isOpen.value) {
      return;
    }
    isOpen.value = false;
    notice.value = "";
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind === "grid") {
      deps.pageSetup.pageBreaks.showFor(sheet.id);
    }
  }

  function setScope(next: string): void {
    scope.value = PRINT_SCOPES.some((each) => each.value === next) ? (next as PrintScope) : "sheet";
  }

  async function print(): Promise<void> {
    if (!canPrint.value) {
      return;
    }
    const pages = chosen.value;
    if (pages.length > PRINT_PAGE_LIMIT) {
      notice.value = PRINT_LIMIT_MESSAGE;
      return;
    }
    notice.value = "";
    const rules = sheetPrintPageRuleCss({ pages: pages.map((each) => each.page) });
    const style = document.createElement("style");
    style.dataset.sheetPrint = "";
    style.textContent = rules.css;
    printing.value = true;
    try {
      printPageNames.value = rules.pageNames;
      printPages.value = pages;
      await nextTick();
      document.head.append(style);
      document.body.classList.add("is-sheet-printing");
      await document.fonts?.ready;
      window.print();
    } finally {
      document.body.classList.remove("is-sheet-printing");
      style.remove();
      printPages.value = [];
      printing.value = false;
    }
  }

  /** Page setup… from the preview: the dialog over it. */
  function openPageSetup(): void {
    deps.pageSetup.dialog.show();
  }

  // Ctrl/Cmd+P anywhere in the view (not while another dialog is up).
  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id !== "sheet.print") {
      return;
    }
    event.preventDefault();
    const other = document.querySelector('[role="dialog"]:not([data-testid="print-preview"])');
    if (other === null) {
      open();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
    onScopeDispose(() => window.removeEventListener("keydown", onKeydown, true));
  }

  return {
    layoutFor,
    fontFamily,
    isOpen,
    scope,
    setScope,
    rangeMode,
    rangeText,
    rangeError,
    job,
    jobFor,
    pageCount,
    chosen,
    canPrint,
    notice,
    printing,
    printPages,
    printPageNames,
    emptyMessage: NOTHING_TO_PRINT_MESSAGE,
    open,
    close,
    print,
    openPageSetup,
  };
}

export type SheetPrint = ReturnType<typeof useSheetPrint>;

/** What the view hands the grid and its surfaces (SPEC-136). */
export interface SheetPrinting {
  pageSetup: PageSetup;
  print: SheetPrint;
  pdf: SheetPdfExport;
}

/** Typed pass-through for a `.vue` runtime prop (espree). */
export function asSheetPrinting(value: unknown): SheetPrinting {
  return value as SheetPrinting;
}
