// The assembly workspace's PDF output (SPEC-110): the page axis SPEC-109's
// single sheet lacked. The tray is cut into pages by media/imagesToPdf's
// planPdfPages — one per page, n-up, or every item on one sheet — under the
// ONE sheet setup the inspector edits (no per-page overrides); the preview
// shows the current page, which the strip and PageUp/PageDown move through.
//
// Before anything is built it states what will happen: per tray item whether
// the file is embedded as its original JPEG or re-encoded (and why), the
// estimate computed from those decisions with the lowest-resolution
// placement named (and a warning when it prints below the sheet's dpi), and
// the refusals — MAX_PDF_ASSEMBLY_PAGES, MAX_PDF_ASSEMBLY_BYTES, page numbers
// that would not fit the bottom margin — before any work starts.
//
// Building is usePdfBuild's: a page at a time, cancellable at every page
// boundary, the output parsed back before anything is kept.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import { getMedia } from "../../lib/documentStore";
import type { MediaModel } from "../../lib/media/model";
import { currentImagePart, savedStack } from "../../lib/media/editModel";
import {
  encodingLabel,
  encodingReason,
  estimateImagePdf,
  itemsPerPage,
  N_UP_CHOICES,
  PAGE_MODES,
  pageNumbersRefusal,
  pdfLimitRefusal,
  planPdfPages,
  sourceFacts,
  type ImagePdfItem,
  type PageMode,
  type PdfAssemblyPage,
  type PdfAssemblyPlan,
  type SourceFacts,
} from "../../lib/media/imagesToPdf";
import type { AssemblyItem, AssemblyPlan, PlacedItem } from "../../lib/core/assemblyLayout";
import type { MediaPart } from "../../lib/core/media";
import { backgroundRefusal } from "../../lib/core/sheetLayout";
import { pageCountLabel, pdfDpiWarning, pdfEstimateLine } from "../../lib/assembly";
import { DEFAULT_ASSEMBLY_TITLE, type AssemblyOutputSettings } from "../../lib/assemblyPersist";
import type { SheetSetupController } from "../../composables/useSheetSetup";
import type { TrayEntry } from "../assembly/useAssemblyTray";
import { usePdfBuild, type PdfAssemblySteps, type PdfBuildController } from "./usePdfBuild";

export { asPdfAssemblySteps, BROWSER_PDF_ASSEMBLY_STEPS, type PdfAssemblySteps } from "./usePdfBuild";

/** Typed pass-through for pages read off a `.vue` prop (runtime-typed
 * `Object`; espree bars the annotation in `.vue` scripts). */
export function asPdfPages(value: unknown): PdfAssemblyPage[] {
  return Array.isArray(value) ? (value as PdfAssemblyPage[]) : [];
}

export interface PdfAssemblyDeps {
  entries: Ref<TrayEntry[]>;
  items: ComputedRef<AssemblyItem[]>;
  /** SPEC-109's plan: its layout is single-page's; its gap, fit, align and
   * cell colour every page's. */
  plan: ComputedRef<AssemblyPlan>;
  sheet: SheetSetupController;
  output: AssemblyOutputSettings;
  /** The tray's selection: selecting an item shows its page. */
  selectedKey?: Ref<string>;
  steps: PdfAssemblySteps;
}

export interface PdfAssemblyController extends PdfBuildController {
  plan: ComputedRef<PdfAssemblyPlan>;
  pages: ComputedRef<PdfAssemblyPage[]>;
  pageCount: ComputedRef<number>;
  /** "3 pages". */
  pageCountLabel: ComputedRef<string>;
  /** The page the preview shows (0-based). */
  current: ComputedRef<number>;
  currentPlaced: ComputedRef<PlacedItem[]>;
  /** The number the preview draws on the current page; 0 when off. */
  previewNumber: ComputedRef<number>;
  goTo: (index: number) => void;
  /** PageUp / PageDown: the previous / next page. */
  onPageKey: (event: KeyboardEvent) => void;
  /** Per tray index: "original JPEG", "re-encoded: cropped", … */
  encodings: ComputedRef<string[]>;
  estimate: ComputedRef<string>;
  /** "" unless a placement prints below the sheet's dpi. */
  dpiWarning: ComputedRef<string>;
  /** "" unless the sheet asked for a transparency a PDF cannot hold. */
  backgroundNotice: ComputedRef<string>;
  refusal: ComputedRef<string>;
  /** Why the last page control input was rejected ("" when accepted). */
  error: Ref<string>;
  canProduce: ComputedRef<boolean>;
  setPageMode: (value: string) => boolean;
  setNUp: (value: string) => boolean;
  setPageNumbers: (on: boolean) => void;
  setAuthor: (value: string) => void;
  setSubject: (value: string) => void;
}

/** A tray entry's current image as the PDF reads it. */
interface PdfSource {
  model: MediaModel;
  part: MediaPart;
  facts: SourceFacts;
}

export function usePdfAssembly(deps: PdfAssemblyDeps): PdfAssemblyController {
  const { sheet, output } = deps;
  const error = ref("");
  const page = ref(0);

  const plan = computed<PdfAssemblyPlan>(() => ({ ...deps.plan.value, pageMode: output.pageMode, nUp: output.nUp }));
  const pages = computed(() => {
    const geometry = sheet.geometry.value;
    return geometry ? planPdfPages(deps.items.value, plan.value, geometry, sheet.setup.value.dpi) : [];
  });
  const pageCount = computed(() => pages.value.length);
  const current = computed(() => Math.min(page.value, Math.max(0, pageCount.value - 1)));
  const currentPlaced = computed(() => pages.value[current.value]?.placed ?? []);

  function goTo(index: number): void {
    page.value = Math.min(Math.max(0, Math.round(index)), Math.max(0, pageCount.value - 1));
  }

  function onPageKey(event: KeyboardEvent): void {
    if ((event.key !== "PageUp" && event.key !== "PageDown") || event.altKey || event.ctrlKey || event.metaKey) {
      return;
    }
    event.preventDefault();
    goTo(current.value + (event.key === "PageDown" ? 1 : -1));
  }

  // Selecting an item in the tray shows the page that carries it.
  if (deps.selectedKey) {
    watch(deps.selectedKey, (key) => {
      const index = deps.entries.value.findIndex((entry) => entry.key === key);
      const at = pages.value.findIndex((candidate) => candidate.itemIndices.includes(index));
      if (at >= 0) {
        goTo(at);
      }
    });
  }

  // What each part is, read once per part (a part's bytes never change).
  const factsByMediaId = new Map<string, SourceFacts | null>();
  function factsOf(part: MediaPart): SourceFacts | null {
    if (!factsByMediaId.has(part.mediaId)) {
      let facts: SourceFacts | null = null;
      try {
        facts = sourceFacts(part);
      } catch {
        facts = null;
      }
      factsByMediaId.set(part.mediaId, facts);
    }
    return factsByMediaId.get(part.mediaId) ?? null;
  }

  const sources = computed<(PdfSource | null)[]>(() =>
    deps.entries.value.map((entry) => {
      const model = getMedia(entry.id);
      const part = model ? currentImagePart(model) : undefined;
      const facts = part ? factsOf(part) : null;
      return model && part && facts ? { model, part, facts } : null;
    }),
  );
  const pdfItems = computed<ImagePdfItem[]>(() =>
    deps.entries.value.map((entry, index) => ({
      id: entry.id,
      title: entry.title,
      widthPx: entry.widthPx,
      heightPx: entry.heightPx,
      // The normalised stack: what currentImagePart decides by.
      edits: sources.value[index] ? savedStack(sources.value[index].model) : undefined,
    })),
  );
  const titles = computed(() => deps.entries.value.map((entry) => entry.title));

  const encodings = computed(() => {
    const labels = deps.entries.value.map(() => "");
    for (const placed of pages.value.flatMap((candidate) => candidate.placed)) {
      const source = sources.value[placed.index];
      const item = pdfItems.value[placed.index];
      // A placement clipped wholly off the page embeds nothing to describe.
      if (source && item && placed.widthCm > 0 && placed.heightCm > 0) {
        labels[placed.index] = encodingLabel(encodingReason(source.facts, item, placed), source.facts.contentType);
      }
    }
    return labels;
  });

  const numbers = computed(() =>
    estimateImagePdf(
      pages.value,
      pdfItems.value,
      sources.value.map((source) => source?.facts ?? null),
      sheet.setup.value.dpi,
      output.pageNumbers,
    ),
  );
  const estimate = computed(() => (sheet.geometry.value ? pdfEstimateLine(numbers.value, titles.value) : ""));
  const dpiWarning = computed(() => pdfDpiWarning(numbers.value, titles.value, sheet.setup.value.dpi));
  const backgroundNotice = computed(() => backgroundRefusal(sheet.setup.value.background, "pdf") ?? "");

  const refusal = computed(() => {
    const entries = deps.entries.value;
    const geometry = sheet.geometry.value;
    if (entries.length === 0) {
      return "Add at least one image to the tray.";
    }
    const unreadable = entries.filter((entry, index) => entry.widthPx <= 0 || entry.heightPx <= 0 || !sources.value[index]);
    if (unreadable.length > 0) {
      const names = unreadable.map((entry) => `“${entry.title}”`).join(", ");
      return `${names} could not be read; remove ${unreadable.length === 1 ? "it" : "them"} from the tray.`;
    }
    if (!geometry) {
      return "The sheet is not ready.";
    }
    return (
      (output.pageNumbers ? pageNumbersRefusal(geometry) : null) ??
      pdfLimitRefusal(numbers.value, titles.value, itemsPerPage(plan.value, entries.length)) ??
      ""
    );
  });
  const build = usePdfBuild({
    refusal: () => refusal.value,
    snapshot: () => {
      const geometry = sheet.geometry.value;
      if (!geometry) {
        return null;
      }
      const parts = new Map<string, MediaPart>();
      for (const [index, source] of sources.value.entries()) {
        if (source) {
          parts.set(deps.entries.value[index].id, source.part);
        }
      }
      const { background, dpi } = sheet.setup.value;
      return {
        pages: pages.value,
        items: pdfItems.value,
        parts,
        sheet: { geometry, background, dpi },
        title: output.title.trim() || DEFAULT_ASSEMBLY_TITLE,
        author: output.author,
        subject: output.subject,
        pageNumbers: output.pageNumbers,
      };
    },
    steps: deps.steps,
  });
  const canProduce = computed(() => !build.busy.value && refusal.value === "");

  function setPageMode(value: string): boolean {
    const mode = PAGE_MODES.find((candidate) => candidate === value);
    if (!mode) {
      error.value = `"${value}" is not a page mode.`;
      return false;
    }
    output.pageMode = mode as PageMode;
    error.value = "";
    return true;
  }

  function setNUp(value: string): boolean {
    const nUp = Number(value);
    if (!N_UP_CHOICES.includes(nUp)) {
      error.value = `Images per page must be one of ${N_UP_CHOICES.join(", ")}.`;
      return false;
    }
    output.nUp = nUp;
    error.value = "";
    return true;
  }

  return {
    ...build,
    plan,
    pages,
    pageCount,
    pageCountLabel: computed(() => pageCountLabel(pageCount.value)),
    current,
    currentPlaced,
    previewNumber: computed(() =>
      output.format === "pdf" && output.pageNumbers && pageCount.value > 0 ? current.value + 1 : 0,
    ),
    goTo,
    onPageKey,
    encodings,
    estimate,
    dpiWarning,
    backgroundNotice,
    refusal,
    error,
    canProduce,
    setPageMode,
    setNUp,
    setPageNumbers: (on: boolean) => {
      output.pageNumbers = on;
    },
    setAuthor: (value: string) => {
      output.author = value;
    },
    setSubject: (value: string) => {
      output.subject = value;
    },
  };
}
