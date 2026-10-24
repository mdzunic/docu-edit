// The assembly plan's persistence (SPEC-109): the workspace's plan — tray
// item ids (and titles), layout, overlay transforms, sheet setup and output
// settings, never pixels — lives under the one `assemblyPlan` key of the
// SPEC-066 preferences store, written a moment after each change (and on
// leaving, reloading or closing the tab), and restored when the workspace
// opens. An item that has since
// been trashed or deleted is dropped from the restored tray, and a one-line
// notice names it. "Start over" clears the key until the next change.
//
// "Combine…" and "New › Combined image" start the workspace by writing a
// fresh plan here before routing to /assemble (startAssembly), so a reload
// of the workspace finds the same tray rather than the selection again.
//
// SPEC-111: a PDF in the selection starts the plan as a PDF of its pages
// (one item per page, each with its turns), and blank pages persist too; a
// PDF since trashed or deleted drops its pages with the same notice, and a
// restored tray holding pages is a PDF plan whatever the stored format.
import { onScopeDispose, ref, watch, type Ref } from "vue";
import { loadPreferences, savePreference } from "../../lib/preferences";
import { cloneSetup, initialSheetSetup } from "../../lib/assemblySheet";
import { droppedItemsNotice } from "../../lib/assembly";
import {
  ASSEMBLY_LAYOUT_DEFAULTS,
  ASSEMBLY_OUTPUT_DEFAULTS,
  freshAssembly,
  parsePersistedAssembly,
  type AssemblyLayoutSettings,
  type AssemblyOutputSettings,
  type PersistedAssembly,
  type PersistedAssemblyItem,
} from "../../lib/assemblyPersist";
import { MAX_ASSEMBLY_ITEMS } from "../../lib/core/assemblyLayout";
import { MAX_MERGE_PAGES } from "../../lib/core/pdfDocument";
import { DEFAULT_MERGE_TITLE } from "../../lib/pdfMerge";
import type { SheetSetupController } from "../../composables/useSheetSetup";
import {
  blankEntry,
  isPageEntry,
  libraryImage,
  libraryPdf,
  pageEntries,
  trayEntry,
  type AssemblyTrayController,
  type TrayEntry,
} from "./useAssemblyTray";

/** How long after a change the plan is written. */
export const PERSIST_DELAY_MS = 250;

// A notice the next workspace shows once — "Combine…" over more images than
// a tray holds says so on arrival.
let arrivalNotice = "";

/** Starts the workspace afresh with the library images and PDFs `ids`, in
 * order (an empty list opens an empty tray): an image is one item, a PDF
 * one item per page, and a selection holding a PDF starts a PDF plan. Ids
 * that are neither are skipped. */
export function startAssembly(ids: readonly string[]): void {
  const items: PersistedAssemblyItem[] = [];
  let images = 0;
  let imagesSeen = 0;
  let pagesLeft = 0;
  let pdfs = 0;
  for (const id of ids) {
    const image = libraryImage(id);
    if (image) {
      imagesSeen += 1;
      if (images < MAX_ASSEMBLY_ITEMS && items.length < MAX_MERGE_PAGES) {
        items.push({ id: image.id, title: image.title });
        images += 1;
      }
      continue;
    }
    const pdf = libraryPdf(id);
    if (pdf) {
      pdfs += 1;
      for (const index of pdf.pages.keys()) {
        if (items.length < MAX_MERGE_PAGES) {
          items.push({ id, title: pdf.source.title, page: index });
        } else {
          pagesLeft += 1;
        }
      }
    }
  }
  const imagesLeft = imagesSeen - images;
  arrivalNotice = [
    imagesLeft > 0 ? `A tray holds at most ${MAX_ASSEMBLY_ITEMS} images; the first ${MAX_ASSEMBLY_ITEMS} of ${imagesSeen} were added.` : "",
    pagesLeft > 0 ? `A PDF plan holds at most ${MAX_MERGE_PAGES} pages; ${pagesLeft} were not added.` : "",
  ]
    .filter(Boolean)
    .join(" ");
  savePreference("assemblyPlan", freshAssembly(items, pdfs > 0 ? { format: "pdf", title: DEFAULT_MERGE_TITLE } : {}));
}

/** The tray entries a stored item restores to — none when its source has
 * gone (or, for a page, no longer has that page). */
function restoredEntry(item: PersistedAssemblyItem): TrayEntry | null {
  if (item.blank) {
    return blankEntry(item.rotateBy);
  }
  if (item.page !== undefined) {
    const pdf = libraryPdf(item.id);
    return pdf ? pageEntries(pdf, [{ index: item.page, rotateBy: item.rotateBy }])[0] ?? null : null;
  }
  const image = libraryImage(item.id);
  return image ? trayEntry(image, item.transform) : null;
}

/** The stored plan, or null when none. A deep copy: parsing builds new
 * objects, so the workspace never edits the store's own. */
export function storedAssembly(): PersistedAssembly | null {
  return parsePersistedAssembly(loadPreferences().assemblyPlan) ?? null;
}

export interface AssemblyPersistenceDeps {
  tray: AssemblyTrayController;
  layout: AssemblyLayoutSettings;
  output: AssemblyOutputSettings;
  sheet: SheetSetupController;
}

export interface AssemblyPersistenceController {
  /** The restored-tray notice, or "" (the arrival notice reads here too). */
  notice: Ref<string>;
  /** Writes the plan now (a pending write is folded in). */
  flush: () => void;
  /** "Start over": an empty tray, default settings, no stored plan. */
  reset: () => void;
}

export function useAssemblyPersistence(deps: AssemblyPersistenceDeps): AssemblyPersistenceController {
  const { tray, layout, output, sheet } = deps;
  const notice = ref("");
  let timer: ReturnType<typeof setTimeout> | null = null;

  function snapshot(): PersistedAssembly {
    return {
      items: tray.entries.value.map((entry) => ({
        id: entry.id,
        title: entry.title,
        transform: { ...entry.transform },
        ...(entry.page ? { page: entry.page.index } : {}),
        ...(entry.blank ? { blank: true } : {}),
        ...(entry.rotateBy ? { rotateBy: entry.rotateBy } : {}),
      })),
      layout: { ...layout },
      sheet: cloneSetup(sheet.setup.value),
      marginsLinked: sheet.marginsLinked.value,
      output: { ...output },
    };
  }

  function cancelPending(): void {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  function flush(): void {
    if (timer !== null) {
      cancelPending();
      savePreference("assemblyPlan", snapshot());
    }
  }

  function schedule(): void {
    cancelPending();
    timer = setTimeout(() => {
      timer = null;
      savePreference("assemblyPlan", snapshot());
    }, PERSIST_DELAY_MS);
  }

  function applySheet(plan: PersistedAssembly | null): void {
    if (plan?.sheet) {
      sheet.setup.value = cloneSetup(plan.sheet);
      sheet.marginsLinked.value = plan.marginsLinked;
      return;
    }
    const start = initialSheetSetup(loadPreferences(), "png");
    sheet.setup.value = start.setup;
    sheet.marginsLinked.value = start.linked;
  }

  function restore(): void {
    const plan = storedAssembly();
    const kept: TrayEntry[] = [];
    const dropped: string[] = [];
    const goneSources = new Set<string>();
    for (const item of plan?.items ?? []) {
      const entry = restoredEntry(item);
      if (entry) {
        kept.push(entry);
      } else if (item.page === undefined) {
        dropped.push(item.title);
      } else if (!goneSources.has(item.id)) {
        // A gone PDF is named once, however many of its pages were planned.
        goneSources.add(item.id);
        dropped.push(item.title);
      }
    }
    tray.replace(kept);
    Object.assign(layout, plan?.layout ?? ASSEMBLY_LAYOUT_DEFAULTS);
    Object.assign(output, plan?.output ?? ASSEMBLY_OUTPUT_DEFAULTS);
    // Only a PDF can carry PDF pages.
    if (kept.some(isPageEntry)) {
      output.format = "pdf";
    }
    applySheet(plan);
    notice.value = [droppedItemsNotice(dropped), arrivalNotice].filter(Boolean).join(" ");
    arrivalNotice = "";
    // A dropped item leaves the store holding a plan the tray no longer
    // matches: write the restored one back.
    if (plan && dropped.length > 0) {
      savePreference("assemblyPlan", snapshot());
    }
  }

  function reset(): void {
    tray.replace([]);
    Object.assign(layout, ASSEMBLY_LAYOUT_DEFAULTS);
    Object.assign(output, ASSEMBLY_OUTPUT_DEFAULTS);
    applySheet(null);
    notice.value = "";
    tray.announcement.value = "Started over with an empty tray.";
    // The changes above scheduled a write; the plan is cleared instead.
    cancelPending();
    savePreference("assemblyPlan", null);
  }

  restore();
  // Sync, so a reset can cancel the write its own changes schedule.
  watch([tray.entries, () => ({ ...layout }), () => ({ ...output }), sheet.setup, sheet.marginsLinked], schedule, {
    deep: true,
    flush: "sync",
  });
  // A reload or a closed tab disposes nothing: write the pending change then.
  if (typeof window !== "undefined") {
    window.addEventListener("pagehide", flush);
    window.addEventListener("beforeunload", flush);
  }
  onScopeDispose(() => {
    flush();
    if (typeof window !== "undefined") {
      window.removeEventListener("pagehide", flush);
      window.removeEventListener("beforeunload", flush);
    }
  });

  return { notice, flush, reset };
}
