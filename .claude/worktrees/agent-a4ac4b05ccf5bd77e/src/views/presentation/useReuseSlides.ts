// "Reuse slides" (SPEC-099): pick slides from another deck in the library and
// insert them after the current selection, through importSlides' repair
// rules. The picker lists every presentation except this one and any record
// still in the preparation-milestone shape (`isLegacyPresentationModel`),
// reads the chosen deck with getPresentation, and offers the two formatting
// choices: "Use destination theme" (the default — references keep resolving
// against this deck's theme, layouts map by kind) or "Keep source formatting"
// (every reference resolved through the source deck's theme). Built by
// useSlideSorter, which owns the report it produces.
import { computed, ref, type Ref, type ShallowRef } from "vue";
import { getPresentation, listDocuments } from "../../lib/documentStore";
import { isLegacyPresentationModel, type PresentationModel } from "../../lib/presentation/model";
import { importSlides, type SlideImportReport } from "../../lib/presentation/slideImport";
import { frameAccessibleName } from "../../components/presentation/slideRender";

export interface ReuseSlidesDeps {
  model: ShallowRef<PresentationModel | null>;
  record: (label?: string) => void;
  commit: () => void;
  reportError: (message: string) => void;
  /** Where inserted slides land (after the current selection). */
  insertAt: () => number;
  /** Called after a successful insert with the import report and where the
   * slides landed. */
  onImported: (report: SlideImportReport, at: number) => void;
}

export type ReuseFormatting = "destination" | "source";

export function useReuseSlides(deps: ReuseSlidesDeps) {
  const open = ref(false);
  const decks: Ref<{ id: string; title: string }[]> = ref([]);
  const deckId = ref("");
  const selected = ref<number[]>([]);
  const formatting = ref<ReuseFormatting>("destination");

  const source = computed(() => (deckId.value ? (getPresentation(deckId.value) ?? null) : null));
  const sourceSlides = computed(() => source.value?.slides ?? []);

  /** Opens the picker on the first other deck in the library. */
  function show(): void {
    const current = deps.model.value?.id;
    decks.value = listDocuments()
      .filter((entry) => entry.kind === "presentation" && entry.id !== current)
      .filter((entry) => {
        const model = getPresentation(entry.id);
        return model !== undefined && !isLegacyPresentationModel(model);
      })
      .map((entry) => ({ id: entry.id, title: entry.title }));
    deckId.value = decks.value[0]?.id ?? "";
    selected.value = [];
    formatting.value = "destination";
    open.value = true;
  }

  function chooseDeck(id: string): void {
    deckId.value = id;
    selected.value = [];
  }

  function toggleSlide(index: number): void {
    selected.value = selected.value.includes(index)
      ? selected.value.filter((other) => other !== index)
      : [...selected.value, index].toSorted((a, b) => a - b);
  }

  function selectAll(): void {
    selected.value = sourceSlides.value.map((_, index) => index);
  }

  function setFormatting(value: string): void {
    formatting.value = value === "source" ? "source" : "destination";
  }

  function cancel(): void {
    open.value = false;
  }

  /** Inserts the picked slides (one undo step, "Reuse slides"). */
  function insert(): void {
    const model = deps.model.value;
    const from = source.value;
    if (!model || !from || selected.value.length === 0) {
      return;
    }
    try {
      deps.record("Reuse slides");
      const at = deps.insertAt();
      const report = importSlides(model, from, selected.value, {
        at,
        formatting: formatting.value,
        frameName: frameAccessibleName,
      });
      deps.commit();
      open.value = false;
      deps.onImported(report, at);
    } catch {
      deps.reportError("Couldn't reuse those slides");
    }
  }

  return {
    open,
    decks,
    deckId,
    selected,
    formatting,
    source,
    sourceSlides,
    show,
    chooseDeck,
    toggleSlide,
    selectAll,
    setFormatting,
    cancel,
    insert,
  };
}

export type ReuseSlides = ReturnType<typeof useReuseSlides>;
