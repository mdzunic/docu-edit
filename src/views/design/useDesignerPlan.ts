// The deck designer's Plan step state (SPEC-115) — the review gate between a
// reply and a build. New flow: the validated DeckPlan as an editable outline
// (every text, every layout from a select over the closed vocabulary,
// reorder, delete, add), and `currentPlan()`, which is exactly what Build
// receives — the plan the user last saw, not what the model first said.
// Redesign: one row per existing slide (its original index, read-only title
// and chosen kind) and the theme select; no add, no delete, no editable text.
import { computed, reactive, ref } from "vue";
import type { PresentationModel, SlideLayoutKind } from "../../lib/presentation/model";
import { STANDARD_LAYOUTS } from "../../lib/presentation/layouts";
import type { DeckDropRecord } from "../../lib/generate/deckDrops";
import {
  MAX_PLANNED_SLIDES,
  PLAN_LAYOUT_KINDS,
  PLAN_SIZES,
  PLAN_THEME_NAMES,
  bulletsToText,
  isPlanLayoutKind,
  textToBullets,
  validatePlan,
  type DeckPlan,
  type DeckRedesign,
  type PlanLayoutKind,
} from "../../lib/design/plan";
import { currentLayoutKind, seedLayoutKind, slideTitleText } from "../../lib/design/applyDesign";

/** One plan-step row of the New flow: its bullets edited as text, one per
 * line, two leading spaces per indent level. */
export interface EditablePlanSlide {
  key: number;
  layout: PlanLayoutKind;
  title: string;
  bullets: string;
  note: string;
}

/** One plan-step row of Redesign: an existing slide, by its original index,
 * with its read-only title and the kind chosen for it (null = keep). */
export interface RedesignRow {
  key: number;
  source: number;
  title: string;
  currentKind: SlideLayoutKind | null;
  layout: PlanLayoutKind | null;
}

export interface DesignerPlanOptions {
  mode: "new" | "redesign";
  /** Redesign: the open deck. */
  model: () => PresentationModel | null;
  /** Whether a plan is on screen (live drops are computed only then). */
  ready: () => boolean;
}

export function useDesignerPlan(options: DesignerPlanOptions) {
  let nextKey = 1;

  // --- New ---------------------------------------------------------------------
  const plan = reactive<{ title: string; theme: string; size: string; slides: EditablePlanSlide[] }>({
    title: "",
    theme: PLAN_THEME_NAMES[0],
    size: PLAN_SIZES[0],
    slides: [],
  });

  function loadPlan(next: DeckPlan): void {
    plan.title = next.title;
    plan.theme = next.theme;
    plan.size = next.size;
    plan.slides = next.slides.map((slide) => ({
      key: nextKey++,
      layout: slide.layout,
      title: slide.title,
      bullets: bulletsToText(slide.bullets),
      note: slide.note ?? "",
    }));
  }

  /** The plan as the user last saw it — what Build receives. */
  function currentPlan(): DeckPlan {
    return {
      title: plan.title,
      theme: plan.theme,
      size: plan.size as DeckPlan["size"],
      slides: plan.slides.map((slide) => ({
        layout: slide.layout,
        title: slide.title,
        bullets: textToBullets(slide.bullets),
        ...(slide.note.trim() === "" ? {} : { note: slide.note }),
      })),
    };
  }

  /** What Build would drop from the edited plan — a title slide's bullets
   * past its subtitle above all — listed under "What we changed" first. */
  const pendingDrops = computed<DeckDropRecord[]>(() => {
    if (options.mode !== "new" || !options.ready() || plan.slides.length === 0) {
      return [];
    }
    try {
      return validatePlan(currentPlan()).dropped;
    } catch {
      return [];
    }
  });

  function setPlanField(field: "title" | "theme" | "size", value: string): void {
    if (field === "theme" && !PLAN_THEME_NAMES.includes(value)) {
      return;
    }
    if (field === "size" && !(PLAN_SIZES as readonly string[]).includes(value)) {
      return;
    }
    plan[field] = value;
  }

  function setSlideField(index: number, field: "title" | "bullets" | "note" | "layout", value: string): void {
    const slide = plan.slides[index];
    if (!slide) {
      return;
    }
    if (field === "layout") {
      if (isPlanLayoutKind(value)) {
        slide.layout = value;
      }
      return;
    }
    slide[field] = value;
  }

  /** Alt+↑/Alt+↓: returns the row's new index. */
  function moveSlide(index: number, delta: number): number {
    const target = index + delta;
    if (index < 0 || index >= plan.slides.length || target < 0 || target >= plan.slides.length) {
      return index;
    }
    const [moved] = plan.slides.splice(index, 1);
    plan.slides.splice(target, 0, moved);
    return target;
  }

  function deleteSlide(index: number): void {
    if (index >= 0 && index < plan.slides.length) {
      plan.slides.splice(index, 1);
    }
  }

  const canAddSlide = computed(() => plan.slides.length < MAX_PLANNED_SLIDES);
  function addSlide(): void {
    if (canAddSlide.value) {
      plan.slides.push({ key: nextKey++, layout: "titleAndBody", title: "", bullets: "", note: "" });
    }
  }

  // --- Redesign ------------------------------------------------------------------
  const rows = ref<RedesignRow[]>([]);
  /** "" = Keep current theme. */
  const redesignTheme = ref("");

  /** Seeds the rows from the open deck — in a validated redesign's order with
   * its kinds, or (null) as the deck stands: each row's kind is the model's
   * choice, else the slide's current kind when it is one of the six, else
   * "Keep current". */
  function seedRows(redesign: DeckRedesign | null): void {
    const model = options.model();
    if (!model) {
      rows.value = [];
      return;
    }
    const order = redesign?.slides ?? model.slides.map((_, index) => ({ source: index, layout: null }));
    rows.value = order.map((entry) => {
      const slide = model.slides[entry.source];
      return {
        key: nextKey++,
        source: entry.source,
        title: slideTitleText(slide),
        currentKind: currentLayoutKind(model, slide) ?? null,
        layout: entry.layout ?? seedLayoutKind(model, slide),
      };
    });
    redesignTheme.value = redesign?.theme ?? "";
  }

  /** The reviewed rows as a DeckRedesign. */
  function currentRedesign(): DeckRedesign {
    return {
      theme: redesignTheme.value === "" ? null : redesignTheme.value,
      slides: rows.value.map((row) => ({ source: row.source, layout: row.layout })),
    };
  }

  function setRowLayout(index: number, value: string): void {
    const row = rows.value[index];
    if (row) {
      row.layout = isPlanLayoutKind(value) ? value : null;
    }
  }

  function moveRow(index: number, delta: number): number {
    const target = index + delta;
    if (index < 0 || index >= rows.value.length || target < 0 || target >= rows.value.length) {
      return index;
    }
    const next = rows.value.slice();
    const [moved] = next.splice(index, 1);
    next.splice(target, 0, moved);
    rows.value = next;
    return target;
  }

  function setRedesignTheme(value: string): void {
    redesignTheme.value = value === "" || PLAN_THEME_NAMES.includes(value) ? value : "";
  }

  /** Each kind under the registry's own name for it ("Title and Content"). */
  const layoutOptions = PLAN_LAYOUT_KINDS.map((kind) => ({
    value: kind,
    label: STANDARD_LAYOUTS.find((def) => def.kind === kind)?.name ?? kind,
  }));

  return {
    plan,
    loadPlan,
    currentPlan,
    pendingDrops,
    setPlanField,
    setSlideField,
    moveSlide,
    deleteSlide,
    addSlide,
    canAddSlide,
    rows,
    redesignTheme,
    seedRows,
    currentRedesign,
    setRowLayout,
    moveRow,
    setRedesignTheme,
    layoutOptions,
    themeNames: [...PLAN_THEME_NAMES],
    sizes: [...PLAN_SIZES],
  };
}
