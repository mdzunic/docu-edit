// The Header & Footer editing session (SPEC-093) — the presentation view's
// wiring for the modal dialog: the open/closed state, the dialog's draft fields
// (seeded from the deck values + this slide's visibility), the six catalogue
// date-format options previewed against today, and the two apply flows (Apply /
// Apply to all). The model mutation itself lives framework-free in
// `lib/presentation/headerFooter.ts`; this composable only sequences
// record → mutate → commit and holds the transient draft (the useDeckSetup /
// useTransitions pattern). The dialog does not edit notes/handout fields.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import {
  DEFAULT_DATE_FORMAT,
  applyHeaderFooter,
  headerFooterDraft,
  type HeaderFooterValues,
} from "../../lib/presentation/headerFooter";
import { DATE_TIME_CATALOGUE, renderDateTime } from "../../lib/core/dateTimeFormats";
import { DEFAULT_LANGUAGE } from "../../lib/core/language";
import { now } from "../../lib/clock";
import type { DateFormatOption } from "./headerFooterFields";

export interface HeaderFooterDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
}

/** The six supported date patterns (SPEC-093 AC-57): the catalogue's date (not
 * time) entries, in catalogue order. */
const DATE_PATTERNS = DATE_TIME_CATALOGUE.filter((entry) => !entry.isTime).map((entry) => entry.format);

export function useHeaderFooter(deps: HeaderFooterDeps) {
  const open = ref(false);

  // The dialog's draft — every control as a plain reactive value.
  const footer = ref(false);
  const footerText = ref("");
  const date = ref(false);
  const dateSource = ref<"auto" | "fixed">("auto");
  const dateFormat = ref(DEFAULT_DATE_FORMAT);
  const dateFixedText = ref("");
  const slideNumber = ref(false);
  const showOnTitleSlide = ref(true);

  function initFromModel(): void {
    const model = deps.model.value;
    const draft = model ? headerFooterDraft(model, deps.currentSlide.value) : null;
    footer.value = draft?.footer ?? false;
    footerText.value = draft?.footerText ?? "";
    date.value = draft?.date ?? false;
    dateSource.value = draft?.dateSource ?? "auto";
    dateFormat.value = draft?.dateFormat ?? DEFAULT_DATE_FORMAT;
    dateFixedText.value = draft?.dateFixedText ?? "";
    slideNumber.value = draft?.slideNumber ?? false;
    showOnTitleSlide.value = draft?.showOnTitleSlide ?? true;
  }

  /** Opens the dialog, seeding the draft from the live model + current slide. */
  function show(): void {
    if (!deps.model.value) {
      return;
    }
    initFromModel();
    open.value = true;
  }
  function close(): void {
    open.value = false;
  }

  /** The current slide's date field is out-of-subset (its wire format maps to no
   * catalogue pattern), so the format select is disabled behind a note — the
   * value is preserved verbatim and never recomputed (AC-63/65). */
  const datePreserved = computed(() => {
    const entry = deps.currentSlide.value?.fields?.find((field) => field.kind === "date");
    return entry !== undefined && (entry.wireType !== undefined || entry.wireStyleName !== undefined);
  });

  /** The six date patterns previewed against today's instant (AC-40/41). */
  const dateFormatOptions = computed<DateFormatOption[]>(() => {
    const instant = now();
    return DATE_PATTERNS.map((format) => ({
      value: format,
      label: renderDateTime(format, instant, DEFAULT_LANGUAGE) ?? format,
    }));
  });

  function values(): HeaderFooterValues {
    return {
      footer: footer.value,
      footerText: footerText.value,
      date: date.value,
      dateSource: dateSource.value,
      dateFormat: dateFormat.value,
      dateFixedText: dateFixedText.value,
      slideNumber: slideNumber.value,
      showOnTitleSlide: showOnTitleSlide.value,
    };
  }

  /** Apply: deck values + this slide's visibility (AC-47). */
  function apply(): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    deps.record("Header and footer");
    applyHeaderFooter(model, values(), { mode: "slide", index: deps.currentIndex.value });
    deps.commit();
    close();
  }

  /** Apply to all: deck values + the same visibility on every slide, dropping
   * per-slide text overrides (AC-48). */
  function applyToAll(): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    deps.record("Apply header and footer to all slides");
    applyHeaderFooter(model, values(), { mode: "all" });
    deps.commit();
    close();
  }

  /** Cancel changes nothing (AC-56). */
  function cancel(): void {
    close();
  }

  return {
    open,
    footer,
    footerText,
    date,
    dateSource,
    dateFormat,
    dateFixedText,
    slideNumber,
    showOnTitleSlide,
    datePreserved,
    dateFormatOptions,
    show,
    close,
    apply,
    applyToAll,
    cancel,
  };
}

export type HeaderFooterController = ReturnType<typeof useHeaderFooter>;
