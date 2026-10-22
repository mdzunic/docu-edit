// The deck-setup editing session (SPEC-090) — the presentation view's wiring
// for the "Deck setup" dialog: local draft state for the Size and Properties
// tabs, unit-aware size parsing/validation, the content-rescale decision, and
// the single Apply that records one undo entry, mutates the model through the
// framework-free `deckSetup` commands and commits (re-render + persist). Drafts
// are local, so Cancel changes nothing. The heavy lifting — the rescale maths,
// the first-slide-number origin and the metadata write-back — lives in
// `lib/presentation/deckSetup.ts`; this composable only sequences
// record → mutate → commit and holds the transient dialog UI state.
import { computed, reactive, ref, type Ref } from "vue";
import {
  asSlideSizeName,
  matchSlideSize,
  MAX_SLIDE_DIMENSION_CM,
  MIN_SLIDE_DIMENSION_CM,
  orientationOf,
  slideSizeFor,
  swapOrientation,
  type Orientation,
  type PageSize,
  type SlideSizeName,
} from "../../lib/core/pageSetup";
import { formatLength, parseLengthToCm } from "../../lib/core/units";
import type { DocumentProperties } from "../../lib/core/documentProperties";
import type { PresentationModel } from "../../lib/presentation/model";
import {
  applyDeckProperties,
  applyDeckSize,
  framesOutsideSlide,
  setFirstSlideNumber as setModelFirstSlideNumber,
  type RescaleMode,
} from "../../lib/presentation/deckSetup";
import { usePreferences } from "../../composables/usePreferences";

export type DeckSetupTab = "size" | "properties";

export interface DeckSetupDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  reportError: (message: string) => void;
}

const MIN_FIRST_SLIDE = 1;
const MAX_FIRST_SLIDE = 9999;

/** The six metadata keys the Properties tab edits — a clone seeds the draft. */
function cloneDeckProperties(properties: DocumentProperties | undefined): DocumentProperties {
  return properties ? JSON.parse(JSON.stringify(properties)) : {};
}

function inRange(cm: number | null): cm is number {
  return cm !== null && cm >= MIN_SLIDE_DIMENSION_CM - 1e-9 && cm <= MAX_SLIDE_DIMENSION_CM + 1e-9;
}

export function useDeckSetup(deps: DeckSetupDeps) {
  const { preferences } = usePreferences();

  const isOpen = ref(false);
  const activeTab = ref<DeckSetupTab>("size");

  const preset = ref<SlideSizeName>("16:9");
  const orientation = ref<Orientation>("landscape");
  const widthText = ref("");
  const heightText = ref("");
  const rescaleMode = ref<RescaleMode>("scale");
  const firstSlideNumber = ref(MIN_FIRST_SLIDE);
  const propertiesDraft = reactive<DocumentProperties>({});

  const unit = computed(() => preferences.measurementUnit);
  const isOdp = computed(() => deps.model.value?.format === "odp");

  /** Resets every draft from the live model — called on each open so a
   * re-opened dialog reflects the current deck (and never a stale draft). */
  function initFromModel(): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    const size: PageSize = { widthCm: model.slideWidthCm, heightCm: model.slideHeightCm };
    preset.value = matchSlideSize(size);
    orientation.value = orientationOf(size);
    widthText.value = formatLength(size.widthCm, unit.value);
    heightText.value = formatLength(size.heightCm, unit.value);
    rescaleMode.value = "scale";
    firstSlideNumber.value = model.firstSlideNumber ?? MIN_FIRST_SLIDE;
    activeTab.value = "size";
    const draft = cloneDeckProperties(model.properties);
    for (const key of Object.keys(propertiesDraft) as (keyof DocumentProperties)[]) {
      delete propertiesDraft[key];
    }
    Object.assign(propertiesDraft, draft);
  }

  function open(): void {
    if (!deps.model.value) {
      return;
    }
    initFromModel();
    isOpen.value = true;
  }
  function close(): void {
    isOpen.value = false;
  }
  function setTab(tab: DeckSetupTab): void {
    activeTab.value = tab;
  }

  // --- size tab -------------------------------------------------------------

  const parsedWidthCm = computed(() => parseLengthToCm(widthText.value, unit.value));
  const parsedHeightCm = computed(() => parseLengthToCm(heightText.value, unit.value));

  const sizeValid = computed(() => inRange(parsedWidthCm.value) && inRange(parsedHeightCm.value));
  const sizeError = computed(() =>
    sizeValid.value
      ? ""
      : `Enter a width and height between ${formatLength(MIN_SLIDE_DIMENSION_CM, unit.value)} and ` +
        `${formatLength(MAX_SLIDE_DIMENSION_CM, unit.value)} ${unit.value}`,
  );

  const targetSize = computed<PageSize | null>(() =>
    sizeValid.value ? { widthCm: parsedWidthCm.value!, heightCm: parsedHeightCm.value! } : null,
  );

  function fillFromPreset(name: SlideSizeName, orient: Orientation): void {
    if (name === "Custom") {
      return;
    }
    const size = slideSizeFor(name, orient);
    widthText.value = formatLength(size.widthCm, unit.value);
    heightText.value = formatLength(size.heightCm, unit.value);
  }

  function setPreset(name: string): void {
    const parsed = asSlideSizeName(name);
    preset.value = parsed;
    fillFromPreset(parsed, orientation.value);
  }

  function setOrientation(next: string): void {
    const value = next === "portrait" ? "portrait" : "landscape";
    if (value === orientation.value) {
      return;
    }
    orientation.value = value;
    if (preset.value !== "Custom") {
      fillFromPreset(preset.value, value);
      return;
    }
    // Custom: swap whatever the fields currently hold.
    if (targetSize.value) {
      const swapped = swapOrientation(targetSize.value);
      widthText.value = formatLength(swapped.widthCm, unit.value);
      heightText.value = formatLength(swapped.heightCm, unit.value);
    }
  }

  function setWidth(text: string): void {
    widthText.value = text;
    preset.value = "Custom";
    syncOrientation();
  }
  function setHeight(text: string): void {
    heightText.value = text;
    preset.value = "Custom";
    syncOrientation();
  }
  function syncOrientation(): void {
    if (targetSize.value) {
      orientation.value = orientationOf(targetSize.value);
    }
  }

  function setRescaleMode(mode: string): void {
    rescaleMode.value = mode === "keep" ? "keep" : "scale";
  }

  const overflowCount = computed(() => {
    const model = deps.model.value;
    if (!model || !targetSize.value) {
      return 0;
    }
    return framesOutsideSlide({
      ...model,
      slideWidthCm: targetSize.value.widthCm,
      slideHeightCm: targetSize.value.heightCm,
    });
  });

  const firstSlideValid = computed(
    () =>
      Number.isInteger(firstSlideNumber.value) &&
      firstSlideNumber.value >= MIN_FIRST_SLIDE &&
      firstSlideNumber.value <= MAX_FIRST_SLIDE,
  );
  function setFirstSlideNumber(value: number): void {
    firstSlideNumber.value = value;
  }

  // --- properties tab -------------------------------------------------------

  const created = computed(() => deps.model.value?.properties?.created ?? "");
  const modified = computed(() => deps.model.value?.properties?.modified ?? "");
  const extraFields = computed(() => deps.model.value?.properties?.extraFields ?? []);

  function setProperty(key: string, value: string): void {
    propertiesDraft[key as keyof DocumentProperties] = value as never;
  }

  // --- apply ----------------------------------------------------------------

  const canApply = computed(() =>
    activeTab.value === "properties" ? true : sizeValid.value && firstSlideValid.value,
  );

  function apply(): void {
    const model = deps.model.value;
    if (!model || !canApply.value) {
      return;
    }
    if (activeTab.value === "size") {
      const size = targetSize.value;
      if (!size) {
        return;
      }
      deps.record("Deck setup");
      applyDeckSize(model, size, rescaleMode.value);
      setModelFirstSlideNumber(model, firstSlideNumber.value);
    } else {
      deps.record("Deck properties");
      applyDeckProperties(model, JSON.parse(JSON.stringify(propertiesDraft)));
    }
    deps.commit();
    close();
  }

  return {
    open,
    isOpen,
    close,
    activeTab,
    setTab,
    preset,
    setPreset,
    orientation,
    setOrientation,
    widthText,
    heightText,
    setWidth,
    setHeight,
    unit,
    sizeError,
    sizeValid,
    rescaleMode,
    setRescaleMode,
    overflowCount,
    firstSlideNumber,
    setFirstSlideNumber,
    firstSlideValid,
    isOdp,
    propertiesDraft,
    setProperty,
    created,
    modified,
    extraFields,
    canApply,
    apply,
  };
}

export type DeckSetupController = ReturnType<typeof useDeckSetup>;
