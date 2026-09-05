// The deck-theme editing session (SPEC-094) — the presentation view's wiring
// for the Design panel: the persisted open/closed state (the
// `presentationThemeOpen` preference), the gallery (current theme first, then
// the six built-ins), card selection + hover preview, Apply as ONE undo entry
// ("Apply theme"), the Customise slot/font edits, Reset back to the pristine
// theme, the odp loss note and the pptx `p:clrMapOvr` opt-out note. The model
// mutation rules live in `lib/presentation/theme.ts`; this composable only
// sequences record → mutate → commit and holds the transient panel state.
import { computed, ref, shallowRef, type ShallowRef } from "vue";
import { BUILT_IN_THEMES, type DocumentTheme, type ThemeColorSlot } from "../../lib/core/theme";
import type { DeckTheme, PresentationModel } from "../../lib/presentation/model";
import {
  applyTheme,
  cloneDeckTheme,
  deckThemeOf,
  setThemeColor,
  setThemeFont,
  slideUsesTheme,
} from "../../lib/presentation/theme";
import { slideClrMapMode } from "../../lib/presentation/pptx/theme";
import { usePreferences } from "../../composables/usePreferences";

export interface DeckThemeDeps {
  model: ShallowRef<PresentationModel | null>;
  record: (label?: string) => void;
  commit: () => void;
}

/** One gallery card: the deck's current theme first (id "current"), then the
 * six built-ins (`builtin-{n}`). */
export interface ThemeCard {
  id: string;
  name: string;
  accents: string[];
  current: boolean;
}

function accentsOf(theme: DocumentTheme): string[] {
  return [
    theme.colors.accent1,
    theme.colors.accent2,
    theme.colors.accent3,
    theme.colors.accent4,
    theme.colors.accent5,
    theme.colors.accent6,
  ];
}

function sameScheme(a: DocumentTheme, b: DocumentTheme): boolean {
  return (
    Object.entries(a.colors).every(([slot, hex]) => b.colors[slot as ThemeColorSlot] === hex) &&
    a.fonts.major === b.fonts.major &&
    a.fonts.minor === b.fonts.minor
  );
}

function themeFor(id: string | null): DocumentTheme | undefined {
  const match = id ? /^builtin-(\d+)$/.exec(id) : null;
  return match ? BUILT_IN_THEMES[Number(match[1])] : undefined;
}

export function useDeckTheme(deps: DeckThemeDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationThemeOpen);
  function toggle(): void {
    setPreference("presentationThemeOpen", !open.value);
  }

  const theme = computed<DocumentTheme>(() => deckThemeOf(deps.model.value));
  const themeName = computed(() => theme.value.name);
  /** True when the deck's scheme matches no built-in — a customised theme. */
  const isCustom = computed(() => !BUILT_IN_THEMES.some((builtIn) => sameScheme(builtIn, theme.value)));

  const gallery = computed<ThemeCard[]>(() => [
    { id: "current", name: themeName.value, accents: accentsOf(theme.value), current: true },
    ...BUILT_IN_THEMES.map((builtIn, index) => ({
      id: `builtin-${index}`,
      name: builtIn.name,
      accents: accentsOf(builtIn),
      current: false,
    })),
  ]);
  const currentThemeId = computed(() => "current");

  // --- selection + hover preview --------------------------------------------
  const selectedId = ref("current");
  function select(id: string): void {
    selectedId.value = id;
  }
  const previewId = ref<string | null>(null);
  function hoverPreview(id: string): void {
    previewId.value = id;
  }
  function clearPreview(): void {
    previewId.value = null;
  }
  /** What the stage renders against: a hovered card's scheme while previewing,
   * else the deck's own theme (every ref re-resolves live, nothing persists). */
  const displayTheme = computed<DocumentTheme | undefined>(
    () => themeFor(previewId.value) ?? deps.model.value?.theme,
  );

  // --- apply / customise / reset --------------------------------------------
  /** The pristine theme stashed before the first Customise edit, so Reset can
   * put it back; cleared by Apply (a new baseline) and by Reset itself. */
  const pristine = shallowRef<DeckTheme | null>(null);

  function commitTheme(label: string, mutate: (model: PresentationModel) => void): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    deps.record(label);
    mutate(model);
    deps.commit();
  }

  /** Applies the selected gallery card (one undo entry, "Apply theme"). */
  function apply(id?: string): void {
    const target = themeFor(id ?? selectedId.value);
    if (!target) {
      return; // the current card re-applied is a no-op
    }
    commitTheme("Apply theme", (model) => applyTheme(model, target));
    pristine.value = null;
    selectedId.value = "current";
    clearPreview();
  }

  function stashPristine(): void {
    const model = deps.model.value;
    if (!pristine.value && model) {
      pristine.value = cloneDeckTheme({ ...deckThemeOf(model) });
    }
  }

  function setSchemeColor(slot: ThemeColorSlot, hex: string): void {
    if (!/^#[0-9a-f]{6}$/i.test(hex)) {
      return;
    }
    stashPristine();
    commitTheme("Edit theme", (model) => setThemeColor(model, slot, hex.toLowerCase()));
  }

  function setFont(which: "major" | "minor", family: string): void {
    if (family.trim().length === 0) {
      return;
    }
    stashPristine();
    commitTheme("Edit theme", (model) => setThemeFont(model, which, family.trim()));
  }

  /** Puts the pristine (pre-Customise) theme back — one undo entry of its own. */
  function reset(): void {
    const baseline = pristine.value;
    if (!baseline) {
      return;
    }
    commitTheme("Reset theme", (model) => applyTheme(model, baseline));
    pristine.value = null;
  }
  const canReset = computed(() => pristine.value !== null);

  // --- honesty notes ---------------------------------------------------------
  /** odp only: the stated loss — transformed refs and theme fonts write their
   * literals and lose their reference on re-open (SPEC-094 AC-32). */
  const lossNote = computed<string | null>(() =>
    deps.model.value?.format === "odp"
      ? "OpenDocument has no theme part: colour variants (lighter/darker) and theme fonts save as their current values and stop following the theme after reopening."
      : null,
  );

  /** 1-based numbers of slides opted out of theme colours: a pptx slide whose
   * `p:clrMapOvr` cannot be parsed writes literals so the rendered colour never
   * lies (AC-23), and the panel says so (AC-24). */
  const optedOutSlides = computed<number[]>(() => {
    const model = deps.model.value;
    if (!model || model.format !== "pptx" || !model.theme) {
      return [];
    }
    const out: number[] = [];
    model.slides.forEach((slide, index) => {
      const preserved = slide.preserved?.["p:clrMapOvr"];
      if (preserved !== undefined && slideClrMapMode(preserved, model.theme?.colorMap) === "literal") {
        out.push(index + 1);
      }
    });
    return out;
  });

  /** Whether any slide actually tracks the theme (panel context only). */
  const anySlideUsesTheme = computed(() => (deps.model.value?.slides ?? []).some(slideUsesTheme));

  return {
    open,
    toggle,
    theme,
    themeName,
    isCustom,
    gallery,
    currentThemeId,
    selectedId,
    select,
    previewId,
    hoverPreview,
    clearPreview,
    displayTheme,
    apply,
    setSchemeColor,
    setFont,
    reset,
    canReset,
    lossNote,
    optedOutSlides,
    anySlideUsesTheme,
  };
}
