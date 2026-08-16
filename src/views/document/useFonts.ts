// SPEC-048 — the browser-side font runtime: registers the bundled metric
// substitutes and any fonts embedded in the open document as @font-face faces
// (so the substitution lib/renderFormatting/textMetrics resolve actually
// paint), tracks when they have loaded (AC-19), and drives the "Fonts" panel
// (AC-7/8/10). All the format-neutral logic lives in lib/fonts/*; this
// composable is the thin Vue/DOM layer that owns FontFace registration,
// reactive report state and the pane's open/close.
import { computed, ref, watch } from "vue";
import { documentFontFamilies } from "../../lib/fonts";
import { BUNDLED_FACES } from "../../lib/fonts/bundledFaces";
import { extractEmbeddedFonts } from "../../lib/fonts/embeddedFonts";
import { documentFontReport, type FontReportEntry } from "../../lib/fonts/report";
import type { EditorContext } from "./context";

export interface FontsDeps extends EditorContext {
  /** Recomputes layout when a newly-loaded face changes text metrics — the
   * pagination's invalidateAll (which also resets the canvas measurer). */
  relayout: () => void;
}

/** Typed pass-through for FontReportEntry[] arriving at FontsPane.vue as a
 * runtime-declared Array prop (espree bars the annotation in a .vue script —
 * same idiom as useAccessibilityCheck's asAccessibilityIssues). */
export function asFontReportEntries(value: unknown): FontReportEntry[] {
  return value as FontReportEntry[];
}

/** The status label FontsPane shows for one entry — the substituted case
 * names the twin (AC-8). Typed here rather than in the espree-parsed .vue. */
export function fontStatusLabel(entry: FontReportEntry): string {
  if (entry.status === "substituted" && entry.substitute) {
    return `Substituted — ${entry.substitute}`;
  }
  const labels: Record<FontReportEntry["status"], string> = {
    available: "Available",
    substituted: "Substituted",
    missing: "Missing",
  };
  return labels[entry.status];
}

const FONT_BASE = `${import.meta.env.BASE_URL}fonts/`;

/** Whether the browser can render `family` from a locally-installed (or
 * already-registered @font-face) font, by comparing a probe string's width
 * against the three generic fallbacks — a family the machine lacks measures
 * identically to its fallback. Returns false where no canvas is available
 * (jsdom under Vitest), so the panel degrades to "substituted/missing"
 * rather than guessing "available". */
function makeInstalledProbe(): (family: string) => boolean {
  if (typeof document === "undefined") {
    return () => false;
  }
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return () => false;
  }
  const TEXT = "mmmmmwwwWWlliI0Oo—AÀ";
  const GENERICS = ["monospace", "serif", "sans-serif"];
  const widthIn = (font: string): number => {
    ctx.font = `40px ${font}`;
    return ctx.measureText(TEXT).width;
  };
  const baselines = GENERICS.map(widthIn);
  return (family) => {
    const quoted = /^[a-zA-Z][a-zA-Z0-9-]*$/.test(family.trim()) ? family : `"${family}"`;
    return GENERICS.some((generic, i) => widthIn(`${quoted}, ${generic}`) !== baselines[i]);
  };
}

export function useFonts(deps: FontsDeps) {
  const { doc, documentId, relayout } = deps;

  const fontsPaneOpen = ref(false);
  /** Bumped whenever faces finish loading or the embedded set changes, so the
   * (non-reactive) availability probe is re-run for the report. */
  const fontsVersion = ref(0);
  const fontsReady = ref(false);
  /** Lower-cased families the open document embeds its own faces for. */
  const embeddedFamilies = ref<Set<string>>(new Set());

  const hasFontFace = typeof FontFace !== "undefined" && typeof document !== "undefined";
  const isInstalled = makeInstalledProbe();
  /** Embedded FontFaces we added for the current document, so they can be
   * removed when another document opens. */
  let registeredEmbedded: FontFace[] = [];

  function afterLoad(): void {
    fontsReady.value = true;
    fontsVersion.value += 1;
    relayout();
  }

  // Register the bundled substitutes once. Without these @font-face faces the
  // substitution stacks (Calibri → Carlito …) would fall straight through to
  // the generic keyword, so this is what makes AC-1/AC-2 actually render.
  if (hasFontFace) {
    const loads = BUNDLED_FACES.map((face) => {
      const ff = new FontFace(face.family, `url("${FONT_BASE}${face.file}")`, {
        weight: face.weight,
        style: face.style,
      });
      document.fonts.add(ff);
      return ff.load().catch(() => undefined);
    });
    void Promise.allSettled(loads).then(afterLoad);
  } else {
    // No FontFace API (jsdom): nothing to load, layout is already "ready".
    fontsReady.value = true;
  }

  /** (Re)registers the fonts embedded in the current document (AC-11/13),
   * skipping any whose licence bits forbid embedding, and records their
   * families for the panel. */
  function registerEmbedded(): void {
    const model = doc.value;
    const next = new Set<string>();
    if (hasFontFace) {
      for (const ff of registeredEmbedded) {
        document.fonts.delete(ff);
      }
      registeredEmbedded = [];
    }
    if (model) {
      const faces = extractEmbeddedFonts(model).filter((face) => face.embeddable && face.family);
      for (const face of faces) {
        next.add(face.family.toLowerCase());
        if (hasFontFace) {
          // Copy into a fresh ArrayBuffer so FontFace owns contiguous bytes.
          const buffer = face.bytes.slice().buffer;
          const ff = new FontFace(face.family, buffer, {
            weight: face.bold ? "bold" : "normal",
            style: face.italic ? "italic" : "normal",
          });
          document.fonts.add(ff);
          registeredEmbedded.push(ff);
          void ff.load().then(afterLoad).catch(() => undefined);
        }
      }
    }
    embeddedFamilies.value = next;
    fontsVersion.value += 1;
  }

  // Re-extract embedded fonts whenever a different document opens.
  watch(documentId, registerEmbedded, { immediate: true });

  /** The Fonts panel's list (AC-7/8/10), recomputed as fonts load and the
   * document changes. */
  const fontReport = computed<FontReportEntry[]>(() => {
    // Touch fontsVersion so the non-reactive install probe is re-read.
    void fontsVersion.value;
    const model = doc.value;
    if (!model) {
      return [];
    }
    return documentFontReport(model, { isInstalled, embeddedFamilies: embeddedFamilies.value });
  });

  function openFontsPane(): void {
    fontsPaneOpen.value = true;
  }
  function closeFontsPane(): void {
    fontsPaneOpen.value = false;
  }

  /** The families the open document uses, for the toolbar dropdown's "Fonts
   * in this document" group (SPEC-013) — recomputes as families come and go. */
  const documentFonts = computed(() => (doc.value ? documentFontFamilies(doc.value) : []));

  return { fontsPaneOpen, fontsReady, fontReport, documentFonts, openFontsPane, closeFontsPane };
}
