// The deck designer's Brief step (SPEC-115): the prompt and its choices, the
// optional source, and the request they describe — exactly as it would be
// handed to the gateway, which wraps it in the body SPEC-113's consent dialog
// shows before anything is sent. The live "what will be sent" count and the
// truncation notice read from it. Redesign has no length and no source: its
// request lists the open deck's slides instead (summariseSlides — never body
// text or notes).
import { computed, ref } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { deckThemeOf } from "../../lib/presentation/theme";
import { summariseSlides } from "../../lib/design/applyDesign";
import type { OutlineSource } from "../../lib/design/outlineDesigner";
import {
  AUDIENCES,
  DESIGN_PURPOSE,
  DESIGN_SYSTEM_PROMPT,
  LENGTHS,
  MAX_SOURCE_CHARS,
  REDESIGN_PURPOSE,
  REDESIGN_SYSTEM_PROMPT,
  TONES,
  designUserMessage,
  messageBytes,
  redesignUserMessage,
  truncateSource,
} from "../../lib/design/prompt";
import type { DesignerLibrary } from "./designerLibrary";
import { outlineFromText, outlineText } from "./outlineSource";

export interface DesignerBriefOptions {
  mode: "new" | "redesign";
  /** Redesign: the open deck. */
  model: () => PresentationModel | null;
  library: DesignerLibrary;
}

/** A request before the gateway sees it. */
export interface DesignerRequest {
  purpose: string;
  system: string;
  user: string;
}

export function useDesignerBrief(options: DesignerBriefOptions) {
  const { mode, library } = options;

  const brief = ref("");
  const audience = ref(AUDIENCES[0].value);
  const length = ref(LENGTHS[1].value);
  const tone = ref(TONES[0].value);
  /** "" (nothing), "paste", or a library entry's id. */
  const sourceChoice = ref("");
  const pastedText = ref("");
  /** Store the brief in the deck's GeneratedPrompt property. */
  const keepPrompt = ref(false);

  const sources = mode === "new" ? library.sources() : [];
  const sourceOptions = [
    { value: "", label: "Nothing" },
    { value: "paste", label: "Pasted text" },
    ...sources.map((entry) => ({
      value: entry.id,
      label: `${entry.title} (${entry.kind === "deck" ? "presentation" : "document"})`,
    })),
  ];

  /** The picked source: its title (pasted text has none), its outline for
   * the model-free designer, and the text a request would carry. */
  const source = computed<{ title: string; outline: OutlineSource; text: string } | null>(() => {
    if (sourceChoice.value === "") {
      return null;
    }
    if (sourceChoice.value === "paste") {
      return pastedText.value.trim() === ""
        ? null
        : { title: "", outline: outlineFromText(pastedText.value), text: pastedText.value };
    }
    const entry = sources.find((candidate) => candidate.id === sourceChoice.value);
    const entryOutline = entry ? library.outlineOf(entry.id) : undefined;
    return entry && entryOutline ? { title: entry.title, outline: entryOutline, text: outlineText(entryOutline) } : null;
  });

  /** The dialog's statement that a long source is cut. */
  const truncation = computed(() => {
    const picked = source.value;
    if (!picked) {
      return null;
    }
    const cut = truncateSource(picked.text);
    return cut.truncated
      ? `The source is ${cut.length.toLocaleString("en-US")} characters long; only the first ${MAX_SOURCE_CHARS.toLocaleString("en-US")} will be sent.`
      : null;
  });

  const request = computed<DesignerRequest>(() => {
    if (mode === "redesign") {
      const model = options.model();
      return {
        purpose: REDESIGN_PURPOSE,
        system: REDESIGN_SYSTEM_PROMPT,
        user: redesignUserMessage({
          brief: brief.value,
          audience: audience.value,
          tone: tone.value,
          deckTitle: model?.title ?? "",
          themeName: deckThemeOf(model).name,
          slides: model ? summariseSlides(model) : [],
        }),
      };
    }
    const picked = source.value;
    return {
      purpose: DESIGN_PURPOSE,
      system: DESIGN_SYSTEM_PROMPT,
      user: designUserMessage({
        brief: brief.value,
        audience: audience.value,
        length: length.value,
        tone: tone.value,
        ...(picked ? { source: { title: picked.title, text: picked.text } } : {}),
      }),
    };
  });

  /** The live "what will be sent" count. */
  const bytes = computed(() => messageBytes(request.value.system, request.value.user));

  const slideCount = computed(() => options.model()?.slides.length ?? 0);
  /** New: a brief or a source to design from. Redesign: slides to arrange. */
  const canSend = computed(() => (mode === "redesign" ? slideCount.value > 0 : brief.value.trim() !== "" || source.value !== null));
  /** The model-free path: a source to outline, or a deck to arrange by hand. */
  const canOutline = computed(() => (mode === "redesign" ? slideCount.value > 0 : source.value !== null));

  return {
    brief,
    audience,
    length,
    tone,
    sourceChoice,
    pastedText,
    keepPrompt,
    sourceOptions,
    source,
    truncation,
    request,
    bytes,
    slideCount,
    canSend,
    canOutline,
    audiences: [...AUDIENCES],
    lengths: [...LENGTHS],
    tones: [...TONES],
  };
}

export type DesignerBrief = ReturnType<typeof useDesignerBrief>;
