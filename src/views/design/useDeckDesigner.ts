// The deck designer's session (SPEC-115): the three-step Brief → Plan → Build
// flow behind DeckDesignerDialog, for both entry points — the library's
// "Presentation from a description…" (a new deck, mode "new") and a deck's
// "Redesign…" (the open deck's order, layouts and theme, mode "redesign").
// The Brief's state is useDesignerBrief's, the plan step's useDesignerPlan's;
// this composable runs the flow between them.
//
// Everything a model says is untrusted data. A reply is read with JSON.parse
// inside a try and nothing else, then checked by deterministic code before the
// plan step shows it (lib/design/reply.ts). A reply that fails gets exactly
// one repair — the reply sent back with the validator's own message — and a
// second failure ends the flow with the error, the raw text and "Try again".
// Build makes no request: the plan the user last saw is validated and built
// through SPEC-106's deckFromDescription, stored through storePresentation as
// a new library entry, and nothing is stored before then — or at all when
// anything fails or is cancelled.
//
// The gateway is reached only through composables/useModelGateway.ts, by an
// import() taken on the first request (designerLibrary.ts): opening the
// dialog, outlining a source or rearranging by hand loads none of it. And no
// request is possible while the deck-designer experiment is off.
import { computed, ref, shallowRef, type Ref } from "vue";
import { useRouter } from "vue-router";
import type { ModelResult } from "../../lib/ai/modelGateway";
import {
  ModelConsentDeniedError,
  ModelNotConfiguredError,
  ModelUnreachableError,
  describeModelError,
  isAbortError,
} from "../../lib/ai/errors";
import type { ProviderMode } from "../../lib/ai/providerConfig";
import { isExperimentEnabled } from "../../lib/experiments";
import { loadPreferences } from "../../lib/preferences";
import { editorPathFor } from "../../lib/documentKinds";
import { uniqueTitle } from "../../lib/documentStore";
import { useToasts, type ToastAction } from "../../composables/useToasts";
import type { PresentationModel } from "../../lib/presentation/model";
import type { DeckDropRecord } from "../../lib/generate/deckDrops";
import { DocumentGenerationError } from "../../lib/generate/errors";
import { validatePlan, type DeckPlan, type DeckRedesign } from "../../lib/design/plan";
import { applyRedesign, isRedesignNoop } from "../../lib/design/applyDesign";
import { buildPlannedDeck } from "../../lib/design/build";
import { OUTLINE_DESIGNER_MODEL, planFromOutline } from "../../lib/design/outlineDesigner";
import { repairUserMessage } from "../../lib/design/prompt";
import { checkPlanReply, checkRedesignReply, type CheckedReply } from "../../lib/design/reply";
import { loadModelGateway, storeLibrary, type DesignerGateway, type DesignerLibrary } from "./designerLibrary";
import { useDesignerBrief } from "./useDesignerBrief";
import { useDesignerPlan } from "./useDesignerPlan";

export type { DesignerGateway, DesignerLibrary, DesignerSourceEntry } from "./designerLibrary";
export type { EditablePlanSlide, RedesignRow } from "./useDesignerPlan";

export type DesignerMode = "new" | "redesign";
export type DesignerStep = "brief" | "plan" | "build";
/** Where the plan step stands: waiting for a reply (streaming, or its one
 * repair), showing a plan the user can edit, or showing why there is none. */
export type PlanPhase = "idle" | "streaming" | "repairing" | "ready" | "failed";
export type BuildPhase = "idle" | "building" | "built" | "failed";

/** The open deck, for Redesign: the view's model and its undo choreography
 * (SPEC-080's record → mutate → commit()). */
export interface RedesignHost {
  model: Ref<PresentationModel | null>;
  record: (label?: string) => void;
  commit: () => void;
  currentIndex?: Ref<number>;
}

export interface DeckDesignerDeps {
  mode: DesignerMode;
  host?: RedesignHost | null;
  onClose: () => void;
  loadGateway?: () => Promise<DesignerGateway>;
  isEnabled?: () => boolean;
  providerMode?: () => ProviderMode;
  library?: DesignerLibrary;
  navigate?: (path: string) => void;
  toast?: (message: string, action?: ToastAction) => void;
  now?: () => Date;
}

/** The steps, as the dialog shows them. */
export const DESIGNER_STEPS: readonly { id: DesignerStep; label: string }[] = Object.freeze([
  { id: "brief", label: "Brief" },
  { id: "plan", label: "Plan" },
  { id: "build", label: "Build" },
]);

function routerNavigate(): (path: string) => void {
  const router = useRouter();
  return (path) => {
    void router.push(path);
  };
}

export function useDeckDesigner(deps: DeckDesignerDeps) {
  const mode = deps.mode;
  const isEnabled = deps.isEnabled ?? (() => isExperimentEnabled("deck-designer"));
  const library = deps.library ?? storeLibrary();
  const navigate = deps.navigate ?? routerNavigate();
  const toast = deps.toast ?? useToasts().addToast;
  const now = deps.now ?? (() => new Date());
  const loadGateway = deps.loadGateway ?? loadModelGateway;
  const deckModel = () => deps.host?.model.value ?? null;

  const step = ref<DesignerStep>("brief");
  const phase = ref<PlanPhase>("idle");
  const providerOff = ref((deps.providerMode ?? (() => loadPreferences().modelProvider.mode))() === "off");

  const briefState = useDesignerBrief({ mode, model: deckModel, library });
  const planState = useDesignerPlan({ mode, model: deckModel, ready: () => phase.value === "ready" });

  const streamText = ref("");
  const error = ref<string | null>(null);
  /** The model's raw reply (the last one) — shown in designer-raw. */
  const raw = ref("");
  /** Why the provider gave no answer, when it is off, unreachable or declined. */
  const offline = ref<string | null>(null);
  const dropped = shallowRef<DeckDropRecord[]>([]);
  const generatedModel = ref("");
  const buildPhase = ref<BuildPhase>("idle");
  const built = shallowRef<{ id: string; title: string; slides: number } | null>(null);

  // --- requests ------------------------------------------------------------------
  let controller: AbortController | null = null;
  let requestToken = 0;

  function stopRequest(): void {
    requestToken += 1;
    controller?.abort();
    controller = null;
  }

  function resetReply(): void {
    error.value = null;
    raw.value = "";
    offline.value = null;
    dropped.value = [];
    streamText.value = "";
  }

  function checkReply(text: string): CheckedReply<DeckPlan | DeckRedesign> {
    return mode === "redesign" ? checkRedesignReply(text, briefState.slideCount.value) : checkPlanReply(text);
  }

  function accept(result: CheckedReply<DeckPlan | DeckRedesign> & { ok: true }, reply: ModelResult): void {
    if (mode === "redesign") {
      planState.seedRows(result.value as DeckRedesign);
    } else {
      planState.loadPlan(result.value as DeckPlan);
    }
    dropped.value = result.dropped;
    generatedModel.value = reply.model;
    raw.value = reply.text;
    phase.value = "ready";
  }

  /** Sends the brief: the reply streams into the plan step as plain text and
   * is checked on arrival — repaired once at most, then accepted or failed. */
  async function send(): Promise<void> {
    if (!briefState.canSend.value) {
      return;
    }
    stopRequest();
    resetReply();
    step.value = "plan";
    if (!isEnabled()) {
      // A disabled experiment cannot make a request, whatever is on screen.
      offline.value = "The deck designer is switched off in Preferences › Labs, so nothing was sent.";
      phase.value = "failed";
      return;
    }
    phase.value = "streaming";
    const token = requestToken;
    const ownController = new AbortController();
    controller = ownController;
    const base = briefState.request.value;
    const onDelta = (chunk: string) => {
      if (token === requestToken) {
        streamText.value += chunk;
      }
    };
    try {
      const gateway = await loadGateway();
      if (token !== requestToken) {
        return;
      }
      let reply = await gateway.stream({ ...base, json: true, signal: ownController.signal }, onDelta);
      let result = checkReply(reply.text);
      if (!result.ok && token === requestToken) {
        // The one repair: the reply goes back with the validator's message.
        phase.value = "repairing";
        streamText.value = "";
        const user = repairUserMessage(base.user, reply.text, result.error.message);
        reply = await gateway.stream({ ...base, user, json: true, signal: ownController.signal }, onDelta);
        result = checkReply(reply.text);
      }
      if (token !== requestToken) {
        return;
      }
      if (result.ok) {
        accept(result, reply);
      } else {
        error.value = result.error.message;
        raw.value = reply.text;
        phase.value = "failed";
      }
    } catch (failure) {
      if (token !== requestToken || isAbortError(failure)) {
        return;
      }
      const unanswered =
        failure instanceof ModelNotConfiguredError ||
        failure instanceof ModelConsentDeniedError ||
        failure instanceof ModelUnreachableError;
      if (unanswered) {
        offline.value = describeModelError(failure);
      } else {
        error.value = describeModelError(failure);
      }
      phase.value = "failed";
    } finally {
      if (controller === ownController) {
        controller = null;
      }
    }
  }

  /** "Try again": a new request from the same brief — never a silent retry. */
  function tryAgain(): void {
    void send();
  }

  /** The model-free path. New: "Outline this document" — the deterministic
   * designer over the picked source. Redesign: the plan step seeded from the
   * open deck as it stands. No request is made either way. */
  function outline(): void {
    if (!briefState.canOutline.value) {
      return;
    }
    stopRequest();
    resetReply();
    if (mode === "redesign") {
      planState.seedRows(null);
      generatedModel.value = "";
    } else {
      const picked = briefState.source.value;
      if (!picked) {
        return;
      }
      const validated = validatePlan(planFromOutline(picked.outline));
      planState.loadPlan(validated.plan);
      dropped.value = validated.dropped;
      generatedModel.value = OUTLINE_DESIGNER_MODEL;
    }
    step.value = "plan";
    phase.value = "ready";
  }

  // --- Build / Apply -------------------------------------------------------------------
  function close(): void {
    stopRequest();
    deps.onClose();
  }

  /** Build deck (New): no model request — the edited plan is validated and
   * built through SPEC-106, then stored as a new library entry, opened, and
   * offered back through an Undo toast that trashes it. */
  function build(): void {
    if (mode !== "new" || phase.value !== "ready" || planState.plan.slides.length === 0) {
      return;
    }
    stopRequest();
    step.value = "build";
    buildPhase.value = "building";
    error.value = null;
    let model: PresentationModel;
    try {
      const { plan: validated } = validatePlan(planState.currentPlan());
      model = buildPlannedDeck(validated, {
        generatedAt: now().toISOString(),
        generatedModel: generatedModel.value || OUTLINE_DESIGNER_MODEL,
        ...(briefState.keepPrompt.value ? { prompt: briefState.brief.value } : {}),
      });
    } catch (failure) {
      if (!(failure instanceof DocumentGenerationError)) {
        throw failure;
      }
      buildPhase.value = "failed";
      error.value = failure.message;
      if (raw.value === "") {
        raw.value = JSON.stringify(planState.currentPlan(), null, 2);
      }
      return;
    }
    model.title = uniqueTitle(model.title, library.titles());
    library.store(model);
    buildPhase.value = "built";
    built.value = { id: model.id, title: model.title, slides: model.slides.length };
    const id = model.id;
    navigate(editorPathFor("presentation", id));
    toast(`Created “${model.title}”`, {
      label: "Undo",
      run: () => {
        library.trash(id);
        navigate("/");
      },
    });
    close();
  }

  /** Apply redesign: one record → mutate → commit(), or nothing at all when
   * nothing would change. No library entry, no navigation, no toast, no
   * provenance. */
  function applyRedesignPlan(): void {
    const host = deps.host;
    const model = host?.model.value;
    if (mode !== "redesign" || !host || !model || phase.value !== "ready") {
      return;
    }
    const redesign = planState.currentRedesign();
    if (redesign.slides.length === model.slides.length && !isRedesignNoop(model, redesign)) {
      const current = host.currentIndex ? model.slides[host.currentIndex.value] : undefined;
      host.record("Redesign");
      applyRedesign(model, redesign);
      if (host.currentIndex && current) {
        host.currentIndex.value = Math.max(0, model.slides.indexOf(current));
      }
      host.commit();
    }
    close();
  }

  // --- moving between steps --------------------------------------------------------------
  /** Back: from the Brief it leaves (nothing is stored); from the plan it
   * returns to the Brief, stopping any request; from Build, to the plan. */
  function back(): void {
    if (step.value === "build") {
      // Build only ever starts from a ready plan, so the plan step is there.
      step.value = "plan";
      buildPhase.value = "idle";
      error.value = null;
      return;
    }
    if (step.value === "plan") {
      stopRequest();
      if (phase.value === "streaming" || phase.value === "repairing") {
        phase.value = "idle";
      }
      step.value = "brief";
      return;
    }
    close();
  }

  /** Cancel: aborts whatever is in flight through its AbortSignal and closes,
   * storing nothing. */
  function cancel(): void {
    close();
  }

  return {
    mode,
    steps: DESIGNER_STEPS,
    step,
    phase,
    providerOff,
    ...briefState,
    ...planState,
    streamText,
    repairing: computed(() => phase.value === "repairing"),
    error,
    raw,
    offline,
    dropped,
    generatedModel,
    buildPhase,
    built,
    send,
    tryAgain,
    outline,
    build,
    applyRedesign: applyRedesignPlan,
    back,
    cancel,
  };
}

export type DeckDesigner = ReturnType<typeof useDeckDesigner>;
