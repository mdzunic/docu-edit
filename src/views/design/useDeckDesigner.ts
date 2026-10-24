// The deck designer's session (SPEC-115): the three-step Brief → Plan → Build
// flow behind DeckDesignerDialog, for both entry points — the library's
// "Presentation from a description…" (a new deck, mode "new") and a deck's
// "Redesign…" (the open deck's order, layouts and theme, mode "redesign").
//
// Everything a model says is untrusted data. A reply is read with JSON.parse
// inside a try and nothing else, then checked by deterministic code before the
// plan step shows it — validatePlan + planToDescription + deckFromDescription
// for a plan, validateRedesign for a redesign. A reply that fails gets exactly
// one repair (the reply sent back with the validator's own message); a second
// failure ends the flow with the error, the raw text and "Try again". Build
// makes no request: the plan the user last saw is validated and built through
// SPEC-106's deckFromDescription, stored through storePresentation as a new
// library entry, and nothing is stored before then — or at all when anything
// fails or is cancelled.
//
// The gateway is reached only through composables/useModelGateway.ts, by an
// import() taken on the first request: opening the dialog, outlining a source
// or rearranging by hand loads none of it. And no request is possible while
// the deck-designer experiment is off.
import { computed, reactive, ref, shallowRef, type Ref } from "vue";
import { useRouter } from "vue-router";
import type { ModelRequest, ModelResult } from "../../lib/ai/modelGateway";
import {
  ModelConsentDeniedError,
  ModelNotConfiguredError,
  ModelUnreachableError,
  describeModelError,
  isAbortError,
} from "../../lib/ai/errors";
import { isExperimentEnabled } from "../../lib/experiments";
import { loadPreferences } from "../../lib/preferences";
import type { ProviderMode } from "../../lib/ai/providerConfig";
import { editorPathFor } from "../../lib/documentKinds";
import {
  getDocument,
  getPresentation,
  listDocuments,
  listTrashedDocuments,
  storePresentation,
  trashDocument,
  uniqueTitle,
} from "../../lib/documentStore";
import { useToasts, type ToastAction } from "../../composables/useToasts";
import type { PresentationModel, SlideLayoutKind } from "../../lib/presentation/model";
import { deckThemeOf } from "../../lib/presentation/theme";
import { deckFromDescription } from "../../lib/generate/deckDescription";
import type { DeckDropRecord } from "../../lib/generate/deckDrops";
import { DocumentGenerationError } from "../../lib/generate/errors";
import {
  MAX_PLANNED_SLIDES,
  PLAN_LAYOUT_KINDS,
  PLAN_SIZES,
  PLAN_THEME_NAMES,
  bulletLevel,
  bulletText,
  isPlanLayoutKind,
  makeBullet,
  planToDescription,
  validatePlan,
  validateRedesign,
  type DeckPlan,
  type DeckRedesign,
  type PlanLayoutKind,
  type PlannedBullet,
} from "../../lib/design/plan";
import {
  MAX_REDESIGN_TITLE_CHARS,
  applyRedesign,
  currentLayoutKind,
  isRedesignNoop,
  seedLayoutKind,
  slideTitleText,
} from "../../lib/design/applyDesign";
import { buildPlannedDeck } from "../../lib/design/build";
import { OUTLINE_DESIGNER_MODEL, planFromOutline, type OutlineSource } from "../../lib/design/outlineDesigner";
import { MAX_INDENT_LEVEL } from "../../lib/presentation/text";
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
  repairUserMessage,
  truncateSource,
} from "../../lib/design/prompt";
import { outlineFromDeck, outlineFromDocument, outlineFromText, outlineText } from "./outlineSource";

export type DesignerMode = "new" | "redesign";
export type DesignerStep = "brief" | "plan" | "build";
/** Where the plan step stands: waiting for a reply (streaming, or its one
 * repair), showing a plan the user can edit, or showing why there is none. */
export type PlanPhase = "idle" | "streaming" | "repairing" | "ready" | "failed";
export type BuildPhase = "idle" | "building" | "built" | "failed";

/** The gateway calls the designer makes — `stream` only: the plan step
 * streams its reply so a slow local model shows progress. */
export interface DesignerGateway {
  stream(request: ModelRequest, onDelta: (chunk: string) => void): Promise<ModelResult>;
}

/** A library entry the New flow can take as its source. */
export interface DesignerSourceEntry {
  id: string;
  title: string;
  kind: "document" | "deck";
}

/** The library as the designer uses it: what may be a source (and its
 * outline), the titles a new deck is numbered against, storing and trashing. */
export interface DesignerLibrary {
  sources(): DesignerSourceEntry[];
  outlineOf(id: string): OutlineSource | undefined;
  titles(): string[];
  store(model: PresentationModel): void;
  trash(id: string): void;
}

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

/** The steps, as the dialog shows them. */
export const DESIGNER_STEPS: readonly { id: DesignerStep; label: string }[] = Object.freeze([
  { id: "brief", label: "Brief" },
  { id: "plan", label: "Plan" },
  { id: "build", label: "Build" },
]);

export function bulletsToText(bullets: PlannedBullet[]): string {
  return bullets.map((bullet) => `${"  ".repeat(bulletLevel(bullet))}${bulletText(bullet)}`).join("\n");
}

export function textToBullets(text: string): PlannedBullet[] {
  return text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "")
    .map((line) => {
      const indent = /^[ \t]*/.exec(line)?.[0].replace(/\t/g, "  ").length ?? 0;
      return makeBullet(line.trim(), Math.min(Math.floor(indent / 2), MAX_INDENT_LEVEL));
    });
}

function routerNavigate(): (path: string) => void {
  const router = useRouter();
  return (path) => {
    void router.push(path);
  };
}

/** The real library: text documents and decks from the store. */
export function storeLibrary(): DesignerLibrary {
  return {
    sources() {
      return listDocuments()
        .filter((entry) => entry.kind === "text" || entry.kind === "presentation")
        .map((entry) => ({ id: entry.id, title: entry.title, kind: entry.kind === "presentation" ? "deck" : "document" }));
    },
    outlineOf(id) {
      const deck = getPresentation(id);
      if (deck) {
        return outlineFromDeck(deck);
      }
      const doc = getDocument(id);
      return doc ? outlineFromDocument(doc) : undefined;
    },
    titles() {
      return [...listDocuments(), ...listTrashedDocuments()].map((entry) => entry.title);
    },
    store: storePresentation,
    trash: trashDocument,
  };
}

async function loadModelGateway(): Promise<DesignerGateway> {
  const { useModelGateway } = await import("../../composables/useModelGateway");
  return useModelGateway();
}

type Checked<T> = { ok: true; value: T; dropped: DeckDropRecord[] } | { ok: false; error: DocumentGenerationError };

/** The one way a reply is read: JSON.parse inside a try. */
function parseReply(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new DocumentGenerationError(`reply: not valid JSON (${error instanceof Error ? error.message : String(error)})`);
  }
}

function checked<T>(run: () => { value: T; dropped: DeckDropRecord[] }): Checked<T> {
  try {
    return { ok: true, ...run() };
  } catch (error) {
    if (error instanceof DocumentGenerationError) {
      return { ok: false, error };
    }
    throw error;
  }
}

export function useDeckDesigner(deps: DeckDesignerDeps) {
  const mode = deps.mode;
  const isEnabled = deps.isEnabled ?? (() => isExperimentEnabled("deck-designer"));
  const providerMode = deps.providerMode ?? (() => loadPreferences().modelProvider.mode);
  const library = deps.library ?? storeLibrary();
  const navigate = deps.navigate ?? routerNavigate();
  const toast = deps.toast ?? useToasts().addToast;
  const now = deps.now ?? (() => new Date());
  const loadGateway = deps.loadGateway ?? loadModelGateway;

  const step = ref<DesignerStep>("brief");

  // --- Brief -------------------------------------------------------------------
  const brief = ref("");
  const audience = ref(AUDIENCES[0].value);
  const length = ref(LENGTHS[1].value);
  const tone = ref(TONES[0].value);
  /** "" (nothing), "paste", or a library entry's id. */
  const sourceChoice = ref("");
  const pastedText = ref("");
  const keepPrompt = ref(false);
  const providerOff = ref(providerMode() === "off");

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

  function deckSummaries() {
    const model = deps.host?.model.value;
    return (model?.slides ?? []).map((slide, index) => ({
      index,
      title: Array.from(slideTitleText(slide)).slice(0, MAX_REDESIGN_TITLE_CHARS).join(""),
      layout: (model ? currentLayoutKind(model, slide) : undefined) ?? "other",
    }));
  }

  /** The request the brief describes, exactly as it would be handed to the
   * gateway (which wraps it in the body the consent dialog shows). */
  const request = computed(() => {
    if (mode === "redesign") {
      const model = deps.host?.model.value;
      return {
        purpose: REDESIGN_PURPOSE,
        system: REDESIGN_SYSTEM_PROMPT,
        user: redesignUserMessage({
          brief: brief.value,
          audience: audience.value,
          tone: tone.value,
          deckTitle: model?.title ?? "",
          themeName: deckThemeOf(model).name,
          slides: deckSummaries(),
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

  const slideCount = computed(() => deps.host?.model.value?.slides.length ?? 0);
  const canSend = computed(() =>
    mode === "redesign" ? slideCount.value > 0 : brief.value.trim() !== "" || source.value !== null,
  );
  const canOutline = computed(() => (mode === "redesign" ? slideCount.value > 0 : source.value !== null));

  // --- Plan --------------------------------------------------------------------
  const phase = ref<PlanPhase>("idle");
  const streamText = ref("");
  /** Whether the reply being read is the one repair. */
  const repairing = ref(false);
  const error = ref<string | null>(null);
  /** The model's raw reply (the last one) — shown in designer-raw. */
  const raw = ref("");
  /** Why the provider gave no answer, when it is off, unreachable or declined. */
  const offline = ref<string | null>(null);
  const dropped = shallowRef<DeckDropRecord[]>([]);
  const generatedModel = ref("");

  let nextKey = 1;
  const plan = reactive<{ title: string; theme: string; size: string; slides: EditablePlanSlide[] }>({
    title: "",
    theme: PLAN_THEME_NAMES[0],
    size: PLAN_SIZES[0],
    slides: [],
  });

  const rows = ref<RedesignRow[]>([]);
  /** "" = Keep current theme. */
  const redesignTheme = ref("");

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
   * past its subtitle above all — listed before Build. */
  const pendingDrops = computed<DeckDropRecord[]>(() => {
    if (mode !== "new" || phase.value !== "ready" || plan.slides.length === 0) {
      return [];
    }
    try {
      return validatePlan(currentPlan()).dropped;
    } catch {
      return [];
    }
  });

  function seedRows(redesign: DeckRedesign | null): void {
    const model = deps.host?.model.value;
    if (!model) {
      rows.value = [];
      return;
    }
    const order = redesign?.slides ?? model.slides.map((_, index) => ({ source: index, layout: null }));
    rows.value = order.map((entry) => {
      const slide = model.slides[entry.source];
      const kind = currentLayoutKind(model, slide);
      return {
        key: nextKey++,
        source: entry.source,
        title: slideTitleText(slide),
        currentKind: kind ?? null,
        layout: entry.layout ?? seedLayoutKind(model, slide),
      };
    });
    redesignTheme.value = redesign?.theme ?? "";
  }

  // --- requests ----------------------------------------------------------------
  let controller: AbortController | null = null;
  let requestToken = 0;

  function stopRequest(): void {
    requestToken += 1;
    controller?.abort();
    controller = null;
  }

  function checkReply(text: string): Checked<DeckPlan | DeckRedesign> {
    const parse = () => parseReply(text);
    if (mode === "redesign") {
      return checked(() => {
        const { redesign, dropped: drops } = validateRedesign(parse(), slideCount.value);
        return { value: redesign, dropped: drops };
      });
    }
    return checked(() => {
      const { plan: validated, dropped: drops } = validatePlan(parse());
      deckFromDescription(planToDescription(validated));
      return { value: validated, dropped: drops };
    });
  }

  function accept(result: Checked<DeckPlan | DeckRedesign> & { ok: true }, reply: ModelResult): void {
    if (mode === "redesign") {
      seedRows(result.value as DeckRedesign);
    } else {
      loadPlan(result.value as DeckPlan);
    }
    dropped.value = result.dropped;
    generatedModel.value = reply.model;
    raw.value = reply.text;
    phase.value = "ready";
  }

  function fail(message: string, text: string): void {
    error.value = message;
    raw.value = text;
    phase.value = "failed";
  }

  /** Sends the brief: the reply streams into the plan step as plain text and
   * is checked on arrival — repaired once at most, then accepted or failed. */
  async function send(): Promise<void> {
    if (!canSend.value) {
      return;
    }
    stopRequest();
    step.value = "plan";
    error.value = null;
    raw.value = "";
    offline.value = null;
    dropped.value = [];
    streamText.value = "";
    repairing.value = false;
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
    const base = request.value;
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
        repairing.value = true;
        phase.value = "repairing";
        streamText.value = "";
        const repair = { ...base, user: repairUserMessage(base.user, reply.text, result.error.message) };
        reply = await gateway.stream({ ...repair, json: true, signal: ownController.signal }, onDelta);
        result = checkReply(reply.text);
      }
      if (token !== requestToken) {
        return;
      }
      if (result.ok) {
        accept(result, reply);
      } else {
        fail(result.error.message, reply.text);
      }
    } catch (failure) {
      if (token !== requestToken || isAbortError(failure)) {
        return;
      }
      if (
        failure instanceof ModelNotConfiguredError ||
        failure instanceof ModelConsentDeniedError ||
        failure instanceof ModelUnreachableError
      ) {
        offline.value = describeModelError(failure);
        phase.value = "failed";
        return;
      }
      fail(describeModelError(failure), "");
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
    if (!canOutline.value) {
      return;
    }
    stopRequest();
    error.value = null;
    raw.value = "";
    offline.value = null;
    streamText.value = "";
    repairing.value = false;
    if (mode === "redesign") {
      seedRows(null);
      dropped.value = [];
      generatedModel.value = "";
    } else {
      const picked = source.value;
      if (!picked) {
        return;
      }
      const validated = validatePlan(planFromOutline(picked.outline));
      loadPlan(validated.plan);
      dropped.value = validated.dropped;
      generatedModel.value = OUTLINE_DESIGNER_MODEL;
    }
    step.value = "plan";
    phase.value = "ready";
  }

  // --- plan editing (New) --------------------------------------------------------
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

  // --- plan editing (Redesign) -----------------------------------------------------
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

  // --- Build / Apply -----------------------------------------------------------------
  const buildPhase = ref<BuildPhase>("idle");
  const built = shallowRef<{ id: string; title: string; slides: number } | null>(null);

  function close(): void {
    stopRequest();
    deps.onClose();
  }

  /** Build deck (New): no model request — the edited plan is validated and
   * built through SPEC-106, then stored as a new library entry. */
  function build(): void {
    if (mode !== "new" || phase.value !== "ready" || plan.slides.length === 0) {
      return;
    }
    stopRequest();
    step.value = "build";
    buildPhase.value = "building";
    error.value = null;
    let model: PresentationModel;
    try {
      const { plan: validated } = validatePlan(currentPlan());
      model = buildPlannedDeck(validated, {
        generatedAt: now().toISOString(),
        generatedModel: generatedModel.value || OUTLINE_DESIGNER_MODEL,
        ...(keepPrompt.value ? { prompt: brief.value } : {}),
      });
    } catch (failure) {
      if (!(failure instanceof DocumentGenerationError)) {
        throw failure;
      }
      buildPhase.value = "failed";
      error.value = failure.message;
      if (raw.value === "") {
        raw.value = JSON.stringify(currentPlan(), null, 2);
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
    const redesign: DeckRedesign = {
      theme: redesignTheme.value === "" ? null : redesignTheme.value,
      slides: rows.value.map((row) => ({ source: row.source, layout: row.layout })),
    };
    if (rows.value.length === model.slides.length && !isRedesignNoop(model, redesign)) {
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

  // --- navigation between steps --------------------------------------------------------
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

  const layoutOptions = PLAN_LAYOUT_KINDS.map((kind) => ({ value: kind, label: kind }));

  return {
    mode,
    steps: DESIGNER_STEPS,
    step,
    // brief
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
    canSend,
    canOutline,
    providerOff,
    audiences: AUDIENCES,
    lengths: LENGTHS,
    tones: TONES,
    // plan
    phase,
    streamText,
    repairing,
    error,
    raw,
    offline,
    dropped,
    pendingDrops,
    generatedModel,
    plan,
    currentPlan,
    layoutOptions,
    themeNames: PLAN_THEME_NAMES,
    sizes: PLAN_SIZES,
    setPlanField,
    setSlideField,
    moveSlide,
    deleteSlide,
    addSlide,
    canAddSlide,
    rows,
    redesignTheme,
    setRowLayout,
    moveRow,
    setRedesignTheme,
    // build
    buildPhase,
    built,
    // actions
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
