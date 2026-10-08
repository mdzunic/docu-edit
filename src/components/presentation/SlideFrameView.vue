<script setup lang="ts">
// One slide frame rendered absolutely-positioned in slide space (SPEC-079):
// text → nested divs/spans, images → <img> from the media resolver, shapes →
// a devicePixelRatio-crisp <canvas> painted by paintShapeIntoCanvas plus a
// centred text overlay, groups → this component recursively, foreign → a
// labelled dashed placeholder box. Runtime prop declarations, not
// defineProps<{...}>() generics — this repo lints .vue scripts with espree,
// which rejects TS-only syntax; every typed transform lives in slideRender.ts.
//
// SPEC-097: a text/shape frame's paragraphs render inside a `.text-body`
// (insets, anchor, columns, wrap and the autofit scale — slideTextFit.ts
// #bodyStyle) on a frame root carrying `data-frame-id` (top-level frames
// only, so `frameTextHost` never finds a filmstrip thumbnail). List markers
// are derived (slideLists.ts), linked runs render as `<a data-testid=
// "slide-link">`, and a "shrink" body re-renders at its own MEASURED scale
// held in a local ref — measuring never writes the model.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, useTemplateRef, watch } from "vue";
import { cmToPx } from "../../lib/core/units";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import { paintChartIntoCanvas } from "../../lib/core/chartCanvas";
import { mediaElementVolume } from "../../lib/presentation/mediaFrames";
import { effectiveAnchor, effectiveAutofit } from "../../lib/presentation/textBody";
import { createCanPlayProbe, mediaRenderDescriptor } from "./mediaRender";
import { bodyStyle, measureFrameFit } from "./slideTextFit";
import {
  asResolver,
  asTheme,
  bodyNoWrap,
  chartVisualOf,
  asSlideFrame,
  defaultParagraphAlign,
  fitScaleDiffers,
  followEditorLink,
  frameLinkRender,
  framePositionStyle,
  groupChildren,
  groupOriginStyle,
  hasMarker,
  imageSrc,
  initialFitScale,
  isEmptyTextPlaceholder,
  isForeignVisible,
  foreignLabel,
  markedParagraphs,
  markerFallbackText,
  markerImageSrc,
  markerStyle,
  paragraphStyle,
  placeholderPrompt,
  runRenders,
  shapeVisualOf,
  noMeasuredScale,
  textBodyFrameOf,
  tableColumnBoundaries,
  tableColumnWidths,
  tableRenderRows,
  tableRowBoundaries,
  tableRowHeights,
  textFrameFill,
  frameAccessibleName,
  frameRole,
} from "./slideRender";

const props = defineProps({
  /** The SlideFrame to render (presentation model). */
  frame: { type: Object, required: true },
  /** Resolves a media id to a `data:` URI (or "" when unresolved). */
  resolveMedia: { type: Function, required: true },
  /** Selection/editing overlays (SPEC-081); only the top-level stage passes
   * these — group children render read-only. */
  selected: { type: Boolean, default: false },
  editing: { type: Boolean, default: false },
  /** SPEC-091: the {row,col} of the table cell in a session (null otherwise),
   * so that cell's `<td>` renders the in-cell contenteditable host. */
  editingCell: { type: Object, default: null },
  /** False for a group's children, which render read-only (no selection,
   * focus or test hooks) until ungroup promotes them (SPEC-082). */
  topLevel: { type: Boolean, default: true },
  /** True while an insertion tool is armed: frames ignore pointer-down so the
   * draw gesture on the surface underneath is not stolen as a move. */
  armed: { type: Boolean, default: false },
  /** SPEC-086 playback: an inline style (opacity/transform/clip-path/transition)
   * layered onto the frame's position during a build preview/show. Null in the
   * editor, where frames always render fully visible (AC-25). */
  playStyle: { type: Object, default: null },
  /** SPEC-095: true in SPEC-084 playback (not the editor), so an autoStart media
   * frame that is currently visible plays under the autoplay policy. */
  playbackMode: { type: Boolean, default: false },
  /** SPEC-094: the deck's theme, so scheme references resolve at render.
   * Null/absent = Office (`DEFAULT_THEME`). */
  theme: { type: Object, default: null },
});

const emit = defineEmits([
  "select",
  "start-edit",
  "start-edit-cell",
  "cell-select",
  "table-border-pointerdown",
  "frame-pointerdown",
  "edit-mounted",
  "edit-unmounted",
]);

const frame = computed(() => asSlideFrame(props.frame));
const deckTheme = computed(() => asTheme(props.theme));
const kind = computed(() => frame.value.kind);
// AC-51/52/53: each top-level frame is a focusable element with a role, an
// accessible name and the `slide-frame` test id; children carry none.
const testId = computed(() => (props.topLevel ? "slide-frame" : undefined));
const role = computed(() => (props.topLevel ? frameRole(frame.value) : undefined));
const ariaName = computed(() => (props.topLevel ? frameAccessibleName(frame.value) : undefined));
const tabIndex = computed(() => (props.topLevel ? 0 : undefined));
const position = computed(() => framePositionStyle(frame.value));
const paragraphs = computed(() => markedParagraphs(frame.value));
const textFill = computed(() => textFrameFill(frame.value, deckTheme.value));
const children = computed(() => groupChildren(frame.value));
const groupOrigin = computed(() => groupOriginStyle(frame.value));
const foreignVisible = computed(() => isForeignVisible(frame.value));
const tableColumns = computed(() => tableColumnWidths(frame.value));
const tableRows = computed(() => tableRenderRows(frame.value, deckTheme.value));
const tableRowHeightsPx = computed(() => tableRowHeights(frame.value));
// AC-8a/b: a selected (not editing) table exposes drag handles on its internal
// column/row boundaries; suppressed while a text session or insertion is active.
const showResizeHandles = computed(
  () => kind.value === "table" && props.topLevel && props.selected && !props.editing && !props.armed,
);
const colBoundaries = computed(() => tableColumnBoundaries(frame.value));
const rowBoundaries = computed(() => tableRowBoundaries(frame.value));
const foreignText = computed(() => foreignLabel(frame.value));
const imgSrc = computed(() => imageSrc(frame.value, asResolver(props.resolveMedia)));
const showPrompt = computed(() => !props.editing && isEmptyTextPlaceholder(frame.value));
const promptText = computed(() => placeholderPrompt(frame.value));
const resolver = computed(() => asResolver(props.resolveMedia));
// `frameTextHost(frameId)` finds a rendered body by this attribute; only the
// interactive stage's top-level frames carry it (never a thumbnail).
const frameIdAttr = computed(() => (props.topLevel ? frame.value.id : undefined));
const paragraphAlign = computed(() => defaultParagraphAlign(frame.value));

// --- links (SPEC-097 AC-47..50) ---------------------------------------------
// Anchors are never navigated by the browser on any surface. In playback the
// click bubbles to PresentationPlayView#onStageClick, which follows it (and
// does not advance); on the editor stage a plain click stays inert (so it
// still selects / double-click still starts a text session) and Cmd/Ctrl+
// click follows; a thumbnail's links are fully inert. Only playback links are
// focusable — elsewhere they would add tab stops inside a frame or thumbnail.
const linkTabIndex = computed(() => (props.playbackMode ? undefined : -1));
const frameLink = computed(() => frameLinkRender(frame.value));
// A whole-frame link is a clickable overlay covering the box: in playback for
// every linkable kind, and on the editor stage for image/shape frames (Cmd/
// Ctrl+click follows; a text frame's own runs already carry the frame link,
// and an overlay there would steal the caret's double-click).
const showFrameLink = computed(
  () =>
    frameLink.value !== null &&
    !props.editing &&
    (props.playbackMode || (props.topLevel && (kind.value === "image" || kind.value === "shape"))),
);
function onLinkClick(event = new MouseEvent("click")) {
  event.preventDefault();
  if (props.playbackMode || !props.topLevel) return;
  if (!(event.metaKey || event.ctrlKey)) return;
  const el = event.currentTarget;
  if (el instanceof Element) followEditorLink(el);
}

// --- text body + autofit (SPEC-097 AC-11..15, 37..40) -----------------------
// The body renders from the STORED scale; after mount/update a "shrink" body
// measures itself (slideTextFit.ts#measureFrameFit) and, if its measurement
// disagrees, re-renders at the measured scale through this LOCAL ref. The
// model is never written here (AC-40/42) — only a command that already called
// markSlideEdited stores a scale. jsdom measures null: keep what is stored.
const measuredScale = shallowRef(noMeasuredScale());
const currentScale = computed(() => measuredScale.value ?? initialFitScale(frame.value));
const textBodyStyle = computed(() => {
  const body = textBodyFrameOf(frame.value);
  return body ? bodyStyle(body, currentScale.value) : {};
});
const noWrap = computed(() => bodyNoWrap(frame.value));
const anchoredBody = computed(() => {
  const body = textBodyFrameOf(frame.value);
  return body !== null && effectiveAnchor(body) !== "top";
});
const textBodyEl = useTemplateRef("textBodyEl");
function remeasure() {
  const body = textBodyFrameOf(frame.value);
  if (!body || props.editing) return;
  if (effectiveAutofit(body) !== "shrink") {
    measuredScale.value = null;
    return;
  }
  const host = textBodyEl.value;
  const fit = measureFrameFit(body, host instanceof HTMLElement ? host : null);
  if (!fit) return;
  measuredScale.value = fitScaleDiffers(fit.scale, initialFitScale(body)) ? fit.scale : null;
}
function scheduleRemeasure() {
  void nextTick(remeasure);
}
let unmounted = false;
onMounted(() => {
  scheduleRemeasure();
  // Web fonts settling change every measurement; measure once more then.
  const fonts = typeof document !== "undefined" ? document.fonts : undefined;
  if (fonts && fonts.ready && typeof fonts.ready.then === "function") {
    void fonts.ready.then(() => {
      if (!unmounted) remeasure();
    });
  }
});
onBeforeUnmount(() => {
  unmounted = true;
});
watch([() => props.frame, deckTheme], scheduleRemeasure, { deep: true });
watch(
  () => props.editing,
  (isEditing) => {
    if (!isEditing) scheduleRemeasure();
  },
);

// --- media frame (SPEC-095) -------------------------------------------------
// Nothing auto-plays in the editor (AC-19); activating the badge (click, or
// Enter on the selected frame) swaps in native controls that play inline
// (AC-20). Local state resets on deselect and — since a slide change hands the
// frame a fresh object identity — on slide change too (AC-21).
const mediaEl = useTemplateRef("mediaEl");
const mediaActivated = ref(false);
const mediaRuntimeError = ref(false);
const mediaBlocked = ref(false);
const mediaDescriptor = computed(() => {
  const f = frame.value;
  if (f.kind !== "media") return null;
  return mediaRenderDescriptor(f, asResolver(props.resolveMedia), createCanPlayProbe(f.media), mediaRuntimeError.value);
});
const mediaSrc = computed(() => {
  const f = frame.value;
  return f.kind === "media" && f.mediaId ? asResolver(props.resolveMedia)(f.mediaId) : "";
});
const mediaElementVol = computed(() => {
  const f = frame.value;
  return f.kind === "media" ? mediaElementVolume(f.volume) : 1;
});
const mediaLoop = computed(() => frame.value.kind === "media" && frame.value.loop === true);
const mediaMuted = computed(() => frame.value.kind === "media" && frame.value.muted === true);
// A frame hidden by a pending entrance build (or after its exit) must not play
// (AC-40/41); its build style removes it from view.
const mediaBuildHidden = computed(() => {
  const s = props.playStyle;
  return !!s && (s.visibility === "hidden" || s.opacity === "0");
});
// In SPEC-084 playback, an autoStart clip that is visible plays; entering
// playback is a user gesture, so play() is allowed (AC-35), but an unmuted clip
// reached without a gesture rejects and degrades to media-blocked (AC-36).
const mediaShouldAutoStart = computed(
  () =>
    props.playbackMode &&
    frame.value.kind === "media" &&
    frame.value.autoStart === true &&
    !mediaBuildHidden.value &&
    !!mediaDescriptor.value?.playable,
);
function activateMedia() {
  if (mediaDescriptor.value?.playable) {
    mediaBlocked.value = false;
    mediaActivated.value = true;
  }
}
function startPlayback() {
  mediaBlocked.value = false;
  mediaActivated.value = true;
  void nextTick(attemptPlay);
}
function attemptPlay() {
  const el = mediaEl.value;
  if (!(el instanceof HTMLMediaElement)) return;
  el.muted = mediaMuted.value;
  el.volume = mediaElementVol.value;
  const played = el.play();
  if (played && typeof played.then === "function") {
    // A muted clip is never blocked (AC-37); an unmuted policy rejection degrades
    // to the media-blocked affordance that plays on the next input (AC-36).
    played.catch((error) => {
      if (error && error.name === "NotAllowedError" && !el.muted) {
        mediaBlocked.value = true;
        mediaActivated.value = false;
      }
    });
  }
}
function onMediaError() {
  mediaRuntimeError.value = true;
  mediaActivated.value = false;
}
function applyMediaVolume(event = new Event("loadedmetadata")) {
  const el = event.target;
  if (el instanceof HTMLMediaElement) {
    el.volume = mediaElementVol.value;
  }
}
watch(mediaShouldAutoStart, (should) => {
  if (should) startPlayback();
}, { immediate: true });
// An exit effect (or any build) that hides the target pauses it (AC-41).
watch(mediaBuildHidden, (hidden) => {
  if (hidden && props.playbackMode) {
    const el = mediaEl.value;
    if (el instanceof HTMLMediaElement) el.pause();
    mediaActivated.value = false;
  }
});
// Deselecting stops editor inline playback (AC-21).
watch(
  () => props.selected,
  (sel) => {
    if (!sel && !props.playbackMode) {
      mediaActivated.value = false;
    }
  },
);

const canvasEl = useTemplateRef("canvasEl");
const chartCanvasEl = useTemplateRef("chartCanvasEl");
const editHost = useTemplateRef("editHost");
const tableFrameEl = useTemplateRef("tableFrameEl");

// The in-cell contenteditable host is created/destroyed as the editing cell
// changes (double-click, Enter, Tab between cells). Hand the freshly-mounted
// host to the editing composable — the same edit-mounted/edit-unmounted
// contract the text frame uses (SPEC-091 AC-6a).
const editingCellKey = computed(() =>
  kind.value === "table" && props.editing && props.editingCell
    ? `${props.editingCell.row}:${props.editingCell.col}`
    : null,
);
watch(editingCellKey, (key) => {
  if (key) {
    void nextTick(() => {
      const el = tableFrameEl.value?.querySelector(".cell-edit");
      if (el instanceof HTMLElement) emit("edit-mounted", el);
    });
  } else {
    emit("edit-unmounted");
  }
});

/** True for the one owner cell currently in an editing session. */
function isEditingCell(cell = { rowIndex: 0, colIndex: 0 }) {
  return (
    props.editing &&
    props.editingCell != null &&
    props.editingCell.row === cell.rowIndex &&
    props.editingCell.col === cell.colIndex
  );
}

function onCellMousedown(row = 0, col = 0, event = new MouseEvent("mousedown")) {
  // Anchor the panel/session on this cell and select the frame; stop here so a
  // click inside the table never bubbles to the surface deselect (AC-9c).
  event.stopPropagation();
  emit("cell-select", frame.value.id, row, col);
}

function onCellDblclick(row = 0, col = 0, event = new MouseEvent("dblclick")) {
  event.stopPropagation();
  emit("start-edit-cell", frame.value.id, row, col, { x: event.clientX, y: event.clientY });
}

function onBoundaryPointerdown(axis = "col", index = 0, event = new PointerEvent("pointerdown")) {
  // Stop here so a boundary drag never also starts a frame move (AC-8a/b).
  event.stopPropagation();
  emit("table-border-pointerdown", frame.value.id, axis, index, event);
}

// The contenteditable host is created/destroyed with `editing`; hand it to the
// editing composable (which populates it and attaches the caret) once mounted.
watch(
  () => props.editing,
  (isEditing) => {
    if (isEditing) {
      void nextTick(() => {
        if (editHost.value) emit("edit-mounted", editHost.value);
      });
    } else {
      emit("edit-unmounted");
    }
  },
);

function onFrameMousedown(event = new MouseEvent("mousedown")) {
  // Select on pointer-down (before focus moves); editing swallows its own.
  // The event rides along so Shift+click toggling (SPEC-096) is not
  // collapsed by this trailing mousedown.
  if (props.editing) return;
  event.stopPropagation();
  emit("select", frame.value.id, event);
}

function onFrameDblclick(event = new MouseEvent("dblclick")) {
  // A text frame opens a caret session; a chart frame opens its data sheet
  // (SPEC-092 AC-14) — the parent's start-edit handler routes by frame kind.
  if (frame.value.kind !== "text" && frame.value.kind !== "chart") return;
  event.stopPropagation();
  // Pass the click's viewport coordinates so the session drops the caret at the
  // click point rather than the end of the frame (AC-1).
  emit("start-edit", frame.value.id, { x: event.clientX, y: event.clientY });
}

function onFramePointerdown(event = new PointerEvent("pointerdown")) {
  // Begin a move drag (SPEC-082); the manipulation composable selects and
  // captures. Editing sessions, group children, and armed insertion (which
  // draws on the surface below) opt out.
  if (props.editing || !props.topLevel || props.armed) return;
  emit("frame-pointerdown", frame.value.id, event);
}

function paint() {
  const el = canvasEl.value;
  const visual = shapeVisualOf(frame.value, deckTheme.value);
  if (!el || !visual) {
    return;
  }
  try {
    paintShapeIntoCanvas(el, visual, cmToPx(frame.value.widthCm), cmToPx(frame.value.heightCm));
  } catch {
    // A host without a real 2D canvas (jsdom) can't paint; rendering must
    // never break the view.
  }
}

// The chart branch paints through the same devicePixelRatio-crisp canvas
// machinery as the shape branch (SPEC-092 AC-11); an unsupported chart draws
// paintChart's dashed placeholder (AC-12).
function paintChartFrame() {
  const el = chartCanvasEl.value;
  const visual = chartVisualOf(frame.value);
  if (!el || !visual) {
    return;
  }
  try {
    paintChartIntoCanvas(el, visual, cmToPx(frame.value.widthCm), cmToPx(frame.value.heightCm));
  } catch {
    // A host without a real 2D canvas (jsdom) can't paint; rendering must
    // never break the view.
  }
}

// paintShapeIntoCanvas/paintChartIntoCanvas size the backing store at
// devicePixelRatio, so the frame stays crisp at dpr ≥ 2 (AC-43); repaint
// whenever the frame changes.
onMounted(paint);
onMounted(paintChartFrame);
watch([() => props.frame, deckTheme], () => {
  paint();
  paintChartFrame();
}, { deep: true });
</script>

<template>
  <!-- Text frame -->
  <div
    v-if="kind === 'text'"
    class="frame text-frame"
    :class="{ 'is-selected': selected, 'is-editing': editing, 'is-linked-frame': showFrameLink }"
    :style="[{ ...position, background: textFill }, playStyle]"
    :data-testid="testId"
    :data-frame-id="frameIdAttr"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @dblclick="onFrameDblclick"
    @pointerdown="onFramePointerdown"
  >
    <!-- Editing session: an empty contenteditable the editing composable
         owns, inside the same text body (insets, anchor, fit scale) so the
         session edits what the stage shows (SPEC-097) -->
    <div
      v-if="editing"
      class="text-body is-editing"
      :class="{ 'no-wrap': noWrap, 'is-anchored': anchoredBody }"
      :style="textBodyStyle"
    >
      <div
        ref="editHost"
        class="ce-host"
        contenteditable="true"
        spellcheck="false"
        data-testid="slide-text-edit"
      />
    </div>
    <!-- Empty placeholder prompt (never in the model/files/clipboard) -->
    <div
      v-else-if="showPrompt"
      class="text-body"
      :style="textBodyStyle"
    >
      <div
        class="ph-prompt"
        data-testid="slide-placeholder-prompt"
      >
        {{ promptText }}
      </div>
    </div>
    <div
      v-else
      ref="textBodyEl"
      class="text-body"
      :class="{ 'no-wrap': noWrap }"
      :style="textBodyStyle"
    >
      <div
        v-for="(entry, index) in paragraphs"
        :key="index"
        class="paragraph"
        :style="paragraphStyle(entry.para, paragraphAlign)"
        data-testid="slide-paragraph"
      >
        <span
          v-if="hasMarker(entry)"
          class="bullet-marker"
          aria-hidden="true"
          :style="markerStyle(entry, deckTheme)"
        ><img
          v-if="markerImageSrc(entry, resolver)"
          class="bullet-image"
          :src="markerImageSrc(entry, resolver)"
          alt=""
        ><template v-else>{{ markerFallbackText(entry) }}</template></span>
        <template
          v-for="(run, runIndex) in runRenders(entry.para, frame, deckTheme)"
          :key="runIndex"
        >
          <a
            v-if="run.link"
            class="slide-link"
            data-testid="slide-link"
            draggable="false"
            :href="run.link.href"
            :title="run.link.tooltip"
            :data-link-href="run.link.external || undefined"
            :data-link-slide="run.link.slide || undefined"
            :tabindex="linkTabIndex"
            :style="run.style"
            @click="onLinkClick"
          >{{ run.text }}</a>
          <span
            v-else
            :style="run.style"
          >{{ run.text }}</span>
        </template>
      </div>
    </div>
    <a
      v-if="showFrameLink && frameLink"
      class="frame-link"
      data-testid="slide-link"
      draggable="false"
      :href="frameLink.href"
      :title="frameLink.tooltip"
      :aria-label="frameLink.tooltip || 'Link'"
      :data-link-href="frameLink.external || undefined"
      :data-link-slide="frameLink.slide || undefined"
      :tabindex="linkTabIndex"
      @click="onLinkClick"
    />
  </div>

  <!-- Image frame -->
  <div
    v-else-if="kind === 'image'"
    class="frame"
    :class="{ 'is-selected': selected, 'is-linked-frame': showFrameLink }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :data-frame-id="frameIdAttr"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <img
      class="frame-image"
      :src="imgSrc"
      alt=""
    >
    <a
      v-if="showFrameLink && frameLink"
      class="frame-link"
      data-testid="slide-link"
      draggable="false"
      :href="frameLink.href"
      :title="frameLink.tooltip"
      :aria-label="frameLink.tooltip || 'Link'"
      :data-link-href="frameLink.external || undefined"
      :data-link-slide="frameLink.slide || undefined"
      :tabindex="linkTabIndex"
      @click="onLinkClick"
    />
  </div>

  <!-- Media frame (SPEC-095) -->
  <div
    v-else-if="kind === 'media' && mediaDescriptor"
    class="frame media-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
    @keydown.enter="activateMedia"
  >
    <div
      class="slide-media-frame"
      data-testid="slide-media-frame"
    >
      <!-- Playing inline / in playback: native controls (AC-20/35) -->
      <video
        v-if="mediaActivated && mediaDescriptor.noun === 'Video'"
        ref="mediaEl"
        class="media-controls"
        data-testid="media-controls"
        controls
        :src="mediaSrc"
        :loop="mediaLoop"
        :muted="mediaMuted"
        :poster="mediaDescriptor.posterSrc || undefined"
        @error="onMediaError"
        @loadedmetadata="applyMediaVolume"
      />
      <audio
        v-else-if="mediaActivated"
        ref="mediaEl"
        class="media-controls"
        data-testid="media-controls"
        controls
        :src="mediaSrc"
        :loop="mediaLoop"
        :muted="mediaMuted"
        @error="onMediaError"
        @loadedmetadata="applyMediaVolume"
      />
      <!-- Idle: poster or neutral plate, plus the badge / degrade captions -->
      <template v-else>
        <img
          v-if="mediaDescriptor.posterSrc"
          class="media-poster"
          data-testid="media-poster"
          :src="mediaDescriptor.posterSrc"
          alt=""
        >
        <div
          v-else
          class="media-plate"
        >
          <span
            class="media-glyph"
            aria-hidden="true"
          >▶</span>
          <span class="media-name">{{ mediaDescriptor.basename }}</span>
        </div>
        <button
          v-if="mediaBlocked"
          type="button"
          class="media-play media-blocked"
          data-testid="media-blocked"
          aria-label="Click to play"
          @click.stop="startPlayback"
        >
          ▶ Click to play
        </button>
        <button
          v-else-if="mediaDescriptor.playable"
          type="button"
          class="media-play"
          data-testid="media-play"
          aria-label="Play"
          @click.stop="activateMedia"
        >
          ▶
        </button>
        <div
          v-else-if="mediaDescriptor.mode === 'link'"
          class="media-caption media-link-placeholder"
          data-testid="media-link-placeholder"
        >
          <span class="media-name">{{ mediaDescriptor.basename }}</span>
          <span>{{ mediaDescriptor.caption }}</span>
        </div>
        <div
          v-else-if="mediaDescriptor.mode === 'oversize'"
          class="media-caption media-oversize"
          data-testid="media-oversize"
        >
          {{ mediaDescriptor.caption }}
        </div>
        <div
          v-else-if="mediaDescriptor.mode === 'unsupported'"
          class="media-caption media-unsupported"
          data-testid="media-unsupported"
        >
          {{ mediaDescriptor.caption }}
        </div>
      </template>
    </div>
  </div>

  <!-- Shape frame -->
  <div
    v-else-if="kind === 'shape'"
    class="frame shape-frame"
    :class="{ 'is-selected': selected, 'is-linked-frame': showFrameLink }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :data-frame-id="frameIdAttr"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <canvas
      ref="canvasEl"
      class="shape-canvas"
    />
    <!-- Shape text: the same paragraph styling as a text frame (markers,
         indents, spacing, links) in a text body anchored "middle" by default;
         a paragraph with no alignment of its own centres (SPEC-079). -->
    <div
      v-if="paragraphs.length > 0"
      ref="textBodyEl"
      class="text-body shape-text"
      :class="{ 'no-wrap': noWrap }"
      :style="textBodyStyle"
    >
      <div
        v-for="(entry, index) in paragraphs"
        :key="index"
        class="paragraph"
        :style="paragraphStyle(entry.para, paragraphAlign)"
      >
        <span
          v-if="hasMarker(entry)"
          class="bullet-marker"
          aria-hidden="true"
          :style="markerStyle(entry, deckTheme)"
        ><img
          v-if="markerImageSrc(entry, resolver)"
          class="bullet-image"
          :src="markerImageSrc(entry, resolver)"
          alt=""
        ><template v-else>{{ markerFallbackText(entry) }}</template></span>
        <template
          v-for="(run, runIndex) in runRenders(entry.para, frame, deckTheme)"
          :key="runIndex"
        >
          <a
            v-if="run.link"
            class="slide-link"
            data-testid="slide-link"
            draggable="false"
            :href="run.link.href"
            :title="run.link.tooltip"
            :data-link-href="run.link.external || undefined"
            :data-link-slide="run.link.slide || undefined"
            :tabindex="linkTabIndex"
            :style="run.style"
            @click="onLinkClick"
          >{{ run.text }}</a>
          <span
            v-else
            :style="run.style"
          >{{ run.text }}</span>
        </template>
      </div>
    </div>
    <a
      v-if="showFrameLink && frameLink"
      class="frame-link"
      data-testid="slide-link"
      draggable="false"
      :href="frameLink.href"
      :title="frameLink.tooltip"
      :aria-label="frameLink.tooltip || 'Link'"
      :data-link-href="frameLink.external || undefined"
      :data-link-slide="frameLink.slide || undefined"
      :tabindex="linkTabIndex"
      @click="onLinkClick"
    />
  </div>

  <!-- Chart frame (SPEC-092): painted into a dpr-crisp canvas, exactly as the
       shape frame; double-click opens the data sheet (the parent decides). -->
  <div
    v-else-if="kind === 'chart'"
    class="frame chart-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
    @dblclick="onFrameDblclick"
  >
    <canvas
      ref="chartCanvasEl"
      class="chart-canvas"
    />
  </div>

  <!-- Group frame: the box carries its real geometry (so a rotate transform
       spins about the group's own centre, SPEC-096 AC-45); the origin shim
       re-expresses the children's absolute slide-space coordinates inside it. -->
  <div
    v-else-if="kind === 'group'"
    class="frame group-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <div
      class="group-origin"
      :style="groupOrigin"
    >
      <SlideFrameView
        v-for="child in children"
        :key="child.id"
        :frame="child"
        :resolve-media="resolveMedia"
        :top-level="false"
        :theme="theme"
      />
    </div>
  </div>

  <!-- Table frame: a real <table> with a colgroup and merged cells (AC-4) -->
  <div
    v-else-if="kind === 'table'"
    ref="tableFrameEl"
    class="frame table-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <table
      class="slide-table"
      data-testid="slide-table"
    >
      <colgroup>
        <col
          v-for="(w, ci) in tableColumns"
          :key="ci"
          :style="{ width: w }"
        >
      </colgroup>
      <tbody>
        <tr
          v-for="(row, ri) in tableRows"
          :key="ri"
          :style="{ height: tableRowHeightsPx[ri] }"
        >
          <td
            v-for="cell in row"
            :key="cell.colIndex"
            class="slide-table-cell"
            :colspan="cell.colSpan"
            :rowspan="cell.rowSpan"
            :style="cell.style"
            :data-testid="`slide-table-cell-${cell.rowIndex}-${cell.colIndex}`"
            @mousedown="onCellMousedown(cell.rowIndex, cell.colIndex, $event)"
            @dblclick="onCellDblclick(cell.rowIndex, cell.colIndex, $event)"
          >
            <!-- Editing session: the same contenteditable host the text frame
                 uses, owned by the editing composable (SPEC-091 AC-6a) -->
            <div
              v-if="isEditingCell(cell)"
              class="ce-host cell-edit"
              contenteditable="true"
              spellcheck="false"
              data-testid="slide-table-cell-edit"
            />
            <template v-else>
              <div
                v-for="(entry, pi) in cell.paragraphs"
                :key="pi"
                class="paragraph"
                :style="paragraphStyle(entry.para)"
              >
                <span
                  v-if="hasMarker(entry)"
                  class="bullet-marker"
                  aria-hidden="true"
                  :style="markerStyle(entry, deckTheme)"
                ><img
                  v-if="markerImageSrc(entry, resolver)"
                  class="bullet-image"
                  :src="markerImageSrc(entry, resolver)"
                  alt=""
                ><template v-else>{{ markerFallbackText(entry) }}</template></span>
                <template
                  v-for="(run, runIndex) in runRenders(entry.para, frame, deckTheme)"
                  :key="runIndex"
                >
                  <a
                    v-if="run.link"
                    class="slide-link"
                    data-testid="slide-link"
                    draggable="false"
                    :href="run.link.href"
                    :title="run.link.tooltip"
                    :data-link-href="run.link.external || undefined"
                    :data-link-slide="run.link.slide || undefined"
                    :tabindex="linkTabIndex"
                    :style="run.style"
                    @click="onLinkClick"
                  >{{ run.text }}</a>
                  <span
                    v-else
                    :style="run.style"
                  >{{ run.text }}</span>
                </template>
              </div>
            </template>
          </td>
        </tr>
      </tbody>
    </table>
    <!-- Drag-resize handles on the internal grid boundaries (AC-8a/b) -->
    <template v-if="showResizeHandles">
      <div
        v-for="b in colBoundaries"
        :key="`col-${b.index}`"
        class="col-resize-handle"
        data-testid="table-col-resize"
        :data-boundary="b.index"
        :style="{ left: `${b.offsetPct}%` }"
        @pointerdown="onBoundaryPointerdown('col', b.index, $event)"
      />
      <div
        v-for="b in rowBoundaries"
        :key="`row-${b.index}`"
        class="row-resize-handle"
        data-testid="table-row-resize"
        :data-boundary="b.index"
        :style="{ top: `${b.offsetPct}%` }"
        @pointerdown="onBoundaryPointerdown('row', b.index, $event)"
      />
    </template>
  </div>

  <!-- Foreign placeholder box -->
  <div
    v-else-if="foreignVisible"
    class="frame foreign-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <span class="foreign-label">{{ foreignText }}</span>
  </div>
</template>

<style scoped>
.frame {
  box-sizing: border-box;
  overflow: visible;
}

.text-frame {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
}

.text-frame.is-editing,
.text-frame.is-selected {
  cursor: text;
}

/* Non-text frames show a move cursor when selected (AC-16); a text box keeps
   its text cursor so a click drops the caret. */
.shape-frame.is-selected,
.frame.is-selected:not(.text-frame) {
  cursor: move;
}

.text-frame.is-editing {
  outline: 2px solid var(--accent, #3b82f6);
  outline-offset: 1px;
}

/* SPEC-097: a text/shape frame's text body — insets (padding), anchor
   (justify-content), columns and the autofit scale arrive inline from
   slideTextFit.ts#bodyStyle. It fills the frame and overflows it visibly
   (autofit "none" renders exactly as SPEC-079 did). */
.text-body {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
  flex-direction: column;
  box-sizing: border-box;
}

.text-body.no-wrap .paragraph,
.text-body.no-wrap .ce-host {
  white-space: pre;
}

.ce-host {
  flex: 1 1 auto;
  min-height: 1em;
  outline: none;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  line-height: calc(1.2 * var(--line-scale, 1));
  /* the renderer's 18pt default (18 * 96/72), at the body's fit scale */
  font-size: calc(24px * var(--fit-scale, 1));
}

/* A middle/bottom-anchored body keeps its anchor while editing: the host
   hugs its content and the body's justify-content places it. */
.text-body.is-editing.is-anchored > .ce-host {
  flex: 0 0 auto;
}

.ce-host :deep(.ce-p) {
  margin: 0;
}

.ph-prompt {
  color: var(--text-secondary, #888);
  font-style: italic;
  font-size: 24px;
  line-height: 1.2;
  pointer-events: none;
  user-select: none;
}

.selection-box {
  position: absolute;
  inset: -2px;
  border: 1px solid var(--accent, #3b82f6);
  pointer-events: none;
}

.handle {
  position: absolute;
  width: 7px;
  height: 7px;
  background: #fff;
  border: 1px solid var(--accent, #3b82f6);
  border-radius: 1px;
}

.handle-1 { top: -4px; left: -4px; }
.handle-2 { top: -4px; left: calc(50% - 3px); }
.handle-3 { top: -4px; right: -4px; }
.handle-4 { top: calc(50% - 3px); right: -4px; }
.handle-5 { bottom: -4px; right: -4px; }
.handle-6 { bottom: -4px; left: calc(50% - 3px); }
.handle-7 { bottom: -4px; left: -4px; }
.handle-8 { top: calc(50% - 3px); left: -4px; }

.paragraph {
  margin: 0;
  line-height: 1.2;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

/* The marker box hangs in the paragraph's negative text-indent; its width,
   size and colour arrive inline (slideRender.ts#markerStyle). */
.bullet-marker {
  white-space: pre;
}

/* A picture bullet is sized to the line it marks. */
.bullet-image {
  height: 1em;
  width: auto;
  vertical-align: -0.12em;
}

/* SPEC-097 links: underlined in the run's own colour (inline), drawn above a
   whole-frame link overlay so a run's own link wins where it sits. */
.slide-link {
  position: relative;
  z-index: 2;
  cursor: inherit;
}

.frame-link {
  position: absolute;
  inset: 0;
  z-index: 1;
  display: block;
  cursor: inherit;
}

.frame-image {
  width: 100%;
  height: 100%;
  object-fit: fill;
  display: block;
}

/* Media frame (SPEC-095) */
.slide-media-frame {
  position: relative;
  width: 100%;
  height: 100%;
  overflow: hidden;
  background: #1f1f1f;
  display: flex;
  align-items: center;
  justify-content: center;
}
.slide-media-frame .media-poster,
.slide-media-frame .media-controls {
  width: 100%;
  height: 100%;
  display: block;
}
.slide-media-frame .media-poster {
  object-fit: contain;
}
.media-plate {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.4em;
  color: #e8e8e8;
  padding: 0.5em;
  text-align: center;
}
.media-plate .media-glyph {
  font-size: 2em;
  line-height: 1;
}
.media-name {
  font-size: 0.75em;
  word-break: break-all;
}
.media-play {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  width: 2.4em;
  height: 2.4em;
  border-radius: 50%;
  border: none;
  background: rgba(0, 0, 0, 0.55);
  color: #fff;
  font-size: 1em;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}
.media-caption {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.15em;
  padding: 0.3em;
  font-size: 0.7em;
  color: #eee;
  background: rgba(0, 0, 0, 0.55);
  text-align: center;
}

.shape-canvas {
  position: absolute;
  top: 0;
  left: 0;
}

.chart-frame {
  overflow: hidden;
}

.chart-canvas {
  position: absolute;
  top: 0;
  left: 0;
}

/* Shape text overlays the canvas; its anchor (default "middle") and insets
   arrive inline through the text body. Clicks fall through to the frame —
   except on a link, which must stay clickable. */
.shape-text {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.shape-text .slide-link {
  pointer-events: auto;
}

.group-frame {
  /* The group box paints nothing itself and occupies its real geometry (the
     inline position style) so a rotate transform spins about its own centre
     (SPEC-096 AC-45). The empty box area stays click-through — only the
     children (whose events bubble back up to this wrapper) hit. */
  overflow: visible;
  pointer-events: none;
}

/* Children hit-test normally; a NESTED group's own wrapper stays
   click-through like the outer one (the later rule wins the specificity tie). */
.group-frame .frame {
  pointer-events: auto;
}

.group-frame .group-frame {
  pointer-events: none;
}

.table-frame {
  overflow: hidden;
}

.slide-table {
  width: 100%;
  height: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}

.slide-table-cell {
  overflow-wrap: anywhere;
  word-break: break-word;
}

/* The in-cell editing host fills the cell and drops the text-frame flex/font
   sizing so it inherits the cell's own type (SPEC-091 AC-6a). */
.slide-table-cell .cell-edit {
  flex: none;
  min-height: 1em;
  font-size: inherit;
  outline: none;
  cursor: text;
}

.table-frame.is-selected {
  outline: 1px solid var(--accent, #3b82f6);
}

/* Thin invisible strips centred on each internal grid line; they carry the
   resize cursor and swallow the pointer-down that starts the drag (AC-8a/b). */
.col-resize-handle,
.row-resize-handle {
  position: absolute;
  pointer-events: auto;
  touch-action: none;
  z-index: 2;
}

.col-resize-handle {
  top: 0;
  bottom: 0;
  width: 8px;
  transform: translateX(-50%);
  cursor: col-resize;
}

.row-resize-handle {
  left: 0;
  right: 0;
  height: 8px;
  transform: translateY(-50%);
  cursor: row-resize;
}

.foreign-frame {
  border: 1px dashed var(--border-subtle, #b0b0b0);
  display: flex;
  align-items: center;
  justify-content: center;
}

.foreign-label {
  color: var(--text-secondary, #666);
  font-size: 0.75rem;
  font-family: system-ui, sans-serif;
}
</style>
