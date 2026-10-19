<script setup lang="ts">
// A two-stage slide transition (SPEC-085): the outgoing and incoming slides
// stacked inside one letterboxed box, animated from an "enter" state to an
// "active" state over the model's duration with inline transform/opacity/
// clip-path. Playback and the panel Preview both mount this so the visuals are
// identical. The parent completes a transition instantly simply by unmounting
// this layer (snapping to the plain incoming stage) — so an input mid-flight
// needs no imperative "finish" call here. Runtime prop declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import { computed, onBeforeUnmount, onMounted, ref } from "vue";
import SlideStage from "./SlideStage.vue";
import { cmToPx } from "../../lib/core/units";
import { transitionDirection, transitionDurationMs, transitionStagePair } from "./transitionStyles";

const props = defineProps({
  /** The slide leaving (rendered beneath/over per the effect). */
  outgoing: { type: Object, default: null },
  /** The slide arriving — its transition drives the animation. */
  incoming: { type: Object, required: true },
  /** The SlideTransition to render (the incoming slide's). */
  transition: { type: Object, required: true },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme, threaded to the stage surfaces. */
  theme: { type: Object, default: null },
  scale: { type: Number, default: 1 },
  /** When true, no animation runs — the layer completes on the next tick
   * (an instant cut) so reduced-motion users get no movement (AC-16). */
  reducedMotion: { type: Boolean, default: false },
  /** SPEC-086 build state per stage (frame id → inline style): playback hides
   * the incoming slide's unbuilt frames and the outgoing slide's exited ones,
   * so the transition shows what the plain stage will. Preview passes none. */
  outgoingFrameStyles: { type: Object, default: () => ({}) },
  incomingFrameStyles: { type: Object, default: () => ({}) },
  /** Playback's per-showing media URLs per stage (see SlideStage). */
  outgoingMediaUrlOverride: { type: Function, default: null },
  incomingMediaUrlOverride: { type: Function, default: null },
});

const emit = defineEmits(["done"]);

const phase = ref("enter");
const direction = computed(() => transitionDirection(props.transition));
const durationMs = computed(() => transitionDurationMs(props.transition));
const pair = computed(() => transitionStagePair(props.transition.type, direction.value, phase.value));

const boxStyle = computed(() => ({
  width: `${cmToPx(props.widthCm) * props.scale}px`,
  height: `${cmToPx(props.heightCm) * props.scale}px`,
}));

/** The inline CSS for one stage: its phase transform/opacity/clip plus the
 * timed transition shorthand so the enter → active change animates. `which` is
 * "outgoing" or "incoming". */
function layerStyle(which = "incoming") {
  const s = which === "outgoing" ? pair.value.outgoing : pair.value.incoming;
  return {
    transform: s.transform ?? "none",
    opacity: s.opacity ?? "1",
    clipPath: s.clipPath ?? "none",
    transition: `transform ${durationMs.value}ms ease, opacity ${durationMs.value}ms ease, clip-path ${durationMs.value}ms ease`,
  };
}

const outgoingStyle = computed(() => ({ ...layerStyle("outgoing"), zIndex: pair.value.incomingOnTop ? 0 : 1 }));
const incomingStyle = computed(() => ({ ...layerStyle("incoming"), zIndex: pair.value.incomingOnTop ? 1 : 0 }));

let doneFired = false;
let fallbackTimer = 0;
let rafA = 0;
let rafB = 0;

// jsdom (unit tests) may lack requestAnimationFrame; fall back to a timer so
// the layer still resolves rather than throwing on mount.
function fallbackRaf(cb = () => {}) {
  return window.setTimeout(cb, 16);
}
const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : fallbackRaf;
const cancelRaf = typeof cancelAnimationFrame === "function" ? cancelAnimationFrame : clearTimeout;

function finish() {
  if (doneFired) return;
  doneFired = true;
  emit("done");
}

onMounted(() => {
  // Animate for any real effect unless motion is reduced. A missing outgoing
  // slide (Preview on slide 1) animates the incoming in over the black stage.
  const animate = !props.reducedMotion && props.transition.type !== "none";
  if (!animate) {
    // Instant cut — let one frame paint the incoming slide, then complete.
    rafA = raf(() => finish());
    return;
  }
  // Commit the enter state, then flip to active on the next frame so the
  // browser animates the change; a timeout past the duration is the reliable
  // completion signal (transitionend can be missed under load).
  rafA = raf(() => {
    rafB = raf(() => {
      phase.value = "active";
      fallbackTimer = window.setTimeout(finish, durationMs.value + 100);
    });
  });
});

onBeforeUnmount(() => {
  if (rafA) cancelRaf(rafA);
  if (rafB) cancelRaf(rafB);
  if (fallbackTimer) clearTimeout(fallbackTimer);
});

function onIncomingTransitionEnd() {
  finish();
}
</script>

<template>
  <div
    class="transition-layer"
    :style="boxStyle"
    data-testid="slide-transition-layer"
    :data-phase="phase"
    :data-effect="transition.type"
  >
    <div
      v-if="outgoing"
      class="transition-stage"
      :style="outgoingStyle"
      data-testid="transition-outgoing"
    >
      <SlideStage
        :slide="outgoing"
        :width-cm="widthCm"
        :height-cm="heightCm"
        :media="media"
        :scale="scale"
        :interactive="false"
        :frame-styles="outgoingFrameStyles"
        :media-url-override="outgoingMediaUrlOverride"
        :theme="theme"
      />
    </div>
    <div
      class="transition-stage"
      :style="incomingStyle"
      data-testid="transition-incoming"
      @transitionend="onIncomingTransitionEnd"
    >
      <SlideStage
        :slide="incoming"
        :width-cm="widthCm"
        :height-cm="heightCm"
        :media="media"
        :scale="scale"
        :interactive="false"
        :frame-styles="incomingFrameStyles"
        :media-url-override="incomingMediaUrlOverride"
        :theme="theme"
      />
    </div>
  </div>
</template>

<style scoped>
.transition-layer {
  position: relative;
  overflow: hidden;
  /* Uncovered regions during a push, and Preview's "from black" on slide 1,
     read against black (SPEC-085 AC-19). */
  background: #000;
}

.transition-stage {
  position: absolute;
  inset: 0;
  will-change: transform, opacity, clip-path;
}
</style>
