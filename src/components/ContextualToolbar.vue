<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";

// The one floating-toolbar chrome (SPEC-061 AC-22): positioning, collision
// clamping against the viewport, Escape dismissal, focus return and ARIA, so
// the shape/image/chart/field object toolbars and the text-selection toolbar
// all share a single implementation and each caller supplies only its own
// controls (AC-23). The caller passes the anchor point (left/top, already
// resolved from the object's handles or the selection rect); this component
// keeps the bar on-screen and, when `placement` is "above", flips below the
// anchor if there is no room above.
//
// Runtime prop/emit declarations, not the generic forms — see AppShell.vue's
// note on the espree constraint.
const props = defineProps({
  /** Desired left edge in viewport px (clamped to stay on-screen). */
  left: { type: Number, default: 0 },
  /** Desired top edge in viewport px (the anchor's own top when placement is
   * "above"; the bar is lifted above it after measuring its height). */
  top: { type: Number, default: 0 },
  /** "free": place exactly at left/top (object toolbars already resolve this),
   * clamped horizontally to the viewport. "above": centre horizontally on
   * left, sit above the anchor top, flip below when it would clip. */
  placement: { type: String, default: "free" },
  /** When placement is "above", the anchor's bottom edge — where the bar flips
   * to when there is no room above. */
  anchorBottom: { type: Number, default: 0 },
  /** Lift the bar by its own height (translateY(-100%)) so `top` marks its
   * BOTTOM edge — how the shape toolbar clears a shape it sits above without
   * the caller knowing the wrapped bar's height (see useShapes). */
  lift: { type: Boolean, default: false },
  /** aria-label for the role="toolbar" container. */
  label: { type: String, required: true },
  /** data-testid mirrored onto the root, so a migrated object toolbar keeps
   * its existing selector (AC-37). */
  testid: { type: String, default: "" },
});

const emit = defineEmits(["dismiss"]);

const root = useTemplateRef("root");
// position:fixed and the pre-measure hidden state live in the stylesheet
// (plain-string `position`/`visibility` are not valid CSSProperties values
// under vue-tsc); only the resolved left/top offsets ride inline.
const style = ref({ left: "0px", top: "0px" });
const ready = ref(false);

const MARGIN = 8;

function reposition() {
  const el = root.value;
  if (!(el instanceof HTMLElement)) {
    return;
  }
  const rect = el.getBoundingClientRect();
  let left = props.left;
  let top = props.top;
  const maxLeft = Math.max(MARGIN, window.innerWidth - rect.width - MARGIN);
  if (props.placement === "above") {
    // Centre horizontally on the anchor point and lift above it; flip below
    // when the bar would run off the top of the viewport.
    left = props.left - rect.width / 2;
    top = props.top - rect.height - MARGIN;
    if (top < MARGIN) {
      top = props.anchorBottom + MARGIN;
    }
    const maxTop = Math.max(MARGIN, window.innerHeight - rect.height - MARGIN);
    top = Math.max(MARGIN, Math.min(top, maxTop));
  }
  // Horizontal clamp keeps every bar on-screen; vertical clamp only for the
  // centred "above" mode (a "free"/lifted caller has resolved its own top).
  left = Math.max(MARGIN, Math.min(left, maxLeft));
  style.value = {
    left: `${left}px`,
    top: `${top}px`,
  };
  ready.value = true;
}

watch(() => [props.left, props.top, props.placement, props.anchorBottom], () => reposition(), {
  flush: "post",
});

// Escape dismissal is wired at the document level (not a root @keydown): the
// bar is not a tab stop, so when it is raised by a mouse selection focus stays
// in the document and a root listener would never fire — the same reason
// UiMenuSurface listens on document.
function onDocKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    emit("dismiss");
  }
}

onMounted(() => {
  reposition();
  window.addEventListener("resize", reposition);
  document.addEventListener("scroll", reposition, true);
  document.addEventListener("keydown", onDocKeydown);
});

onBeforeUnmount(() => {
  window.removeEventListener("resize", reposition);
  document.removeEventListener("scroll", reposition, true);
  document.removeEventListener("keydown", onDocKeydown);
});
</script>

<template>
  <div
    ref="root"
    class="contextual-toolbar"
    :class="{ ready, 'is-lifted': lift }"
    role="toolbar"
    :aria-label="label"
    :data-testid="testid || undefined"
    :style="style"
  >
    <slot />
  </div>
</template>

<style scoped>
.contextual-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  max-width: min(90vw, 640px);
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  /* Hidden until reposition() has measured and placed it, so it never paints
     for a frame at the wrong spot. */
  visibility: hidden;
}

.contextual-toolbar.ready {
  visibility: visible;
}

/* `top` marks the bottom edge; the bar pulls itself up by its own height. */
.contextual-toolbar.is-lifted {
  transform: translateY(-100%);
}
</style>
