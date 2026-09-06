<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import AppIcon from "./AppIcon.vue";

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
  /** Toolbar Redesign — the leading chip naming what is selected ("Image",
   * "Rectangle", "Column chart"). The handoff gives every object bar this
   * chip so the bar says what it is acting on before it offers an action.
   * Empty (the default) renders no chip: the text-selection bar acts on the
   * selection itself and needs no name. */
  context: { type: String, default: "" },
  /** Icon id for the context chip. Optional: it defaults to the context name
   * lower-cased, which is already the icon id for every object that has one
   * ("Image" → image, "Chart" → chart, "Shape" → shape). A name with no
   * matching glyph — "Field" — simply renders no icon, because AppIcon draws
   * nothing for an unknown id. */
  contextIcon: { type: String, default: "" },
  /** "accent" outlines the whole bar in the accent — the handoff's treatment
   * for the text-selection bar, which has no naming chip to carry the accent
   * instead. */
  tone: { type: String, default: "default" },
});

const emit = defineEmits(["dismiss"]);

const root = useTemplateRef("root");
// position:fixed and the pre-measure hidden state live in the stylesheet
// (plain-string `position`/`visibility` are not valid CSSProperties values
// under vue-tsc); only the resolved left/top offsets ride inline.
const style = ref({ left: "0px", top: "0px" });
const ready = ref(false);
// Whether an "above" bar had to flip below its anchor — the pointer arrow
// swaps ends so it keeps pointing AT the selection rather than away from it.
const flipped = ref(false);
const chipIcon = computed(() => props.contextIcon || props.context.toLowerCase());

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
    // when the bar would run off the top of the viewport OR overlap the
    // toolbar strip (SPEC-061 "The text selection toolbar"). For a selection
    // near the top of the first page the "above" slot is exactly where the
    // strip's controls sit, and a bar parked there steals their clicks — a
    // reach for the font-size stepper lands on the bar's own Insert link.
    left = props.left - rect.width / 2;
    top = props.top - rect.height - MARGIN;
    const strip = document.querySelector('[data-testid="doc-toolbar-strip"]');
    const stripRect = strip ? strip.getBoundingClientRect() : null;
    const clearTop =
      stripRect && stripRect.height > 0 ? Math.max(MARGIN, stripRect.bottom + MARGIN) : MARGIN;
    flipped.value = top < clearTop;
    if (top < clearTop) {
      // Flip to just below the anchor — never to an arbitrary floor. When the
      // anchor itself is scrolled out above the canvas this degenerates to the
      // viewport-margin clamp below, same as it always has; a floor at the
      // strip's bottom would instead park the orphaned bar over the first
      // lines of content, stealing their clicks.
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
    :class="{
      ready,
      'is-lifted': lift,
      'is-accent': tone === 'accent',
      'has-arrow': placement === 'above',
      'is-flipped': flipped,
    }"
    role="toolbar"
    :aria-label="label"
    :data-testid="testid || undefined"
    :style="style"
  >
    <span
      v-if="context"
      class="context-chip"
    >
      <AppIcon
        v-if="chipIcon"
        :name="chipIcon"
      />
      <span class="context-name">{{ context }}</span>
    </span>
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
  gap: var(--space-1);
  padding: var(--space-1);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  /* Every floating bar carries the strip's control metrics, so a control
     looks the same whether it is docked in the strip or raised over the
     selection (SPEC-061 AC-22's "one chrome" taken through to the controls). */
  --control-height: var(--height-toolbar-control);
  --control-min-width: var(--height-toolbar-control);
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

/* The text-selection bar has no naming chip, so the accent rides its outline. */
.contextual-toolbar.is-accent {
  border-color: var(--accent);
}

/* The chip names the selected object and bleeds to the bar's edges, so it
   reads as part of the bar's frame rather than a control inside it. */
.context-chip {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  align-self: stretch;
  margin: calc(-1 * var(--space-1)) var(--space-1) calc(-1 * var(--space-1))
    calc(-1 * var(--space-1));
  padding: var(--space-2) var(--space-3);
  background: var(--surface-selected);
  color: var(--text-accent);
  border-inline-end: var(--border-width) solid var(--accent);
}

.context-name {
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  line-height: var(--line-tight);
  text-transform: uppercase;
  white-space: nowrap;
}

/* A small rotated square points the bar at the range it belongs to. It sits
   on whichever edge faces the selection, so a bar that flipped below its
   anchor still points up at it. */
.contextual-toolbar.has-arrow::after {
  content: "";
  position: absolute;
  left: var(--space-6);
  width: var(--space-4);
  height: var(--space-4);
  background: var(--surface-base);
  transform: rotate(45deg);
}

.contextual-toolbar.has-arrow:not(.is-flipped)::after {
  bottom: calc(-1 * var(--space-2));
  border-inline-end: var(--border-width) solid var(--border-default);
  border-block-end: var(--border-width) solid var(--border-default);
}

.contextual-toolbar.has-arrow.is-flipped::after {
  top: calc(-1 * var(--space-2));
  border-inline-start: var(--border-width) solid var(--border-default);
  border-block-start: var(--border-width) solid var(--border-default);
}

.contextual-toolbar.is-accent.has-arrow:not(.is-flipped)::after {
  border-inline-end-color: var(--accent);
  border-block-end-color: var(--accent);
}

.contextual-toolbar.is-accent.has-arrow.is-flipped::after {
  border-inline-start-color: var(--accent);
  border-block-start-color: var(--accent);
}
</style>
