<script setup lang="ts">
// SPEC-082's object toolbar: Insert (Text box, Image…, Shapes gallery),
// Arrange (z-order), Style (fill/stroke popover) and Ungroup. All controls are
// keyboard-operable; the secondary group (Arrange/Style/Ungroup) collapses into
// a ⋯ overflow menu when the strip is too narrow for the 720px e2e viewport,
// the SPEC-081 overflow rule (AC-1/2/3). Reuses the product-neutral
// ShapeGallery (it renders from core/shapeIcons — no text-product imports) and
// FrameStylePopover. Runtime prop/emit declarations per the espree constraint.
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import ShapeGallery from "../ShapeGallery.vue";
import FrameStylePopover from "./FrameStylePopover.vue";
import { asShapeGeometry } from "../../lib/core/shapes";
import { asSlideFrame } from "./slideRender";
import { PRESENTATION_IMAGE_ACCEPT } from "../../views/presentation/useFrameInsertion";

const props = defineProps({
  /** The selected frame, or null. Enables Arrange/Style and (for a group)
   * Ungroup, and feeds the style popover. */
  selectedFrame: { type: Object, default: null },
  /** The armed insertion tool ("textbox" | "shape"), for active styling. */
  armedKind: { type: String, default: "" },
});

const emit = defineEmits([
  "insert-text-box",
  "insert-image",
  "insert-shape",
  "arrange",
  "ungroup",
  "set-fill",
  "set-stroke",
]);

const hasSelection = computed(() => props.selectedFrame !== null);
const selectedKind = computed(() => (props.selectedFrame ? asSlideFrame(props.selectedFrame).kind : ""));
const canStyle = computed(() => ["shape", "text", "image"].includes(selectedKind.value));
const isGroup = computed(() => selectedKind.value === "group");

// --- popovers --------------------------------------------------------------
const galleryOpen = ref(false);
const arrangeOpen = ref(false);
const styleOpen = ref(false);

function pickShape(geometry = "rect") {
  emit("insert-shape", asShapeGeometry(geometry));
  galleryOpen.value = false;
}

function arrange(target = "front") {
  emit("arrange", target);
  arrangeOpen.value = false;
}

// --- image picker ----------------------------------------------------------
const fileInput = useTemplateRef("fileInput");
function openPicker() {
  fileInput.value?.click();
}
function onFileChange(event = new Event("change")) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement)) return;
  const file = input.files?.[0];
  if (file) emit("insert-image", file);
  input.value = ""; // allow re-picking the same file
}

const ARRANGE = [
  { target: "front", label: "Bring to front" },
  { target: "forward", label: "Bring forward" },
  { target: "backward", label: "Send backward" },
  { target: "back", label: "Send to back" },
];

// --- overflow (⋯) — the secondary group collapses when narrow (AC-3) --------
const root = useTemplateRef("root");
const compact = ref(false);
const overflowOpen = ref(false);
const OVERFLOW_BELOW_PX = 620;
function updateCompact() {
  compact.value = (root.value?.clientWidth ?? 9999) < OVERFLOW_BELOW_PX;
  if (!compact.value) overflowOpen.value = false;
}
const observer = typeof ResizeObserver !== "undefined" ? new ResizeObserver(updateCompact) : null;
onMounted(() => {
  if (observer && root.value) observer.observe(root.value);
});
onBeforeUnmount(() => observer?.disconnect());

const accept = PRESENTATION_IMAGE_ACCEPT;
</script>

<template>
  <div
    ref="root"
    class="insert-bar"
    role="toolbar"
    aria-label="Insert and arrange"
    data-testid="slides-insert-bar"
  >
    <input
      ref="fileInput"
      type="file"
      class="sr-only"
      :accept="accept"
      data-testid="slides-image-input"
      @change="onFileChange"
    >

    <div class="tb-group">
      <button
        type="button"
        class="tb-btn"
        :aria-pressed="armedKind === 'textbox'"
        data-testid="insert-text-box"
        @click="emit('insert-text-box')"
      >
        Text box
      </button>
      <button
        type="button"
        class="tb-btn"
        data-testid="insert-image"
        @click="openPicker"
      >
        Image…
      </button>
      <div class="popover-host">
        <button
          type="button"
          class="tb-btn"
          aria-haspopup="menu"
          :aria-expanded="galleryOpen"
          :aria-pressed="armedKind === 'shape'"
          data-testid="insert-shape"
          @click="galleryOpen = !galleryOpen"
        >
          Shapes ▾
        </button>
        <div
          v-if="galleryOpen"
          class="popover shapes-popover"
          data-testid="shape-gallery-popover"
        >
          <ShapeGallery @pick="pickShape" />
        </div>
        <div
          v-if="galleryOpen"
          class="popover-backdrop"
          @click="galleryOpen = false"
        />
      </div>
    </div>

    <!-- Secondary group: inline, or in the ⋯ menu when compact -->
    <Teleport
      to="#insert-overflow-target"
      :disabled="!compact"
    >
      <div class="tb-secondary">
        <div class="popover-host">
          <button
            type="button"
            class="tb-btn"
            aria-haspopup="menu"
            :aria-expanded="arrangeOpen"
            :disabled="!hasSelection"
            data-testid="arrange-menu"
            @click="arrangeOpen = !arrangeOpen"
          >
            Arrange ▾
          </button>
          <div
            v-if="arrangeOpen"
            class="popover"
            role="menu"
            data-testid="arrange-popover"
          >
            <button
              v-for="a in ARRANGE"
              :key="a.target"
              type="button"
              role="menuitem"
              class="menu-item"
              :data-arrange="a.target"
              @click="arrange(a.target)"
            >
              {{ a.label }}
            </button>
          </div>
          <div
            v-if="arrangeOpen"
            class="popover-backdrop"
            @click="arrangeOpen = false"
          />
        </div>

        <div class="popover-host">
          <button
            type="button"
            class="tb-btn"
            aria-haspopup="dialog"
            :aria-expanded="styleOpen"
            :disabled="!canStyle"
            data-testid="style-button"
            @click="styleOpen = !styleOpen"
          >
            Style ▾
          </button>
          <div
            v-if="styleOpen && selectedFrame"
            class="popover"
            data-testid="style-popover-host"
          >
            <FrameStylePopover
              :frame="selectedFrame"
              @set-fill="(fill) => emit('set-fill', fill)"
              @set-stroke="(stroke) => emit('set-stroke', stroke)"
            />
          </div>
          <div
            v-if="styleOpen"
            class="popover-backdrop"
            @click="styleOpen = false"
          />
        </div>

        <button
          v-if="isGroup"
          type="button"
          class="tb-btn"
          data-testid="ungroup-button"
          @click="emit('ungroup')"
        >
          Ungroup
        </button>
      </div>
    </Teleport>

    <div
      v-show="compact"
      class="tb-group tb-overflow"
    >
      <button
        type="button"
        class="tb-btn"
        data-testid="insert-overflow"
        aria-label="More"
        aria-haspopup="menu"
        :aria-expanded="overflowOpen"
        @click="overflowOpen = !overflowOpen"
      >
        ⋯
      </button>
      <div
        v-show="overflowOpen"
        class="popover"
        role="menu"
        data-testid="insert-overflow-menu"
      >
        <div id="insert-overflow-target" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.insert-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
}

.tb-group,
.tb-secondary {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.tb-btn {
  height: 2rem;
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-sm);
}

.tb-btn:hover:not(:disabled) {
  background: var(--surface-sunken);
}

.tb-btn[aria-pressed="true"] {
  background: var(--accent-soft, var(--surface-sunken));
  border-color: var(--accent, var(--border-default));
}

.tb-btn:disabled {
  opacity: 0.5;
  cursor: default;
}

.popover-host {
  position: relative;
  display: inline-flex;
}

.popover {
  position: absolute;
  left: 0;
  top: calc(100% + var(--space-1));
  z-index: 50;
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.shapes-popover {
  width: 20rem;
  max-height: 22rem;
  overflow: auto;
}

.popover-backdrop {
  position: fixed;
  inset: 0;
  z-index: 49;
}

.menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  white-space: nowrap;
}

.menu-item:hover,
.menu-item:focus-visible {
  background: var(--surface-sunken);
}

.tb-overflow {
  position: relative;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
