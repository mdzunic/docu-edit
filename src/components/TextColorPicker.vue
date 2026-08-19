<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import { parseColorInput, TEXT_COLOR_PALETTE } from "../lib/core/colors";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The selection's text colour ("#rrggbb"); "" when unset (the default
   * colour) or for the blank/indeterminate mixed state, which `mixed`
   * distinguishes (SPEC-013 AC). */
  value: { type: String, default: "" },
  mixed: { type: Boolean, default: false },
  /** Recently applied colours, most recent first (SPEC-013 AC). */
  recent: { type: Array, default: () => [] },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);

const open = ref(false);
const customDraft = ref("");
const customInvalid = ref(false);
const triggerEl = useTemplateRef("triggerEl");
const menuSurface = useTemplateRef("menuSurface");

const PALETTE = TEXT_COLOR_PALETTE;

// The toolbar clips overflow on its vertical axis, so the menu is rendered
// position:fixed (escaped below the anchor) instead of at a clipped offset.
// The measured element is UiMenuSurface's rendered box ($el).
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuSurface.value?.$el ?? null,
});

// The runtime Array prop arrives untyped; String() gives the template the
// string list it needs (espree bars a PropType<string[]> annotation here).
const recentColors = computed(() => props.recent.map(String));

function toggle() {
  if (props.disabled) {
    return;
  }
  open.value = !open.value;
  customDraft.value = "";
  customInvalid.value = false;
}

// Opened by the Format menu's "Text colour" entry (SPEC-049): the control is
// kept in the toolbar but also reachable from the menu.
function openMenu() {
  if (!props.disabled) {
    open.value = true;
    customDraft.value = "";
    customInvalid.value = false;
  }
}

defineExpose({ openMenu });

function pick(hex = "") {
  emit("select", hex);
  open.value = false;
}

function applyCustom() {
  const hex = parseColorInput(customDraft.value);
  if (hex === undefined) {
    customInvalid.value = true;
    return;
  }
  pick(hex);
}

function handleCustomInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    customDraft.value = target.value;
    customInvalid.value = false;
  }
}

function handleCustomKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    applyCustom();
  }
}

// Outside-click and Escape dismissal (and focus return to the trigger) are
// owned by UiMenuSurface — the shared, tested implementation — so this
// component no longer hand-rolls a document listener or a root Escape handler.
</script>

<template>
  <div class="color-picker">
    <button
      ref="triggerEl"
      type="button"
      class="picker-button"
      data-testid="text-color-button"
      title="Text colour"
      aria-label="Text colour"
      aria-haspopup="true"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggle"
    >
      <span class="glyph">A</span>
      <!-- The bar reflects the caret's colour; a striped neutral for the
           indeterminate mixed state (SPEC-013 AC). -->
      <span
        class="color-bar"
        data-testid="text-color-current"
        :class="{ 'color-bar-mixed': mixed }"
        :style="mixed ? {} : { background: value || '#000000' }"
        :data-color="mixed ? 'mixed' : (value || '')"
      />
    </button>
    <UiMenuSurface
      ref="menuSurface"
      :open="open"
      :anchor="triggerEl"
      class="picker-menu"
      data-testid="text-color-menu"
      role="dialog"
      aria-label="Text colour"
      :style="menuStyle"
      @dismiss="open = false"
    >
      <!-- Swatches prevent mousedown so the editor selection stays live;
           the custom input must NOT — it needs focus to be typed into. -->
      <div
        class="swatch-grid"
        data-testid="text-color-palette"
      >
        <button
          v-for="hex in PALETTE"
          :key="hex"
          type="button"
          class="swatch"
          data-testid="text-color-swatch"
          :data-color="hex"
          :title="hex"
          :aria-label="`Colour ${hex}`"
          :style="{ background: hex }"
          :class="{ active: !mixed && hex === value }"
          @mousedown.prevent
          @click="pick(hex)"
        />
      </div>
      <div
        v-if="recentColors.length > 0"
        class="section-label"
      >
        Recent
      </div>
      <div
        v-if="recentColors.length > 0"
        class="swatch-grid"
        data-testid="text-color-recent"
      >
        <button
          v-for="hex in recentColors"
          :key="`recent:${hex}`"
          type="button"
          class="swatch"
          data-testid="text-color-recent-swatch"
          :data-color="hex"
          :title="hex"
          :aria-label="`Recent colour ${hex}`"
          :style="{ background: hex }"
          @mousedown.prevent
          @click="pick(hex)"
        />
      </div>
      <div class="section-label">
        Custom
      </div>
      <div class="custom-row">
        <input
          class="custom-input"
          data-testid="text-color-custom-input"
          type="text"
          placeholder="#rrggbb or rgb(r,g,b)"
          aria-label="Custom colour, hex or RGB"
          :aria-invalid="customInvalid ? 'true' : undefined"
          :value="customDraft"
          @input="handleCustomInput"
          @keydown="handleCustomKeydown"
        >
        <button
          type="button"
          class="custom-apply"
          data-testid="text-color-custom-apply"
          @mousedown.prevent
          @click="applyCustom"
        >
          Apply
        </button>
      </div>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.color-picker {
  position: relative;
  display: inline-flex;
}

.picker-button {
  font: inherit;
  font-size: var(--text-md);
  width: var(--height-control);
  height: var(--height-control);
  padding: 0;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.picker-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.glyph {
  line-height: 1;
}

.color-bar {
  width: 1.125rem;
  height: 0.25rem;
  margin-top: var(--space-1);
  border-radius: var(--radius-sm);
}

.color-bar-mixed {
  background: repeating-linear-gradient(
    45deg,
    var(--border-strong) 0 2px,
    var(--surface-sunken) 2px 4px
  );
}

/* position/left/top come inline from useFloatingMenu; the box (border, radius,
   shadow, padding) is UiMenuSurface's. */
.picker-menu {
  z-index: 30;
}

.swatch-grid {
  display: grid;
  grid-template-columns: repeat(10, 1.125rem);
  gap: var(--space-2);
}

.swatch {
  width: 1.125rem;
  height: 1.125rem;
  padding: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.swatch.active {
  outline: calc(2 * var(--border-width)) solid var(--accent);
  outline-offset: var(--border-width);
}

.section-label {
  margin: var(--space-3) 0 var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.custom-row {
  display: flex;
  gap: var(--space-2);
}

.custom-input {
  font: inherit;
  font-size: var(--text-md);
  flex: 1;
  min-width: 0;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.custom-apply {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}
</style>
