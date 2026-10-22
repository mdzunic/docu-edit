<script setup lang="ts">
// The sheet toolbar's borders (SPEC-123 "Borders"): a split button — the
// main half (`sheet-borders-button`) re-applies the last preset; the arrow
// (`sheet-borders-menu-button`) opens `sheet-borders-menu` with the eight
// presets (`sheet-border-<preset>`), "Line colour" (`sheet-border-line-color`:
// Automatic and the palette) and "Line style" (`sheet-border-line-style`:
// the 13 SheetBorderStyles) — the line the presets draw, kept for the view
// session by useFormatting. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { ref } from "vue";
import { TEXT_COLOR_PALETTE } from "../../lib/core/colors";
import { LINE_STYLES } from "../../lib/spreadsheet/formatCells";
import AppIcon from "../AppIcon.vue";
import SheetMenuButton from "./SheetMenuButton.vue";

defineProps({
  lineStyle: { type: String, default: "thin" },
  /** "" = automatic. */
  lineColor: { type: String, default: "" },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["apply-last", "preset", "line-style", "line-color"]);

const PRESETS = [
  { id: "bottom", label: "Bottom border" },
  { id: "top", label: "Top border" },
  { id: "left", label: "Left border" },
  { id: "right", label: "Right border" },
  { id: "all", label: "All borders" },
  { id: "outside", label: "Outside borders" },
  { id: "thick-outside", label: "Thick outside borders" },
  { id: "none", label: "No border" },
];

const colorsOpen = ref(false);

function onLineStyle(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("line-style", target.value);
  }
}
</script>

<template>
  <span class="split">
    <button
      type="button"
      class="main"
      data-testid="sheet-borders-button"
      aria-label="Borders"
      title="Borders"
      :disabled="disabled"
      @mousedown.prevent
      @click="emit('apply-last')"
    >
      <AppIcon name="borders" />
    </button>
    <SheetMenuButton
      v-slot="{ close }"
      label="Border options"
      testid="sheet-borders-menu-button"
      menu-testid="sheet-borders-menu"
      :disabled="disabled"
    >
      <button
        v-for="preset in PRESETS"
        :key="preset.id"
        type="button"
        role="menuitem"
        :data-testid="`sheet-border-${preset.id}`"
        @mousedown.prevent
        @click="emit('preset', preset.id); close()"
      >
        {{ preset.label }}
      </button>
      <hr class="separator">
      <button
        type="button"
        role="menuitem"
        data-testid="sheet-border-line-color"
        :aria-expanded="colorsOpen ? 'true' : 'false'"
        @mousedown.prevent
        @click="colorsOpen = !colorsOpen"
      >
        Line colour
        <span
          class="chip"
          :style="{ background: lineColor || '#000000' }"
        />
      </button>
      <div
        v-if="colorsOpen"
        class="swatch-grid"
      >
        <button
          type="button"
          class="automatic"
          data-testid="sheet-border-line-color-automatic"
          @mousedown.prevent
          @click="emit('line-color', null)"
        >
          Automatic
        </button>
        <button
          v-for="hex in TEXT_COLOR_PALETTE"
          :key="hex"
          type="button"
          class="swatch"
          data-testid="sheet-border-line-color-swatch"
          :data-color="hex"
          :aria-label="`Line colour ${hex}`"
          :style="{ background: hex }"
          @mousedown.prevent
          @click="emit('line-color', hex)"
        />
      </div>
      <label class="line-style">
        Line style
        <select
          data-testid="sheet-border-line-style"
          :value="lineStyle"
          @change="onLineStyle"
        >
          <option
            v-for="style in LINE_STYLES"
            :key="style"
            :value="style"
          >{{ style }}</option>
        </select>
      </label>
    </SheetMenuButton>
  </span>
</template>

<style scoped>
.split {
  display: inline-flex;
  align-items: center;
}

.main {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: var(--control-min-width, var(--height-control));
  height: var(--control-height, var(--height-control));
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.main:not(:disabled):hover {
  background: var(--state-hover);
}

.main:disabled {
  opacity: 0.5;
  cursor: default;
}

.separator {
  border: 0;
  border-top: var(--border-width) solid var(--border-subtle);
  margin: var(--space-1) 0;
}

.chip {
  width: 1rem;
  height: 0.5rem;
  border-radius: var(--radius-sm);
}

.swatch-grid {
  display: grid;
  grid-template-columns: repeat(10, 1.125rem);
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
}

.automatic {
  grid-column: 1 / -1;
  font: inherit;
  font-size: var(--text-sm);
  background: transparent;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.swatch {
  width: 1.125rem;
  height: 1.125rem;
  padding: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.line-style {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-md);
}
</style>
