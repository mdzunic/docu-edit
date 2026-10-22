<script setup lang="ts">
// The sheet toolbar's fill colour (SPEC-123): a split button — the bucket
// (`sheet-fill-button`) applies the last-used fill again; the arrow
// (`sheet-fill-menu-button`) opens `sheet-fill-menu` with "No fill"
// (`sheet-fill-none`), the theme rows (ThemeSwatchRow, `sheet-fill-theme`; a
// theme swatch stores its ThemeColorRef, which the xlsx writer turns into
// theme + tint), the TEXT_COLOR_PALETTE swatches (`sheet-fill-swatch`) and
// SPEC-013's recent colours. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { TEXT_COLOR_PALETTE } from "../../lib/core/colors";
import AppIcon from "../AppIcon.vue";
import ThemeSwatchRow from "../ThemeSwatchRow.vue";
import SheetMenuButton from "./SheetMenuButton.vue";

const props = defineProps({
  /** The fill the bucket applies, as hex ("" = none). */
  color: { type: String, default: "" },
  theme: { type: Object, default: null },
  recent: { type: Array, default: () => [] },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["apply-last", "select"]);

const recentColors = computed(() => props.recent.map(String));

function pick(color = Object(null), close = () => {}) {
  emit("select", color);
  close();
}
</script>

<template>
  <span class="split">
    <button
      type="button"
      class="main"
      data-testid="sheet-fill-button"
      aria-label="Fill colour"
      title="Fill colour"
      :disabled="disabled"
      @mousedown.prevent
      @click="emit('apply-last')"
    >
      <AppIcon name="fill-color" />
      <span
        class="bar"
        :style="{ background: color || 'transparent' }"
      />
    </button>
    <SheetMenuButton
      v-slot="{ close }"
      label="Fill colour options"
      testid="sheet-fill-menu-button"
      menu-testid="sheet-fill-menu"
      role="dialog"
      :disabled="disabled"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="sheet-fill-none"
        @mousedown.prevent
        @click="pick(null, close)"
      >
        No fill
      </button>
      <ThemeSwatchRow
        :theme="theme"
        testid="sheet-fill-theme"
        @select="pick($event, close)"
      />
      <div class="section-label">
        Standard
      </div>
      <div class="swatch-grid">
        <button
          v-for="hex in TEXT_COLOR_PALETTE"
          :key="hex"
          type="button"
          class="swatch"
          data-testid="sheet-fill-swatch"
          :data-color="hex"
          :title="hex"
          :aria-label="`Fill ${hex}`"
          :style="{ background: hex }"
          @mousedown.prevent
          @click="pick(hex, close)"
        />
      </div>
      <template v-if="recentColors.length > 0">
        <div class="section-label">
          Recent
        </div>
        <div class="swatch-grid">
          <button
            v-for="hex in recentColors"
            :key="`recent:${hex}`"
            type="button"
            class="swatch"
            data-testid="sheet-fill-recent-swatch"
            :data-color="hex"
            :title="hex"
            :aria-label="`Recent fill ${hex}`"
            :style="{ background: hex }"
            @mousedown.prevent
            @click="pick(hex, close)"
          />
        </div>
      </template>
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
  flex-direction: column;
  align-items: center;
  justify-content: center;
  min-width: var(--control-min-width, var(--height-control));
  height: var(--control-height, var(--height-control));
  padding: 0 var(--space-1);
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

.bar {
  width: 1.125rem;
  height: 0.25rem;
  border-radius: var(--radius-sm);
  border: var(--border-width) solid var(--border-subtle);
}

.section-label {
  margin: var(--space-3) 0 var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-secondary);
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
</style>
