<script setup lang="ts">
import { computed } from "vue";
import {
  resolveColor,
  THEME_SWATCH_SLOTS,
  THEME_SWATCH_VARIANTS,
  themeSlotLabel,
} from "../lib/core/theme";

// The theme row every document colour surface shares (SPEC-077 AC-20/AC-22):
// ten scheme swatches (dk1, lt1, dk2, lt2, accent1–6) with the five
// lightness variants beneath each. A swatch emits the ThemeColorRef it
// stands for — never the hex it merely displays — so the chosen colour
// follows the document theme (AC-23). TextColorPicker keeps its own inline
// copy of this grid (shipped by SPEC-094 with its own testids); this
// component is the same grid for the dialogs and panels that had none.

// Runtime prop/emit declarations, not the defineProps<{...}>() generics —
// see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The DocumentTheme the swatches resolve against for display. */
  theme: { type: Object, default: null },
  /** Prefix for data-testids: `<prefix>-row` and `<prefix>-swatch`. */
  testid: { type: String, default: "theme-swatch" },
});

const emit = defineEmits(["select"]);

const rows = computed(() => {
  const theme = props.theme;
  if (!theme) {
    return [];
  }
  const doc = Object(theme);
  const base = THEME_SWATCH_SLOTS.map((slot) => ({
    key: slot,
    name: themeSlotLabel(slot),
    hex: resolveColor({ theme: slot }, doc),
    reference: { theme: slot },
  }));
  const variantRows = THEME_SWATCH_VARIANTS.map((variant) =>
    THEME_SWATCH_SLOTS.map((slot) => {
      const reference = Object({ theme: slot });
      if (variant.lumMod !== undefined) reference.lumMod = variant.lumMod;
      if (variant.lumOff !== undefined) reference.lumOff = variant.lumOff;
      return {
        key: `${slot}:${variant.label}`,
        name: `${themeSlotLabel(slot)}, ${variant.label}`,
        hex: resolveColor(reference, doc),
        reference,
      };
    }),
  );
  return [base, ...variantRows];
});
</script>

<template>
  <div
    v-if="rows.length > 0"
    class="theme-swatch-rows"
  >
    <div class="section-label">
      Theme colours
    </div>
    <div
      v-for="(row, rowIndex) in rows"
      :key="`theme-row-${rowIndex}`"
      class="swatch-grid theme-row"
      :data-testid="`${testid}-row`"
    >
      <button
        v-for="cell in row"
        :key="cell.key"
        type="button"
        class="swatch"
        :data-testid="`${testid}-swatch`"
        :data-theme-slot="cell.reference.theme"
        :data-color="cell.hex"
        :title="cell.name"
        :aria-label="cell.name"
        :style="{ background: cell.hex }"
        @mousedown.prevent
        @click="emit('select', cell.reference)"
      />
    </div>
  </div>
</template>

<style scoped>
.section-label {
  font-size: var(--text-sm);
  color: var(--text-muted);
  margin-bottom: var(--space-1);
}

.swatch-grid {
  display: grid;
  grid-template-columns: repeat(10, 1.125rem);
  gap: var(--space-2);
}

.theme-row + .theme-row {
  margin-top: var(--space-1);
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
