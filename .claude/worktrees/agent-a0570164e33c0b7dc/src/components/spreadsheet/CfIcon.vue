<script setup lang="ts">
// One conditional-format icon (SPEC-129 "Icon artwork"): the glyph cfIcons.ts
// authors, as an inline svg labelled for assistive technology — `role="img"`,
// the icon's `aria-label` ("Green up arrow") and `data-icon` ("3Arrows:2") —
// drawn in a cell (SheetCellLayer.vue) and in the rules manager's previews.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { CF_GLYPH_SIZE, CF_GLYPH_STROKE, cfGlyphPaths, cfIconAt, cfPaintFill } from "../../lib/spreadsheet/cfIcons";

const props = defineProps({
  /** The icon set's name (IconSetName). */
  set: { type: String, required: true },
  /** The icon's index in its set, lowest first. */
  index: { type: Number, required: true },
  /** The rendered side, px. */
  sizePx: { type: Number, default: 16 },
});

const icon = computed(() => cfIconAt(props.set, props.index));
const paths = computed(() => (icon.value ? cfGlyphPaths(icon.value.glyph) : []));
</script>

<template>
  <svg
    v-if="icon"
    class="sheet-cf-icon"
    data-testid="sheet-cf-icon"
    role="img"
    :aria-label="icon.label"
    :data-icon="`${set}:${index}`"
    :viewBox="`0 0 ${CF_GLYPH_SIZE} ${CF_GLYPH_SIZE}`"
    :width="sizePx"
    :height="sizePx"
  >
    <path
      v-for="(path, i) in paths"
      :key="i"
      :d="path.d"
      :fill="cfPaintFill(path.paint, icon.color)"
      :stroke="path.paint === 'none' ? icon.color : undefined"
      :stroke-width="path.paint === 'none' ? CF_GLYPH_STROKE : undefined"
    />
  </svg>
</template>

<style scoped>
.sheet-cf-icon {
  flex: none;
  display: block;
}
</style>
