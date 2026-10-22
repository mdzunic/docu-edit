<script setup lang="ts">
// The cell styles gallery (SPEC-123 "Named cell styles"): the built-ins in
// Excel's groups — "Good, Bad and Neutral", "Data and model", "Titles and
// headings" — then "Custom" for the workbook's own names, one
// `sheet-cell-style` per style (`data-style` its name), each drawn in its own
// style as it would apply (sheetToolbar.ts's cellStyleGroups). Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { fontStack } from "../../lib/core/fontSubstitution";

const props = defineProps({
  /** cellStyleGroups' groups. */
  groups: { type: Array, default: () => [] },
});

const emit = defineEmits(["select"]);

const list = computed(() => props.groups.map((group) => Object(group)));

/** A gallery entry's text, its family wrapped in the substitution stack. */
function textStyle(style = Object()) {
  const family = style["font-family"];
  return typeof family === "string" && family !== "" ? { ...style, "font-family": fontStack(family) } : style;
}
</script>

<template>
  <div
    class="gallery"
    data-testid="sheet-cell-styles-gallery"
  >
    <section
      v-for="group in list"
      :key="group.label"
      class="group"
    >
      <h3 class="group-label">
        {{ group.label }}
      </h3>
      <div class="entries">
        <button
          v-for="entry in group.styles"
          :key="entry.name"
          type="button"
          role="menuitem"
          class="entry"
          data-testid="sheet-cell-style"
          :data-style="entry.name"
          :aria-label="entry.name"
          :style="entry.box"
          @mousedown.prevent
          @click="emit('select', entry.name)"
        >
          <span
            class="entry-text"
            :style="textStyle(entry.text)"
          >{{ entry.name }}</span>
        </button>
      </div>
    </section>
  </div>
</template>

<style scoped>
.gallery {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-2);
  min-width: 26rem;
}

.group-label {
  margin: 0 0 var(--space-2);
  font-size: var(--text-sm);
  font-weight: var(--weight-button);
  color: var(--text-secondary);
}

.entries {
  display: grid;
  grid-template-columns: repeat(4, 6rem);
  gap: var(--space-2);
}

/* Each entry is drawn in its own style (inline); this is the ground under it. */
.entry {
  display: flex;
  flex-direction: column;
  justify-content: center;
  height: 2rem;
  padding: 0;
  overflow: hidden;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-subtle);
  cursor: pointer;
}

.entry:hover {
  outline: calc(2 * var(--border-width)) solid var(--accent);
}

.entry-text {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
</style>
