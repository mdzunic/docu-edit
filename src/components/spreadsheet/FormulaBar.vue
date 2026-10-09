<script setup lang="ts">
// The read-only formula bar (SPEC-117): above the grid, the Name Box — a
// read-only input showing the active cell's address (a merge's top-left), or
// "3R × 2C" while a drag is in progress — and the active cell's content
// (its General text with line breaks preserved; SPEC-118 refines it to the
// edit text). Nothing here edits: SPEC-121 makes both editable. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
const props = defineProps({
  /** The Name Box text: "B3", or "3R × 2C" during a drag. */
  address: { type: String, default: "" },
  /** The active cell's content. */
  content: { type: String, default: "" },
});
</script>

<template>
  <div
    class="formula-bar"
    data-testid="formula-bar"
  >
    <input
      class="name-box"
      data-testid="name-box"
      type="text"
      readonly
      aria-label="Name Box"
      :value="props.address"
    >
    <span
      class="formula-bar-fx"
      aria-hidden="true"
    >fx</span>
    <div
      class="formula-bar-content"
      data-testid="formula-bar-content"
      role="textbox"
      aria-readonly="true"
      aria-label="Cell content"
      aria-multiline="true"
      v-text="props.content"
    />
  </div>
</template>

<style scoped>
.formula-bar {
  display: flex;
  flex: none;
  align-items: stretch;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
}

.name-box {
  flex: none;
  width: 7rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: inherit;
  font: inherit;
}

.name-box:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

.formula-bar-fx {
  flex: none;
  align-self: center;
  color: var(--text-secondary);
  font-style: italic;
}

/* Line breaks preserved; a multi-line value is clipped to a few lines. */
.formula-bar-content {
  flex: 1 1 auto;
  min-width: 0;
  max-height: 4.5em;
  padding: var(--space-1) var(--space-2);
  overflow: auto;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
</style>
