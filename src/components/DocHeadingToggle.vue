<script setup lang="ts">
// The collapse chevron rendered next to a heading in the document body
// (SPEC-037 AC-16). It sits INSIDE the heading element but carries no text
// node and `data-canvas-ignore`, so it contributes nothing to the canvas's
// model-text (see lib/domCaret.canvasModelText) and never shifts caret
// offsets — the glyph is a CSS pseudo-element. Its own scoped styles travel
// with it, so it renders identically in the single- and multi-column paths.
// Runtime prop/emit declarations (espree parses .vue scripts — see
// AppShell.vue).
defineProps({
  collapsed: { type: Boolean, default: false },
});

defineEmits(["toggle"]);
</script>

<template>
  <button
    type="button"
    class="doc-heading-toggle"
    :class="{ 'is-collapsed': collapsed }"
    contenteditable="false"
    data-canvas-ignore="1"
    tabindex="-1"
    data-testid="doc-heading-toggle"
    :aria-label="collapsed ? 'Expand heading' : 'Collapse heading'"
    :aria-expanded="collapsed ? 'false' : 'true'"
    @mousedown.prevent.stop
    @click.stop="$emit('toggle')"
  />
</template>

<style scoped>
/* Empty button — the chevron is a pseudo-element so no text node lands inside
   the editable heading. Inline so it sits at the start of the heading line
   without adding block height beyond the line box. */
.doc-heading-toggle {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 1.1em;
  height: 1.1em;
  margin-inline-end: 0.15em;
  padding: 0;
  border: 0;
  border-radius: 2px;
  background: transparent;
  color: #5f6368;
  font-size: 0.7em;
  line-height: 1;
  vertical-align: middle;
  cursor: pointer;
  user-select: none;
}

.doc-heading-toggle::before {
  content: "▾";
}

.doc-heading-toggle.is-collapsed::before {
  content: "▸";
}

.doc-heading-toggle:hover {
  background: #e0e0e0;
}
</style>
