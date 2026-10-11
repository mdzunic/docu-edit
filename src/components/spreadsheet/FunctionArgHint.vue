<script setup lang="ts">
// SPEC-127's argument hint (`function-arg-hint`, a tooltip): the signature
// of the innermost call being typed (functionHelp.ts's signatureParts), its
// current argument bold (`function-arg-current`); past maxArgs nothing is.
// It sits just below the editor; the editor names it in aria-describedby.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { ARG_HINT_ID, asSignatureParts } from "../../views/spreadsheet/useFunctionAssist";

const props = defineProps({
  /** The signature in pieces: `{ text, current? }`. */
  parts: { type: Array, default: () => [] },
  /** The fixed position below the editor. */
  anchorStyle: { type: Object, default: () => ({}) },
});

const parts = computed(() => asSignatureParts(props.parts));
</script>

<template>
  <div
    :id="ARG_HINT_ID"
    class="function-arg-hint"
    data-testid="function-arg-hint"
    role="tooltip"
    :style="props.anchorStyle"
  >
    <template
      v-for="(part, index) in parts"
      :key="index"
    >
      <b
        v-if="part.current"
        data-testid="function-arg-current"
      >{{ part.text }}</b>
      <span v-else>{{ part.text }}</span>
    </template>
  </div>
</template>

<style scoped>
.function-arg-hint {
  position: fixed;
  z-index: 40;
  max-width: 40rem;
  padding: var(--space-1) var(--space-2);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-1);
  font-family: var(--font-ui);
  font-size: var(--text-xs);
  color: var(--text-primary);
  white-space: nowrap;
  pointer-events: none;
}
</style>
