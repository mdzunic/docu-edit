<script setup lang="ts">
// SPEC-067's highlight rendering, shared: a run of text segments, the
// matched ones marked. The library title (HighlightedTitle) marks the
// characters its subsequence search matched; SPEC-114's full-text results
// mark the words a snippet's match covers — the same <mark>, the same look.
// Runtime prop declarations (the .vue ESLint parser rejects TS-only syntax —
// see AppShell.vue).
import { computed } from "vue";
import { asHighlightSegments } from "../../lib/libraryView";

const props = defineProps({
  /** `{ text, match }` segments, in order. */
  segments: { type: Array, required: true },
  /** The test id each <mark> carries. */
  markTestid: { type: String, default: "title-match" },
});

const segments = computed(() => asHighlightSegments(props.segments));
</script>

<template>
  <span>
    <template
      v-for="(segment, index) in segments"
      :key="index"
    >
      <mark
        v-if="segment.match"
        class="match"
        :data-testid="props.markTestid"
      >{{ segment.text }}</mark>
      <template v-else>{{ segment.text }}</template>
    </template>
  </span>
</template>

<style scoped>
.match {
  background: var(--surface-selected);
  color: var(--text-accent);
  border-radius: var(--radius-sm);
  padding: 0 calc(var(--space-1) / 2);
}
</style>
