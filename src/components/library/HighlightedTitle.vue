<script setup lang="ts">
import { computed } from "vue";
import HighlightedText from "./HighlightedText.vue";
import { highlightSegments, subsequenceMatch } from "../../lib/libraryView";

// Renders a document title with the characters matched by the current search
// highlighted (SPEC-067 AC-9). Falls back to the plain title when there is no
// search or no match. Runtime prop declarations (the .vue ESLint parser rejects
// defineProps<{...}>() and other TS-only syntax — see AppShell.vue).
const props = defineProps({
  title: { type: String, required: true },
  search: { type: String, default: "" },
});

const segments = computed(() => {
  const query = props.search.trim();
  if (query.length === 0) {
    return [{ text: props.title, match: false }];
  }
  const indices = subsequenceMatch(query, props.title);
  return indices ? highlightSegments(props.title, indices) : [{ text: props.title, match: false }];
});
</script>

<template>
  <HighlightedText :segments="segments" />
</template>
