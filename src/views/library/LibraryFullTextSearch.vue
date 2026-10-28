<script setup lang="ts">
// SPEC-114's entry point in the library's main column: the gate for the
// experimental full-text results (LibrarySearchResults.vue). While the
// library-search experiment is off it renders nothing and loads nothing; on,
// it loads the surface and passes its intents through — `showing` (results
// stand in for the rows) and `open` (a result to open at its match).
// `lib` is the LibraryController (runtime-typed Object — see useLibrary.ts).
import ExperimentGate from "../../components/ExperimentGate.vue";

const props = defineProps({
  lib: { type: Object, required: true },
});

const emit = defineEmits(["showing", "open"]);

const loadResults = () => import("./LibrarySearchResults.vue");
</script>

<template>
  <ExperimentGate
    experiment="library-search"
    :load="loadResults"
    :lib="props.lib"
    @showing="emit('showing', $event)"
    @open="emit('open', $event)"
  />
</template>
