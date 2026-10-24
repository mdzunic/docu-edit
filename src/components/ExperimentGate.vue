<script setup lang="ts">
import { computed, defineAsyncComponent } from "vue";
import { useExperiments } from "../composables/useExperiments";
import { isExperimentId } from "../lib/experiments";

// SPEC-113's gate for an experimental entry point. While the experiment is
// off it renders nothing and loads nothing: `load` — the surface's own
// `() => import("…")` — is not called until the flag is on, so a disabled
// feature's code is never fetched and is not in the initial bundle. Switched
// on, the surface loads and renders (attributes fall through to it);
// switched off again, it unmounts. The surface carries its own
// ExperimentBadge, placed where its layout wants it.
//
// Runtime prop declarations for the espree constraint AppShell.vue explains.
const props = defineProps({
  /** The ExperimentId that gates this entry point. */
  experiment: { type: String, required: true, validator: (value) => isExperimentId(value) },
  /** `() => import("./TheSurface.vue")` — called at most once, and only
   * while the experiment is on. */
  load: { type: Function, required: true },
});

const { isEnabled } = useExperiments();

let surface = null;

const active = computed(() => {
  if (!isEnabled(props.experiment)) {
    return null;
  }
  surface ??= defineAsyncComponent(() => props.load());
  return surface;
});
</script>

<template>
  <component
    :is="active"
    v-if="active"
  />
</template>
