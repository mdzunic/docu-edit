<script setup lang="ts">
import { computed } from "vue";
import { ICON_PATHS } from "./icons/paths";

// The single icon renderer (SPEC-058). Every toolbar/menu glyph goes through
// here: a 20×20 stroke-only <svg> that inherits its colour from the button via
// stroke="currentColor", is aria-hidden (the button owns the accessible name)
// and non-focusable. An unknown name renders nothing at all — a v-if guard —
// so a typo shows up as a blank button in a unit test rather than shipping.
//
// Runtime prop declaration (not defineProps<{...}>()) for the espree/ESLint
// constraint the other SFCs note; vue-tsc still infers the type.
const props = defineProps({
  /** Icon id — a key of ICON_PATHS. Unknown names render nothing. */
  name: { type: String, required: true },
});

const paths = computed(() => ICON_PATHS[props.name] ?? null);
</script>

<template>
  <svg
    v-if="paths"
    class="app-icon"
    viewBox="0 0 20 20"
    width="20"
    height="20"
    fill="none"
    stroke="currentColor"
    stroke-width="1.5"
    stroke-linecap="round"
    stroke-linejoin="round"
    aria-hidden="true"
    focusable="false"
  >
    <path
      v-for="(d, i) in paths"
      :key="i"
      :d="d"
    />
  </svg>
</template>

<style scoped>
.app-icon {
  display: block;
  flex: none;
}
</style>
