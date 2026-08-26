// The library's grid rendering (SPEC-067 AC-1, AC-3..7): one card per visible
// entry — thumbnail, select checkbox, star, highlighted title, relative time,
// format badge and page/word counts — plus the overflow menu. Card semantics
// (modifier-click selects, plain click opens) live here; opening, menus and
// keyboard handling stay with HomeView, which owns the cross-view state they
// touch. Product-neutral: renders entries and emits intents, mirroring
// LibraryList for the grid view.
// `lib` is the LibraryController from useLibrary (runtime-typed Object — the
// .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
<script setup lang="ts">
import LibraryCard from "../../components/library/LibraryCard.vue";
import { libraryRowMenuItems } from "../../lib/libraryMenu";

const props = defineProps({
  lib: { type: Object, required: true },
  menuFor: { type: String, default: "" },
  highlightId: { type: String, default: "" },
});

const emit = defineEmits(["open", "select-modifiers", "keydown", "menu-toggle", "menu-close", "menu-select", "focus"]);
</script>

<template>
  <div
    class="library-grid"
    role="grid"
    data-testid="library-grid"
    @keydown="emit('keydown', $event)"
  >
    <LibraryCard
      v-for="(entry, index) in props.lib.visible.value"
      :key="entry.id"
      :entry="entry"
      :search="props.lib.search.value"
      :now="props.lib.now.value"
      :selected="props.lib.isSelected(entry.id)"
      :focused="props.lib.rowTabindex(entry.id, index) === 0"
      :menu-open="props.menuFor === entry.id"
      :highlight="entry.id === props.highlightId"
      :menu-items="libraryRowMenuItems(entry.starred)"
      @open="emit('open', $event)"
      @select="(payload) => emit('select-modifiers', payload)"
      @toggle-check="props.lib.toggleSelection"
      @toggle-star="props.lib.toggleStar"
      @menu-toggle="emit('menu-toggle', $event)"
      @menu-close="emit('menu-close')"
      @menu-select="(payload) => emit('menu-select', payload)"
      @focus="emit('focus', $event)"
    />
  </div>
</template>

<style scoped>
.library-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(14rem, 1fr));
  gap: var(--space-5);
}

/* SPEC-068 AC-37 — the compact tier (<640px, matching useResponsiveTier):
   cards reflow to a single column. The search box's own full-width row
   lives with it in LibraryControls. */
@media (max-width: 639.98px) {
  .library-grid {
    grid-template-columns: 1fr;
  }
}
</style>
