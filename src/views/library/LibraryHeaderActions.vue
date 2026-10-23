// The library header's create/open actions (SPEC-068 AC-38): at the compact
// tier they collapse to one primary "New document" button plus an overflow
// menu holding the rest. Product-neutral — renders the buttons and an overflow
// RowMenu, emitting intents; HomeView wires each to a concrete product action.
// The other create intents sit right after "New document" in both layouts:
// "New presentation" (SPEC-080), then "New spreadsheet" (SPEC-121), then
// "New combined image" (SPEC-109's New › Combined image, an empty assembly
// tray) — inert while the library holds no image (`canCombine`).
// `compact` mirrors useResponsiveTier; the overflow item ids are what each
// button would otherwise be.
<script setup lang="ts">
import { computed, ref } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import RowMenu from "../../components/library/RowMenu.vue";

const props = defineProps({
  compact: { type: Boolean, default: false },
  /** Whether the library holds an image to combine. */
  canCombine: { type: Boolean, default: false },
});

const emit = defineEmits([
  "new-document",
  "new-presentation",
  "new-spreadsheet",
  "new-combined-image",
  "new-from-template",
  "open-document",
]);

const menuOpen = ref(false);
const COMBINED_IMAGE_ACTION = { id: "new-combined-image", label: "New combined image", testid: "new-combined-image" };
const OVERFLOW_ACTIONS = computed(() => [
  { id: "new-presentation", label: "New presentation", testid: "new-presentation" },
  { id: "new-spreadsheet", label: "New spreadsheet", testid: "new-spreadsheet" },
  ...(props.canCombine ? [COMBINED_IMAGE_ACTION] : []),
  { id: "new-from-template", label: "New from template…", testid: "new-from-template" },
  { id: "open-document", label: "Open document", testid: "open-document" },
]);

function runOverflowAction(id = "") {
  menuOpen.value = false;
  if (id === "new-presentation") {
    emit("new-presentation");
  } else if (id === "new-spreadsheet") {
    emit("new-spreadsheet");
  } else if (id === "new-combined-image") {
    emit("new-combined-image");
  } else if (id === "new-from-template") {
    emit("new-from-template");
  } else if (id === "open-document") {
    emit("open-document");
  }
}
</script>

<template>
  <UiButton
    variant="primary"
    data-testid="new-document"
    @click="emit('new-document')"
  >
    New document
  </UiButton>
  <template v-if="!props.compact">
    <UiButton
      data-testid="new-presentation"
      @click="emit('new-presentation')"
    >
      New presentation
    </UiButton>
    <UiButton
      data-testid="new-spreadsheet"
      @click="emit('new-spreadsheet')"
    >
      New spreadsheet
    </UiButton>
    <UiButton
      data-testid="new-combined-image"
      :disabled="!props.canCombine"
      :title="props.canCombine ? undefined : 'Add an image to the library to combine images'"
      @click="emit('new-combined-image')"
    >
      New combined image
    </UiButton>
    <UiButton
      data-testid="new-from-template"
      @click="emit('new-from-template')"
    >
      New from template…
    </UiButton>
    <UiButton
      data-testid="open-document"
      @click="emit('open-document')"
    >
      Open document
    </UiButton>
  </template>
  <RowMenu
    v-else
    :open="menuOpen"
    label="More actions"
    :items="OVERFLOW_ACTIONS"
    button-testid="library-actions-overflow"
    menu-testid="library-actions-overflow-menu"
    @toggle="menuOpen = !menuOpen"
    @close="menuOpen = false"
    @select="runOverflowAction"
  />
</template>
