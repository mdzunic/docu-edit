// The library header's create/open actions (SPEC-068 AC-38): at the compact
// tier they collapse to one primary "New document" button plus an overflow
// menu holding the rest. Product-neutral — renders the buttons and an overflow
// RowMenu, emitting intents; HomeView wires each to a concrete product action.
// `compact` mirrors useResponsiveTier; the overflow item ids are what each
// button would otherwise be.
<script setup lang="ts">
import { ref } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import RowMenu from "../../components/library/RowMenu.vue";

const props = defineProps({
  compact: { type: Boolean, default: false },
});

const emit = defineEmits(["new-document", "new-presentation", "new-from-template", "open-document"]);

const menuOpen = ref(false);
const OVERFLOW_ACTIONS = [
  { id: "new-presentation", label: "New presentation", testid: "new-presentation" },
  { id: "new-from-template", label: "New from template…", testid: "new-from-template" },
  { id: "open-document", label: "Open document", testid: "open-document" },
];

function runOverflowAction(id = "") {
  menuOpen.value = false;
  if (id === "new-presentation") {
    emit("new-presentation");
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
