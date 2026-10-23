<script setup lang="ts">
// "Add from library…" (SPEC-109): a picker of the library's image entries —
// never documents, decks, workbooks or PDFs — each with its title and
// format. Ticked images are added to the tray in the order the picker
// lists them; the tray says when its limit cuts the list short. Rows are
// read from the library's listing; no image payload is loaded to show them.
// Runtime prop declarations (espree lints .vue scripts).
import { computed, ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { listDocuments } from "../../lib/documentStore";
import { isImageFormat } from "../../lib/libraryMenu";
import { formatBadgeLabel } from "../../lib/format";
import { MAX_ASSEMBLY_ITEMS } from "../../lib/core/assemblyLayout";

const props = defineProps({
  /** How many more images the tray can take. */
  room: { type: Number, default: MAX_ASSEMBLY_ITEMS },
});

const emit = defineEmits(["add", "close"]);

const images = listDocuments().filter((entry) => isImageFormat(entry.format));
const chosen = ref(new Set());

const chosenIds = computed(() => images.filter((entry) => chosen.value.has(entry.id)).map((entry) => entry.id));

function toggle(id = "") {
  const next = new Set(chosen.value);
  if (next.has(id)) {
    next.delete(id);
  } else {
    next.add(id);
  }
  chosen.value = next;
}
</script>

<template>
  <BaseDialog
    title="Add images from the library"
    testid="assembly-picker"
    width="min(32rem, 90vw)"
    @dismiss="emit('close')"
  >
    <ul
      v-if="images.length > 0"
      class="picker-list"
      aria-label="Images in the library"
    >
      <li
        v-for="entry in images"
        :key="entry.id"
      >
        <label
          class="picker-item"
          data-testid="picker-item"
        >
          <input
            type="checkbox"
            data-testid="picker-check"
            :checked="chosen.has(entry.id)"
            @change="toggle(entry.id)"
          >
          <span class="picker-title">{{ entry.title }}</span>
          <span class="picker-format">{{ formatBadgeLabel(entry.format) }}</span>
        </label>
      </li>
    </ul>
    <p
      v-else
      class="picker-empty"
      data-testid="assembly-picker-empty"
    >
      There are no images in the library yet.
    </p>
    <p
      v-if="chosenIds.length > props.room"
      class="picker-note"
      data-testid="assembly-picker-limit"
    >
      The tray has room for {{ props.room }} more; only the first {{ props.room }} will be added.
    </p>
    <template #actions>
      <UiButton
        data-testid="assembly-picker-cancel"
        @click="emit('close')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="assembly-picker-add"
        :disabled="chosenIds.length === 0"
        @click="emit('add', chosenIds)"
      >
        Add {{ chosenIds.length > 0 ? chosenIds.length : "" }}
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.picker-list {
  list-style: none;
  margin: 0;
  padding: 0;
  max-height: 50vh;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.picker-item {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.picker-item:hover {
  background: var(--surface-sunken);
}

.picker-title {
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.picker-format,
.picker-empty,
.picker-note {
  color: var(--text-secondary);
}

.picker-empty,
.picker-note {
  margin: 0;
}
</style>
