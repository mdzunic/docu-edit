// The library's shared overlay surface (SPEC-067/062/066/003): the hidden
// file input, the open-error alert, and every dialog/palette HomeView owns —
// rename, permanent-delete, empty-trash, template gallery and the command
// palette. A dumb presentational child: it renders the right overlay for each
// prop and re-emits each action as an intent; HomeView keeps all the wiring.
// It also owns the file input and exposes `pick()` so HomeView can open the
// chooser without reaching into a child's DOM.
<script setup lang="ts">
import { useTemplateRef } from "vue";
import RenameDialog from "../../components/RenameDialog.vue";
import ConfirmDeleteDialog from "../../components/ConfirmDeleteDialog.vue";
import TemplateGalleryDialog from "../../components/TemplateGalleryDialog.vue";
import CommandPalette from "../../components/CommandPalette.vue";

const props = defineProps({
  openFileAccept: { type: String, default: "" },
  openError: { type: String, default: "" },
  renameFor: { type: String, default: "" },
  renameTarget: { type: String, default: "" },
  permanentDeleteFor: { type: String, default: "" },
  permanentDeleteTitle: { type: String, default: "" },
  confirmingEmptyTrash: { type: Boolean, default: false },
  trashCount: { type: Number, default: 0 },
  showGallery: { type: Boolean, default: false },
  commandPaletteOpen: { type: Boolean, default: false },
  libraryCommands: { type: Array, default: () => [] },
  libraryCommandIds: { type: Array, default: () => [] },
  paletteRecentIds: { type: Array, default: () => [] },
});

const emit = defineEmits([
  "rename-save",
  "rename-cancel",
  "permanent-delete-confirm",
  "permanent-delete-cancel",
  "empty-trash-confirm",
  "empty-trash-cancel",
  "template-choose-blank",
  "template-choose-builtin",
  "template-choose-user",
  "template-choose-file",
  "template-close",
  "command-run",
  "command-dismiss",
  "file-chosen",
]);

const fileInput = useTemplateRef("fileInput");

function pick() {
  fileInput.value?.click();
}

function onFileChosen() {
  const file = fileInput.value?.files?.[0];
  if (fileInput.value) {
    fileInput.value.value = "";
  }
  if (!file) {
    return;
  }
  emit("file-chosen", file);
}

defineExpose({ pick });
</script>

<template>
  <input
    ref="fileInput"
    type="file"
    tabindex="-1"
    aria-hidden="true"
    :accept="props.openFileAccept"
    data-testid="open-document-input"
    class="visually-hidden"
    @change="onFileChosen"
  >
  <p
    v-if="props.openError"
    class="open-error"
    role="alert"
    data-testid="open-document-error"
  >
    {{ props.openError }}
  </p>
  <RenameDialog
    v-if="props.renameFor"
    :current-title="props.renameTarget"
    @save="emit('rename-save', $event)"
    @cancel="emit('rename-cancel')"
  />
  <ConfirmDeleteDialog
    v-if="props.permanentDeleteFor"
    :document-title="props.permanentDeleteTitle"
    kind="document permanently"
    @confirm="emit('permanent-delete-confirm')"
    @cancel="emit('permanent-delete-cancel')"
  />
  <ConfirmDeleteDialog
    v-else-if="props.confirmingEmptyTrash"
    :document-title="`${props.trashCount} document(s) in the trash`"
    kind="everything in the trash"
    @confirm="emit('empty-trash-confirm')"
    @cancel="emit('empty-trash-cancel')"
  />
  <TemplateGalleryDialog
    v-if="props.showGallery"
    @choose-blank="emit('template-choose-blank')"
    @choose-builtin="emit('template-choose-builtin', $event)"
    @choose-user="emit('template-choose-user', $event)"
    @choose-file="emit('template-choose-file', $event)"
    @close="emit('template-close')"
  />
  <CommandPalette
    v-if="props.commandPaletteOpen"
    :commands="props.libraryCommands"
    :common-ids="props.libraryCommandIds"
    :recent-ids="props.paletteRecentIds"
    @run="emit('command-run', $event)"
    @dismiss="emit('command-dismiss')"
  />
</template>

<style scoped>
.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

.open-error {
  margin: var(--space-4) auto 0;
  color: var(--text-danger);
  max-width: 32rem;
  text-align: center;
}
</style>
