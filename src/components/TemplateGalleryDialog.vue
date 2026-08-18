<script setup lang="ts">
// The template gallery (SPEC-046 AC-1, AC-6, AC-7, AC-8, AC-9, AC-10):
// offered alongside "Blank document" from the library's "New document"
// flow. Lists the six shipped templates plus any user templates ("Save as
// template"), each with a short preview, lets a user template be renamed or
// deleted in place, and offers importing a .dotx/.ott file as a template.
// BaseDialog owns the shell and — new with the migration — Escape-to-close;
// its #header-extra slot carries the import and ✕ close buttons on the
// title row. The nested rename/delete dialogs are template siblings of the
// gallery's BaseDialog (a fragment root), not children of its panel, so the
// gallery's focus trap never reaches into them — they were siblings of the
// panel before the migration too.
import { computed, ref, useTemplateRef } from "vue";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { deleteTemplate, listUserTemplates, renameTemplate } from "../lib/templateStore";
import { openTemplateFile } from "../lib/openDocument";
import BaseDialog from "./BaseDialog.vue";
import RenameDialog from "./RenameDialog.vue";
import ConfirmDeleteDialog from "./ConfirmDeleteDialog.vue";

const emit = defineEmits(["choose-blank", "choose-builtin", "choose-user", "choose-file", "close"]);

const fileInput = useTemplateRef("fileInput");
const importError = ref("");

function pickTemplateFile() {
  importError.value = "";
  fileInput.value?.click();
}

async function handleTemplateFileChosen() {
  const file = fileInput.value?.files?.[0];
  if (fileInput.value) {
    fileInput.value.value = "";
  }
  if (!file) {
    return;
  }
  try {
    const model = await openTemplateFile(file);
    importError.value = "";
    emit("choose-file", model);
  } catch (error) {
    importError.value = error instanceof Error ? error.message : String(error);
  }
}

const userTemplatesVersion = ref(0);
const userTemplates = computed(() => {
  void userTemplatesVersion.value;
  return listUserTemplates();
});

const renameFor = ref("");
const deleteFor = ref("");
const renameTarget = computed(() => userTemplates.value.find((entry) => entry.id === renameFor.value));
const deleteTarget = computed(() => userTemplates.value.find((entry) => entry.id === deleteFor.value));

function startRename(id = "") {
  renameFor.value = id;
}

function applyRename(newTitle = "") {
  renameTemplate(renameFor.value, newTitle);
  renameFor.value = "";
  userTemplatesVersion.value += 1;
}

function startDelete(id = "") {
  deleteFor.value = id;
}

function confirmDelete() {
  deleteTemplate(deleteFor.value);
  deleteFor.value = "";
  userTemplatesVersion.value += 1;
}
</script>

<template>
  <BaseDialog
    title="New document"
    testid="template-gallery-dialog"
    @dismiss="emit('close')"
  >
    <template #header-extra>
      <div class="header-actions">
        <button
          type="button"
          class="import-button"
          data-testid="template-gallery-import"
          @click="pickTemplateFile"
        >
          Import template file…
        </button>
        <button
          type="button"
          class="close-button"
          aria-label="Close"
          data-testid="template-gallery-close"
          @click="emit('close')"
        >
          ✕
        </button>
      </div>
    </template>
    <input
      ref="fileInput"
      type="file"
      accept=".dotx,.ott"
      data-testid="template-gallery-import-input"
      class="visually-hidden"
      @change="handleTemplateFileChosen"
    >
    <p
      v-if="importError"
      class="import-error"
      role="alert"
      data-testid="template-gallery-import-error"
    >
      {{ importError }}
    </p>
    <div class="gallery-grid">
      <button
        type="button"
        class="template-card"
        data-testid="template-card-blank"
        @click="emit('choose-blank')"
      >
        <span class="template-card-title">Blank document</span>
        <span class="template-card-preview">Start from an empty page.</span>
      </button>
      <button
        v-for="template in BUILTIN_TEMPLATES"
        :key="template.id"
        type="button"
        class="template-card"
        data-testid="template-card"
        :data-template-id="template.id"
        @click="emit('choose-builtin', template.id)"
      >
        <span class="template-card-title">{{ template.title }}</span>
        <span class="template-card-preview">{{ template.preview }}</span>
      </button>
    </div>
    <template v-if="userTemplates.length > 0">
      <h3 class="section-heading">
        Your templates
      </h3>
      <div
        class="gallery-grid"
        data-testid="user-template-list"
      >
        <div
          v-for="template in userTemplates"
          :key="template.id"
          class="template-card user-template-card"
          data-testid="user-template-card"
          :data-template-id="template.id"
        >
          <button
            type="button"
            class="template-card-body"
            @click="emit('choose-user', template.id)"
          >
            <span class="template-card-title">{{ template.title }}</span>
            <span class="template-card-preview">User template</span>
          </button>
          <div class="user-template-actions">
            <button
              type="button"
              data-testid="user-template-rename"
              aria-label="Rename template"
              @click="startRename(template.id)"
            >
              Rename
            </button>
            <button
              type="button"
              data-testid="user-template-delete"
              aria-label="Delete template"
              @click="startDelete(template.id)"
            >
              Delete
            </button>
          </div>
        </div>
      </div>
    </template>
  </BaseDialog>
  <RenameDialog
    v-if="renameTarget"
    :current-title="renameTarget.title"
    @save="applyRename"
    @cancel="renameFor = ''"
  />
  <ConfirmDeleteDialog
    v-if="deleteTarget"
    kind="template"
    :document-title="deleteTarget.title"
    @confirm="confirmDelete"
    @cancel="deleteFor = ''"
  />
</template>

<style scoped>
/* Fixed 42rem panel with the tall-dialog scroll cap. */
:deep(.dialog) {
  width: 42rem;
  max-height: 85vh;
  overflow-y: auto;
}

.close-button {
  font: inherit;
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
  font-size: var(--text-lg);
  padding: var(--space-2) var(--space-3);
}

.header-actions {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.import-button {
  font: inherit;
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
  font-size: var(--text-md);
}

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

.import-error {
  margin: 0 0 var(--space-4);
  color: var(--text-danger);
  font-size: var(--text-md);
}

.section-heading {
  margin: var(--space-6) 0 var(--space-4);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
}

.gallery-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr));
  gap: var(--space-4);
}

.template-card {
  font: inherit;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-4);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  cursor: pointer;
}

.template-card:hover {
  border-color: var(--surface-brand);
  background: var(--surface-raised);
}

.template-card-title {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.template-card-preview {
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.user-template-card {
  padding: 0;
  overflow: hidden;
  cursor: default;
}

.template-card-body {
  font: inherit;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-4);
  border: none;
  background: transparent;
  cursor: pointer;
  width: 100%;
  box-sizing: border-box;
}

.user-template-actions {
  display: flex;
  border-top: var(--border-width) solid var(--border-subtle);
}

.user-template-actions button {
  font: inherit;
  flex: 1;
  padding: var(--space-3) var(--space-3);
  border: none;
  background: var(--surface-base);
  cursor: pointer;
  font-size: var(--text-md);
  color: var(--text-primary);
}

.user-template-actions button:hover {
  background: var(--surface-canvas);
}

.user-template-actions button + button {
  border-left: var(--border-width) solid var(--border-subtle);
}
</style>
