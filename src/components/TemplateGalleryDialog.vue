<script setup lang="ts">
// The template gallery (SPEC-046 AC-1, AC-6, AC-7, AC-8, AC-9, AC-10):
// offered alongside "Blank document" from the library's "New document"
// flow. Lists the six shipped templates plus any user templates ("Save as
// template"), each with a short preview, lets a user template be renamed or
// deleted in place, and offers importing a .dotx/.ott file as a template.
import { computed, ref, useTemplateRef } from "vue";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { deleteTemplate, listUserTemplates, renameTemplate } from "../lib/templateStore";
import { openTemplateFile } from "../lib/openDocument";
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
  <div
    class="dialog-backdrop"
    data-testid="template-gallery-dialog"
    @click.self="emit('close')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="New document from template"
    >
      <div class="dialog-header">
        <h2 class="dialog-title">
          New document
        </h2>
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
      </div>
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
    </div>
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
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  width: 42rem;
  max-width: 90vw;
  max-height: 85vh;
  overflow-y: auto;
  box-sizing: border-box;
}

.dialog-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 1rem;
}

.dialog-title {
  margin: 0;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.close-button {
  font: inherit;
  border: none;
  background: transparent;
  cursor: pointer;
  color: #5f6368;
  font-size: 1rem;
  padding: 0.25rem 0.5rem;
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.import-button {
  font: inherit;
  padding: 0.375rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-size: 0.8125rem;
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

.import-error {
  margin: 0 0 0.75rem;
  color: #b3261e;
  font-size: 0.8125rem;
}

.section-heading {
  margin: 1.25rem 0 0.75rem;
  font-size: 0.875rem;
  font-weight: 600;
  color: #5f6368;
}

.gallery-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr));
  gap: 0.75rem;
}

.template-card {
  font: inherit;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
  padding: 0.75rem;
  border: 1px solid #dadce0;
  border-radius: 6px;
  background: #fff;
  cursor: pointer;
}

.template-card:hover {
  border-color: #2c3e50;
  background: #f8f9fa;
}

.template-card-title {
  font-weight: 600;
  color: #202124;
}

.template-card-preview {
  font-size: 0.8125rem;
  color: #5f6368;
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
  gap: 0.25rem;
  padding: 0.75rem;
  border: none;
  background: transparent;
  cursor: pointer;
  width: 100%;
  box-sizing: border-box;
}

.user-template-actions {
  display: flex;
  border-top: 1px solid #f0f0f0;
}

.user-template-actions button {
  font: inherit;
  flex: 1;
  padding: 0.375rem 0.5rem;
  border: none;
  background: #fff;
  cursor: pointer;
  font-size: 0.8125rem;
  color: #3c4043;
}

.user-template-actions button:hover {
  background: #f1f3f4;
}

.user-template-actions button + button {
  border-left: 1px solid #f0f0f0;
}
</style>
