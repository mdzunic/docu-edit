<script setup lang="ts">
// The Name Manager's New… / Edit… (SPEC-133): `name-editor` on BaseDialog,
// titled "Define name" or "Edit name", bound to useNames' editor —
// `name-editor-name`, `name-editor-scope` (Workbook or a sheet; fixed when
// editing), `name-editor-comment`, `name-editor-refers-to` and
// `name-editor-error` (validated live as the user types; OK
// `name-editor-ok` disabled while there is an error; Enter in a one-line
// field is OK; `name-editor-cancel` and Escape cancel). Refers to points
// (SPEC-122) the way SPEC-132's Select Data does: "Select cells"
// (`name-editor-pick`) collapses the dialogs to `name-editor-picking`, whose
// text follows the grid's selection on any sheet, until Done
// (`name-editor-pick-done`). Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, useTemplateRef } from "vue";
import { asNames } from "../../views/spreadsheet/useNames";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useNames' return (its `editor` is open). */
  names: { type: Object, required: true },
});

const n = computed(() => asNames(props.names));
const fields = computed(() => n.value.editor.value?.fields ?? { name: "", scope: "", comment: "", refersTo: "" });
const editingExisting = computed(() => n.value.editor.value?.mode === "edit");
const error = computed(() => n.value.editorError.value);
const refersTo = useTemplateRef("refersTo");

function valueOf(event = new Event("input")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement || el instanceof HTMLTextAreaElement ? el.value : "";
}

function pick() {
  n.value.startPick(refersTo.value?.selectionStart ?? undefined);
}
</script>

<template>
  <div
    v-if="n.picking.value"
    class="name-editor-picking"
    role="region"
    aria-label="Select the cells the name refers to"
    data-testid="name-editor-picking"
  >
    <span data-testid="name-editor-picking-text">{{ fields.refersTo }}</span>
    <UiButton
      data-testid="name-editor-pick-done"
      @click="n.endPick()"
    >
      Done
    </UiButton>
  </div>
  <BaseDialog
    v-else
    :title="editingExisting ? 'Edit name' : 'Define name'"
    testid="name-editor"
    width="28rem"
    @dismiss="n.cancelEditor()"
  >
    <form
      class="name-editor-form"
      @submit.prevent="n.confirmEditor()"
    >
      <label class="name-editor-field">
        <span>Name</span>
        <input
          type="text"
          class="name-editor-control"
          data-testid="name-editor-name"
          autofocus
          spellcheck="false"
          :aria-invalid="error !== '' && fields.name !== '' ? 'true' : undefined"
          :value="fields.name"
          @input="n.setField('name', valueOf($event))"
        >
      </label>
      <label class="name-editor-field">
        <span>Scope</span>
        <select
          class="name-editor-control"
          data-testid="name-editor-scope"
          :disabled="editingExisting"
          :value="fields.scope"
          @change="n.setField('scope', valueOf($event))"
        >
          <option
            v-for="scope in n.scopes.value"
            :key="scope.value"
            :value="scope.value"
          >
            {{ scope.label }}
          </option>
        </select>
      </label>
      <label class="name-editor-field">
        <span>Comment</span>
        <textarea
          class="name-editor-control name-editor-comment"
          data-testid="name-editor-comment"
          rows="2"
          :value="fields.comment"
          @input="n.setField('comment', valueOf($event))"
        />
      </label>
      <label class="name-editor-field">
        <span>Refers to</span>
        <span class="name-editor-refers">
          <input
            ref="refersTo"
            type="text"
            class="name-editor-control"
            data-testid="name-editor-refers-to"
            spellcheck="false"
            :value="fields.refersTo"
            @input="n.setField('refersTo', valueOf($event))"
          >
          <UiButton
            variant="secondary"
            data-testid="name-editor-pick"
            @click="pick"
          >
            Select cells
          </UiButton>
        </span>
      </label>
      <p
        class="name-editor-error"
        role="alert"
        data-testid="name-editor-error"
      >
        {{ error }}
      </p>
      <button
        type="submit"
        hidden
      />
    </form>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="name-editor-cancel"
        @click="n.cancelEditor()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="name-editor-ok"
        :disabled="error !== ''"
        @click="n.confirmEditor()"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.name-editor-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.name-editor-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-md);
}

.name-editor-control {
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.name-editor-control[aria-invalid="true"] {
  border-color: var(--danger);
}

.name-editor-comment {
  padding: var(--space-1) var(--space-2);
  resize: vertical;
}

.name-editor-refers {
  display: flex;
  gap: var(--space-2);
}

.name-editor-refers .name-editor-control {
  flex: 1 1 auto;
  min-width: 0;
}

.name-editor-error {
  min-height: 1.25em;
  margin: 0;
  color: var(--danger);
  font-size: var(--text-sm);
}

.name-editor-picking {
  position: fixed;
  top: var(--space-6);
  left: 50%;
  z-index: 50;
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-4);
  transform: translateX(-50%);
  background: var(--surface-base);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-3);
  color: var(--text-primary);
  font-size: var(--text-sm);
}
</style>
