<script setup lang="ts">
// The link dialog on a cell (SPEC-135): `sheet-link-dialog` on BaseDialog,
// titled "Insert link" or "Edit link", bound to useCellLinks' `dialog` — the
// text to display (`sheet-link-text`, the active cell's display text), the
// kind (`sheet-link-kind-web` / `sheet-link-kind-place`), a web address
// (`sheet-link-address`) or a place: a visible sheet (`sheet-link-sheet`)
// and a cell or range on it (`sheet-link-ref`, checked with parseRange), or
// one of the workbook's visible defined names (`sheet-link-names`, which
// wins while chosen); the tooltip (`sheet-link-tooltip`) and the error line
// (`sheet-link-error`: a dangerous scheme is refused with LinkDialog.vue's
// "This address type isn't allowed in links", a bad reference, or the cell
// entry's refusal). Save (`sheet-link-save`, disabled until the target is
// valid; Enter in a field), Remove (`sheet-link-remove`, an existing link
// only) and Cancel (`sheet-link-cancel`, Escape). Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asCellLinks } from "../../views/spreadsheet/useCellLinks";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useCellLinks' return (its `dialog` is open). */
  links: { type: Object, required: true },
});

const l = computed(() => asCellLinks(props.links));
const state = computed(() => l.value.dialog.value);
const fields = computed(() => state.value?.fields ?? { text: "", kind: "web", address: "", sheet: "", ref: "", name: "", tooltip: "" });

function valueOf(event = new Event("input")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement ? el.value : "";
}

function setText(key = "text", event = new Event("input")) {
  if (key === "text" || key === "address" || key === "sheet" || key === "ref" || key === "name" || key === "tooltip") {
    l.value.setField(key, valueOf(event));
  }
}

function setKind(kind = "web") {
  l.value.setField("kind", kind === "place" ? "place" : "web");
}
</script>

<template>
  <BaseDialog
    :title="state?.editing ? 'Edit link' : 'Insert link'"
    testid="sheet-link-dialog"
    width="28rem"
    @dismiss="l.cancel()"
  >
    <form
      class="sheet-link-form"
      @submit.prevent="l.save()"
    >
      <label class="sheet-link-field">
        <span>Text to display</span>
        <input
          type="text"
          class="sheet-link-control"
          data-testid="sheet-link-text"
          :value="fields.text"
          @input="setText('text', $event)"
        >
      </label>
      <fieldset class="sheet-link-kinds">
        <legend>Link to</legend>
        <label>
          <input
            type="radio"
            name="sheet-link-kind"
            value="web"
            data-testid="sheet-link-kind-web"
            :checked="fields.kind === 'web'"
            @change="setKind('web')"
          >
          Web address
        </label>
        <label>
          <input
            type="radio"
            name="sheet-link-kind"
            value="place"
            data-testid="sheet-link-kind-place"
            :checked="fields.kind === 'place'"
            @change="setKind('place')"
          >
          Place in this workbook
        </label>
      </fieldset>
      <label
        v-if="fields.kind === 'web'"
        class="sheet-link-field"
      >
        <span>Address</span>
        <input
          type="text"
          class="sheet-link-control"
          data-testid="sheet-link-address"
          placeholder="https://"
          autofocus
          spellcheck="false"
          :aria-invalid="l.error.value !== '' ? 'true' : undefined"
          :value="fields.address"
          @input="setText('address', $event)"
        >
      </label>
      <template v-else>
        <div class="sheet-link-place">
          <label class="sheet-link-field">
            <span>Sheet</span>
            <select
              class="sheet-link-control"
              data-testid="sheet-link-sheet"
              :value="fields.sheet"
              @change="setText('sheet', $event)"
            >
              <option
                v-for="sheet in state?.sheets ?? []"
                :key="sheet"
                :value="sheet"
              >
                {{ sheet }}
              </option>
            </select>
          </label>
          <label class="sheet-link-field">
            <span>Cell reference</span>
            <input
              type="text"
              class="sheet-link-control"
              data-testid="sheet-link-ref"
              autofocus
              spellcheck="false"
              :aria-invalid="l.error.value !== '' ? 'true' : undefined"
              :value="fields.ref"
              @input="setText('ref', $event)"
            >
          </label>
        </div>
        <label class="sheet-link-field">
          <span>Or a defined name</span>
          <select
            class="sheet-link-control"
            data-testid="sheet-link-names"
            :disabled="(state?.names.length ?? 0) === 0"
            :value="fields.name"
            @change="setText('name', $event)"
          >
            <option value="">
              (none)
            </option>
            <option
              v-for="name in state?.names ?? []"
              :key="name.value"
              :value="name.value"
            >
              {{ name.label }}
            </option>
          </select>
        </label>
      </template>
      <label class="sheet-link-field">
        <span>Tooltip</span>
        <input
          type="text"
          class="sheet-link-control"
          data-testid="sheet-link-tooltip"
          :value="fields.tooltip"
          @input="setText('tooltip', $event)"
        >
      </label>
      <p
        class="sheet-link-error"
        role="alert"
        data-testid="sheet-link-error"
      >
        {{ l.error.value }}
      </p>
      <button
        type="submit"
        hidden
      />
    </form>
    <template #actions>
      <UiButton
        v-if="state?.editing"
        variant="secondary"
        data-testid="sheet-link-remove"
        @click="l.remove()"
      >
        Remove link
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="sheet-link-cancel"
        @click="l.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-link-save"
        :disabled="!l.canSave.value"
        @click="l.save()"
      >
        Save
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.sheet-link-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.sheet-link-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  min-width: 0;
  color: var(--text-primary);
  font-size: var(--text-md);
}

.sheet-link-control {
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.sheet-link-control[aria-invalid="true"] {
  border-color: var(--danger);
}

.sheet-link-kinds {
  display: flex;
  gap: var(--space-4);
  margin: 0;
  padding: 0;
  border: none;
  color: var(--text-primary);
  font-size: var(--text-md);
}

.sheet-link-kinds legend {
  float: left;
  margin-right: var(--space-2);
  padding: 0;
}

.sheet-link-place {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-2);
}

.sheet-link-error {
  min-height: 1.25em;
  margin: 0;
  color: var(--danger);
  font-size: var(--text-sm);
}
</style>
