<script setup lang="ts">
// The "File → Properties" dialog (SPEC-041): General/Description/Statistics/
// Custom Properties tabs, plus two dialog-level controls that apply
// regardless of which tab is active — the document identity (used by
// comments and track changes, SPEC-038/SPEC-039) and the "Remove personal
// information on save" toggle (AC-12a/AC-15: neither ambiguity in the intake
// names a specific tab for these, so they live outside the tabstrip like a
// settings panel's shared footer).
//
// Every field is committed live (no separate Save/Cancel): each edit clones
// the local `propertiesDraft`, plain-JSON-serialised so nothing reactive
// ever leaks into DocumentModel, and emits it as one `update-properties`
// event — the parent assigns it straight onto `doc.value.properties`, the
// same "mutate the stored model directly" contract DocumentView.vue's own
// `handleTitleUpdate` already uses. Runtime prop/emit declarations, and
// every helper below takes plain untyped/default-valued parameters rather
// than type annotations or generics — see AppShell.vue's espree comment.
// The few places that need a real static type for vue-tsc to infer
// (propertiesDraft's shape, the per-row/new-property drafts) get it from a
// typed factory function declared in a plain `.ts` module instead — the
// same trick documentModel.ts's noActiveNote/noBlocks use.
import { computed, nextTick, onMounted, reactive, ref, useTemplateRef, watch } from "vue";
import { asCustomPropertyType, cloneDocumentProperties, CUSTOM_PROPERTY_TYPES } from "../lib/documentModel";
import {
  canonicalCustomPropertyValue,
  createCustomProperty,
  customPropertyNameError,
  customPropertyValueError,
  defaultCustomPropertyValue,
  emptyCustomPropertyRowDrafts,
  emptyNewCustomPropertyDraft,
} from "../lib/customProperties";
import { documentFormatLabel, formatFileSize, formatTimestamp } from "../lib/format";
import { docxFileName } from "../lib/docx/writer";
import { odtFileName } from "../lib/odf/writer";
import { mdFileName } from "../lib/md/writer";
import { CORE_EXTRA_FIELD_LABELS } from "../lib/docx/docProps";
import { META_EXTRA_FIELD_LABELS } from "../lib/odf/meta";

const props = defineProps({
  model: { type: Object, required: true },
  /** lib/wordCount's CountStats for the document body (SPEC-036, AC-4). */
  stats: { type: Object, required: true },
  pages: { type: Number, required: true },
  /** Bytes, or null while still being computed / for a document that has
   * never been saved to a file (AC-18a). */
  fileSize: { type: [Number, null], default: null },
  /** The settable identity used by comments and track changes
   * (lib/commentIdentity.ts) — AC-15. */
  authorIdentity: { type: String, required: true },
});

const emit = defineEmits(["close", "update-properties", "update-author-identity"]);

const TABS = [
  { id: "general", label: "General" },
  { id: "description", label: "Description" },
  { id: "statistics", label: "Statistics" },
  { id: "custom", label: "Custom Properties" },
];
const activeTab = ref("general");
const tablist = useTemplateRef("tablist");

function focusTab(id = "") {
  void nextTick(() => {
    const el = tablist.value?.querySelector(`[data-testid="properties-tab-${id}"]`);
    if (el instanceof HTMLElement) {
      el.focus();
    }
  });
}

function onTabKeydown(event = new KeyboardEvent("keydown")) {
  const index = TABS.findIndex((tab) => tab.id === activeTab.value);
  let nextIndex = -1;
  if (event.key === "ArrowRight") {
    nextIndex = (index + 1) % TABS.length;
  } else if (event.key === "ArrowLeft") {
    nextIndex = (index - 1 + TABS.length) % TABS.length;
  } else if (event.key === "Home") {
    nextIndex = 0;
  } else if (event.key === "End") {
    nextIndex = TABS.length - 1;
  } else {
    return;
  }
  event.preventDefault();
  activeTab.value = TABS[nextIndex].id;
  focusTab(activeTab.value);
}

// --- General tab (AC-2, AC-18a) --------------------------------------------

const modelRef = computed(() => props.model);

const fileName = computed(() => {
  const model = modelRef.value;
  if (!model.format) {
    return "";
  }
  if (model.format === "odt") {
    return odtFileName(model.title);
  }
  if (model.format === "md") {
    return mdFileName(model.title);
  }
  return docxFileName(model.title);
});
const formatLabel = computed(() => documentFormatLabel(modelRef.value.format));
const sizeLabel = computed(() => {
  if (!modelRef.value.format) {
    return "";
  }
  return props.fileSize === null ? "Calculating…" : formatFileSize(props.fileSize);
});
const createdLabel = computed(() => {
  const created = modelRef.value.properties?.created;
  return created ? formatTimestamp(new Date(created)) : "";
});
const modifiedLabel = computed(() => {
  const modified = modelRef.value.properties?.modified;
  return modified ? formatTimestamp(new Date(modified)) : "";
});

// --- Description tab + custom properties + toggle (AC-3, AC-5..14) --------

const propertiesDraft = reactive(cloneDocumentProperties(modelRef.value.properties));

function commit() {
  emit("update-properties", JSON.parse(JSON.stringify(propertiesDraft)));
}

const extraFields = computed(() => modelRef.value.properties?.extraFields ?? []);
function extraFieldLabel(name = "") {
  return CORE_EXTRA_FIELD_LABELS[name] ?? META_EXTRA_FIELD_LABELS[name] ?? name;
}

const removePersonalInfo = computed({
  get: () => propertiesDraft.removePersonalInfoOnSave === true,
  set: (value) => {
    propertiesDraft.removePersonalInfoOnSave = value;
    commit();
  },
});
function onToggleRemovePersonalInfo(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    removePersonalInfo.value = target.checked;
  }
}

// The document identity (AC-15/AC-16): a local draft, not a direct
// `:value="authorIdentity"` binding — same reason CommentsPane.vue's own
// authorDraft isn't (see its header comment): an unrelated re-render mid-
// keystroke would otherwise snap the field back via the DOM's live-value
// reconciliation. Committed on change/Enter, matching CommentsPane's own
// convention for the same field.
const identityDraft = ref(props.authorIdentity);
watch(
  () => props.authorIdentity,
  (name) => {
    identityDraft.value = name;
  },
);
function commitIdentity() {
  emit("update-author-identity", identityDraft.value);
}

// --- Custom Properties tab (AC-7..11) --------------------------------------

const unsupportedCustomProperties = computed(() => modelRef.value.properties?.unsupportedCustomProperties ?? []);

const newProperty = reactive(emptyNewCustomPropertyDraft());
watch(
  () => newProperty.type,
  (type) => {
    newProperty.value = defaultCustomPropertyValue(asCustomPropertyType(type));
  },
);
const newNameError = computed(() =>
  newProperty.name === ""
    ? ""
    : customPropertyNameError(newProperty.name, propertiesDraft.customProperties.map((property) => property.name)),
);
const newValueError = computed(() => customPropertyValueError(newProperty.type, newProperty.value));
const canAddProperty = computed(
  () => newProperty.name.trim() !== "" && newNameError.value === "" && newValueError.value === "",
);
function addCustomProperty() {
  if (!canAddProperty.value) {
    return;
  }
  propertiesDraft.customProperties.push(createCustomProperty(newProperty.name, newProperty.type, newProperty.value));
  newProperty.name = "";
  newProperty.type = "text";
  newProperty.value = "";
  commit();
}

/** Per-row in-progress text, keyed by array index, so an invalid keystroke
 * shows its error without touching `propertiesDraft` (AC-11) — mirrors the
 * identity field's draft-vs-committed split above, per-row. */
const rowDrafts = reactive(emptyCustomPropertyRowDrafts());

function committedRow(index = 0) {
  return propertiesDraft.customProperties[index];
}
function rowName(index = 0) {
  return rowDrafts[index]?.name ?? committedRow(index).name;
}
function rowValue(index = 0) {
  return rowDrafts[index]?.value ?? committedRow(index).value;
}
function otherNames(index = 0) {
  return propertiesDraft.customProperties.filter((_property, i) => i !== index).map((property) => property.name);
}
function rowNameError(index = 0) {
  return customPropertyNameError(rowName(index), otherNames(index));
}
function rowValueError(index = 0) {
  return customPropertyValueError(committedRow(index).type, rowValue(index));
}
function onRowNameInput(index = 0, value = "") {
  rowDrafts[index] = { name: value, value: rowValue(index) };
  if (customPropertyNameError(value, otherNames(index)) === "") {
    committedRow(index).name = value.trim();
    commit();
  }
}
function onRowValueInput(index = 0, value = "") {
  rowDrafts[index] = { name: rowName(index), value };
  const type = committedRow(index).type;
  if (customPropertyValueError(type, value) === "") {
    committedRow(index).value = canonicalCustomPropertyValue(type, value);
    commit();
  }
}
function onRowNameEvent(index = 0, event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    onRowNameInput(index, target.value);
  }
}
function onRowValueEvent(index = 0, event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement || target instanceof HTMLSelectElement) {
    onRowValueInput(index, target.value);
  }
}
function removeCustomProperty(index = 0) {
  propertiesDraft.customProperties.splice(index, 1);
  // Rows after the removed one shift left in `customProperties`; reindex
  // their in-progress drafts to match, or a later row's uncommitted edit
  // would land on the wrong row after removal.
  const shifted = emptyCustomPropertyRowDrafts();
  for (const key of Object.keys(rowDrafts)) {
    const i = Number(key);
    if (i > index) {
      shifted[i - 1] = rowDrafts[i];
    } else if (i < index) {
      shifted[i] = rowDrafts[i];
    }
    delete rowDrafts[i];
  }
  Object.assign(rowDrafts, shifted);
  commit();
}

const closeButton = useTemplateRef("closeButton");
onMounted(async () => {
  await nextTick();
  closeButton.value?.focus();
});
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="properties-dialog"
    @click.self="emit('close')"
    @keydown.escape="emit('close')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Document properties"
    >
      <h2 class="dialog-title">
        Document properties
      </h2>

      <div
        ref="tablist"
        class="tablist"
        role="tablist"
        aria-label="Document properties tabs"
      >
        <button
          v-for="tab in TABS"
          :id="`properties-tab-${tab.id}`"
          :key="tab.id"
          type="button"
          role="tab"
          :aria-selected="activeTab === tab.id ? 'true' : 'false'"
          :aria-controls="`properties-panel-${tab.id}`"
          :tabindex="activeTab === tab.id ? 0 : -1"
          class="tab"
          :class="{ active: activeTab === tab.id }"
          :data-testid="`properties-tab-${tab.id}`"
          @click="activeTab = tab.id"
          @keydown="onTabKeydown"
        >
          {{ tab.label }}
        </button>
      </div>

      <div class="tab-panels">
        <div
          v-show="activeTab === 'general'"
          id="properties-panel-general"
          role="tabpanel"
          aria-labelledby="properties-tab-general"
          data-testid="properties-panel-general"
        >
          <dl class="fields-grid">
            <dt>File name</dt>
            <dd data-testid="properties-general-filename">
              {{ fileName }}
            </dd>
            <dt>Format</dt>
            <dd data-testid="properties-general-format">
              {{ formatLabel }}
            </dd>
            <dt>Size</dt>
            <dd data-testid="properties-general-size">
              {{ sizeLabel }}
            </dd>
            <dt>Created</dt>
            <dd data-testid="properties-general-created">
              {{ createdLabel }}
            </dd>
            <dt>Modified</dt>
            <dd data-testid="properties-general-modified">
              {{ modifiedLabel }}
            </dd>
          </dl>
        </div>

        <div
          v-show="activeTab === 'description'"
          id="properties-panel-description"
          role="tabpanel"
          aria-labelledby="properties-tab-description"
          data-testid="properties-panel-description"
        >
          <label
            class="dialog-label"
            for="properties-title"
          >Title</label>
          <input
            id="properties-title"
            v-model="propertiesDraft.title"
            class="dialog-input"
            data-testid="properties-title-input"
            @change="commit"
          >

          <label
            class="dialog-label"
            for="properties-subject"
          >Subject</label>
          <input
            id="properties-subject"
            v-model="propertiesDraft.subject"
            class="dialog-input"
            data-testid="properties-subject-input"
            @change="commit"
          >

          <label
            class="dialog-label"
            for="properties-keywords"
          >Keywords</label>
          <input
            id="properties-keywords"
            v-model="propertiesDraft.keywords"
            class="dialog-input"
            data-testid="properties-keywords-input"
            @change="commit"
          >

          <label
            class="dialog-label"
            for="properties-comments"
          >Comments</label>
          <textarea
            id="properties-comments"
            v-model="propertiesDraft.comments"
            class="dialog-input dialog-textarea"
            data-testid="properties-comments-input"
            @change="commit"
          />

          <label
            class="dialog-label"
            for="properties-author"
          >Author</label>
          <input
            id="properties-author"
            v-model="propertiesDraft.author"
            class="dialog-input"
            data-testid="properties-author-input"
            @change="commit"
          >

          <label
            class="dialog-label"
            for="properties-last-modified-by"
          >Last modified by</label>
          <input
            id="properties-last-modified-by"
            v-model="propertiesDraft.lastModifiedBy"
            class="dialog-input"
            data-testid="properties-last-modified-by-input"
            @change="commit"
          >

          <div
            v-if="extraFields.length > 0"
            class="extra-fields"
            data-testid="properties-extra-fields"
          >
            <h3 class="section-title">
              Other properties
            </h3>
            <p class="hint">
              Read from the file; not editable here, but preserved when you save.
            </p>
            <dl class="fields-grid">
              <template
                v-for="field in extraFields"
                :key="field.name"
              >
                <dt>{{ extraFieldLabel(field.name) }}</dt>
                <dd>{{ field.value }}</dd>
              </template>
            </dl>
          </div>
        </div>

        <div
          v-show="activeTab === 'statistics'"
          id="properties-panel-statistics"
          role="tabpanel"
          aria-labelledby="properties-tab-statistics"
          data-testid="properties-panel-statistics"
        >
          <dl class="fields-grid">
            <dt>Pages</dt>
            <dd data-testid="properties-statistics-pages">
              {{ pages }}
            </dd>
            <dt>Words</dt>
            <dd data-testid="properties-statistics-words">
              {{ stats.words }}
            </dd>
            <dt>Characters (no spaces)</dt>
            <dd data-testid="properties-statistics-characters-without-spaces">
              {{ stats.charactersWithoutSpaces }}
            </dd>
            <dt>Characters (with spaces)</dt>
            <dd data-testid="properties-statistics-characters-with-spaces">
              {{ stats.charactersWithSpaces }}
            </dd>
            <dt>Paragraphs</dt>
            <dd data-testid="properties-statistics-paragraphs">
              {{ stats.paragraphs }}
            </dd>
            <dt>Sentences</dt>
            <dd data-testid="properties-statistics-sentences">
              {{ stats.sentences }}
            </dd>
          </dl>
        </div>

        <div
          v-show="activeTab === 'custom'"
          id="properties-panel-custom"
          role="tabpanel"
          aria-labelledby="properties-tab-custom"
          data-testid="properties-panel-custom"
        >
          <div
            v-for="(property, index) in propertiesDraft.customProperties"
            :key="index"
            class="custom-property-row"
            data-testid="custom-property-row"
          >
            <input
              class="dialog-input"
              :value="rowName(index)"
              data-testid="custom-property-name"
              aria-label="Custom property name"
              @input="onRowNameEvent(index, $event)"
            >
            <span class="custom-property-type">{{ property.type }}</span>
            <input
              v-if="property.type !== 'boolean'"
              class="dialog-input"
              :value="rowValue(index)"
              :placeholder="property.type === 'date' ? 'YYYY-MM-DD' : ''"
              data-testid="custom-property-value"
              aria-label="Custom property value"
              @input="onRowValueEvent(index, $event)"
            >
            <select
              v-else
              class="dialog-input"
              :value="rowValue(index)"
              data-testid="custom-property-value"
              aria-label="Custom property value"
              @change="onRowValueEvent(index, $event)"
            >
              <option value="true">
                Yes
              </option>
              <option value="false">
                No
              </option>
            </select>
            <button
              type="button"
              data-testid="custom-property-remove"
              aria-label="Remove custom property"
              @click="removeCustomProperty(index)"
            >
              Remove
            </button>
            <p
              v-if="rowNameError(index)"
              class="dialog-error"
              data-testid="custom-property-name-error"
            >
              {{ rowNameError(index) }}
            </p>
            <p
              v-if="rowValueError(index)"
              class="dialog-error"
              data-testid="custom-property-value-error"
            >
              {{ rowValueError(index) }}
            </p>
          </div>

          <div
            v-for="property in unsupportedCustomProperties"
            :key="property.name"
            class="custom-property-row unsupported"
            data-testid="custom-property-unsupported-row"
          >
            <span>{{ property.name }}</span>
            <span class="custom-property-type">{{ property.rawType }} (unsupported type, read-only)</span>
          </div>

          <h3 class="section-title">
            Add a custom property
          </h3>
          <div class="custom-property-row">
            <input
              v-model="newProperty.name"
              class="dialog-input"
              placeholder="Name"
              data-testid="custom-property-new-name"
              @keydown.enter="addCustomProperty"
            >
            <select
              v-model="newProperty.type"
              class="dialog-input"
              data-testid="custom-property-new-type"
            >
              <option
                v-for="type in CUSTOM_PROPERTY_TYPES"
                :key="type"
                :value="type"
              >
                {{ type }}
              </option>
            </select>
            <input
              v-if="newProperty.type !== 'boolean'"
              v-model="newProperty.value"
              class="dialog-input"
              :placeholder="newProperty.type === 'date' ? 'YYYY-MM-DD' : 'Value'"
              data-testid="custom-property-new-value"
              @keydown.enter="addCustomProperty"
            >
            <select
              v-else
              v-model="newProperty.value"
              class="dialog-input"
              data-testid="custom-property-new-value"
            >
              <option value="true">
                Yes
              </option>
              <option value="false">
                No
              </option>
            </select>
            <button
              type="button"
              class="primary"
              data-testid="custom-property-add"
              :disabled="!canAddProperty"
              @click="addCustomProperty"
            >
              Add
            </button>
          </div>
          <p
            v-if="newNameError"
            class="dialog-error"
            data-testid="custom-property-new-name-error"
          >
            {{ newNameError }}
          </p>
          <p
            v-if="newProperty.value && newValueError"
            class="dialog-error"
            data-testid="custom-property-new-value-error"
          >
            {{ newValueError }}
          </p>
        </div>
      </div>

      <div class="dialog-footer">
        <label
          class="dialog-label"
          for="properties-identity"
        >Your name (used for comments and tracked changes)</label>
        <input
          id="properties-identity"
          v-model="identityDraft"
          class="dialog-input"
          data-testid="properties-identity-input"
          @change="commitIdentity"
          @keydown.enter="commitIdentity"
        >

        <label class="checkbox-field">
          <input
            type="checkbox"
            :checked="removePersonalInfo"
            data-testid="properties-remove-personal-info-toggle"
            @change="onToggleRemovePersonalInfo"
          >
          <span>Remove personal information on save</span>
        </label>
        <p
          v-if="removePersonalInfo"
          class="hint"
          data-testid="properties-remove-personal-info-note"
        >
          Removes the author and last-modified-by fields; replaces comment and
          tracked-change authors with "Author 1", "Author 2", etc.; and
          strips revision timestamps.
        </p>
      </div>

      <div class="dialog-actions">
        <button
          ref="closeButton"
          type="button"
          class="primary"
          data-testid="properties-close"
          @click="emit('close')"
        >
          Close
        </button>
      </div>
    </div>
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
  display: flex;
  flex-direction: column;
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  width: 34rem;
  max-width: 90vw;
  max-height: 85vh;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.tablist {
  display: flex;
  gap: 0.25rem;
  border-bottom: 1px solid #e0e0e0;
  margin-bottom: 1rem;
}

.tab {
  font: inherit;
  padding: 0.5rem 0.75rem;
  border: none;
  border-bottom: 2px solid transparent;
  background: transparent;
  cursor: pointer;
  color: #5f6368;
}

.tab.active {
  color: #2c3e50;
  border-bottom-color: #2c3e50;
  font-weight: 600;
}

.tab-panels {
  overflow-y: auto;
  flex: 1 1 auto;
  min-height: 0;
}

.fields-grid {
  display: grid;
  grid-template-columns: 10rem 1fr;
  gap: 0.375rem 1rem;
  margin: 0 0 1rem;
  font-size: 0.8125rem;
}

.fields-grid dt {
  color: #5f6368;
}

.fields-grid dd {
  margin: 0;
  color: #202124;
  word-break: break-word;
}

.dialog-label {
  display: block;
  margin: 0.5rem 0 0.25rem;
  font-size: 0.8125rem;
  color: #5f6368;
}

.dialog-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.dialog-textarea {
  min-height: 4rem;
  resize: vertical;
}

.dialog-error {
  margin: 0.25rem 0 0;
  color: #b3261e;
  font-size: 0.75rem;
}

.section-title {
  margin: 1rem 0 0.25rem;
  font-size: 0.8125rem;
  font-weight: 600;
  color: #3c4043;
}

.hint {
  margin: 0 0 0.5rem;
  font-size: 0.75rem;
  color: #5f6368;
}

.extra-fields {
  margin-top: 1rem;
  border-top: 1px solid #e0e0e0;
  padding-top: 0.5rem;
}

.custom-property-row {
  display: grid;
  grid-template-columns: 1fr 5.5rem 1fr auto;
  gap: 0.375rem;
  align-items: center;
  margin-bottom: 0.5rem;
}

.custom-property-row.unsupported {
  grid-template-columns: 1fr 1fr;
  color: #5f6368;
  font-size: 0.8125rem;
}

.custom-property-type {
  font-size: 0.75rem;
  color: #5f6368;
  text-transform: capitalize;
}

.dialog-footer {
  border-top: 1px solid #e0e0e0;
  padding-top: 0.75rem;
  margin-top: 0.5rem;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.8125rem;
  color: #3c4043;
  margin-top: 0.75rem;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary,
.custom-property-row button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.custom-property-row button {
  font: inherit;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.custom-property-row button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
