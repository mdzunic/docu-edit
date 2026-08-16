<script setup lang="ts">
// "Insert → Field…" (SPEC-053 AC-1): lists exactly the Field-catalogue types
// (cross-references have their own dialog), exposes per-type options, and
// shows a live preview of the value it will insert — computed through the same
// evaluator the document uses, so the preview and the inserted result agree.
// Merge chips never appear here (SPEC-046 owns those). Runtime prop/emit
// declarations per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { asDocumentModel } from "../lib/documentModel";
import { asFieldSpec, asStringList, DATE_TIME_CATALOGUE, evaluateField } from "../lib/fields";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** The open document model, for the live preview's evaluation. */
  model: { type: Object, required: true },
  /** The clock instant (ms) the preview evaluates against. */
  nowMs: { type: Number, default: 0 },
  /** The document's file name, for the fileName field's preview. */
  fileName: { type: String, default: "" },
  /** The document's custom property names (SPEC-041), for the docProperty
     field's picker. */
  propertyNames: { type: Array, default: () => [] },
});

const emit = defineEmits(["insert", "cancel"]);

// The catalogue this dialog offers, in menu order — kind plus display label.
const TYPES = [
  { kind: "date", label: "Date" },
  { kind: "time", label: "Time" },
  { kind: "author", label: "Author" },
  { kind: "docTitle", label: "Title" },
  { kind: "docSubject", label: "Subject" },
  { kind: "docKeywords", label: "Keywords" },
  { kind: "docProperty", label: "Document property" },
  { kind: "fileName", label: "File name" },
  { kind: "page", label: "Page number" },
  { kind: "numpages", label: "Total pages" },
  { kind: "wordCount", label: "Word count" },
];

const kind = ref("date");
const dateFormat = ref(DATE_TIME_CATALOGUE[0].format);
const timeFormat = ref(DATE_TIME_CATALOGUE.find((entry) => entry.isTime)?.format ?? "HH:mm");
const includePath = ref(false);
const numberFormat = ref("decimal");
const propertyNames = computed(() => asStringList(props.propertyNames));
const propertyName = ref(propertyNames.value[0] ?? "");

const dateFormats = DATE_TIME_CATALOGUE.filter((entry) => !entry.isTime);
const timeFormats = DATE_TIME_CATALOGUE.filter((entry) => entry.isTime);

/** The field spec the current selections describe (without `cached` — the
   parent computes that at insert time). */
const draftField = computed(() => {
  switch (kind.value) {
    case "date":
      return { kind: "date", format: dateFormat.value, cached: "" };
    case "time":
      return { kind: "time", format: timeFormat.value, cached: "" };
    case "docProperty":
      return { kind: "docProperty", name: propertyName.value, cached: "" };
    case "fileName":
      return includePath.value ? { kind: "fileName", path: true, cached: "" } : { kind: "fileName", cached: "" };
    case "page":
      return { kind: "page", format: numberFormat.value };
    case "numpages":
      return { kind: "numpages", format: numberFormat.value };
    default:
      return { kind: kind.value, cached: "" };
  }
});

const preview = computed(() => {
  try {
    return evaluateField(asFieldSpec(draftField.value), {
      model: asDocumentModel(props.model),
      now: new Date(props.nowMs),
      fileName: props.fileName,
      pageNumber: 1,
      totalPages: 1,
    });
  } catch {
    return "";
  }
});

function insert() {
  emit("insert", draftField.value);
}
</script>

<template>
  <BaseDialog
    title="Field"
    testid="field-dialog"
    width="28rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-body">
      <label class="row">
        <span class="row-label">Field type</span>
        <select
          v-model="kind"
          class="row-input"
          data-testid="field-type"
          autofocus
        >
          <option
            v-for="type in TYPES"
            :key="type.kind"
            :value="type.kind"
          >
            {{ type.label }}
          </option>
        </select>
      </label>

      <label
        v-if="kind === 'date'"
        class="row"
      >
        <span class="row-label">Format</span>
        <select
          v-model="dateFormat"
          class="row-input"
          data-testid="field-date-format"
        >
          <option
            v-for="entry in dateFormats"
            :key="entry.format"
            :value="entry.format"
          >
            {{ entry.format }}
          </option>
        </select>
      </label>

      <label
        v-if="kind === 'time'"
        class="row"
      >
        <span class="row-label">Format</span>
        <select
          v-model="timeFormat"
          class="row-input"
          data-testid="field-time-format"
        >
          <option
            v-for="entry in timeFormats"
            :key="entry.format"
            :value="entry.format"
          >
            {{ entry.format }}
          </option>
        </select>
      </label>

      <label
        v-if="kind === 'docProperty'"
        class="row"
      >
        <span class="row-label">Property</span>
        <select
          v-model="propertyName"
          class="row-input"
          data-testid="field-property-name"
        >
          <option
            v-for="name in propertyNames"
            :key="name"
            :value="name"
          >
            {{ name }}
          </option>
        </select>
      </label>

      <label
        v-if="kind === 'fileName'"
        class="row check-row"
      >
        <input
          v-model="includePath"
          type="checkbox"
          data-testid="field-filename-path"
        >
        Include full path
      </label>

      <label
        v-if="kind === 'page' || kind === 'numpages'"
        class="row"
      >
        <span class="row-label">Number format</span>
        <select
          v-model="numberFormat"
          class="row-input"
          data-testid="field-number-format"
        >
          <option value="decimal">
            1, 2, 3
          </option>
          <option value="lowerRoman">
            i, ii, iii
          </option>
          <option value="lowerAlpha">
            a, b, c
          </option>
        </select>
      </label>

      <div class="preview">
        <span class="row-label">Preview</span>
        <span
          class="preview-value"
          data-testid="field-preview"
        >{{ preview || "—" }}</span>
      </div>
    </div>

    <template #actions>
      <button
        type="button"
        data-testid="field-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="field-insert"
        @click="insert"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field-body {
  display: flex;
  flex-direction: column;
  gap: 0.75rem;
}

.row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  font-size: 0.9375rem;
}

.check-row {
  justify-content: flex-start;
  gap: 0.5rem;
}

.row-label {
  color: #5f6368;
}

.row-input {
  font: inherit;
  min-width: 14rem;
}

.preview {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  padding-top: 0.5rem;
  border-top: 1px solid #e0e0e0;
}

.preview-value {
  font-weight: 600;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}
</style>
