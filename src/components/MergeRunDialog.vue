<script setup lang="ts">
// The mail-merge run dialog (SPEC-046 AC-21..24, AC-35..48): attach a CSV
// or JSON data source, preview the first record, see the column-to-field
// mapping, choose output, and run — producing a .zip (one document per
// record) or a single page-break-concatenated document. Template errors
// block the run before any document is produced; data errors (missing
// fields, a non-array {{#each}} target) never fail it, landing in the
// post-run report instead. BaseDialog owns the shell and — new with the
// migration — Escape-to-close, which this dialog never bound; its
// #header-extra slot carries the ✕ close button on the title row.
import { computed, ref, shallowRef, toRaw, useTemplateRef } from "vue";
import {
  asMergeDelivery,
  asMergeFormat,
  asParsedDataSourceOrNull,
  buildFieldMappingReport,
  MergeTemplateError,
  noMergeRunReport,
  parseCsvSource,
  parseJsonSource,
  previewFirstRecord,
  readFileAsText,
  runMerge,
  validateTemplate,
} from "../lib/merge";
import { asDocumentModel } from "../lib/documentModel";
import { blockText } from "../lib/textEditing";
import { DOCX_MIME_TYPE } from "../lib/docx/writer";
import { ODT_MIMETYPE } from "../lib/odf/writer";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  model: { type: Object, required: true },
  /** The source already attached elsewhere in this session (SPEC-046 "The
   * named field list" — attaching once benefits every field list), if any. */
  initialSource: { type: Object, default: null },
});

const emit = defineEmits(["attach-source", "close"]);

// Runtime Object props carry no element type (the espree constraint bars
// PropType<...> here) — these typed pass-throughs give the script below
// typed access, the same seam CommentsPane.vue's asVisibleCommentThreads
// gives it. toRaw strips any reactive wrapping Vue may have applied to the
// object prop: merge.ts's structuredClone calls need the plain object a
// reactive Proxy is not (Node/browser structured-clone cannot clone a Proxy).
const model = computed(() => asDocumentModel(toRaw(props.model)));
const initialSource = computed(() => asParsedDataSourceOrNull(toRaw(props.initialSource)));

const fileInput = useTemplateRef("fileInput");
// shallowRef, not ref: a plain ref() would deep-reactive-wrap the assigned
// ParsedDataSource (same reasoning as DocumentView.vue's `doc` — an opaque
// value read/replaced whole, never mutated field-by-field through the
// template).
const source = shallowRef(initialSource.value);
const sourceFileName = ref(initialSource.value ? "(previously attached)" : "");
const sourceError = ref("");

function pickSourceFile() {
  sourceError.value = "";
  fileInput.value?.click();
}

async function handleSourceFileChosen() {
  const file = fileInput.value?.files?.[0];
  if (fileInput.value) {
    fileInput.value.value = "";
  }
  if (!file) {
    return;
  }
  try {
    const text = await readFileAsText(file);
    const isJson = /\.json$/i.test(file.name);
    const parsed = isJson ? parseJsonSource(text) : parseCsvSource(text);
    source.value = parsed;
    sourceFileName.value = file.name;
    sourceError.value = "";
    emit("attach-source", parsed);
  } catch (error) {
    sourceError.value = error instanceof Error ? error.message : "Could not read that file";
  }
}

const templateErrors = computed(() => validateTemplate(model.value));

const mappingReport = computed(() => buildFieldMappingReport(source.value, model.value));

const previewLines = computed(() => {
  if (!source.value || source.value.records.length === 0 || templateErrors.value.length > 0) {
    return [];
  }
  try {
    const resolved = previewFirstRecord(model.value, source.value);
    return resolved.sections[0].blocks.map((block) => blockText(block)).filter((text) => text.length > 0);
  } catch (error) {
    return error instanceof MergeTemplateError ? [] : [];
  }
});

const format = ref(model.value.format === "odt" ? "odt" : "docx");
const delivery = ref("zip");
const nameField = ref("");

const scalarSourceFields = computed(() => {
  const attached = source.value;
  return attached ? attached.columns.filter((column) => !attached.arrayColumns.includes(column)) : [];
});

const canRun = computed(() => !!source.value && source.value.records.length > 0 && templateErrors.value.length === 0);

const running = ref(false);
const runReport = shallowRef(noMergeRunReport());
const runError = ref("");

function downloadBytes(bytes = new Uint8Array(), fileName = "", mimeType = "") {
  const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

async function run() {
  if (!source.value || !canRun.value) {
    return;
  }
  running.value = true;
  runError.value = "";
  try {
    const result = await runMerge(model.value, source.value, {
      format: asMergeFormat(format.value),
      delivery: asMergeDelivery(delivery.value),
      nameField: nameField.value || undefined,
    });
    const mimeType = result.delivery === "zip" ? "application/zip" : format.value === "odt" ? ODT_MIMETYPE : DOCX_MIME_TYPE;
    downloadBytes(result.bytes, result.fileName, mimeType);
    runReport.value = result.report;
  } catch (error) {
    runError.value = error instanceof Error ? error.message : "The merge could not run";
  } finally {
    running.value = false;
  }
}
</script>

<template>
  <BaseDialog
    title="Mail merge"
    testid="merge-run-dialog"
    @dismiss="emit('close')"
  >
    <template #header-extra>
      <button
        type="button"
        class="close-button"
        aria-label="Close"
        data-testid="merge-run-close"
        @click="emit('close')"
      >
        ✕
      </button>
    </template>

    <section class="merge-section">
      <h3 class="section-heading">
        Data source
      </h3>
      <button
        type="button"
        data-testid="merge-run-attach"
        @click="pickSourceFile"
      >
        {{ source ? "Change file…" : "Attach CSV or JSON…" }}
      </button>
      <input
        ref="fileInput"
        type="file"
        accept=".csv,.json"
        data-testid="merge-run-attach-input"
        class="visually-hidden"
        @change="handleSourceFileChosen"
      >
      <span
        v-if="sourceFileName"
        class="source-file-name"
        data-testid="merge-run-source-name"
      >{{ sourceFileName }}</span>
      <p
        v-if="sourceError"
        class="dialog-error"
        role="alert"
        data-testid="merge-run-source-error"
      >
        {{ sourceError }}
      </p>
      <p
        v-if="source"
        class="record-count"
        data-testid="merge-run-record-count"
      >
        {{ source.records.length }} record{{ source.records.length === 1 ? "" : "s" }}
      </p>
    </section>

    <section
      v-if="templateErrors.length > 0"
      class="merge-section"
    >
      <h3 class="section-heading error-heading">
        Template errors — the run is blocked
      </h3>
      <ul
        class="error-list"
        data-testid="merge-run-template-errors"
      >
        <li
          v-for="(error, index) in templateErrors"
          :key="index"
        >
          {{ error.message }}
        </li>
      </ul>
    </section>

    <section
      v-if="source"
      class="merge-section"
    >
      <h3 class="section-heading">
        Field mapping
      </h3>
      <ul
        class="mapping-list"
        data-testid="merge-run-mapping"
      >
        <li
          v-for="column in mappingReport.columns"
          :key="column.name"
          :data-mapped="column.mapped"
        >
          {{ column.name }} — {{ column.mapped ? "mapped" : "unmapped" }}
        </li>
      </ul>
      <p v-if="mappingReport.unmappedChips.length > 0">
        Chips with no matching column: {{ mappingReport.unmappedChips.join(", ") }}
      </p>
    </section>

    <section
      v-if="previewLines.length > 0"
      class="merge-section"
    >
      <h3 class="section-heading">
        Preview — first record
      </h3>
      <div
        class="preview-box"
        data-testid="merge-run-preview"
      >
        <p
          v-for="(line, index) in previewLines"
          :key="index"
        >
          {{ line }}
        </p>
      </div>
    </section>

    <section class="merge-section">
      <h3 class="section-heading">
        Output
      </h3>
      <label class="option-row">
        Format
        <select
          v-model="format"
          data-testid="merge-run-format"
        >
          <option value="docx">
            Word (.docx)
          </option>
          <option value="odt">
            OpenDocument (.odt)
          </option>
        </select>
      </label>
      <label class="option-row">
        Delivery
        <select
          v-model="delivery"
          data-testid="merge-run-delivery"
        >
          <option value="zip">
            One .zip, one document per record
          </option>
          <option value="single">
            Single document, one record per page
          </option>
        </select>
      </label>
      <label
        v-if="delivery === 'zip'"
        class="option-row"
      >
        Name field (optional)
        <select
          v-model="nameField"
          data-testid="merge-run-name-field"
        >
          <option value="">
            None — 001.docx, 002.docx, …
          </option>
          <option
            v-for="field in scalarSourceFields"
            :key="field"
            :value="field"
          >
            {{ field }}
          </option>
        </select>
      </label>
    </section>

    <p
      v-if="runError"
      class="dialog-error"
      role="alert"
      data-testid="merge-run-error"
    >
      {{ runError }}
    </p>

    <section
      v-if="runReport"
      class="merge-section"
    >
      <h3 class="section-heading">
        Merge report
      </h3>
      <p data-testid="merge-run-report-summary">
        {{ runReport.recordCount }} record{{ runReport.recordCount === 1 ? "" : "s" }} merged.
      </p>
      <ul
        v-if="runReport.missingFields.length > 0"
        data-testid="merge-run-report-missing"
      >
        <li
          v-for="(entry, index) in runReport.missingFields"
          :key="index"
        >
          Record {{ entry.recordIndex + 1 }}: missing "{{ entry.field }}"
        </li>
      </ul>
      <ul
        v-if="runReport.arrayFieldErrors.length > 0"
        data-testid="merge-run-report-array-errors"
      >
        <li
          v-for="(entry, index) in runReport.arrayFieldErrors"
          :key="index"
        >
          Record {{ entry.recordIndex + 1 }}: "{{ entry.field }}" is not a list
        </li>
      </ul>
    </section>

    <template #actions>
      <button
        type="button"
        data-testid="merge-run-cancel"
        @click="emit('close')"
      >
        Close
      </button>
      <button
        type="button"
        class="primary"
        data-testid="merge-run-run"
        :disabled="!canRun || running"
        @click="run"
      >
        {{ running ? "Running…" : "Run merge" }}
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* This dialog is a fixed 32rem wide (not the base's min-width knob) and
   tall enough to need the panel-level scroll cap. */
:deep(.dialog) {
  width: 32rem;
  max-height: 85vh;
  overflow-y: auto;
}

/* Its header row sat tighter above the first section than the base's 1rem. */
:deep(.dialog-header) {
  margin-bottom: 0.5rem;
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

.merge-section {
  margin: 1rem 0;
}

.section-heading {
  margin: 0 0 0.5rem;
  font-size: 0.875rem;
  font-weight: 600;
  color: #5f6368;
}

.error-heading {
  color: #b3261e;
}

.error-list {
  margin: 0;
  padding-left: 1.25rem;
  color: #b3261e;
  font-size: 0.875rem;
}

.mapping-list {
  margin: 0;
  padding-left: 1.25rem;
  font-size: 0.875rem;
}

.mapping-list li[data-mapped="false"] {
  color: #b3261e;
}

.preview-box {
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 0.5rem 0.75rem;
  max-height: 10rem;
  overflow-y: auto;
  font-size: 0.875rem;
}

.preview-box p {
  margin: 0.25rem 0;
}

.option-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  margin: 0.5rem 0;
  font-size: 0.875rem;
}

.option-row select {
  font: inherit;
}

.source-file-name {
  margin-left: 0.5rem;
  color: #5f6368;
  font-size: 0.8125rem;
}

.record-count {
  margin: 0.5rem 0 0;
  font-size: 0.8125rem;
  color: #5f6368;
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

.dialog-error {
  margin: 0.5rem 0;
  color: #b3261e;
  font-size: 0.8125rem;
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

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
