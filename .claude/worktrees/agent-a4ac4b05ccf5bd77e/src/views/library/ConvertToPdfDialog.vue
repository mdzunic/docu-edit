<script setup lang="ts">
// "Convert to PDF…" (SPEC-108): a BaseDialog shell around the shared sheet
// setup panel and its preview, a title field (defaulting to the image's
// title) and three actions — Save to library, Download, Cancel. All logic
// is useConvertToPdf's; a successful save toasts and emits `saved` with the
// new document's id, so HomeView only refreshes the list and closes.
// Runtime prop declarations (espree lints .vue scripts).
import BaseDialog from "../../components/BaseDialog.vue";
import UiButton from "../../components/ui/UiButton.vue";
import SheetSetupPanel from "../../components/assembly/SheetSetupPanel.vue";
import SheetPreview from "../../components/assembly/SheetPreview.vue";
import { useToasts } from "../../composables/useToasts";
import { asConvertSteps, useConvertToPdf } from "./useConvertToPdf";

const props = defineProps({
  /** The library id of the image to convert. */
  sourceId: { type: String, required: true },
  /** Test seam: the browser-bound conversion steps (canvas, decoding). */
  steps: { type: Object, default: undefined },
});

const emit = defineEmits(["close", "saved"]);

const { addToast } = useToasts();
const conv = useConvertToPdf({ sourceId: props.sourceId, steps: asConvertSteps(props.steps) });
const sheet = conv.sheet;

async function onSave() {
  const saved = await conv.save();
  if (saved) {
    addToast(`Saved “${saved.title}” to the library`);
    emit("saved", saved.id);
  }
}

async function onDownload() {
  await conv.download();
}
</script>

<template>
  <BaseDialog
    title="Convert to PDF"
    testid="convert-pdf-dialog"
    width="min(46rem, 90vw)"
    @dismiss="emit('close')"
  >
    <div class="convert-body">
      <div class="convert-setup">
        <label class="convert-title">
          <span>Title</span>
          <input
            v-model="conv.title.value"
            type="text"
            data-testid="convert-pdf-title"
            autofocus
          >
        </label>
        <p
          v-if="conv.isSvg.value"
          class="convert-note"
          data-testid="convert-pdf-svg-notice"
        >
          This SVG is rasterised at its placed size and the chosen resolution; the PDF holds pixels, not vectors.
        </p>
        <SheetSetupPanel :controller="sheet" />
      </div>
      <SheetPreview
        class="convert-preview"
        :geometry="sheet.geometry.value"
        :placement="sheet.placement.value"
        :unit="sheet.unit.value"
        :background="conv.previewBackground.value"
        :image-url="conv.previewUrl.value"
      />
    </div>
    <p
      v-if="conv.blockedReason.value"
      class="convert-note"
      data-testid="convert-pdf-blocked"
    >
      {{ conv.blockedReason.value }}
    </p>
    <p
      v-if="conv.failure.value"
      class="convert-failure"
      data-testid="convert-pdf-failure"
      role="alert"
    >
      {{ conv.failure.value }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="convert-pdf-cancel"
        @click="emit('close')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="convert-pdf-download"
        :disabled="!conv.canConvert.value"
        @click="onDownload"
      >
        Download
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="convert-pdf-save"
        :disabled="!conv.canConvert.value"
        @click="onSave"
      >
        Save to library
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.convert-body {
  display: grid;
  grid-template-columns: minmax(0, 1fr) minmax(0, 16rem);
  gap: var(--space-5);
  max-height: 70vh;
  overflow-y: auto;
  margin-bottom: var(--space-3);
}

.convert-setup {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.convert-title {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.convert-title input {
  font: inherit;
  font-size: var(--text-md);
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.convert-title input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.convert-note,
.convert-failure {
  margin: 0;
  font-size: var(--text-sm);
}

.convert-note {
  color: var(--text-secondary);
}

.convert-failure {
  color: var(--text-danger);
}

@media (max-width: 40rem) {
  .convert-body {
    grid-template-columns: minmax(0, 1fr);
  }
}
</style>
