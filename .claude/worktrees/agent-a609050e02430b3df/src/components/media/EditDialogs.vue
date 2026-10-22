<script setup lang="ts">
// The image editor's dialogs (SPEC-112): Revert to original (naming the
// edits it discards), Flatten and discard original (named as permanent and
// not undoable), the unsaved-edits check on Done, and Export — SPEC-108's
// format and background rules for a flattened PNG or JPEG download.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { asImageEditor } from "../../views/media/useImageEditor";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const a = computed(() => e.value.actions);
const editCount = computed(() => e.value.history.value.length);
const outputSize = computed(() => e.value.commands.outputSize.value);

function text(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement || target instanceof HTMLSelectElement ? target.value : "";
}

function chooseFormat(event = new Event("change")) {
  a.value.exportFormat.value = text(event) === "jpeg" ? "jpeg" : "png";
}

function setQuality(event = new Event("change")) {
  const value = Number(text(event));
  a.value.exportQuality.value = Number.isFinite(value) ? Math.min(100, Math.max(1, Math.round(value))) : 90;
}
</script>

<template>
  <BaseDialog
    v-if="e.revertPending.value"
    title="Revert to original"
    testid="edit-revert-dialog"
    width="26rem"
    @dismiss="e.cancelRevert()"
  >
    <p
      class="message"
      data-testid="edit-revert-message"
    >
      Discard {{ editCount }} {{ editCount === 1 ? "edit" : "edits" }} — {{ e.revertSummary.value }} — and
      return to the original image? You can undo this until you leave the editor.
    </p>
    <template #actions>
      <UiButton
        data-testid="edit-revert-cancel"
        @click="e.cancelRevert()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="edit-revert-confirm"
        @click="e.confirmRevert()"
      >
        Revert to original
      </UiButton>
    </template>
  </BaseDialog>

  <BaseDialog
    v-if="a.flattenPending.value"
    title="Flatten and discard original"
    testid="edit-flatten-dialog"
    width="28rem"
    @dismiss="a.cancelFlatten()"
  >
    <p
      class="message"
      data-testid="edit-flatten-message"
    >
      This permanently replaces the stored original with the edited image. Every
      edit is flattened into the pixels and whatever the redactions hide is
      destroyed. This is permanent and cannot be undone.
    </p>
    <template #actions>
      <UiButton
        data-testid="edit-flatten-cancel"
        @click="a.cancelFlatten()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="danger"
        data-testid="edit-flatten-confirm"
        @click="a.confirmFlatten()"
      >
        Flatten and discard original
      </UiButton>
    </template>
  </BaseDialog>

  <BaseDialog
    v-if="e.discardPending.value"
    title="Discard unsaved edits?"
    testid="edit-discard-dialog"
    width="24rem"
    @dismiss="e.cancelDiscard()"
  >
    <p class="message">
      Your edits since the last save will be lost. The stored image is not changed.
    </p>
    <template #actions>
      <UiButton
        data-testid="edit-discard-cancel"
        @click="e.cancelDiscard()"
      >
        Keep editing
      </UiButton>
      <UiButton
        variant="danger"
        data-testid="edit-discard-confirm"
        @click="e.confirmDiscard()"
      >
        Discard
      </UiButton>
    </template>
  </BaseDialog>

  <BaseDialog
    v-if="a.exportOpen.value"
    title="Export image"
    testid="edit-export-dialog"
    width="24rem"
    @dismiss="a.cancelExport()"
  >
    <div class="form">
      <label class="field">
        <span>Format</span>
        <select
          data-testid="edit-export-format"
          :value="a.exportFormat.value"
          @change="chooseFormat"
        >
          <option value="png">PNG</option>
          <option value="jpeg">JPEG</option>
        </select>
      </label>
      <label
        v-if="a.exportFormat.value === 'jpeg'"
        class="field"
      >
        <span>Quality</span>
        <input
          type="range"
          data-testid="edit-export-quality"
          min="1"
          max="100"
          :value="a.exportQuality.value"
          @input="setQuality"
        >
        <input
          type="number"
          class="number"
          data-testid="edit-export-quality-number"
          aria-label="JPEG quality"
          min="1"
          max="100"
          :value="a.exportQuality.value"
          @change="setQuality"
        >
      </label>
      <label class="field">
        <span>Background</span>
        <select
          data-testid="edit-export-background"
          :value="a.exportBackground.value === 'transparent' ? 'transparent' : 'colour'"
          @change="a.exportBackground.value = text($event) === 'transparent' ? 'transparent' : '#ffffff'"
        >
          <option value="transparent">Transparent</option>
          <option value="colour">Colour</option>
        </select>
        <input
          v-if="a.exportBackground.value !== 'transparent'"
          type="color"
          data-testid="edit-export-background-color"
          aria-label="Background colour"
          :value="a.exportBackground.value"
          @change="a.exportBackground.value = text($event)"
        >
      </label>
      <p
        v-if="a.exportBackgroundNotice()"
        class="notice"
        data-testid="edit-export-background-notice"
        role="status"
      >
        {{ a.exportBackgroundNotice() }}
      </p>
      <p
        class="size"
        data-testid="edit-export-size"
      >
        {{ outputSize ? `${outputSize.widthPx} × ${outputSize.heightPx} px, flattened` : "" }}
      </p>
    </div>
    <template #actions>
      <UiButton
        data-testid="edit-export-cancel"
        @click="a.cancelExport()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="edit-export-download"
        :disabled="a.busy.value"
        @click="a.exportImage()"
      >
        Download
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.message {
  margin: 0;
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  color: var(--text-primary);
}

.field {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.field > span {
  min-width: 6rem;
}

.number {
  width: 4.5rem;
  font: inherit;
}

.notice,
.size {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>
