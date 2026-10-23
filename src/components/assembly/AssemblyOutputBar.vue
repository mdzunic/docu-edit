<script setup lang="ts">
// The assembly output bar (SPEC-109): the format (PNG, or JPEG with its
// quality), the title the saved image takes, the live estimate of the
// output's dimensions and bytes, and the two actions — Save to library and
// Download — plus Start over. When JPEG cannot honour a transparent sheet it
// says so here, before anything is made, and offers the colour it paints
// instead; a plan too large for this browser is refused here, naming the
// limit, with the dpi that would fit one click away. While producing it
// reports "Item n of N" and offers Cancel. All logic is the useAssembly
// controller passed in. Runtime prop declarations (espree lints .vue
// scripts).
import { ref } from "vue";
import UiButton from "../ui/UiButton.vue";
import { useToasts } from "../../composables/useToasts";
import { ASSEMBLY_FORMATS } from "../../lib/assemblyPersist";
import { ASSEMBLY_FORMAT_LABELS } from "../../lib/assembly";

const props = defineProps({
  /** The useAssembly() controller. */
  assembly: { type: Object, required: true },
});

const emit = defineEmits(["saved"]);

const a = props.assembly;
const out = a.out;
const { addToast } = useToasts();
const saved = ref("");

function valueOf(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLSelectElement || el instanceof HTMLInputElement ? el.value : "";
}

function commit(event = new Event("change"), apply = (value = "") => Boolean(value), current = () => "") {
  const el = event.target;
  if (el instanceof HTMLSelectElement || el instanceof HTMLInputElement) {
    apply(el.value);
    el.value = current();
  }
}

async function onSave() {
  saved.value = "";
  const model = await out.save();
  if (model) {
    saved.value = `Saved “${model.title}” to the library.`;
    addToast(`Saved “${model.title}” to the library`);
    emit("saved", model.id);
  }
}

async function onDownload() {
  saved.value = "";
  await out.download();
}

function onReset() {
  saved.value = "";
  a.reset();
}
</script>

<template>
  <section
    class="assembly-output"
    data-testid="assembly-output"
    aria-label="Output"
  >
    <div class="output-settings">
      <label class="output-field">
        <span>Format</span>
        <select
          data-testid="assembly-format"
          :value="a.output.format"
          @change="commit($event, a.setFormat, () => a.output.format)"
        >
          <option
            v-for="format in ASSEMBLY_FORMATS"
            :key="format"
            :value="format"
          >
            {{ ASSEMBLY_FORMAT_LABELS[format] }}
          </option>
        </select>
      </label>
      <label
        v-if="a.output.format === 'jpeg'"
        class="output-field"
      >
        <span>Quality</span>
        <input
          type="number"
          min="1"
          max="100"
          step="1"
          data-testid="assembly-quality"
          :value="String(a.output.quality)"
          @change="commit($event, a.setQuality, () => String(a.output.quality))"
        >
      </label>
      <label class="output-field output-title">
        <span>Title</span>
        <input
          type="text"
          data-testid="assembly-title"
          :value="a.output.title"
          @input="a.setTitle(valueOf($event))"
        >
      </label>
      <p
        class="output-estimate"
        data-testid="assembly-estimate"
        aria-live="polite"
      >
        {{ out.estimate.value }}
      </p>
    </div>

    <div
      v-if="out.backgroundNotice.value"
      class="output-notice"
    >
      <p
        data-testid="assembly-background-notice"
        role="status"
      >
        {{ out.backgroundNotice.value }}
      </p>
      <label class="output-field output-inline">
        <span>Fill colour</span>
        <input
          type="color"
          data-testid="assembly-matte"
          :value="a.output.matte"
          @change="a.setMatte(valueOf($event))"
        >
      </label>
    </div>

    <div
      v-if="out.refusal.value && a.tray.entries.value.length > 0"
      class="output-refusal"
    >
      <p
        data-testid="assembly-refusal"
        role="alert"
      >
        {{ out.refusal.value }}
      </p>
      <UiButton
        v-if="out.fixDpi.value"
        size="sm"
        data-testid="assembly-use-dpi"
        @click="a.sheet.setDpi(out.fixDpi.value)"
      >
        Use {{ out.fixDpi.value }} dpi
      </UiButton>
    </div>

    <div class="output-actions">
      <p
        v-if="out.busy.value"
        class="output-progress"
        data-testid="assembly-progress"
        role="status"
      >
        {{ out.progress.value }}
      </p>
      <UiButton
        v-if="out.busy.value"
        data-testid="assembly-cancel"
        @click="out.cancel()"
      >
        Cancel
      </UiButton>
      <p
        v-if="out.failure.value"
        class="output-failure"
        data-testid="assembly-failure"
        role="alert"
      >
        {{ out.failure.value }}
      </p>
      <p
        v-else-if="saved"
        class="output-saved"
        data-testid="assembly-saved"
        role="status"
      >
        {{ saved }}
      </p>
      <UiButton
        variant="ghost"
        data-testid="assembly-reset"
        :disabled="out.busy.value"
        @click="onReset"
      >
        Start over
      </UiButton>
      <UiButton
        data-testid="assembly-download"
        :disabled="!out.canProduce.value"
        @click="onDownload"
      >
        Download
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="assembly-save"
        :disabled="!out.canProduce.value"
        @click="onSave"
      >
        Save to library
      </UiButton>
    </div>
  </section>
</template>

<style scoped>
.assembly-output {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.output-settings,
.output-actions,
.output-notice,
.output-refusal {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-3);
}

.output-actions {
  justify-content: flex-end;
  align-items: center;
}

.output-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.output-inline {
  flex-direction: row;
  align-items: center;
}

.output-field span {
  color: var(--text-secondary);
}

.output-field select,
.output-field input {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.output-field input[type="number"] {
  width: 5rem;
}

.output-title input {
  width: 14rem;
}

.output-field select:focus-visible,
.output-field input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.output-estimate,
.output-notice p,
.output-refusal p,
.output-progress,
.output-failure,
.output-saved {
  margin: 0;
}

.output-estimate {
  color: var(--text-secondary);
  margin-left: auto;
  align-self: center;
}

.output-notice p {
  color: var(--text-secondary);
  font-weight: var(--weight-medium);
}

.output-refusal p,
.output-failure {
  color: var(--text-danger);
}

.output-progress,
.output-saved {
  color: var(--text-secondary);
  margin-right: auto;
}
</style>
