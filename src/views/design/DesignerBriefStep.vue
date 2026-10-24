<script setup lang="ts">
// SPEC-115: the deck designer's Brief step — the prompt (designer-brief), the
// audience/length/tone selects, the optional source (designer-source: nothing,
// pasted text, or a library document or deck chosen by id), and a live count
// of what a request would carry. Redesign has no length and no source — the
// source is the open deck. Nothing here sends anything: the request goes, if
// at all, only after SPEC-113's consent dialog has shown its exact body.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed, useId } from "vue";
import UiField from "../../components/ui/UiField.vue";
import UiSelect from "../../components/ui/UiSelect.vue";
import { asDeckDesigner, fieldChecked, fieldValue } from "./designerFields";

const props = defineProps({
  /** The useDeckDesigner() session. */
  controller: { type: Object, required: true },
});

const c = asDeckDesigner(props.controller);
const id = useId();
const isNew = computed(() => c.mode === "new");

function setBrief(event = new Event("input")) {
  c.brief.value = fieldValue(event);
}
function setAudience(value = "") {
  c.audience.value = value;
}
function setLength(value = "") {
  c.length.value = value;
}
function setTone(value = "") {
  c.tone.value = value;
}
function setSource(value = "") {
  c.sourceChoice.value = value;
}
function setPasted(event = new Event("input")) {
  c.pastedText.value = fieldValue(event);
}
function setKeepPrompt(event = new Event("change")) {
  c.keepPrompt.value = fieldChecked(event);
}
</script>

<template>
  <div class="brief">
    <UiField
      :label="isNew ? 'Describe the deck you want' : 'How should this deck change?'"
      :for="`${id}-brief`"
      :hint="isNew ? 'For example: a 10-slide update for the board on the Rotterdam migration, one number per slide, restrained.' : 'For example: tell the story in the order a customer meets it, with plainer layouts.'"
    >
      <textarea
        :id="`${id}-brief`"
        class="brief-text"
        data-testid="designer-brief"
        rows="4"
        :value="c.brief.value"
        @input="setBrief"
      />
    </UiField>

    <div class="brief-choices">
      <UiField
        label="Audience"
        :for="`${id}-audience`"
      >
        <UiSelect
          :id="`${id}-audience`"
          data-testid="designer-audience"
          :model-value="c.audience.value"
          :options="c.audiences"
          @update:model-value="setAudience"
        />
      </UiField>
      <UiField
        v-if="isNew"
        label="Length"
        :for="`${id}-length`"
      >
        <UiSelect
          :id="`${id}-length`"
          data-testid="designer-length"
          :model-value="c.length.value"
          :options="c.lengths"
          @update:model-value="setLength"
        />
      </UiField>
      <UiField
        label="Tone"
        :for="`${id}-tone`"
      >
        <UiSelect
          :id="`${id}-tone`"
          data-testid="designer-tone"
          :model-value="c.tone.value"
          :options="c.tones"
          @update:model-value="setTone"
        />
      </UiField>
    </div>

    <template v-if="isNew">
      <UiField
        label="Source"
        :for="`${id}-source`"
        hint="Only the source you pick here is sent, and only its title names it."
      >
        <UiSelect
          :id="`${id}-source`"
          data-testid="designer-source"
          :model-value="c.sourceChoice.value"
          :options="c.sourceOptions"
          @update:model-value="setSource"
        />
      </UiField>
      <textarea
        v-if="c.sourceChoice.value === 'paste'"
        class="brief-text"
        data-testid="designer-source-text"
        aria-label="Pasted source text"
        rows="5"
        :value="c.pastedText.value"
        @input="setPasted"
      />
      <p
        v-if="c.truncation.value"
        class="brief-note"
        data-testid="designer-truncated"
      >
        {{ c.truncation.value }}
      </p>
      <label class="brief-keep">
        <input
          type="checkbox"
          data-testid="designer-keep-prompt"
          :checked="c.keepPrompt.value"
          @change="setKeepPrompt"
        >
        <span>Keep the prompt with the deck — it is stored in the file's GeneratedPrompt property and will travel with the
          file wherever it goes.</span>
      </label>
    </template>

    <p
      class="brief-bytes"
      data-testid="designer-bytes"
      aria-live="polite"
    >
      What will be sent: {{ c.bytes.value.toLocaleString("en-US") }} bytes — the design instructions and
      {{ isNew ? (c.source.value ? "your brief and the source text" : "your brief") : "your brief and each slide's number, title and layout (no body text, no notes)" }}.
      You will see the exact request before anything is sent.
    </p>
    <p
      v-if="c.providerOff.value"
      class="brief-note"
      data-testid="designer-provider-off"
    >
      No model provider is set up (Preferences › Labs › Model provider), so nothing can be sent.
      {{ isNew ? "“Outline this document” builds a plan from the source without a model." : "“Rearrange by hand” opens the plan with this deck as it is." }}
    </p>
  </div>
</template>

<style scoped>
.brief {
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.brief-text {
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-control);
  line-height: var(--line-normal);
  resize: vertical;
}

.brief-choices {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-4);
}

.brief-keep {
  display: flex;
  align-items: flex-start;
  gap: var(--space-2);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.brief-bytes,
.brief-note {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}
</style>
