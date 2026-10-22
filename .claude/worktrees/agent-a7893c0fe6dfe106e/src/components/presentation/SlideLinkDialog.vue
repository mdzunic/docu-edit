<script setup lang="ts">
// The slide link dialog (SPEC-097 AC-46), over BaseDialog: a web address
// (`link-url`) OR a slide of this deck (`link-slide`, a <select> labelled by
// each slide's slideOutline title), an optional tooltip, and Apply / Remove /
// Cancel. Dangerous URLs (sanitizeHref's schemes) are rejected inline before
// Apply, and the controller re-checks through cleanLinkTarget. A dumb
// :controller-taking component like HeaderFooterDialog.vue — every command
// lives in the useSlideTextDepth controller (its useSlideLinks half). Runtime
// prop declarations, no type syntax (espree); typed helpers live in
// slideTextControls.ts.
import { computed, ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { eventString, linkDraftOf, linkTargetOf, linkUrlError } from "./slideTextControls";

const props = defineProps({
  /** The useSlideTextDepth() controller (linkDialog, slideChoices, applyLink,
   * removeLink, closeLinkDialog). */
  controller: { type: Object, required: true },
});

const c = props.controller;

const draft = linkDraftOf(c.linkDialog.value?.initial);
const url = ref(draft.url);
const slide = ref(draft.slide);
const tooltip = ref(draft.tooltip);
const applyError = ref("");

/** Whether the range/frame already carries a link (Remove has work to do). */
const hasLink = computed(() => draft.url !== "" || draft.slide !== "");
const urlError = computed(() => (slide.value === "" ? linkUrlError(url.value) : ""));
const error = computed(() => urlError.value || applyError.value);
const canApply = computed(() => (slide.value !== "" || url.value.trim() !== "") && !urlError.value);

function onUrl(event = new Event("input")) {
  url.value = eventString(event);
  applyError.value = "";
  // A typed address replaces a chosen slide — one target at a time.
  if (url.value.trim() !== "") slide.value = "";
}
function onSlide(event = new Event("change")) {
  slide.value = eventString(event);
  applyError.value = "";
  if (slide.value !== "") url.value = "";
}
function onTooltip(event = new Event("input")) {
  tooltip.value = eventString(event);
}
function apply() {
  if (!canApply.value) return;
  applyError.value = c.applyLink(linkTargetOf(url.value, slide.value, tooltip.value));
}
</script>

<template>
  <BaseDialog
    title="Link"
    testid="slide-link-dialog"
    width="24rem"
    @dismiss="c.closeLinkDialog()"
  >
    <form
      class="link-form"
      @submit.prevent="apply"
    >
      <label class="field">
        <span class="field-label">Web address</span>
        <input
          type="text"
          inputmode="url"
          class="text-input"
          placeholder="https://"
          autofocus
          data-testid="link-url"
          :value="url"
          :aria-invalid="urlError !== ''"
          @input="onUrl"
        >
      </label>
      <label class="field">
        <span class="field-label">Or a slide in this deck</span>
        <select
          class="text-input"
          data-testid="link-slide"
          :value="slide"
          @change="onSlide"
        >
          <option value="">
            No slide
          </option>
          <option
            v-for="choice in c.slideChoices.value"
            :key="choice.number"
            :value="String(choice.number)"
          >
            {{ choice.label }}
          </option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Tooltip</span>
        <input
          type="text"
          class="text-input"
          data-testid="link-tooltip"
          :value="tooltip"
          @input="onTooltip"
        >
      </label>
      <p
        v-if="error"
        class="error"
        role="alert"
        data-testid="link-error"
      >
        {{ error }}
      </p>
    </form>
    <template #actions>
      <UiButton
        data-testid="link-remove"
        :disabled="!hasLink"
        @click="c.removeLink()"
      >
        Remove link
      </UiButton>
      <UiButton
        data-testid="link-cancel"
        @click="c.closeLinkDialog()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="link-apply"
        :disabled="!canApply"
        @click="apply"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.link-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.field-label {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.text-input {
  height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-default);
  color: var(--text-primary);
  font: inherit;
}

.error {
  margin: 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
}
</style>
