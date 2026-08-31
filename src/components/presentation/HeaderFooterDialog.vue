<script setup lang="ts">
// The "Header and footer" dialog (SPEC-093): a BaseDialog-shelled form that is a
// dumb :controller-taking component like TransitionsPanel.vue / DeckSetupDialog
// .vue — every value and command lives in the `useHeaderFooter` controller
// passed in, this component is presentation + event forwarding only. Runtime
// prop declarations and untyped helpers, not the generic forms (espree lints
// .vue scripts); typed transforms live in headerFooterFields.ts.
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { asDateFormatOptions, eventChecked, eventValue } from "../../views/presentation/headerFooterFields";

const props = defineProps({
  /** The useHeaderFooter() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  /** The 1-based number of the slide the dialog opened on, for context. */
  slideNumber: { type: Number, default: 1 },
});

const c = props.controller;

function setSource(value = "auto") {
  c.dateSource.value = value;
}
</script>

<template>
  <BaseDialog
    title="Header and footer"
    testid="header-footer-dialog"
    width="26rem"
    @dismiss="c.cancel()"
  >
    <div class="hf-form">
      <!-- Footer -->
      <label class="row">
        <input
          type="checkbox"
          data-testid="footer-show"
          :checked="c.footer.value"
          @change="c.footer.value = eventChecked($event)"
        >
        <span>Footer</span>
      </label>
      <input
        type="text"
        class="text-input"
        data-testid="footer-text"
        aria-label="Footer text"
        :value="c.footerText.value"
        :disabled="!c.footer.value"
        @input="c.footerText.value = eventValue($event)"
      >

      <!-- Date and time -->
      <label class="row">
        <input
          type="checkbox"
          data-testid="footer-date-show"
          :checked="c.date.value"
          @change="c.date.value = eventChecked($event)"
        >
        <span>Date and time</span>
      </label>
      <fieldset
        class="date-group"
        :disabled="!c.date.value"
      >
        <label class="radio">
          <input
            type="radio"
            name="hf-date-source"
            data-testid="footer-date-auto"
            :checked="c.dateSource.value === 'auto'"
            @change="setSource('auto')"
          >
          <span>Update automatically</span>
        </label>
        <select
          class="text-input"
          data-testid="footer-date-format"
          aria-label="Date format"
          :value="c.dateFormat.value"
          :disabled="c.dateSource.value !== 'auto' || c.datePreserved.value"
          @change="c.dateFormat.value = eventValue($event)"
        >
          <option
            v-for="option in asDateFormatOptions(c.dateFormatOptions.value)"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
        <p
          v-if="c.datePreserved.value"
          class="note"
          data-testid="footer-date-format-preserved"
        >
          This slide's date uses a format kept from the original file, so it is
          preserved as-is.
        </p>
        <label class="radio">
          <input
            type="radio"
            name="hf-date-source"
            data-testid="footer-date-fixed"
            :checked="c.dateSource.value === 'fixed'"
            @change="setSource('fixed')"
          >
          <span>Fixed</span>
        </label>
        <input
          type="text"
          class="text-input"
          data-testid="footer-date-fixed-text"
          aria-label="Fixed date text"
          :value="c.dateFixedText.value"
          :disabled="c.dateSource.value !== 'fixed'"
          @input="c.dateFixedText.value = eventValue($event)"
        >
      </fieldset>

      <!-- Slide number -->
      <label class="row">
        <input
          type="checkbox"
          data-testid="footer-slide-number"
          :checked="c.slideNumber.value"
          @change="c.slideNumber.value = eventChecked($event)"
        >
        <span>Slide number (this slide: {{ slideNumber }})</span>
      </label>

      <!-- Title-slide policy -->
      <label class="row">
        <input
          type="checkbox"
          data-testid="footer-hide-on-title"
          :checked="!c.showOnTitleSlide.value"
          @change="c.showOnTitleSlide.value = !eventChecked($event)"
        >
        <span>Don't show on title slide</span>
      </label>
    </div>

    <template #actions>
      <UiButton
        data-testid="footer-cancel"
        @click="c.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="footer-apply-all"
        @click="c.applyToAll()"
      >
        Apply to all
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="footer-apply"
        @click="c.apply()"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.hf-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.text-input {
  width: 100%;
  box-sizing: border-box;
}

.date-group {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  border: var(--border-width) solid var(--border);
  border-radius: var(--radius-1);
  padding: var(--space-3);
  margin: 0;
}

.date-group[disabled] {
  opacity: 0.55;
}

.radio {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.note {
  margin: 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}
</style>
