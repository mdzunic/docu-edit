<script setup lang="ts">
// Dialog field wrapper (SPEC-058): a label above a control, with an optional
// hint or error line below. `for` ties the <label> to the control's id so the
// label is clickable and announced; when an error is present it is wired
// through aria via the caller's aria-describedby (the error line carries a
// stable id derived from `for`). The control itself is the default slot.
const props = defineProps({
  label: { type: String, required: true },
  /** Secondary helper text under the control. */
  hint: { type: String, default: "" },
  /** Error text; when set it replaces the hint and is styled as danger. */
  error: { type: String, default: "" },
  /** id of the control this labels — becomes the <label for>. */
  for: { type: String, default: "" },
});
</script>

<template>
  <div class="ui-field">
    <label
      class="ui-field-label"
      :for="props.for || undefined"
    >
      {{ props.label }}
    </label>
    <slot />
    <p
      v-if="props.error"
      :id="props.for ? `${props.for}-error` : undefined"
      class="ui-field-error"
      role="alert"
    >
      {{ props.error }}
    </p>
    <p
      v-else-if="props.hint"
      :id="props.for ? `${props.for}-hint` : undefined"
      class="ui-field-hint"
    >
      {{ props.hint }}
    </p>
  </div>
</template>

<style scoped>
.ui-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.ui-field-label {
  color: var(--text-secondary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  font-weight: var(--weight-medium);
  line-height: var(--line-tight);
}

.ui-field-hint {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.ui-field-error {
  margin: 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}
</style>
