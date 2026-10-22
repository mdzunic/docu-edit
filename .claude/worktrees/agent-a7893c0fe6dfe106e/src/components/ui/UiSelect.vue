<script setup lang="ts">
import { computed } from "vue";

// Restyled native <select> (SPEC-058). Kept a real <select> so the platform's
// keyboard and screen-reader behaviour is unchanged; only the chrome is
// tokenised. v-model works through modelValue/update:modelValue. Options may
// be plain strings or { value, label } records.
const props = defineProps({
  modelValue: { type: [String, Number], default: "" },
  /** Either string[] or Array<{ value, label }>. */
  options: { type: Array, default: () => [] },
  /** aria-label for the select when no visible <label for> wraps it. */
  label: { type: String, default: "" },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["update:modelValue"]);

// Items arrive untyped (the runtime Array prop). Reflect.get reads value/label
// off object options without a TS cast, which espree bars in .vue scripts.
const normalised = computed(() =>
  props.options.map((o) =>
    o !== null && typeof o === "object"
      ? {
          value: Reflect.get(o, "value"),
          label: String(Reflect.get(o, "label") ?? Reflect.get(o, "value")),
        }
      : { value: o, label: String(o) },
  ),
);

function onChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("update:modelValue", target.value);
  }
}
</script>

<template>
  <select
    class="ui-select"
    :value="props.modelValue"
    :aria-label="props.label || undefined"
    :disabled="props.disabled"
    @change="onChange"
  >
    <option
      v-for="opt in normalised"
      :key="String(opt.value)"
      :value="opt.value"
    >
      {{ opt.label }}
    </option>
  </select>
</template>

<style scoped>
.ui-select {
  height: var(--height-control);
  padding: 0 var(--space-3);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-control);
  line-height: var(--line-tight);
  cursor: pointer;
  transition: border-color var(--motion-fast) var(--easing);
}

.ui-select:not(:disabled):hover {
  border-color: var(--border-strong);
}

.ui-select:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.ui-select:disabled {
  cursor: default;
  opacity: 0.5;
}
</style>
