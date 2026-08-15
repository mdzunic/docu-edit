<script setup lang="ts">
import { computed } from "vue";
import { STANDARD_FONT_FAMILIES } from "../lib/core/fonts";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The selection's font family; "" renders the blank/indeterminate state
   * for a mixed selection (SPEC-013 AC). */
  value: { type: String, default: "" },
  /** Families the open document uses, for the "Fonts in this document"
   * group (SPEC-013 AC). */
  documentFonts: { type: Array, default: () => [] },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);

// The runtime Array prop arrives untyped; String() gives the template the
// string list it needs (espree bars a PropType<string[]> annotation here).
const docFonts = computed(() => props.documentFonts.map(String));

// The value may name a family outside both groups (e.g. an opened file's
// exotic font at the caret before documentFonts refreshes); a fallback
// option keeps the select from silently showing the wrong entry.
const knownValue = computed(
  () =>
    props.value === "" ||
    docFonts.value.includes(props.value) ||
    STANDARD_FONT_FAMILIES.includes(props.value),
);

/** Each list item renders in its own typeface (SPEC-013 AC); the quoted
 * family plus a generic fallback for families not installed locally. */
function optionStyle(family = "") {
  return { fontFamily: `'${family}', sans-serif` };
}

function handleChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement && target.value !== "") {
    emit("select", target.value);
  }
}
</script>

<template>
  <select
    class="font-family-select"
    data-testid="font-family-select"
    title="Font family"
    aria-label="Font family"
    :disabled="disabled"
    :value="value"
    :style="value ? optionStyle(value) : {}"
    @change="handleChange"
  >
    <!-- The blank option gives the mixed state something to select without
         offering "" as a pickable entry. -->
    <option
      value=""
      disabled
      hidden
    />
    <optgroup
      label="Fonts in this document"
      data-testid="font-family-doc-group"
    >
      <option
        v-for="family in docFonts"
        :key="`doc:${family}`"
        :value="family"
        :style="optionStyle(family)"
      >
        {{ family }}
      </option>
    </optgroup>
    <optgroup
      label="Fonts"
      data-testid="font-family-all-group"
    >
      <option
        v-for="family in STANDARD_FONT_FAMILIES"
        :key="family"
        :value="family"
        :style="optionStyle(family)"
      >
        {{ family }}
      </option>
    </optgroup>
    <option
      v-if="!knownValue"
      :value="value"
      :style="optionStyle(value)"
    >
      {{ value }}
    </option>
  </select>
</template>

<style scoped>
.font-family-select {
  font: inherit;
  font-size: 0.875rem;
  height: 1.75rem;
  max-width: 11rem;
  padding: 0 0.25rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.font-family-select:disabled {
  opacity: 0.5;
  cursor: default;
}
</style>
