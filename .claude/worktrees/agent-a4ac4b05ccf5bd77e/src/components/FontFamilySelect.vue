<script setup lang="ts">
import { computed } from "vue";
import { STANDARD_FONT_FAMILIES } from "../lib/core/fonts";
import UiControlCaption from "./ui/UiControlCaption.vue";

// A native <select>, deliberately NOT rebuilt on UiMenuSurface (SPEC-058
// AC-39): the primitive exists so no popup hand-rolls its own dismissal, and
// a native select hand-rolls nothing — the browser owns its popup, its
// outside-click and its Escape. Rebuilding it as a custom dropdown would
// also break the SPEC-013 e2e suite, which drives this control with
// selectOption() — and AC-61 freezes those suites. (UiSelect is no home for
// it either: this control needs optgroups and per-option font styling, which
// UiSelect's flat options list cannot express.)
//
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
  <span class="font-field">
    <UiControlCaption
      class="kicker"
      text="Font"
    />
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
  </span>
</template>

<style scoped>
/* The handoff's field: one hairline box carrying an uppercase accent kicker
   (FONT) and the value. The native <select> keeps its own semantics and
   testid — it simply loses its border to the wrapper, so the pair reads as a
   single instrument rather than a label glued to a control. */
.font-field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  height: var(--control-height, var(--height-control));
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.font-field:has(.font-family-select:disabled) {
  opacity: 0.5;
}

.kicker {
  color: var(--text-accent);
  flex: none;
}

.font-family-select {
  font: inherit;
  font-size: var(--text-md);
  height: 100%;
  /* Fixed, not content-sized: intrinsic select width follows the widest
     option's text metrics, which differ per platform (macOS renders wider
     than the CI container) and per the selected family (the closed control
     renders in it) — both made the strip's fit at the 1280px budget
     (SPEC-049 AC-4 / SPEC-058 AC-48) platform- and caret-dependent. */
  width: 9rem;
  padding: 0;
  border: 0;
  background: transparent;
  cursor: pointer;
}

.font-family-select:disabled {
  cursor: default;
}
</style>
