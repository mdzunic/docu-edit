<script setup lang="ts">
// SPEC-056 equation wrapper. Renders an EquationSpec as selectable DOM text
// (AC-28/29): glyphs from SPEC-026's catalogue in the document's serif
// fallback, non-glyph parts as CSS rules — no math font, no external asset.
// A construct outside the closed set is shown as a read-only box with the
// fixed message (AC-24). The whole equation exposes its alt text to assistive
// technology (AC-21) via role="math" + aria-label. Runtime props + as*
// casts, not TS generics — espree cannot parse TS in a .vue script.
import { computed } from "vue";
import { asEquationSpec } from "../lib/documentModel";
import { layoutRow } from "../lib/math/render";
import { equationAltText } from "../lib/math/linearForm";
import MathBox from "./MathBox.vue";

const props = defineProps({
  /** A documentModel EquationSpec. */
  equation: { type: Object, required: true },
});

const UNSUPPORTED_MESSAGE = "We can't edit this equation yet — it's kept exactly as it was";

const spec = computed(() => asEquationSpec(props.equation));

// An equation is "unsupported" when its whole body is a single opaqueMath
// node — the preservation case shown read-only (AC-24).
const unsupported = computed(
  () => spec.value.body.length === 1 && spec.value.body[0].kind === "opaqueMath",
);

const opaqueLinear = computed(() => {
  const only = spec.value.body[0];
  return only && only.kind === "opaqueMath" ? only.linear : "";
});

const altText = computed(() => equationAltText(spec.value));

const box = computed(() => layoutRow(spec.value.body));
</script>

<template>
  <span
    v-if="unsupported"
    class="equation equation-unsupported"
    :class="{ 'equation-block': spec.display === 'block' }"
    role="math"
    :aria-label="altText || undefined"
    data-testid="equation-unsupported"
    :title="UNSUPPORTED_MESSAGE"
  ><span class="equation-unsupported-linear">{{ opaqueLinear }}</span><span
    class="equation-unsupported-note"
  >{{ UNSUPPORTED_MESSAGE }}</span></span>

  <span
    v-else
    class="equation"
    :class="{ 'equation-block': spec.display === 'block' }"
    role="math"
    :aria-label="altText || undefined"
    data-testid="equation"
  ><MathBox :box="box" /></span>
</template>

<style scoped>
/* An inline equation straddles the line's centre rather than sitting on the
   text baseline: an inline-flex box baselines on its bottom edge, so without
   this a fraction's bar and denominator droop below the run (AC-43). */
.equation {
  display: inline-flex;
  align-items: center;
  vertical-align: middle;
}
.equation-block {
  display: flex;
  justify-content: center;
  width: 100%;
}
.equation-unsupported {
  gap: 0.5em;
  padding: 0.1em 0.4em;
  border: 1px dashed #b0b0b0;
  border-radius: 3px;
  background: #f5f5f5;
  color: #555;
  cursor: not-allowed;
}
.equation-unsupported-linear {
  font-family: "Liberation Serif", serif;
  color: #222;
}
.equation-unsupported-note {
  font-size: 0.75em;
  font-style: italic;
}
</style>
