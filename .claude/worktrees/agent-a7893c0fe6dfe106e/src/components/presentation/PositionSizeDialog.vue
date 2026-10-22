<script setup lang="ts">
// The "Position and size…" dialog (SPEC-096): numeric editing of the anchor
// frame's x / y / width / height in the measurementUnit preference, plus
// rotation in whole degrees and an aspect-lock toggle — the keyboard/touch
// path SPEC-068 requires for every drag gesture this spec adds. Each field
// commits on blur or Enter as ONE undo entry (the parent's applyPosition);
// clamping (MIN_FRAME_SIZE_CM / FRAME_MIN_ON_SLIDE_CM) lives in the
// commands. Runtime prop/emit declarations per the espree constraint.
import { ref, watch } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { formatLength, parseLengthToCm, asMeasurementUnit } from "../../lib/core/units";
import { normalizeRotationDeg } from "../../lib/core/shapes";
import { asSlideFrame } from "./slideRender";

const props = defineProps({
  /** The anchor SlideFrame the dialog edits. */
  frame: { type: Object, required: true },
  /** The measurement unit ("cm" | "mm" | "in" | "pt"). */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["apply", "dismiss"]);

const xText = ref("");
const yText = ref("");
const widthText = ref("");
const heightText = ref("");
const rotationText = ref("0");
const aspectLock = ref(false);

function unitOf() {
  return asMeasurementUnit(props.unit);
}

// Seed (and re-seed after each apply) from the live frame.
watch(
  () => props.frame,
  () => {
    const f = asSlideFrame(props.frame);
    const u = unitOf();
    xText.value = formatLength(f.xCm, u);
    yText.value = formatLength(f.yCm, u);
    widthText.value = formatLength(f.widthCm, u);
    heightText.value = formatLength(f.heightCm, u);
    rotationText.value = String(normalizeRotationDeg(f.rotationDeg ?? 0));
  },
  { immediate: true, deep: true },
);

/** The full patch the current field texts describe, in cm/degrees. A field
 * that does not parse falls back to the frame's current value. */
function patch() {
  const f = asSlideFrame(props.frame);
  const u = unitOf();
  return {
    xCm: parseLengthToCm(xText.value, u) ?? f.xCm,
    yCm: parseLengthToCm(yText.value, u) ?? f.yCm,
    widthCm: parseLengthToCm(widthText.value, u) ?? f.widthCm,
    heightCm: parseLengthToCm(heightText.value, u) ?? f.heightCm,
    rotationDeg: normalizeRotationDeg(Number(rotationText.value) || 0),
  };
}

/** One field commit (blur or Enter) = one apply = one undo entry. With the
 * aspect lock on, editing width scales height to keep the ratio (and the
 * reverse), from the frame's current ratio. */
function commitField(field = "") {
  const f = asSlideFrame(props.frame);
  const u = unitOf();
  let p = patch();
  if (aspectLock.value && f.widthCm > 0 && f.heightCm > 0) {
    const ratio = f.widthCm / f.heightCm;
    if (field === "width") {
      p = { ...p, heightCm: p.widthCm / ratio };
    } else if (field === "height") {
      p = { ...p, widthCm: p.heightCm * ratio };
    }
  }
  emit("apply", p);
  // Re-seed from the (clamped) frame after the command ran.
  const g = asSlideFrame(props.frame);
  xText.value = formatLength(g.xCm, u);
  yText.value = formatLength(g.yCm, u);
  widthText.value = formatLength(g.widthCm, u);
  heightText.value = formatLength(g.heightCm, u);
  rotationText.value = String(normalizeRotationDeg(g.rotationDeg ?? 0));
}
</script>

<template>
  <BaseDialog
    title="Position and size"
    testid="slides-position-dialog"
    width="22rem"
    @dismiss="emit('dismiss')"
  >
    <div class="fields">
      <label class="field">
        Position X ({{ unit }})
        <input
          v-model="xText"
          data-testid="position-x"
          autofocus
          @blur="commitField('x')"
          @keydown.enter.prevent="commitField('x')"
        >
      </label>
      <label class="field">
        Position Y ({{ unit }})
        <input
          v-model="yText"
          data-testid="position-y"
          @blur="commitField('y')"
          @keydown.enter.prevent="commitField('y')"
        >
      </label>
      <label class="field">
        Width ({{ unit }})
        <input
          v-model="widthText"
          data-testid="position-width"
          @blur="commitField('width')"
          @keydown.enter.prevent="commitField('width')"
        >
      </label>
      <label class="field">
        Height ({{ unit }})
        <input
          v-model="heightText"
          data-testid="position-height"
          @blur="commitField('height')"
          @keydown.enter.prevent="commitField('height')"
        >
      </label>
      <label class="field">
        Rotation (°)
        <input
          v-model="rotationText"
          data-testid="position-rotation"
          inputmode="numeric"
          @blur="commitField('rotation')"
          @keydown.enter.prevent="commitField('rotation')"
        >
      </label>
      <label class="lock">
        <input
          v-model="aspectLock"
          type="checkbox"
          data-testid="position-aspect-lock"
        >
        Lock aspect ratio
      </label>
    </div>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="position-close"
        @click="emit('dismiss')"
      >
        Close
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.fields {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-3);
}

.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.field input {
  height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-default);
  color: var(--text-primary);
}

.lock {
  grid-column: 1 / -1;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-primary);
}
</style>
