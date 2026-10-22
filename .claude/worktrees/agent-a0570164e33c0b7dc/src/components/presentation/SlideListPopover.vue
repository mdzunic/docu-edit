<script setup lang="ts">
// The slide list popover (SPEC-097 AC-25..29), opened from the toolbar's
// `slides-list-menu`: the eight core/lists.ts marker styles as
// `list-style-{style}` buttons previewed through markerText(style, 1),
// "Custom bullet…" (a grid of glyphs from core/specialCharacters.ts),
// "Picture bullet…" (a hidden image input — PRESENTATION_IMAGE_MIME_TYPES) and
// the start-at / indent / hanging number inputs. A slim rendering of the same
// vocabulary DefineListDialog.vue offers the word processor (that dialog is
// bound to the text model). Every control is a native button/input, so the
// whole popover is keyboard-operable; Escape closes it and the toolbar returns
// focus to the trigger. SPEC-081's overflow rule: on open the popover clamps
// its height to the viewport below it (scrolling inside) and shifts left when
// it would run off the right edge, so nothing falls out of reach at the 720px
// e2e viewport.
//
// Runtime prop/emit declarations and untyped helpers — espree lints .vue
// scripts; the typed vocabulary lives in slideTextControls.ts.
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { asFormatState } from "./slideToolbarState";
import { BULLET_CHAR_CHOICES, LIST_STYLE_CHOICES, eventNumber, inputValue } from "./slideTextControls";

const props = defineProps({
  /** The selection's SelectionFormatState (marks the current style). */
  state: { type: Object, default: null },
  /** `{ indentCm, hangingCm }` of the touched paragraphs ("" when mixed). */
  geometry: { type: Object, default: null },
  /** The picture-bullet input's accept list. */
  accept: { type: String, default: "" },
  /** The last list command's error message ("" for none). */
  error: { type: String, default: "" },
});

const emit = defineEmits(["style", "bullet-char", "bullet-image", "start-at", "indent", "hanging", "close"]);

const s = computed(() => asFormatState(props.state));
const currentStyle = computed(() => s.value?.bulletStyle ?? "");
const startAtValue = computed(() => inputValue(s.value?.startAt));
const indentValue = computed(() => inputValue(props.geometry?.indentCm));
const hangingValue = computed(() => inputValue(props.geometry?.hangingCm));

const charsOpen = ref(false);
const styles = LIST_STYLE_CHOICES;
const chars = BULLET_CHAR_CHOICES;

const root = useTemplateRef("root");
const fileInput = useTemplateRef("fileInput");

// SPEC-081's overflow rule: clamp to the viewport below the trigger.
const VIEWPORT_MARGIN_PX = 8;
const clampStyle = ref({});
onMounted(async () => {
  await nextTick();
  const el = root.value;
  if (!el) return;
  const rect = el.getBoundingClientRect();
  const maxHeight = Math.max(160, window.innerHeight - rect.top - VIEWPORT_MARGIN_PX);
  const overflowRight = rect.right - (window.innerWidth - VIEWPORT_MARGIN_PX);
  clampStyle.value = {
    maxHeight: `${maxHeight}px`,
    ...(overflowRight > 0 ? { transform: `translateX(-${Math.ceil(overflowRight)}px)` } : {}),
  };
  const first = el.querySelector("button");
  if (first instanceof HTMLElement) first.focus();
});

function onKeydown(event = new KeyboardEvent("keydown")) {
  // The stage's window keys (nudge, Escape-deselect) must not also act on
  // keys typed into the popover's controls.
  event.stopPropagation();
  if (event.key === "Escape") {
    event.preventDefault();
    emit("close");
  }
}

function pickStyle(style = "disc") {
  emit("style", style);
  emit("close");
}
function pickChar(char = "") {
  emit("bullet-char", char);
  emit("close");
}
function openPicture() {
  fileInput.value?.click();
}
function onPicture(event = new Event("change")) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement)) return;
  const file = input.files?.[0];
  input.value = "";
  if (file) {
    emit("bullet-image", file);
    emit("close");
  }
}
function onNumber(name = "start-at", event = new Event("change")) {
  const value = eventNumber(event);
  if (!Number.isFinite(value)) return;
  if (name === "start-at") emit("start-at", value);
  else if (name === "indent") emit("indent", value);
  else emit("hanging", value);
}
</script>

<template>
  <div
    ref="root"
    class="list-popover"
    role="dialog"
    aria-label="List style"
    data-testid="list-style-popover"
    :style="clampStyle"
    @keydown="onKeydown"
  >
    <div
      class="style-grid"
      role="group"
      aria-label="Marker style"
    >
      <button
        v-for="choice in styles"
        :key="choice.style"
        type="button"
        class="style-button"
        :class="{ 'is-current': currentStyle === choice.style }"
        :aria-pressed="currentStyle === choice.style"
        :aria-label="choice.label"
        :title="choice.label"
        :data-testid="`list-style-${choice.style}`"
        @mousedown.prevent
        @click="pickStyle(choice.style)"
      >
        <span
          class="style-preview"
          aria-hidden="true"
        >{{ choice.preview }}</span>
      </button>
    </div>

    <button
      type="button"
      class="menu-item"
      data-testid="list-bullet-char"
      :aria-expanded="charsOpen"
      @mousedown.prevent
      @click="charsOpen = !charsOpen"
    >
      Custom bullet…
    </button>
    <div
      v-if="charsOpen"
      class="char-grid"
      role="group"
      aria-label="Bullet character"
      data-testid="list-bullet-char-grid"
    >
      <button
        v-for="choice in chars"
        :key="choice.char"
        type="button"
        class="char-button"
        :aria-label="choice.name"
        :title="choice.name"
        :data-char="choice.char"
        data-testid="list-bullet-char-option"
        @mousedown.prevent
        @click="pickChar(choice.char)"
      >
        {{ choice.char }}
      </button>
    </div>

    <button
      type="button"
      class="menu-item"
      data-testid="list-bullet-image"
      @mousedown.prevent
      @click="openPicture"
    >
      Picture bullet…
    </button>
    <input
      ref="fileInput"
      type="file"
      tabindex="-1"
      aria-hidden="true"
      class="sr-only"
      :accept="accept"
      data-testid="list-bullet-image-input"
      @change="onPicture"
    >

    <p
      v-if="error"
      class="list-error"
      role="alert"
      data-testid="list-error"
    >
      {{ error }}
    </p>

    <div class="list-fields">
      <label class="list-field">
        <span>Start at</span>
        <input
          type="number"
          min="1"
          max="32767"
          step="1"
          data-testid="list-start-at"
          :value="startAtValue"
          @change="onNumber('start-at', $event)"
        >
      </label>
      <label class="list-field">
        <span>Indent (cm)</span>
        <input
          type="number"
          min="0"
          max="50"
          step="0.05"
          data-testid="list-indent"
          :value="indentValue"
          @change="onNumber('indent', $event)"
        >
      </label>
      <label class="list-field">
        <span>Hanging (cm)</span>
        <input
          type="number"
          min="0"
          max="50"
          step="0.05"
          data-testid="list-hanging"
          :value="hangingValue"
          @change="onNumber('hanging', $event)"
        >
      </label>
    </div>
  </div>
</template>

<style scoped>
.list-popover {
  position: absolute;
  left: 0;
  top: calc(100% + var(--space-1));
  z-index: 50;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  width: 15rem;
  overflow-y: auto;
  padding: var(--space-2);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.style-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: var(--space-1);
}

.style-button,
.char-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: var(--height-control);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-default);
  color: var(--text-primary);
  cursor: pointer;
}

.style-button.is-current {
  background: var(--surface-selected);
  border-color: var(--accent);
  color: var(--text-accent);
}

.style-button:hover,
.char-button:hover,
.menu-item:hover {
  background: var(--state-hover);
}

.style-preview {
  font-size: var(--text-md);
}

.char-grid {
  display: grid;
  grid-template-columns: repeat(8, 1fr);
  gap: var(--space-1);
  max-height: 9rem;
  overflow-y: auto;
}

.char-button {
  height: 1.75rem;
  padding: 0;
}

.menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.style-button:focus-visible,
.char-button:focus-visible,
.menu-item:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.list-error {
  margin: 0;
  color: var(--text-danger, var(--text-secondary));
}

.list-fields {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: var(--space-2);
}

.list-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-secondary);
}

.list-field input {
  width: 100%;
  min-width: 0;
  height: 1.9rem;
  padding: 0 var(--space-1);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-default);
  color: var(--text-primary);
}
</style>
