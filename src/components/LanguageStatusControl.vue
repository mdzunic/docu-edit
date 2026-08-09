<script setup lang="ts">
// The status bar's language control (SPEC-035 AC-15, AC-15.1): shows the
// current language, and opens a popover to change it — self-contained
// (button + popover in one component) the same way HighlightPicker.vue is,
// anchored with the same useFloatingMenu escape-the-clipped-strip pattern.
import { computed, onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { asLanguageOptions } from "../lib/language";

const props = defineProps({
  currentLanguageTag: { type: String, default: "" },
  currentLanguageLabel: { type: String, default: "" },
  languageOptions: { type: Array, default: () => [] },
  /** Whether a real (non-collapsed) text selection is active — "Do not
   * check spelling" only applies to a selection (SPEC-035 AC-11). */
  selectionActive: { type: Boolean, default: false },
  selectionNoProof: { type: Boolean, default: false },
});

const emit = defineEmits(["select", "manage-dictionary", "toggle-no-proof"]);

const options = asLanguageOptions(props.languageOptions);

const open = ref(false);
const root = useTemplateRef("root");
const triggerEl = useTemplateRef("triggerEl");
const menuEl = useTemplateRef("menuEl");

// The status bar sits at the bottom of the viewport, so the popover opens
// upward (placement: "above") — "below" (the default, right for toolbar
// buttons near the top) would open off-screen here.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuEl.value,
  align: "center",
  placement: "above",
});

function toggle() {
  open.value = !open.value;
}

function pick(tag = "") {
  emit("select", tag);
  open.value = false;
}

const noProofChecked = computed({
  get: () => props.selectionNoProof,
  set: (value) => emit("toggle-no-proof", value),
});

function manageDictionary() {
  open.value = false;
  emit("manage-dictionary");
}

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (open.value && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    open.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    open.value = false;
  }
}

document.addEventListener("mousedown", handleDocumentMousedown);
onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleDocumentMousedown);
});
</script>

<template>
  <div
    ref="root"
    class="language-control"
    @keydown="handleKeydown"
  >
    <button
      ref="triggerEl"
      type="button"
      class="language-button"
      data-testid="language-status-button"
      aria-haspopup="true"
      :aria-expanded="open ? 'true' : 'false'"
      :aria-label="`Language: ${currentLanguageLabel}`"
      @click="toggle"
    >
      {{ currentLanguageLabel }}
    </button>
    <div
      v-if="open"
      ref="menuEl"
      class="language-menu"
      data-testid="language-menu"
      role="menu"
      aria-label="Language"
      :style="menuStyle"
    >
      <button
        v-for="option in options"
        :key="option.tag"
        type="button"
        role="menuitemradio"
        class="language-option"
        :class="{ active: option.tag === currentLanguageTag }"
        :aria-checked="option.tag === currentLanguageTag ? 'true' : 'false'"
        :data-testid="`language-option-${option.tag}`"
        @click="pick(option.tag)"
      >
        {{ option.label }}
      </button>
      <template v-if="selectionActive">
        <div
          class="language-menu-separator"
          role="separator"
        />
        <label class="no-proof-option">
          <input
            v-model="noProofChecked"
            type="checkbox"
            data-testid="language-no-proof"
          >
          Do not check spelling
        </label>
      </template>
      <div
        class="language-menu-separator"
        role="separator"
      />
      <button
        type="button"
        role="menuitem"
        class="language-option"
        data-testid="language-manage-dictionary"
        @click="manageDictionary"
      >
        Custom dictionary…
      </button>
    </div>
  </div>
</template>

<style scoped>
.language-control {
  position: relative;
  display: inline-flex;
}

.language-button {
  font: inherit;
  font-size: inherit;
  padding: 0;
  border: none;
  background: transparent;
  color: inherit;
  cursor: pointer;
}

.language-button:hover,
.language-button:focus-visible {
  text-decoration: underline;
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   status bar's own clipping the same way HighlightPicker's does). */
.language-menu {
  z-index: 30;
  min-width: 10rem;
  padding: 0.25rem;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.language-option {
  font: inherit;
  font-size: 0.8125rem;
  display: block;
  width: 100%;
  padding: 0.375rem 0.625rem;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  text-align: left;
}

.language-option:hover {
  background: #f1f3f4;
}

.language-option.active {
  background: #d2e3fc;
}

.language-menu-separator {
  height: 1px;
  margin: 0.25rem 0;
  background: #e0e0e0;
}

.no-proof-option {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.375rem 0.625rem;
  font-size: 0.8125rem;
  color: #202124;
  cursor: pointer;
}
</style>
