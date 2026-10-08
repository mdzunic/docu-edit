<script setup lang="ts">
// The presentation formatting toolbar (SPEC-081): font family / size / colour
// reuse the product-neutral document widgets (they already live in
// src/components and read the shared core/fonts + core/colors), and the
// bold/italic/underline/align/bullet/indent controls are toolbar-local buttons.
// Every control reflects the selection's state (mixed → indeterminate) and is
// disabled when nothing is targeted. Alignment/bullet/indent collapse into a ⋯
// overflow menu when the strip is too narrow for them (the e2e-viewport lesson:
// controls must never fall off the edge unreachable).
//
// SPEC-097 adds the numbered toggle (`slides-numbering`, beside
// `slides-bullet`), the list menu (`slides-list-menu` → SlideListPopover) and,
// inside a text session, the link button (`slides-link`, also Cmd/Ctrl+K). The
// link button renders only while a session is open: the insert bar's
// whole-frame `slides-link` renders only outside one, so the test id is never
// on two elements at once (both ribbon strips stay mounted).
//
// Runtime prop/emit declarations, not the generic forms — espree lints .vue
// scripts (see AppShell.vue's note).
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import FontFamilySelect from "../FontFamilySelect.vue";
import FontSizeControl from "../FontSizeControl.vue";
import TextColorPicker from "../TextColorPicker.vue";
import UiToggleButton from "../ui/UiToggleButton.vue";
import SlideListPopover from "./SlideListPopover.vue";
import { alignIsActive, asFormatState, togglePressed } from "./slideToolbarState";

const props = defineProps({
  /** The selection's formatting state (SelectionFormatState) or null when no
   * text frame is targeted. */
  state: { type: Object, default: null },
  /** Whether the formatting controls act on anything (a selected/edited text
   * frame). */
  enabled: { type: Boolean, default: false },
  /** Recently applied text colours, for the colour picker. */
  recentColors: { type: Array, default: () => [] },
  /** The deck theme (SPEC-094): renders the colour picker's theme row. */
  theme: { type: Object, default: null },
  /** SPEC-097: the touched paragraphs' `{ indentCm, hangingCm }` for the list
   * popover ("" when mixed). */
  listGeometry: { type: Object, default: null },
  /** SPEC-097: the last list command's error, shown in the popover. */
  listError: { type: String, default: "" },
  /** SPEC-097: the picture-bullet input's accept list. */
  imageAccept: { type: String, default: "" },
  /** SPEC-097: a text session is open, so the link button shows. */
  linkEnabled: { type: Boolean, default: false },
});

const emit = defineEmits([
  "bold",
  "italic",
  "underline",
  "font",
  "size",
  "color",
  "align",
  "bullet",
  "indent",
  "outdent",
  // SPEC-097
  "numbering",
  "list-style",
  "bullet-char",
  "bullet-image",
  "start-at",
  "list-indent",
  "list-hanging",
  "link",
]);

// SPEC-081 AC-18: the presentation font-size control clamps to 6–96 pt, a
// narrower range than the word processor's 1–409 — passed to the shared
// FontSizeControl so both typed entry and the steppers respect it.
const PRESENTATION_FONT_SIZE_MIN = 6;
const PRESENTATION_FONT_SIZE_MAX = 96;

const s = computed(() => asFormatState(props.state));

// --- control values from the folded selection state ------------------------
const fontValue = computed(() => (typeof s.value?.fontFamily === "string" ? s.value.fontFamily : ""));
const sizeValue = computed(() => (typeof s.value?.sizePt === "number" ? s.value.sizePt : undefined));
const colorValue = computed(() => (typeof s.value?.color === "string" ? s.value.color : ""));
const colorMixed = computed(() => s.value?.color === "mixed");

// togglePressed folds to the aria strings; UiToggleButton takes the tri-state
// as a boolean or the literal "mixed", so normalise here.
function pressed(field = "bold") {
  const value = togglePressed(s.value, field);
  return value === "mixed" ? "mixed" : value === "true";
}
function alignActive(value = "") {
  return alignIsActive(s.value, value);
}

// The shared icon set (src/components/icons/paths.ts) — the same glyphs the
// word processor's toolbar draws, so the two editors read as one product.
// `caption` is the Toolbar Redesign's uppercase word under the icon; the
// deck's strip wears the same grammar as the document's.
const ALIGNMENTS = [
  { value: "left", icon: "align-left", label: "Align left", caption: "Left" },
  { value: "center", icon: "align-center", label: "Align centre", caption: "Centre" },
  { value: "right", icon: "align-right", label: "Align right", caption: "Right" },
  { value: "justify", icon: "align-justify", label: "Justify", caption: "Fill" },
];

// --- the list popover (SPEC-097) -------------------------------------------
const listOpen = ref(false);
const listHost = useTemplateRef("listHost");
/** Closes the popover and hands focus back to its trigger (keyboard users
 * land where they opened it). */
function closeList() {
  listOpen.value = false;
  const trigger = listHost.value?.querySelector('[data-testid="slides-list-menu"]');
  if (trigger instanceof HTMLElement) trigger.focus();
}

// --- overflow (⋯) ----------------------------------------------------------
const root = useTemplateRef("root");
const compact = ref(false);
const overflowOpen = ref(false);
const OVERFLOW_BELOW_PX = 620;

function updateCompact() {
  // A width of 0 means the strip is not being laid out at all — it is the
  // ribbon tab that is not showing (v-show sets display:none), not a narrow
  // window. Treating that as "compact" would activate the overflow Teleport
  // below, whose target lives inside this same subtree, and Vue's patch
  // crashes on the next unmount (insertBefore of null). Hold the last known
  // answer until the strip is measurable again.
  const width = root.value?.clientWidth ?? 9999;
  if (width === 0) {
    return;
  }
  compact.value = width < OVERFLOW_BELOW_PX;
  if (!compact.value) overflowOpen.value = false;
}
const observer = typeof ResizeObserver !== "undefined" ? new ResizeObserver(updateCompact) : null;
onMounted(() => {
  if (observer && root.value) observer.observe(root.value);
});
onBeforeUnmount(() => observer?.disconnect());
</script>

<template>
  <div
    ref="root"
    class="slides-toolbar"
    role="toolbar"
    aria-label="Slide formatting"
    data-testid="slides-toolbar"
  >
    <div
      class="tb-band"
      role="group"
      aria-label="Type"
    >
      <div
        class="tb-group"
        data-testid="slides-font-family"
      >
        <FontFamilySelect
          :value="fontValue"
          :disabled="!enabled"
          @select="(family) => emit('font', family)"
        />
      </div>
      <div
        class="tb-group"
        data-testid="slides-font-size"
      >
        <FontSizeControl
          :value="sizeValue"
          :min="PRESENTATION_FONT_SIZE_MIN"
          :max="PRESENTATION_FONT_SIZE_MAX"
          :disabled="!enabled"
          @change="(pt) => emit('size', pt)"
        />
      </div>
    </div>

    <div
      class="tb-band"
      role="group"
      aria-label="Character"
    >
      <UiToggleButton
        icon="bold"
        label="Bold"
        caption="Bold"
        data-testid="slides-bold"
        :pressed="pressed('bold')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('bold')"
      />
      <UiToggleButton
        icon="italic"
        label="Italic"
        caption="Italic"
        data-testid="slides-italic"
        :pressed="pressed('italic')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('italic')"
      />
      <UiToggleButton
        icon="underline"
        label="Underline"
        caption="Under"
        data-testid="slides-underline"
        :pressed="pressed('underline')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('underline')"
      />
    </div>

    <div
      class="tb-band"
      role="group"
      aria-label="Colour"
    >
      <div
        class="tb-group"
        data-testid="slides-text-color"
      >
        <TextColorPicker
          :value="colorValue"
          :mixed="colorMixed"
          :recent="recentColors"
          :disabled="!enabled"
          :theme="theme"
          @select="(color) => emit('color', color)"
        />
      </div>
    </div>

    <!-- Alignment / bullet / indent: inline, or in the ⋯ menu when compact -->
    <Teleport
      :to="'#slides-overflow-target'"
      :disabled="!compact"
    >
      <div
        class="tb-secondary tb-band"
        role="group"
        aria-label="Paragraph"
      >
        <div
          class="tb-group"
          role="group"
          aria-label="Alignment"
          data-testid="slides-align"
        >
          <UiToggleButton
            v-for="a in ALIGNMENTS"
            :key="a.value"
            :icon="a.icon"
            :label="a.label"
            :caption="a.caption"
            :pressed="alignActive(a.value)"
            :data-align="a.value"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('align', a.value)"
          />
        </div>

        <div class="tb-group">
          <UiToggleButton
            icon="bullet-list"
            label="Bulleted list"
            caption="Bullets"
            data-testid="slides-bullet"
            :pressed="pressed('bullet')"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('bullet')"
          />
          <UiToggleButton
            icon="numbered-list"
            label="Numbered list"
            caption="Numbers"
            data-testid="slides-numbering"
            :pressed="pressed('numbered')"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('numbering')"
          />
          <div
            ref="listHost"
            class="popover-host"
          >
            <UiToggleButton
              icon="multilevel-list"
              label="List style"
              caption="Lists"
              data-testid="slides-list-menu"
              aria-haspopup="dialog"
              :aria-expanded="listOpen"
              :pressed="listOpen"
              :disabled="!enabled"
              @mousedown.prevent
              @click="listOpen = !listOpen"
            />
            <SlideListPopover
              v-if="listOpen && enabled"
              :state="state"
              :geometry="listGeometry"
              :accept="imageAccept"
              :error="listError"
              @style="(style) => emit('list-style', style)"
              @bullet-char="(char) => emit('bullet-char', char)"
              @bullet-image="(file) => emit('bullet-image', file)"
              @start-at="(value) => emit('start-at', value)"
              @indent="(cm) => emit('list-indent', cm)"
              @hanging="(cm) => emit('list-hanging', cm)"
              @close="closeList"
            />
            <div
              v-if="listOpen && enabled"
              class="popover-backdrop"
              @click="listOpen = false"
            />
          </div>
          <UiToggleButton
            icon="indent-decrease"
            label="Decrease indent"
            caption="Outdent"
            data-testid="slides-outdent"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('outdent')"
          />
          <UiToggleButton
            icon="indent-increase"
            label="Increase indent"
            caption="Indent"
            data-testid="slides-indent"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('indent')"
          />
          <UiToggleButton
            v-if="linkEnabled"
            icon="link"
            label="Link (Ctrl+K)"
            caption="Link"
            data-testid="slides-link"
            @mousedown.prevent
            @click="emit('link')"
          />
        </div>
      </div>
    </Teleport>

    <div
      v-show="compact"
      class="tb-group tb-overflow"
    >
      <UiToggleButton
        icon="more-horizontal"
        label="More formatting"
        data-testid="slides-overflow"
        aria-haspopup="menu"
        :aria-expanded="overflowOpen"
        @click="overflowOpen = !overflowOpen"
      />
      <div
        v-show="overflowOpen"
        class="tb-overflow-menu"
        role="menu"
        data-testid="slides-overflow-menu"
      >
        <div id="slides-overflow-target" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.slides-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
}

.tb-group {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.tb-secondary {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

/* Bands mirror the word processor's toolbar (DocumentToolbar): grouped
   controls with a separator rule between bands. The band names stay as
   aria-labels for AT; they are not painted (the menus above name things). */
.tb-band {
  display: flex;
  align-items: center;
  gap: var(--space-1);
  padding-inline: var(--space-3);
}

.tb-band:first-of-type {
  padding-inline-start: 0;
}

.tb-band + .tb-band {
  border-inline-start: var(--border-width) solid var(--border-subtle);
}

/* Inside the overflow menu the teleported band drops its strip chrome. */
.tb-overflow-menu .tb-band {
  padding: 0;
  border: none;
}


.tb-overflow {
  position: relative;
}

.popover-host {
  position: relative;
  display: inline-flex;
}

.popover-backdrop {
  position: fixed;
  inset: 0;
  z-index: 49;
}

.tb-overflow-menu {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-1));
  z-index: 40;
  padding: var(--space-2);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}
</style>
