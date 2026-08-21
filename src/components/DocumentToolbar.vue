<script setup lang="ts">
import { nextTick, onMounted, ref, useTemplateRef, watch } from "vue";
import AppIcon from "./AppIcon.vue";
import UiMenuSurface from "./ui/UiMenuSurface.vue";
import UndoRedoControls from "./UndoRedoControls.vue";
import StylePicker from "./StylePicker.vue";
import FontFamilySelect from "./FontFamilySelect.vue";
import FontSizeControl from "./FontSizeControl.vue";
import FormatToolbar from "./FormatToolbar.vue";
import TextColorPicker from "./TextColorPicker.vue";
import HighlightPicker from "./HighlightPicker.vue";
import ParagraphToolbar from "./ParagraphToolbar.vue";
import { useToolbarOverflow } from "../composables/useToolbarOverflow";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// The banded document toolbar (SPEC-061). It groups the strip's controls into
// eight fixed bands separated by subtle rules (AC-3, AC-4), collapses the
// lowest-priority bands whole into a trailing "More formatting" menu when they
// do not fit (AC-12..16), and gives the strip a single roving tab stop
// (AC-31..34). A control that moves into the overflow menu is the very same
// instance, relocated by <Teleport>, not a copy (AC-16). DocumentView owns the
// commands; this component only lays them out and re-emits their events.
//
// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits
// generics — see AppShell.vue's note on the espree constraint.
const props = defineProps({
  /** Px reserved on the trailing edge for open side panes, so opening a pane
   * collapses the strip like narrowing the window (AC-15). */
  reservedRight: { type: Number, default: 0 },
  /** undoRedoState: { canUndo, undoLabel, canRedo, redoLabel }. */
  undoRedo: { type: Object, required: true },
  styleEntries: { type: Array, required: true },
  /** styleState: { id, name }. */
  styleState: { type: Object, required: true },
  /** toolbarFormatValues: fontFamily, fontSizePt, color, colorMixed, … */
  formatValues: { type: Object, required: true },
  documentFonts: { type: Array, default: () => [] },
  recentColors: { type: Array, default: () => [] },
  formatStates: { type: Object, required: true },
  /** paragraphState: { alignment, spaceBefore, direction }. */
  paragraphState: { type: Object, required: true },
  /** listState: { bulleted, numbered, marker }. */
  listState: { type: Object, required: true },
  paragraphPagination: { type: Object, required: true },
});

const emit = defineEmits([
  "undo",
  "redo",
  "apply-style",
  "update-style",
  "reset-style",
  "font-family-select",
  "font-size-change",
  "toggle-character-format",
  "text-color-select",
  "highlight-select",
  "align",
  "indent",
  "spacing-preset",
  "custom-spacing",
  "toggle-space-before",
  "toggle-direction",
  "toggle-list",
  "list-marker",
  "toggle-pagination",
]);

const BAND_ORDER = [1, 2, 3, 4, 5, 6, 7, 8];
// Indexed by band number (index 0 unused) so BAND_NAMES[band] type-checks
// without a TS index signature the .vue espree parser would reject.
const BAND_NAMES = [
  "",
  "History",
  "Style",
  "Type",
  "Character",
  "Colour",
  "Paragraph",
  "Lists",
  "Paragraph extras",
];

const strip = useTemplateRef("strip");
const overflowButton = useTemplateRef("overflowButton");
const overflowMenu = useTemplateRef("overflowMenu");
const textColorPicker = useTemplateRef("textColorPicker");
const highlightPicker = useTemplateRef("highlightPicker");
const menuOpen = ref(false);
const reserved = ref(props.reservedRight);
watch(
  () => props.reservedRight,
  (value) => {
    reserved.value = value;
  },
);

const { collapsedCount, measuring, isCollapsed } = useToolbarOverflow({
  strip: () => strip.value,
  bandOrder: BAND_ORDER,
  reserved,
});

// A collapsible band's controls teleport into the matching section of the
// (kept-mounted) overflow menu only while the band is collapsed; otherwise the
// Teleport is disabled and they render in place in the strip. The target is a
// selector string so no element ref plumbing is needed.
function teleportDisabled(band = 0) {
  return measuring.value || !isCollapsed(band);
}

const { menuStyle } = useFloatingMenu({
  open: () => menuOpen.value,
  anchor: () => overflowButton.value,
  menu: () => overflowMenu.value?.$el ?? null,
  align: "start",
});

watch(collapsedCount, (count) => {
  if (count === 0) {
    menuOpen.value = false;
  }
  void updateRoving();
});

function toggleMenu() {
  menuOpen.value = !menuOpen.value;
}

// The Format menu opens the toolbar's own colour/highlight popups through
// these, rather than duplicating the swatch surfaces (fixed data-testids).
function openTextColorMenu() {
  textColorPicker.value?.openMenu();
}
function openHighlightMenu() {
  highlightPicker.value?.openMenu();
}
defineExpose({ openTextColorMenu, openHighlightMenu });

// ---- Roving focus (AC-31..34) --------------------------------------------
// Returns the strip's focusable controls in DOM order (the overflow button
// last), excluding anything inside a popup surface or a collapsed band. Kept
// as Element[] with per-use `instanceof HTMLElement` guards, per the app's
// espree constraint (no TS casts in .vue scripts).
function stripControls() {
  const el = strip.value;
  if (!el) {
    return [];
  }
  const candidates = Array.from(
    el.querySelectorAll(
      "button:not([disabled]), input:not([disabled]), select:not([disabled]), [role='button']:not([aria-disabled='true'])",
    ),
  );
  return candidates.filter((control) => {
    if (control.closest(".ui-menu-surface") || control.closest("[role='menu']")) {
      return false;
    }
    return (
      control === overflowButton.value ||
      (control instanceof HTMLElement && control.offsetParent !== null)
    );
  });
}

function clearMenuTabindex() {
  const el = overflowMenu.value?.$el;
  if (el instanceof HTMLElement) {
    for (const control of Array.from(el.querySelectorAll("button, input, select"))) {
      control.removeAttribute("tabindex");
    }
  }
}

// Make exactly `active` the tab stop (tabindex 0); every other strip control
// gets -1. Inlined at each call site rather than a helper, so the array param
// avoids the never[] the espree-safe (annotation-free) style would infer.
async function updateRoving() {
  await nextTick();
  const controls = stripControls();
  if (controls.length === 0) {
    return;
  }
  const active = controls[0];
  for (const control of controls) {
    if (control instanceof HTMLElement) {
      control.tabIndex = control === active ? 0 : -1;
    }
  }
  clearMenuTabindex();
}

function onFocusin(event = new FocusEvent("focusin")) {
  const active = event.target;
  if (active instanceof HTMLElement) {
    const controls = stripControls();
    if (controls.includes(active)) {
      for (const control of controls) {
        if (control instanceof HTMLElement) {
          control.tabIndex = control === active ? 0 : -1;
        }
      }
    }
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  const key = event.key;
  if (key !== "ArrowLeft" && key !== "ArrowRight" && key !== "Home" && key !== "End") {
    return;
  }
  const active = document.activeElement;
  if (
    active instanceof HTMLElement &&
    (active.tagName === "INPUT" ||
      active.tagName === "SELECT" ||
      active.isContentEditable ||
      active.closest("[role='menu']"))
  ) {
    return;
  }
  const controls = stripControls();
  if (controls.length === 0) {
    return;
  }
  const index = active ? controls.indexOf(active) : -1;
  let next = index;
  if (key === "Home") {
    next = 0;
  } else if (key === "End") {
    next = controls.length - 1;
  } else if (key === "ArrowRight") {
    next = index < 0 ? 0 : Math.min(controls.length - 1, index + 1);
  } else if (key === "ArrowLeft") {
    next = index < 0 ? controls.length - 1 : Math.max(0, index - 1);
  }
  const target = controls[next];
  if (target instanceof HTMLElement) {
    event.preventDefault();
    for (const control of controls) {
      if (control instanceof HTMLElement) {
        control.tabIndex = control === target ? 0 : -1;
      }
    }
    target.focus();
  }
}

onMounted(() => {
  void updateRoving();
});
</script>

<template>
  <div
    ref="strip"
    class="doc-toolbar-strip"
    data-testid="doc-toolbar-strip"
    @keydown="onKeydown"
    @focusin="onFocusin"
  >
    <!-- Band 1 — History -->
    <div
      v-show="!isCollapsed(1) || measuring"
      class="band"
      data-band="1"
      role="group"
      :aria-label="BAND_NAMES[1]"
    >
      <UndoRedoControls
        :can-undo="undoRedo.canUndo"
        :undo-label="undoRedo.undoLabel"
        :can-redo="undoRedo.canRedo"
        :redo-label="undoRedo.redoLabel"
        @undo="emit('undo')"
        @redo="emit('redo')"
      />
    </div>
    <!-- Band 2 — Style -->
    <div
      v-show="!isCollapsed(2) || measuring"
      class="band"
      data-band="2"
      role="group"
      :aria-label="BAND_NAMES[2]"
    >
      <StylePicker
        :entries="styleEntries"
        :current-id="styleState.id"
        :current-name="styleState.name"
        @apply="emit('apply-style', $event)"
        @update-style="emit('update-style', $event)"
        @reset-style="emit('reset-style', $event)"
      />
    </div>
    <!-- Band 3 — Type -->
    <div
      v-show="!isCollapsed(3) || measuring"
      class="band"
      data-band="3"
      role="group"
      :aria-label="BAND_NAMES[3]"
    >
      <FontFamilySelect
        :value="formatValues.fontFamily"
        :document-fonts="documentFonts"
        @select="emit('font-family-select', $event)"
      />
      <FontSizeControl
        :value="formatValues.fontSizePt ?? undefined"
        @change="emit('font-size-change', $event)"
      />
    </div>
    <!-- Band 4 — Character -->
    <div
      v-show="!isCollapsed(4) || measuring"
      class="band"
      data-band="4"
      role="group"
      :aria-label="BAND_NAMES[4]"
    >
      <FormatToolbar
        :states="formatStates"
        @toggle="emit('toggle-character-format', $event)"
      />
    </div>
    <!-- Band 5 — Colour -->
    <div
      v-show="!isCollapsed(5) || measuring"
      class="band"
      data-band="5"
      role="group"
      :aria-label="BAND_NAMES[5]"
    >
      <Teleport
        to="#overflow-band-5"
        :disabled="teleportDisabled(5)"
      >
        <TextColorPicker
          ref="textColorPicker"
          :value="formatValues.color"
          :mixed="formatValues.colorMixed"
          :recent="recentColors"
          @select="emit('text-color-select', $event)"
        />
        <HighlightPicker
          ref="highlightPicker"
          :value="formatValues.highlight"
          :mixed="formatValues.highlightMixed"
          @select="emit('highlight-select', $event)"
        />
      </Teleport>
    </div>
    <!-- Band 6 — Paragraph (alignment + line spacing) -->
    <div
      v-show="!isCollapsed(6) || measuring"
      class="band"
      data-band="6"
      role="group"
      :aria-label="BAND_NAMES[6]"
    >
      <Teleport
        to="#overflow-band-6"
        :disabled="teleportDisabled(6)"
      >
        <ParagraphToolbar
          section="align"
          :alignment="paragraphState.alignment"
          :space-before="paragraphState.spaceBefore"
          :direction="paragraphState.direction"
          :bulleted="listState.bulleted"
          :numbered="listState.numbered"
          :list-marker="listState.marker"
          :pagination="paragraphPagination"
          @align="emit('align', $event)"
          @indent="emit('indent', $event)"
          @spacing-preset="emit('spacing-preset', $event)"
          @custom-spacing="emit('custom-spacing')"
          @toggle-space-before="emit('toggle-space-before')"
          @toggle-direction="emit('toggle-direction')"
          @toggle-list="emit('toggle-list', $event)"
          @list-marker="emit('list-marker', $event)"
          @toggle-pagination="emit('toggle-pagination', $event)"
        />
      </Teleport>
    </div>
    <!-- Band 7 — Lists (indents + list toggles) -->
    <div
      v-show="!isCollapsed(7) || measuring"
      class="band"
      data-band="7"
      role="group"
      :aria-label="BAND_NAMES[7]"
    >
      <Teleport
        to="#overflow-band-7"
        :disabled="teleportDisabled(7)"
      >
        <ParagraphToolbar
          section="lists"
          :alignment="paragraphState.alignment"
          :space-before="paragraphState.spaceBefore"
          :direction="paragraphState.direction"
          :bulleted="listState.bulleted"
          :numbered="listState.numbered"
          :list-marker="listState.marker"
          :pagination="paragraphPagination"
          @align="emit('align', $event)"
          @indent="emit('indent', $event)"
          @spacing-preset="emit('spacing-preset', $event)"
          @custom-spacing="emit('custom-spacing')"
          @toggle-space-before="emit('toggle-space-before')"
          @toggle-direction="emit('toggle-direction')"
          @toggle-list="emit('toggle-list', $event)"
          @list-marker="emit('list-marker', $event)"
          @toggle-pagination="emit('toggle-pagination', $event)"
        />
      </Teleport>
    </div>
    <!-- Band 8 — Paragraph extras (direction + pagination) -->
    <div
      v-show="!isCollapsed(8) || measuring"
      class="band"
      data-band="8"
      role="group"
      :aria-label="BAND_NAMES[8]"
    >
      <Teleport
        to="#overflow-band-8"
        :disabled="teleportDisabled(8)"
      >
        <ParagraphToolbar
          section="extras"
          :alignment="paragraphState.alignment"
          :space-before="paragraphState.spaceBefore"
          :direction="paragraphState.direction"
          :bulleted="listState.bulleted"
          :numbered="listState.numbered"
          :list-marker="listState.marker"
          :pagination="paragraphPagination"
          @align="emit('align', $event)"
          @indent="emit('indent', $event)"
          @spacing-preset="emit('spacing-preset', $event)"
          @custom-spacing="emit('custom-spacing')"
          @toggle-space-before="emit('toggle-space-before')"
          @toggle-direction="emit('toggle-direction')"
          @toggle-list="emit('toggle-list', $event)"
          @list-marker="emit('list-marker', $event)"
          @toggle-pagination="emit('toggle-pagination', $event)"
        />
      </Teleport>
    </div>
    <button
      v-show="collapsedCount > 0"
      ref="overflowButton"
      type="button"
      class="overflow-button"
      data-testid="toolbar-overflow-button"
      aria-label="More formatting"
      title="More formatting"
      aria-haspopup="menu"
      :aria-expanded="menuOpen ? 'true' : 'false'"
      @mousedown.prevent
      @click="toggleMenu"
    >
      <AppIcon name="more-horizontal" />
    </button>
    <UiMenuSurface
      ref="overflowMenu"
      :open="menuOpen"
      keep-mounted
      :anchor="overflowButton"
      :return-focus="false"
      class="overflow-menu"
      role="menu"
      aria-label="More formatting"
      data-testid="toolbar-overflow-menu"
      :style="menuStyle"
      @dismiss="menuOpen = false"
    >
      <div
        v-for="band in [5, 6, 7, 8]"
        v-show="isCollapsed(band)"
        :key="band"
        class="overflow-section"
        :data-overflow-band="band"
      >
        <p class="overflow-heading">
          {{ BAND_NAMES[band] }}
        </p>
        <div
          :id="`overflow-band-${band}`"
          class="overflow-band-target"
        />
      </div>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
/* The strip is one non-wrapping row that never scrolls horizontally (AC-13):
   it clips its own overflow and collapses bands to stay within its width. */
.doc-toolbar-strip {
  flex: 1 1 auto;
  display: flex;
  align-items: center;
  min-width: 0;
  width: 100%;
  overflow: hidden;
  column-gap: 0;
}

.band {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  flex: none;
  padding-inline: var(--space-3);
}

.band[data-band="1"] {
  padding-inline-start: 0;
}

/* AC-3: a 1px --border-subtle rule with --space-3 on each side between bands.
   The padding-inline on every band supplies the space each side; the rule is
   the leading border on all but the first band. */
.band[data-band]:not([data-band="1"]) {
  border-inline-start: var(--border-width) solid var(--border-subtle);
}

.overflow-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: none;
  width: var(--height-control);
  height: var(--height-control);
  margin-inline-start: var(--space-3);
  padding: 0;
  color: var(--text-secondary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.overflow-button:hover {
  background: var(--state-hover);
  color: var(--text-primary);
}

.overflow-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   strip's clipping). */
.overflow-menu {
  z-index: 40;
  min-width: 14rem;
  max-height: min(28rem, 80vh);
  overflow-y: auto;
}

.overflow-section + .overflow-section {
  margin-top: var(--space-2);
  padding-top: var(--space-2);
  border-top: var(--border-width) solid var(--border-subtle);
}

.overflow-heading {
  margin: 0 0 var(--space-1);
  padding: 0 var(--space-2);
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

/* A collapsed band lays out its controls as a wrapping row inside the menu. */
.overflow-band-target {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: 0 var(--space-2);
}
</style>
