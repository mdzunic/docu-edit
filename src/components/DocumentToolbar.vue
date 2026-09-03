<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import AppIcon from "./AppIcon.vue";
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
import { useTier } from "../composables/useResponsiveTier";

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
  /** The document theme (SPEC-077 AC-20/22): both colour pickers render the
   * theme swatch row when present. */
  theme: { type: Object, default: null },
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

// SPEC-068 AC-9 — at the compact tier the strip reduces to one essential band:
// undo, redo (band 1), the style picker (band 2), bold + italic (the reduced
// band 4) and the overflow button. Everything else — the font Type band, the
// remaining character formats, and the Colour/Paragraph/Lists/extras bands —
// relocates into the overflow menu (and stays reachable in the menu sheet and
// command palette, AC-10). This is a curated collapse layered ABOVE the
// width-driven overflow: those bands are forced into overflow at compact
// regardless of measured width, so the essential band is stable at 390px.
const tier = useTier();
const compact = () => tier.value === "compact";
const COMPACT_OVERFLOW_BANDS = new Set([3, 5, 6, 7, 8]);
// The character formats the reduced band 4 keeps inline at compact; the rest
// teleport into the overflow menu's Character section.
const COMPACT_INLINE_FORMATS = ["bold", "italic"];
const COMPACT_OVERFLOW_FORMATS = ["underline", "strikethrough", "subscript", "superscript"];

/** Whether a band is collapsed into the overflow menu — the union of the
 * width-driven plan and the compact tier's forced set. */
function bandCollapsed(band = 0) {
  return (compact() && COMPACT_OVERFLOW_BANDS.has(band)) || isCollapsed(band);
}

/** The overflow button is present whenever anything is collapsed — which at
 * compact is always the case. */
function overflowVisible() {
  return collapsedCount.value > 0 || measuring.value || compact();
}

// A collapsible band's controls teleport into the matching section of the
// (kept-mounted) overflow menu only while the band is collapsed; otherwise the
// Teleport is disabled and they render in place in the strip. The target is a
// selector string so no element ref plumbing is needed.
function teleportDisabled(band = 0) {
  // Guard on the target actually existing: a disabled Teleport renders inline,
  // so we never enable one whose destination is not yet in the DOM (the
  // insertBefore-of-null crash). Re-evaluated every render via the reactive
  // measuring/collapsedCount reads above.
  if (measuring.value || !bandCollapsed(band)) {
    return true;
  }
  return !document.getElementById(`overflow-band-${band}`);
}

// The reduced band 4's overflow half (SPEC-068): the character formats beyond
// bold/italic teleport into the overflow menu's Character section, but only at
// the compact tier — at every other tier band 4 renders all six inline.
function extrasTeleportDisabled() {
  if (measuring.value || !compact()) {
    return true;
  }
  return !document.getElementById("overflow-band-4x");
}

const { menuStyle } = useFloatingMenu({
  open: () => menuOpen.value,
  anchor: () => overflowButton.value,
  menu: () => overflowMenu.value,
  align: "start",
});

watch(collapsedCount, (count) => {
  if (count === 0) {
    menuOpen.value = false;
  }
  void updateRoving();
});

// The font-size steppers start disabled (no size known at mount) and enable
// once a size resolves. Re-normalize tabindex when that happens so a freshly
// enabled stepper does not keep a default (0) tab stop (AC-32).
watch(
  () => props.formatValues.fontSizePt,
  () => void updateRoving(),
);

function toggleMenu() {
  menuOpen.value = !menuOpen.value;
}

// Dismiss the overflow menu on Escape or an outside pointerdown (the strip's
// own popup dismissal — the overflow menu is a plain fixed div, not a
// UiMenuSurface, so it never becomes a Teleport destination that could crash
// the patch when a band collapses into it).
function onDocPointerDown(event = new PointerEvent("pointerdown")) {
  if (!menuOpen.value) {
    return;
  }
  const target = event.target;
  if (!(target instanceof Node)) {
    return;
  }
  const menu = overflowMenu.value;
  const button = overflowButton.value;
  if ((menu instanceof HTMLElement && menu.contains(target)) || (button && button.contains(target))) {
    return;
  }
  menuOpen.value = false;
}

function onDocKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape" && menuOpen.value) {
    menuOpen.value = false;
  }
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
    // offsetParent is null for a v-show="display:none" control — the hidden
    // overflow button (collapsedCount === 0) and any collapsed band's controls
    // fall out here, so End lands on the last *visible* control and, when the
    // overflow button is shown, on it (AC-33, AC-34).
    return control instanceof HTMLElement && control.offsetParent !== null;
  });
}

function clearMenuTabindex() {
  const el = overflowMenu.value;
  if (el instanceof HTMLElement) {
    for (const control of Array.from(el.querySelectorAll("button, input, select"))) {
      control.removeAttribute("tabindex");
    }
  }
}

// Every focusable-in-principle strip control, whether or not it is currently
// enabled or visible — the set updateRoving() drives to -1 so a control that
// is disabled now (e.g. the font-size steppers before a size is known) does not
// resurface as a second tabindex=0 tab stop the moment it enables (AC-32).
function allStripControls() {
  const el = strip.value;
  if (!el) {
    return [];
  }
  return Array.from(el.querySelectorAll("button, input, select, [role='button']")).filter(
    (control) => !control.closest(".ui-menu-surface") && !control.closest("[role='menu']"),
  );
}

// Make exactly one control the tab stop (tabindex 0) and force every other
// strip control — including disabled/hidden ones — to -1, so the strip is a
// single tab stop (AC-32). The active stop is the currently-focused control if
// focus is already inside the strip; otherwise the first reachable control.
async function updateRoving() {
  await nextTick();
  const reachable = stripControls();
  if (reachable.length === 0) {
    return;
  }
  const focused = document.activeElement;
  const active =
    focused instanceof HTMLElement && reachable.includes(focused) ? focused : reachable[0];
  for (const control of allStripControls()) {
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
  // An open popup (the common-sizes dropdown, a colour swatch grid) owns its own
  // arrow handling — never rove out from inside one.
  if (active instanceof HTMLElement && active.closest("[role='menu']")) {
    return;
  }
  // The font-size field is the one strip control that owns Left/Right for its
  // own caret. Roving still has to pass through it (AC-31/AC-32), so an arrow
  // only leaves the field once the caret sits at the boundary in the travel
  // direction — the APG toolbar text-field convention. Home/End always jump to
  // the strip ends regardless.
  if ((key === "ArrowLeft" || key === "ArrowRight") && active instanceof HTMLInputElement) {
    const start = active.selectionStart ?? 0;
    const end = active.selectionEnd ?? 0;
    const atStart = start === 0 && end === 0;
    const atEnd = start === active.value.length && end === active.value.length;
    if ((key === "ArrowLeft" && !atStart) || (key === "ArrowRight" && !atEnd)) {
      return;
    }
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
    // Entering the font-size field, drop the caret at the edge we arrived from
    // so a further arrow in the same direction ejects in one press rather than
    // first collapsing a browser select-all.
    if (target instanceof HTMLInputElement && (key === "ArrowLeft" || key === "ArrowRight")) {
      const caret = key === "ArrowLeft" ? 0 : target.value.length;
      target.setSelectionRange(caret, caret);
    }
  }
}

onMounted(() => {
  void updateRoving();
  document.addEventListener("pointerdown", onDocPointerDown, true);
  document.addEventListener("keydown", onDocKeydown);
});

onBeforeUnmount(() => {
  document.removeEventListener("pointerdown", onDocPointerDown, true);
  document.removeEventListener("keydown", onDocKeydown);
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
    <!-- The overflow menu is teleported to <body>, OUT of this strip's own
         subtree. A collapsing band then teleports its controls into a target
         (#overflow-band-N) that lives in <body>, never a sibling subtree of
         the same component — teleporting within one component's own root
         crashes Vue's patch (insertBefore of null). The menu is position:fixed
         and hidden via `visibility` (not display:none, which also breaks the
         inbound Teleport) so its targets stay laid out while it is closed. -->
    <Teleport to="body">
      <div
        ref="overflowMenu"
        class="overflow-menu ui-menu-surface"
        :class="{ 'is-open': menuOpen }"
        role="menu"
        aria-label="More formatting"
        data-testid="toolbar-overflow-menu"
        :style="menuStyle"
      >
        <!-- Compact tier only (SPEC-068): the Type band and the character
             formats beyond bold/italic relocate here so the essential band
             stays undo/redo/style/bold/italic. -->
        <div
          v-show="compact() && bandCollapsed(3)"
          class="overflow-section"
          data-overflow-band="3"
        >
          <p class="overflow-heading">
            {{ BAND_NAMES[3] }}
          </p>
          <div
            id="overflow-band-3"
            class="overflow-band-target"
          />
        </div>
        <div
          v-show="compact()"
          class="overflow-section"
          data-overflow-band="4x"
        >
          <p class="overflow-heading">
            {{ BAND_NAMES[4] }}
          </p>
          <div
            id="overflow-band-4x"
            class="overflow-band-target"
          />
        </div>
        <div
          v-for="band in [5, 6, 7, 8]"
          v-show="bandCollapsed(band)"
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
      </div>
    </Teleport>
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
    <!-- Band 3 — Type. Relocates to the overflow menu at the compact tier. -->
    <div
      v-show="!bandCollapsed(3) || measuring"
      class="band"
      data-band="3"
      role="group"
      :aria-label="BAND_NAMES[3]"
    >
      <Teleport
        to="#overflow-band-3"
        :disabled="teleportDisabled(3)"
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
      </Teleport>
    </div>
    <!-- Band 4 — Character. At compact only bold/italic stay in the essential
         band; the rest teleport into the overflow menu's Character section. -->
    <div
      v-show="!isCollapsed(4) || measuring"
      class="band"
      data-band="4"
      role="group"
      :aria-label="BAND_NAMES[4]"
    >
      <FormatToolbar
        :states="formatStates"
        :only="compact() && !measuring ? COMPACT_INLINE_FORMATS : undefined"
        @toggle="emit('toggle-character-format', $event)"
      />
      <Teleport
        to="#overflow-band-4x"
        :disabled="extrasTeleportDisabled()"
      >
        <FormatToolbar
          v-if="compact() && !measuring"
          :states="formatStates"
          :only="COMPACT_OVERFLOW_FORMATS"
          @toggle="emit('toggle-character-format', $event)"
        />
      </Teleport>
    </div>
    <!-- Band 5 — Colour -->
    <div
      v-show="!bandCollapsed(5) || measuring"
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
          :theme="theme"
          @select="emit('text-color-select', $event)"
        />
        <HighlightPicker
          ref="highlightPicker"
          :value="formatValues.highlight"
          :mixed="formatValues.highlightMixed"
          :theme="theme"
          @select="emit('highlight-select', $event)"
        />
      </Teleport>
    </div>
    <!-- Band 6 — Paragraph (alignment + line spacing) -->
    <div
      v-show="!bandCollapsed(6) || measuring"
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
      v-show="!bandCollapsed(7) || measuring"
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
      v-show="!bandCollapsed(8) || measuring"
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
      v-show="overflowVisible()"
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
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  padding: var(--space-2);
  /* Closed: hidden but still laid out (see the template note on Teleport). */
  visibility: hidden;
  pointer-events: none;
}

.overflow-menu.is-open {
  visibility: visible;
  pointer-events: auto;
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
