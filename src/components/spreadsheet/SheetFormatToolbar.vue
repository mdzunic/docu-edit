<script setup lang="ts">
// The sheet formatting strip (SPEC-123 "The toolbar"): eight fixed bands with
// SPEC-061's chrome — AppIcon glyphs, tooltips naming the SHEET_SHORTCUTS
// binding (shortcutLabel), ONE roving tab stop (arrows, Home and End move
// it) and 1px rules between bands — collapsing bands 8 → 5 into a trailing
// "More formatting" menu through the unchanged useToolbarOverflow /
// planCollapse (at a 1440px viewport nothing collapses). A collapsed band's
// controls are the same instances, teleported (DocumentToolbar's pattern).
//   1 History — the format painter (undo/redo stay in the app bar)
//   2 Type — FontFamilySelect, FontSizeControl (1–409 pt)
//   3 Character — bold, italic, underline (single), strikethrough
//   4 Colour — TextColorPicker (the workbook theme), SheetFillPicker,
//     SheetBordersMenu
//   5 Number — the number-format dropdown, currency, percent, comma, decimals
//   6 Alignment — left/centre/right, top/middle/bottom, wrap, SPEC-124's
//     merge slot
//   7 Indent & orientation — indent −/+, rotation
//   8 Styles — the cell styles gallery, Clear
// Everything reads `formatting.toolbar` (the ACTIVE cell's effective style)
// and runs `formatting.run`; the painter is useFormatPainter's. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import { useFloatingMenu } from "../../composables/useFloatingMenu";
import { useToolbarOverflow } from "../../composables/useToolbarOverflow";
import { isMacPlatform } from "../../lib/core/platform";
import { shortcutLabel } from "../../lib/shortcuts";
import { SHEET_SHORTCUTS } from "../../views/spreadsheet/sheetShortcuts";
import { asFormatPainter } from "../../views/spreadsheet/useFormatPainter";
import { asFormatting } from "../../views/spreadsheet/useFormatting";
import AppIcon from "../AppIcon.vue";
import FontFamilySelect from "../FontFamilySelect.vue";
import FontSizeControl from "../FontSizeControl.vue";
import TextColorPicker from "../TextColorPicker.vue";
import UiToggleButton from "../ui/UiToggleButton.vue";
import CellStylesGallery from "./CellStylesGallery.vue";
import SheetBordersMenu from "./SheetBordersMenu.vue";
import SheetFillPicker from "./SheetFillPicker.vue";
import SheetMenuButton from "./SheetMenuButton.vue";
import SheetMergeButton from "./SheetMergeButton.vue";

const props = defineProps({
  /** useFormatting's return. */
  formatting: { type: Object, required: true },
  /** useFormatPainter's return. */
  painter: { type: Object, required: true },
  /** SPEC-124: useMerges' return — the band-6 merge split button; absent:
   * the slot stays empty. */
  merges: { type: Object, default: null },
});

const formatting = computed(() => asFormatting(props.formatting));
const painter = computed(() => asFormatPainter(props.painter));
const s = computed(() => formatting.value.toolbar.value);
const run = (command = Object()) => formatting.value.run(command);
const pressed = (format = "") => Reflect.get(s.value, format) === true;

const BAND_ORDER = [1, 2, 3, 4, 5, 6, 7, 8];
const BAND_NAMES = ["", "History", "Type", "Character", "Colour", "Number", "Alignment", "Indent & orientation", "Styles"];

const isMac = isMacPlatform();
const tip = (label = "", id = "") => {
  const keys = shortcutLabel(id, isMac, SHEET_SHORTCUTS);
  return keys ? `${label} (${keys})` : label;
};

const TOGGLES = [
  { format: "bold", label: tip("Bold", "sheet.bold"), icon: "bold", testid: "sheet-bold" },
  { format: "italic", label: tip("Italic", "sheet.italic"), icon: "italic", testid: "sheet-italic" },
  { format: "underline", label: tip("Underline", "sheet.underline"), icon: "underline", testid: "sheet-underline" },
  { format: "strikethrough", label: tip("Strikethrough", "sheet.strikethrough"), icon: "strikethrough", testid: "sheet-strikethrough" },
];
const ALIGNS = [
  { value: "left", label: "Align left", icon: "align-left" },
  { value: "center", label: "Centre", icon: "align-center" },
  { value: "right", label: "Align right", icon: "align-right" },
];
const VALIGNS = [
  { value: "top", label: "Top align", icon: "align-top", testid: "sheet-valign-top" },
  { value: "center", label: "Middle align", icon: "align-middle", testid: "sheet-valign-middle" },
  { value: "bottom", label: "Bottom align", icon: "align-bottom", testid: "sheet-valign-bottom" },
];
const ROTATIONS = [
  { degrees: 45, label: "Angle counterclockwise", testid: "sheet-rotation-ccw" },
  { degrees: 135, label: "Angle clockwise", testid: "sheet-rotation-cw" },
  { degrees: 255, label: "Vertical text", testid: "sheet-rotation-vertical" },
  { degrees: 90, label: "Rotate text up", testid: "sheet-rotation-up" },
  { degrees: 180, label: "Rotate text down", testid: "sheet-rotation-down" },
];
const NUMBER_TOOLS = [
  { command: { kind: "currency" }, label: "Accounting number format", icon: "currency", testid: "sheet-currency" },
  { command: { kind: "percent" }, label: "Percent style", icon: "percent", testid: "sheet-percent" },
  { command: { kind: "comma" }, label: "Comma style", icon: "comma-style", testid: "sheet-comma" },
  { command: { kind: "decimals", delta: -1 }, label: "Decrease decimal", icon: "decimals-decrease", testid: "sheet-decimals-decrease" },
  { command: { kind: "decimals", delta: 1 }, label: "Increase decimal", icon: "decimals-increase", testid: "sheet-decimals-increase" },
];
const INDENT_TOOLS = [
  { command: { kind: "indent", delta: -1 }, label: "Decrease indent", icon: "indent-decrease", testid: "sheet-indent-decrease" },
  { command: { kind: "indent", delta: 1 }, label: "Increase indent", icon: "indent-increase", testid: "sheet-indent-increase" },
];
const CLEARS = [
  { what: "all", label: "Clear all", testid: "sheet-clear-all" },
  { what: "formats", label: "Clear formats", testid: "sheet-clear-formats" },
  { what: "contents", label: "Clear contents", testid: "sheet-clear-contents" },
];

// ---- overflow (useToolbarOverflow, unchanged) ------------------------------
const strip = useTemplateRef("strip");
const overflowButton = useTemplateRef("overflowButton");
const overflowMenu = useTemplateRef("overflowMenu");
const menuOpen = ref(false);
const { collapsedCount, measuring, isCollapsed } = useToolbarOverflow({ strip: () => strip.value, bandOrder: BAND_ORDER, reserved: ref(0), captionsWanted: () => false });
const { menuStyle } = useFloatingMenu({ open: () => menuOpen.value, anchor: () => overflowButton.value, menu: () => overflowMenu.value, align: "start" });

function teleportDisabled(band = 0) {
  return measuring.value || !isCollapsed(band) || !document.getElementById(`sheet-overflow-band-${band}`);
}

watch(collapsedCount, (count) => {
  if (count === 0) {
    menuOpen.value = false;
  }
  void updateRoving();
});

function onDocPointerDown(event = new PointerEvent("pointerdown")) {
  const target = event.target;
  if (!menuOpen.value || !(target instanceof Node)) {
    return;
  }
  const menu = overflowMenu.value;
  if ((menu instanceof HTMLElement && menu.contains(target)) || overflowButton.value?.contains(target) || (target instanceof Element && target.closest(".ui-menu-surface"))) {
    return;
  }
  menuOpen.value = false;
}

// ---- roving tab stop (SPEC-061) -------------------------------------------
function controls(all = false) {
  const el = strip.value;
  if (!el) {
    return [];
  }
  const selector = all ? "button, input, select" : "button:not([disabled]), input:not([disabled]), select:not([disabled])";
  return Array.from(el.querySelectorAll(selector)).filter(
    (control) => control instanceof HTMLElement && !control.closest("[role='menu'], [role='dialog'], .ui-menu-surface") && (all || control.offsetParent !== null),
  );
}

async function updateRoving() {
  await nextTick();
  const reachable = controls();
  const focused = document.activeElement;
  const active = reachable.find((control) => control === focused) ?? reachable[0];
  for (const control of controls(true)) {
    if (control instanceof HTMLElement) {
      control.tabIndex = control === active ? 0 : -1;
    }
  }
}

function onFocusin(event = new FocusEvent("focusin")) {
  const target = event.target;
  if (target instanceof HTMLElement && controls().includes(target)) {
    for (const control of controls(true)) {
      if (control instanceof HTMLElement) {
        control.tabIndex = control === target ? 0 : -1;
      }
    }
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  const key = event.key;
  const active = document.activeElement;
  if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(key) || (active instanceof Element && active.closest("[role='menu'], [role='dialog']"))) {
    return;
  }
  if ((key === "ArrowLeft" || key === "ArrowRight") && active instanceof HTMLInputElement) {
    const atStart = active.selectionStart === 0 && active.selectionEnd === 0;
    const atEnd = active.selectionStart === active.value.length;
    if ((key === "ArrowLeft" && !atStart) || (key === "ArrowRight" && !atEnd)) {
      return;
    }
  }
  const list = controls();
  const index = active instanceof HTMLElement ? list.indexOf(active) : -1;
  const next = key === "Home" ? 0 : key === "End" ? list.length - 1 : key === "ArrowRight" ? Math.min(list.length - 1, index + 1) : Math.max(0, index - 1);
  const target = list[next];
  if (target instanceof HTMLElement) {
    event.preventDefault();
    target.focus();
  }
}

watch(() => s.value.disabled, () => void updateRoving());

onMounted(() => {
  void updateRoving();
  document.addEventListener("pointerdown", onDocPointerDown, true);
});

onBeforeUnmount(() => {
  document.removeEventListener("pointerdown", onDocPointerDown, true);
});
</script>

<template>
  <div
    ref="strip"
    class="sheet-toolbar"
    role="toolbar"
    aria-label="Sheet toolbar"
    data-testid="sheet-format-toolbar"
    @keydown="onKeydown"
    @focusin="onFocusin"
  >
    <Teleport to="body">
      <div
        ref="overflowMenu"
        class="overflow-menu ui-menu-surface"
        :class="{ 'is-open': menuOpen }"
        role="menu"
        aria-label="More formatting"
        data-testid="sheet-toolbar-overflow-menu"
        :style="menuStyle"
      >
        <div
          v-for="band in [5, 6, 7, 8]"
          v-show="isCollapsed(band)"
          :key="band"
          class="overflow-section"
        >
          <p class="overflow-heading">
            {{ BAND_NAMES[band] }}
          </p>
          <div
            :id="`sheet-overflow-band-${band}`"
            class="overflow-band-target"
          />
        </div>
      </div>
    </Teleport>
    <div
      class="band"
      data-band="1"
      role="group"
      :aria-label="BAND_NAMES[1]"
    >
      <button
        type="button"
        class="tool"
        data-testid="sheet-format-painter"
        aria-label="Format painter (double-click to keep it on)"
        title="Format painter (double-click to keep it on)"
        :aria-pressed="painter.mode.value !== '' ? 'true' : 'false'"
        :data-locked="painter.mode.value === 'locked' ? 'true' : undefined"
        :class="{ 'is-pressed': painter.mode.value !== '' }"
        :disabled="s.disabled || (painter.disabled.value && painter.mode.value === '')"
        @mousedown.prevent
        @click="painter.click()"
        @dblclick="painter.dblclick()"
      >
        <AppIcon name="format-painter" />
      </button>
    </div>
    <div
      class="band"
      data-band="2"
      role="group"
      :aria-label="BAND_NAMES[2]"
    >
      <FontFamilySelect
        :value="s.fontFamily"
        :document-fonts="s.documentFonts"
        :disabled="s.disabled"
        @select="run({ kind: 'font-family', family: $event })"
      />
      <FontSizeControl
        :value="s.fontSizePt"
        :min="1"
        :max="409"
        :disabled="s.disabled"
        @change="run({ kind: 'font-size', sizePt: $event })"
      />
    </div>
    <div
      class="band"
      data-band="3"
      role="group"
      :aria-label="BAND_NAMES[3]"
    >
      <UiToggleButton
        v-for="toggle in TOGGLES"
        :key="toggle.format"
        :icon="toggle.icon"
        :label="toggle.label"
        :pressed="pressed(toggle.format)"
        :disabled="s.disabled"
        :data-testid="toggle.testid"
        @mousedown.prevent
        @click="run({ kind: 'toggle', format: toggle.format })"
      />
    </div>
    <div
      class="band"
      data-band="4"
      role="group"
      :aria-label="BAND_NAMES[4]"
    >
      <TextColorPicker
        :value="s.textColor"
        :recent="s.recentColors"
        :theme="s.theme"
        :disabled="s.disabled"
        @select="run({ kind: 'text-color', color: $event })"
      />
      <SheetFillPicker
        :color="s.fillColor"
        :theme="s.theme"
        :recent="s.recentColors"
        :disabled="s.disabled"
        @apply-last="run({ kind: 'fill-last' })"
        @select="run({ kind: 'fill', color: $event })"
      />
      <SheetBordersMenu
        :line-style="s.borderLineStyle"
        :line-color="s.borderLineColor"
        :disabled="s.disabled"
        @apply-last="run({ kind: 'border-last' })"
        @preset="run({ kind: 'border', preset: $event })"
        @line-style="run({ kind: 'border-line-style', style: $event })"
        @line-color="run({ kind: 'border-line-color', color: $event })"
      />
    </div>
    <div
      v-show="!isCollapsed(5) || measuring"
      class="band"
      data-band="5"
      role="group"
      :aria-label="BAND_NAMES[5]"
    >
      <Teleport
        to="#sheet-overflow-band-5"
        :disabled="teleportDisabled(5)"
      >
        <SheetMenuButton
          v-slot="{ close }"
          label="Number format"
          :text="s.numberFormatLabel"
          testid="sheet-number-format"
          menu-testid="sheet-number-format-menu"
          :disabled="s.disabled"
        >
          <button
            v-for="item in s.numberFormats"
            :key="item.id"
            type="button"
            role="menuitem"
            :data-testid="`sheet-number-format-${item.id}`"
            @mousedown.prevent
            @click="run({ kind: 'number-format', preset: item.id }); close()"
          >
            <span>{{ item.label }}</span>
            <span class="item-hint">{{ item.preview }}</span>
          </button>
          <button
            type="button"
            role="menuitem"
            data-testid="sheet-number-format-more"
            @mousedown.prevent
            @click="close(); run({ kind: 'format-cells', tab: 'number' })"
          >
            More formats…
          </button>
        </SheetMenuButton>
        <button
          v-for="tool in NUMBER_TOOLS"
          :key="tool.testid"
          type="button"
          class="tool"
          :data-testid="tool.testid"
          :aria-label="tool.label"
          :title="tool.label"
          :disabled="s.disabled"
          @mousedown.prevent
          @click="run(tool.command)"
        >
          <AppIcon :name="tool.icon" />
        </button>
      </Teleport>
    </div>
    <div
      v-show="!isCollapsed(6) || measuring"
      class="band"
      data-band="6"
      role="group"
      :aria-label="BAND_NAMES[6]"
    >
      <Teleport
        to="#sheet-overflow-band-6"
        :disabled="teleportDisabled(6)"
      >
        <UiToggleButton
          v-for="align in ALIGNS"
          :key="align.value"
          :icon="align.icon"
          :label="align.label"
          :pressed="s.horizontal === align.value"
          :disabled="s.disabled"
          :data-testid="`sheet-align-${align.value}`"
          @mousedown.prevent
          @click="run({ kind: 'align', horizontal: align.value })"
        />
        <UiToggleButton
          v-for="align in VALIGNS"
          :key="align.value"
          :icon="align.icon"
          :label="align.label"
          :pressed="s.vertical === align.value"
          :disabled="s.disabled"
          :data-testid="align.testid"
          @mousedown.prevent
          @click="run({ kind: 'valign', vertical: align.value })"
        />
        <UiToggleButton
          icon="wrap-text"
          label="Wrap text"
          :pressed="s.wrap"
          :disabled="s.disabled"
          data-testid="sheet-wrap"
          @mousedown.prevent
          @click="run({ kind: 'wrap' })"
        />
        <!-- SPEC-124's Merge & Centre split button (`sheet-merge`). -->
        <span
          class="merge-slot"
          data-slot="sheet-merge"
        >
          <SheetMergeButton
            v-if="props.merges"
            :merges="props.merges"
          />
        </span>
      </Teleport>
    </div>
    <div
      v-show="!isCollapsed(7) || measuring"
      class="band"
      data-band="7"
      role="group"
      :aria-label="BAND_NAMES[7]"
    >
      <Teleport
        to="#sheet-overflow-band-7"
        :disabled="teleportDisabled(7)"
      >
        <button
          v-for="tool in INDENT_TOOLS"
          :key="tool.testid"
          type="button"
          class="tool"
          :data-testid="tool.testid"
          :aria-label="tool.label"
          :title="tool.label"
          :disabled="s.disabled"
          @mousedown.prevent
          @click="run(tool.command)"
        >
          <AppIcon :name="tool.icon" />
        </button>
        <SheetMenuButton
          v-slot="{ close }"
          label="Text rotation"
          icon="text-rotation"
          testid="sheet-rotation"
          menu-testid="sheet-rotation-menu"
          :disabled="s.disabled"
        >
          <button
            v-for="rotation in ROTATIONS"
            :key="rotation.degrees"
            type="button"
            role="menuitemradio"
            :aria-checked="s.textRotation === rotation.degrees ? 'true' : 'false'"
            :data-testid="rotation.testid"
            @mousedown.prevent
            @click="run({ kind: 'rotation', degrees: rotation.degrees }); close()"
          >
            {{ rotation.label }}
          </button>
          <button
            type="button"
            role="menuitem"
            data-testid="sheet-rotation-format"
            @mousedown.prevent
            @click="close(); run({ kind: 'format-cells', tab: 'alignment' })"
          >
            Format cell alignment…
          </button>
        </SheetMenuButton>
      </Teleport>
    </div>
    <div
      v-show="!isCollapsed(8) || measuring"
      class="band"
      data-band="8"
      role="group"
      :aria-label="BAND_NAMES[8]"
    >
      <Teleport
        to="#sheet-overflow-band-8"
        :disabled="teleportDisabled(8)"
      >
        <SheetMenuButton
          v-slot="{ close }"
          label="Cell styles"
          icon="cell-styles"
          testid="sheet-cell-styles"
          menu-testid="sheet-cell-styles-menu"
          :disabled="s.disabled"
        >
          <CellStylesGallery
            :groups="s.cellStyles"
            @select="close(); run({ kind: 'cell-style', name: $event })"
          />
        </SheetMenuButton>
        <SheetMenuButton
          v-slot="{ close }"
          label="Clear"
          icon="clear-formatting"
          testid="sheet-clear"
          menu-testid="sheet-clear-menu"
          :disabled="s.disabled"
        >
          <button
            v-for="clear in CLEARS"
            :key="clear.what"
            type="button"
            role="menuitem"
            :data-testid="clear.testid"
            @mousedown.prevent
            @click="close(); run({ kind: 'clear', what: clear.what })"
          >
            {{ clear.label }}
          </button>
        </SheetMenuButton>
      </Teleport>
    </div>
    <button
      v-show="collapsedCount > 0 || measuring"
      ref="overflowButton"
      type="button"
      class="tool overflow-button"
      data-testid="toolbar-overflow-button"
      aria-label="More formatting"
      title="More formatting"
      aria-haspopup="menu"
      :aria-expanded="menuOpen ? 'true' : 'false'"
      @mousedown.prevent
      @click="menuOpen = !menuOpen"
    >
      <AppIcon name="more-horizontal" />
    </button>
  </div>
</template>

<style scoped>
/* One non-wrapping row that never scrolls: it collapses bands instead. */
.sheet-toolbar {
  flex: none;
  display: flex;
  align-items: center;
  min-width: 0;
  width: 100%;
  overflow: hidden;
  padding: var(--space-1) var(--space-2);
  box-sizing: border-box;
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
  --toolbar-caption-display: none;
  --control-height: var(--height-control);
}

.band {
  display: inline-flex;
  align-items: center;
  gap: 0;
  flex: none;
  padding-inline: var(--space-2);
}

.band[data-band="1"] {
  padding-inline-start: 0;
}

/* SPEC-061's 1px rule between bands. */
.band[data-band]:not([data-band="1"]) {
  border-inline-start: var(--border-width) solid var(--border-subtle);
}

.tool {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: var(--height-control);
  height: var(--control-height, var(--height-control));
  padding: 0;
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.tool:not(:disabled):hover {
  background: var(--state-hover);
}

.tool:disabled {
  opacity: 0.5;
  cursor: default;
}

.tool:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.tool.is-pressed {
  background: var(--surface-selected);
}

.merge-slot {
  display: inline-flex;
}

.overflow-button {
  margin-inline-start: var(--space-2);
}

.overflow-menu {
  z-index: 40;
  --control-height: var(--height-control);
  min-width: 14rem;
  padding: var(--space-2);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  visibility: hidden;
  pointer-events: none;
}

.overflow-menu.is-open {
  visibility: visible;
  pointer-events: auto;
}

.overflow-heading {
  margin: 0 0 var(--space-1);
  font-size: var(--text-caption);
  text-transform: uppercase;
  color: var(--text-accent);
}

.overflow-band-target {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}
</style>
