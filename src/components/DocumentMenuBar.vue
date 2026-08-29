<script setup lang="ts">
// The document menu bar (SPEC-049): File/Edit/View/Insert/Format/Tools, in
// that order, hosting the non-continuous-use commands relocated off the
// toolbar strip. It is a stable map of the product: a command that does not
// apply is disabled, never hidden (AC-7); the toolbar keeps only the
// continuous-use controls (AC-3).
//
// This component is dumb rendering plus one command channel: every menu
// state it shows arrives as a prop, every nullary command leaves as a
// `command` event carrying a string id (DocumentView.vue owns the dispatch
// table), and the handful of picks that carry a payload — the table-size
// grid, break kinds, note kinds, the shape/chart galleries, the multilevel
// presets, the image file and the format painter's click/dblclick pair —
// each have a dedicated event. The only behaviour owned here is menu-local:
// the table grid's hover state and the hidden image picker the "Image…"
// command clicks.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed, ref, useTemplateRef } from "vue";
import MenuBar from "./MenuBar.vue";
import { useTier } from "../composables/useResponsiveTier";
import MenuItem from "./MenuItem.vue";
import MenuCheckItem from "./MenuCheckItem.vue";
import MenuRadioItem from "./MenuRadioItem.vue";
import MenuSubmenu from "./MenuSubmenu.vue";
import ShapeGallery from "./ShapeGallery.vue";
import ChartGallery from "./ChartGallery.vue";
import { GALLERY_PRESETS as MULTILEVEL_PRESETS } from "../lib/core/listDefinitions";
import { PICKER_MAX_ROWS, PICKER_MAX_COLS } from "../lib/core/tables";
import { ACCEPTED_IMAGE_ACCEPT_ATTR } from "../lib/core/media";
import { shortcutLabel } from "../lib/shortcuts";
import { ZOOM_PRESETS } from "../lib/zoom";

// SPEC-063: whether a preset's View → Zoom radio shows checked — only when no
// fit mode is active and the live percentage matches the preset (AC-34).
function zoomPresetChecked(preset = 0, percent = 100, fitMode = "none") {
  return fitMode === "none" && Math.round(percent) === preset;
}

const props = defineProps({
  /** False renders every editing command disabled (present, not hidden). */
  editable: { type: Boolean, default: false },
  /** The view's undo/redo state: { canUndo, undoLabel, canRedo, redoLabel }. */
  undoRedo: {
    type: Object,
    default: () => ({ canUndo: false, undoLabel: "", canRedo: false, redoLabel: "" }),
  },
  autocorrectSmartQuotes: { type: Boolean, default: false },
  autocorrectDashes: { type: Boolean, default: false },
  /** View → Rulers: true when both rulers are shown, "mixed" when only one
   * is (SPEC-027 AC-14, SPEC-064 AC-7). */
  showRuler: { type: [Boolean, String], default: true },
  /** The per-ruler toggles' own states. */
  showHorizontalRuler: { type: Boolean, default: true },
  showVerticalRuler: { type: Boolean, default: true },
  outlinePaneOpen: { type: Boolean, default: false },
  stylesPaneOpen: { type: Boolean, default: false },
  commentsPaneOpen: { type: Boolean, default: false },
  commentsOnlyView: { type: Boolean, default: false },
  trackChangesRecording: { type: Boolean, default: false },
  trackChangesPaneOpen: { type: Boolean, default: false },
  canConvertLiteralChips: { type: Boolean, default: false },
  /** "View → Field shading" checked state (SPEC-053 AC-11). */
  fieldShading: { type: Boolean, default: false },
  /** The appearance preference: "system" | "light" | "dark" (SPEC-059). */
  theme: { type: String, default: "system" },
  /** The theme actually in effect: "light" | "dark" (SPEC-059) — gates the
   * dim-page item, which is dark-only. */
  resolvedTheme: { type: String, default: "light" },
  /** Whether the dim-page opt-in is on (SPEC-059 AC-13/14). */
  dimPage: { type: Boolean, default: false },
  /** Whether "Insert -> Chart from table" applies (a table holds the active
   * cell-range selection) — computed by the view, which owns the model. */
  tableSelectedForChart: { type: Boolean, default: false },
  /** Whether "Insert -> Caption" applies (an image/chart/table is selected) —
   * computed by the view (SPEC-055 AC-1). */
  captionTargetSelected: { type: Boolean, default: false },
  /** "Auto-caption new images/tables" checked states (SPEC-055 AC-18). */
  autoCaptionImages: { type: Boolean, default: false },
  autoCaptionTables: { type: Boolean, default: false },
  /** The format painter's armed state: "" | "once" | "repeat". */
  painterMode: { type: String, default: "" },
  /** SPEC-063 view mode: "print" | "web" (Focus is a separate flag). */
  viewMode: { type: String, default: "print" },
  /** SPEC-063 Focus mode active. */
  focusMode: { type: Boolean, default: false },
  /** SPEC-063 live zoom percentage and fit mode, for the View → Zoom checks. */
  zoomPercent: { type: Number, default: 100 },
  zoomFitMode: { type: String, default: "none" },
  /** SPEC-072 Format → Drawing: "" when align/distribute are available, else
   * the disabled-with-reason string for the current selection (AC-17). */
  drawingAlignReason: { type: String, default: "" },
  drawingDistributeReason: { type: String, default: "" },
  /** The remembered single-drawing align reference: "margin" | "page" |
   * "column" (AC-3). */
  alignToReference: { type: String, default: "margin" },
  /** Whether the Column reference is offered (section has >1 column, AC-3). */
  columnReferenceAvailable: { type: Boolean, default: false },
});

const emit = defineEmits([
  "command",
  "insert-table-size",
  "insert-break",
  "insert-section-break",
  "insert-note",
  "shape-pick",
  "connector-pick",
  "freeform-pick",
  "chart-pick",
  "multilevel-preset",
  "insert-image-file",
  "painter-click",
  "painter-dblclick",
  "drawing-align-to",
]);

// SPEC-072 Format → Drawing → Align / Distribute items. The command suffix
// doubles as the `drawing-<id>` testid.
const DRAWING_ALIGN_ITEMS = [
  { id: "align-left", label: "Align left" },
  { id: "align-center", label: "Align centre" },
  { id: "align-right", label: "Align right" },
  { id: "align-top", label: "Align top" },
  { id: "align-middle", label: "Align middle" },
  { id: "align-bottom", label: "Align bottom" },
];
const DRAWING_DISTRIBUTE_ITEMS = [
  { id: "distribute-horizontal", label: "Distribute horizontally" },
  { id: "distribute-vertical", label: "Distribute vertically" },
];
const ALIGN_TO_REFERENCES = [
  { value: "margin", label: "Margin" },
  { value: "page", label: "Page" },
  { value: "column", label: "Column" },
];

// SPEC-068 AC-12/13 — features that genuinely cannot be operated by touch on a
// small screen stay LISTED in their menu but disabled, with a stated reason,
// below 640px. The menu is a stable map of the product (SPEC-049): a command is
// disabled, never hidden. Resizing past 640px re-enables them with no reload.
const tier = useTier();
const compact = computed(() => tier.value === "compact");
const COMPACT_UNAVAILABLE_REASON = "not available on this screen size";
/** The reason string for an item that is unavailable only because of the
 * compact tier (empty at every wider tier so the shortcut shows instead). */
const compactReason = computed(() => (compact.value ? COMPACT_UNAVAILABLE_REASON : ""));

// File → Export's formats, in menu order: Word (SPEC-006), OpenDocument
// (SPEC-007), Markdown (SPEC-050), PDF (SPEC-034 AC-14). The id doubles as the
// `file.export-<id>` command suffix and the item's `download-<id>` testid.
const EXPORT_FORMATS = [
  { id: "docx", label: "Word (.docx)" },
  { id: "odt", label: "OpenDocument Text (.odt)" },
  { id: "md", label: "Markdown (.md)" },
  { id: "pdf", label: "PDF (.pdf)" },
];

const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "insert", label: "Insert" },
  { id: "format", label: "Format" },
  { id: "tools", label: "Tools" },
];

// The Insert menu's "Image…" command drives the same hidden file picker the
// old toolbar button did (SPEC-024).
const menuImageInput = useTemplateRef("menuImageInput");

function triggerImageInsert() {
  if (props.editable) {
    menuImageInput.value?.click();
  }
}

function handleMenuImageChange(event = new Event("change")) {
  const target = event.target;
  const file = target instanceof HTMLInputElement ? target.files?.[0] : undefined;
  if (file) {
    emit("insert-image-file", file);
  }
  if (target instanceof HTMLInputElement) {
    // Clear so picking the same file twice still fires a change event.
    target.value = "";
  }
}

// The Insert menu's table-size grid (SPEC-022): the same quick-pick the old
// toolbar dropdown offered, now a section of the Insert menu.
const TABLE_GRID_ROWS = Array.from({ length: PICKER_MAX_ROWS }, (_, index) => index + 1);
const TABLE_GRID_COLS = Array.from({ length: PICKER_MAX_COLS }, (_, index) => index + 1);
const tableGridHover = ref({ rows: 0, cols: 0 });

function hoverTableGrid(rows = 0, cols = 0) {
  tableGridHover.value = { rows, cols };
}

function chooseTableGrid(rows = 0, cols = 0) {
  emit("insert-table-size", rows, cols);
  // The pick closes the menu, so clear the highlight the pointer left behind:
  // reopening Insert should show an unlit grid, not the last size chosen.
  hoverTableGrid(0, 0);
}
</script>

<template>
  <MenuBar
    :menus="MENUS"
    aria-label="Document menu"
  >
    <template #file>
      <!-- Export (SPEC-006/007/034/050): a copy of the document in a chosen
           format, downloaded. Lives here rather than as an app-bar button so
           the bar keeps only the save state and the Close control. -->
      <MenuSubmenu
        label="Export"
        testid="file-menu-export"
      >
        <MenuItem
          v-for="format in EXPORT_FORMATS"
          :key="format.id"
          :label="format.label"
          :testid="`download-${format.id}`"
          @select="emit('command', `file.export-${format.id}`)"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <MenuItem
        label="Print"
        testid="file-menu-print"
        :shortcut="shortcutLabel('file.print')"
        @select="emit('command', 'file.print')"
      />
      <MenuItem
        label="Page setup"
        testid="page-setup-button"
        :disabled="!editable"
        @select="emit('command', 'file.page-setup')"
      />
      <MenuItem
        label="Properties…"
        testid="file-menu-properties"
        @select="emit('command', 'file.properties')"
      />
      <MenuItem
        label="Version history"
        testid="file-menu-version-history"
        @select="emit('command', 'file.version-history')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Save as template"
        testid="file-menu-save-as-template"
        @select="emit('command', 'file.save-as-template')"
      />
      <MenuItem
        label="Mail merge…"
        testid="file-menu-mail-merge"
        :disabled="compact"
        :reason="compactReason"
        @select="emit('command', 'file.mail-merge')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Delete document"
        testid="delete-document"
        :disabled="!editable"
        @select="emit('command', 'file.delete')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Preferences…"
        testid="file-menu-preferences"
        :shortcut="shortcutLabel('app.preferences')"
        @select="emit('command', 'app.preferences')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Close"
        testid="file-menu-close"
        @select="emit('command', 'file.close')"
      />
    </template>

    <template #edit>
      <MenuItem
        label="Undo"
        testid="edit-menu-undo"
        :shortcut="shortcutLabel('edit.undo')"
        :disabled="!editable || !undoRedo.canUndo"
        @select="emit('command', 'edit.undo')"
      />
      <MenuItem
        label="Redo"
        testid="edit-menu-redo"
        :shortcut="shortcutLabel('edit.redo')"
        :disabled="!editable || !undoRedo.canRedo"
        @select="emit('command', 'edit.redo')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Cut"
        testid="edit-menu-cut"
        :disabled="!editable"
        @select="emit('command', 'edit.cut')"
      />
      <MenuItem
        label="Copy"
        testid="edit-menu-copy"
        @select="emit('command', 'edit.copy')"
      />
      <MenuItem
        label="Paste"
        testid="edit-menu-paste"
        :disabled="!editable"
        @select="emit('command', 'edit.paste')"
      />
      <MenuItem
        label="Paste without formatting"
        testid="edit-menu-paste-plain"
        :shortcut="shortcutLabel('edit.pastePlain')"
        :disabled="!editable"
        @select="emit('command', 'edit.pastePlain')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Select all"
        testid="edit-menu-select-all"
        :shortcut="shortcutLabel('edit.selectAll')"
        @select="emit('command', 'edit.selectAll')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Find and replace"
        testid="edit-menu-find-replace"
        :shortcut="shortcutLabel('find.openReplace')"
        @select="emit('command', 'find.openReplace')"
      />
      <hr class="menu-separator">
      <div class="menu-section-label">
        Autocorrect
      </div>
      <MenuCheckItem
        label="Smart quotes"
        testid="autocorrect-smart-quotes"
        :checked="autocorrectSmartQuotes"
        :disabled="!editable"
        @select="emit('command', 'edit.autocorrect-smart-quotes')"
      />
      <MenuCheckItem
        label="Em and en dashes"
        testid="autocorrect-dashes"
        :checked="autocorrectDashes"
        :disabled="!editable"
        @select="emit('command', 'edit.autocorrect-dashes')"
      />
    </template>

    <template #view>
      <!-- SPEC-063: the three view modes as menuitemradio entries (AC-32). -->
      <MenuSubmenu
        label="Layout"
        testid="view-menu-layout"
      >
        <MenuRadioItem
          label="Print layout"
          testid="layout-print"
          :checked="viewMode === 'print' && !focusMode"
          @select="emit('command', 'view.layout-print')"
        />
        <MenuRadioItem
          label="Web layout"
          testid="layout-web"
          :checked="viewMode === 'web' && !focusMode"
          @select="emit('command', 'view.layout-web')"
        />
        <MenuRadioItem
          label="Focus"
          testid="layout-focus"
          :shortcut="shortcutLabel('view.focus')"
          :checked="focusMode"
          @select="emit('command', 'view.focus')"
        />
      </MenuSubmenu>
      <!-- SPEC-063: the zoom ladder plus Fit width / Fit page (AC-34). -->
      <MenuSubmenu
        label="Zoom"
        testid="view-menu-zoom"
      >
        <MenuRadioItem
          v-for="preset in ZOOM_PRESETS"
          :key="preset"
          :label="`${preset}%`"
          :testid="`zoom-${preset}`"
          :checked="zoomPresetChecked(preset, zoomPercent, zoomFitMode)"
          @select="emit('command', `view.zoom-${preset}`)"
        />
        <hr class="menu-separator">
        <MenuRadioItem
          label="Fit width"
          testid="zoom-fit-width"
          :checked="zoomFitMode === 'fit-width'"
          @select="emit('command', 'view.zoom-fit-width')"
        />
        <MenuRadioItem
          label="Fit page"
          testid="zoom-fit-page"
          :checked="zoomFitMode === 'fit-page'"
          @select="emit('command', 'view.zoom-fit-page')"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <!-- Rulers: one control for both (SPEC-064 AC-7), then each on its own. -->
      <MenuCheckItem
        label="Rulers"
        testid="toggle-ruler-button"
        :checked="showRuler"
        @select="emit('command', 'view.ruler')"
      />
      <MenuCheckItem
        label="Horizontal ruler"
        testid="toggle-horizontal-ruler-button"
        :checked="showHorizontalRuler"
        @select="emit('command', 'view.horizontal-ruler')"
      />
      <MenuCheckItem
        label="Vertical ruler"
        testid="toggle-vertical-ruler-button"
        :checked="showVerticalRuler"
        @select="emit('command', 'view.vertical-ruler')"
      />
      <hr class="menu-separator">
      <MenuCheckItem
        label="Outline pane"
        testid="view-menu-outline"
        :checked="outlinePaneOpen"
        @select="emit('command', 'view.outline-pane')"
      />
      <MenuCheckItem
        label="Styles pane"
        testid="view-menu-styles-pane"
        :checked="stylesPaneOpen"
        @select="emit('command', 'view.styles-pane')"
      />
      <MenuCheckItem
        label="Comments pane"
        testid="view-menu-comments-pane"
        :checked="commentsPaneOpen"
        @select="emit('command', 'view.comments-pane')"
      />
      <MenuCheckItem
        label="Comments-only view"
        testid="view-menu-comments-only"
        :checked="commentsOnlyView"
        @select="emit('command', 'view.comments-only')"
      />
      <MenuCheckItem
        label="Field shading"
        testid="view-menu-field-shading"
        :checked="fieldShading"
        @select="emit('command', 'view.field-shading')"
      />
      <hr class="menu-separator">
      <MenuSubmenu
        label="Appearance"
        testid="view-menu-appearance"
      >
        <MenuRadioItem
          label="System"
          testid="appearance-system"
          :checked="theme === 'system'"
          @select="emit('command', 'view.theme-system')"
        />
        <MenuRadioItem
          label="Light"
          testid="appearance-light"
          :checked="theme === 'light'"
          @select="emit('command', 'view.theme-light')"
        />
        <MenuRadioItem
          label="Dark"
          testid="appearance-dark"
          :checked="theme === 'dark'"
          @select="emit('command', 'view.theme-dark')"
        />
        <hr class="menu-separator">
        <MenuCheckItem
          label="Dim page in dark theme"
          testid="appearance-dim-page"
          hint="Screen only"
          :checked="dimPage"
          :disabled="resolvedTheme !== 'dark'"
          @select="emit('command', 'view.dim-page')"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <MenuCheckItem
        label="Track changes"
        testid="view-menu-track-changes"
        :shortcut="shortcutLabel('edit.toggleTrackChanges')"
        :checked="trackChangesRecording"
        :disabled="!editable"
        @select="emit('command', 'edit.toggleTrackChanges')"
      />
      <MenuCheckItem
        label="Track changes pane"
        testid="view-menu-track-changes-pane"
        :checked="trackChangesPaneOpen"
        @select="emit('command', 'view.track-changes-pane')"
      />
    </template>

    <template #insert>
      <div class="menu-section-label">
        Table
      </div>
      <div class="menu-table-grid-wrap">
        <div
          class="menu-table-grid"
          data-testid="insert-table-grid"
          @mouseleave="hoverTableGrid(0, 0)"
        >
          <div
            v-for="rows in TABLE_GRID_ROWS"
            :key="rows"
            class="menu-table-grid-row"
          >
            <button
              v-for="cols in TABLE_GRID_COLS"
              :key="cols"
              type="button"
              class="menu-table-grid-cell"
              :class="{ active: rows <= tableGridHover.rows && cols <= tableGridHover.cols }"
              :data-testid="`table-grid-cell-${rows}-${cols}`"
              :aria-label="`Insert ${rows} by ${cols} table`"
              :disabled="!editable"
              tabindex="-1"
              @mouseenter="hoverTableGrid(rows, cols)"
              @mousedown.prevent
              @click="chooseTableGrid(rows, cols)"
            />
          </div>
        </div>
        <div
          class="menu-table-grid-label"
          data-testid="insert-table-grid-label"
        >
          {{ tableGridHover.rows > 0 ? `${tableGridHover.rows} × ${tableGridHover.cols}` : "Insert table" }}
        </div>
      </div>
      <MenuItem
        label="Insert table…"
        testid="insert-table-dialog-item"
        :disabled="!editable"
        @select="emit('command', 'insert.table-dialog')"
      />
      <MenuItem
        label="Image…"
        testid="insert-image-button"
        :disabled="!editable"
        @select="triggerImageInsert"
      />
      <hr class="menu-separator">
      <div class="menu-section-label">
        Shape
      </div>
      <!-- Drawing creation gestures need a canvas that does not exist at the
           compact width (SPEC-068): the gallery is replaced by a disabled,
           reasoned entry, while existing drawings still render/select/move. -->
      <ShapeGallery
        v-if="!compact"
        @pick="emit('shape-pick', $event)"
        @pick-connector="emit('connector-pick', $event)"
        @pick-freeform="emit('freeform-pick', $event)"
      />
      <MenuItem
        v-else
        label="Shapes"
        testid="insert-shapes-disabled"
        disabled
        :reason="compactReason"
      />
      <MenuItem
        label="Text box"
        testid="insert-textbox-button"
        :disabled="!editable || compact"
        :reason="compactReason"
        @select="emit('command', 'insert.textbox')"
      />
      <hr class="menu-separator">
      <div class="menu-section-label">
        Chart
      </div>
      <ChartGallery @pick="emit('chart-pick', $event)" />
      <MenuItem
        label="Chart from table"
        testid="insert-chart-from-table-button"
        :disabled="!editable || !tableSelectedForChart"
        @select="emit('command', 'insert.chart-from-table')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Page break"
        testid="break-page"
        :shortcut="shortcutLabel('insert.pageBreak')"
        :disabled="!editable"
        @select="emit('insert-break', 'page')"
      />
      <MenuItem
        label="Column break"
        testid="break-column"
        :disabled="!editable"
        @select="emit('insert-break', 'column')"
      />
      <MenuItem
        label="Section break (next page)"
        testid="break-section-next"
        :disabled="!editable"
        @select="emit('insert-section-break', 'nextPage')"
      />
      <MenuItem
        label="Section break (continuous)"
        testid="break-section-continuous"
        :disabled="!editable"
        @select="emit('insert-section-break', 'continuous')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Header"
        testid="hf-menu-header"
        :disabled="!editable"
        @select="emit('command', 'insert.header')"
      />
      <MenuItem
        label="Footer"
        testid="hf-menu-footer"
        :disabled="!editable"
        @select="emit('command', 'insert.footer')"
      />
      <MenuItem
        label="Page number"
        testid="hf-menu-page-number"
        :disabled="!editable"
        @select="emit('command', 'insert.page-number')"
      />
      <MenuItem
        label="Header and footer options"
        testid="hf-menu-options"
        :disabled="!editable"
        @select="emit('command', 'insert.header-footer-options')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Merge field…"
        testid="insert-menu-merge-field"
        :disabled="!editable || compact"
        :reason="compactReason"
        @select="emit('command', 'insert.merge-field')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Footnote"
        testid="notes-menu-footnote"
        :shortcut="shortcutLabel('insert.footnote')"
        :disabled="!editable"
        @select="emit('insert-note', 'footnote')"
      />
      <MenuItem
        label="Endnote"
        testid="notes-menu-endnote"
        :shortcut="shortcutLabel('insert.endnote')"
        :disabled="!editable"
        @select="emit('insert-note', 'endnote')"
      />
      <MenuItem
        label="Footnote and endnote settings"
        testid="notes-menu-settings"
        :disabled="!editable"
        @select="emit('command', 'insert.note-settings')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Equation"
        testid="insert-equation-button"
        :shortcut="shortcutLabel('insert.equation')"
        :disabled="!editable"
        @select="emit('command', 'insert.equation')"
      />
      <MenuItem
        label="Display equation"
        testid="insert-display-equation-button"
        :disabled="!editable"
        @select="emit('command', 'insert.display-equation')"
      />
      <MenuItem
        label="Numbered equation"
        testid="insert-numbered-equation-button"
        :disabled="!editable"
        @select="emit('command', 'insert.numbered-equation')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Field…"
        testid="insert-field-button"
        :disabled="!editable"
        @select="emit('command', 'insert.field')"
      />
      <MenuItem
        label="Date and time"
        testid="insert-date-time-button"
        :disabled="!editable"
        @select="emit('command', 'insert.date-time')"
      />
      <MenuItem
        label="Cross-reference"
        testid="insert-cross-reference-button"
        :disabled="!editable"
        @select="emit('command', 'insert.cross-reference')"
      />
      <MenuItem
        label="Caption"
        testid="insert-caption-button"
        :disabled="!editable || !captionTargetSelected"
        @select="emit('command', 'insert.caption')"
      />
      <MenuItem
        label="Table of contents"
        testid="insert-toc-button"
        :disabled="!editable"
        @select="emit('command', 'insert.toc')"
      />
      <MenuItem
        label="Table of figures"
        testid="insert-table-of-figures-button"
        :disabled="!editable"
        @select="emit('command', 'insert.table-of-figures')"
      />
      <MenuCheckItem
        label="Auto-caption new images"
        testid="auto-caption-images"
        :checked="autoCaptionImages"
        :disabled="!editable"
        @select="emit('command', 'format.auto-caption-images')"
      />
      <MenuCheckItem
        label="Auto-caption new tables"
        testid="auto-caption-tables"
        :checked="autoCaptionTables"
        :disabled="!editable"
        @select="emit('command', 'format.auto-caption-tables')"
      />
      <MenuItem
        label="Special character"
        testid="insert-special-character-button"
        :disabled="!editable"
        @select="emit('command', 'insert.special-character')"
      />
      <MenuItem
        label="Tab character"
        testid="insert-tab-character-button"
        :disabled="!editable"
        @select="emit('command', 'insert.tab-character')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Link"
        testid="toolbar-insert-link"
        :shortcut="shortcutLabel('insert.link')"
        :disabled="!editable"
        @select="emit('command', 'insert.link')"
      />
      <MenuItem
        label="Bookmark"
        testid="toolbar-insert-bookmark"
        :disabled="!editable"
        @select="emit('command', 'insert.bookmark')"
      />
      <MenuItem
        label="Bookmark list"
        testid="toolbar-bookmarks"
        @select="emit('command', 'insert.bookmark-list')"
      />
      <MenuItem
        label="Comment"
        testid="toolbar-insert-comment"
        :shortcut="shortcutLabel('insert.comment')"
        :disabled="!editable"
        @select="emit('command', 'insert.comment')"
      />
    </template>

    <template #format>
      <MenuItem
        label="Styles pane"
        testid="format-menu-styles-pane"
        @select="emit('command', 'format.styles-pane')"
      />
      <MenuItem
        label="Import styles from another document…"
        testid="format-menu-import-styles"
        :disabled="!editable"
        @select="emit('command', 'format.import-styles')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Clear formatting"
        testid="clear-formatting-button"
        :shortcut="shortcutLabel('format.clearFormatting')"
        :disabled="!editable"
        @select="emit('command', 'format.clearFormatting')"
      />
      <MenuItem
        label="Tabs…"
        testid="format-menu-tabs"
        :disabled="!editable"
        @select="emit('command', 'format.tabs')"
      />
      <MenuItem
        label="Edit points"
        testid="format-menu-edit-points"
        :disabled="!editable"
        @select="emit('command', 'format.edit-points')"
      />
      <!-- Format painter keeps its double-click-to-lock gesture
           (SPEC-020 AC-14): data-menu-keep-open holds the menu open
           across the click/click/dblclick sequence, and its armed state
           shows as a live check (SPEC-049 AC-11). -->
      <button
        type="button"
        role="menuitemcheckbox"
        class="menu-item"
        data-testid="format-painter-button"
        data-menu-keep-open
        title="Format painter (double-click to keep it on)"
        :aria-checked="painterMode !== '' ? 'true' : 'false'"
        :disabled="!editable"
        tabindex="-1"
        @mousedown.prevent
        @click="emit('painter-click')"
        @dblclick="emit('painter-dblclick')"
      >
        <span
          class="menu-check"
          aria-hidden="true"
        >{{ painterMode !== '' ? '✓' : '' }}</span>
        <span class="menu-item-label">Format painter</span>
      </button>
      <hr class="menu-separator">
      <MenuItem
        label="Columns"
        testid="columns-button"
        :disabled="!editable"
        @select="emit('command', 'format.columns')"
      />
      <hr class="menu-separator">
      <!-- SPEC-072: align, distribute, the align-to reference and the numeric
           position dialog for a drawing selection. A new command names its
           menu (SPEC-049). -->
      <MenuSubmenu
        label="Drawing"
        testid="format-menu-drawing"
      >
        <MenuSubmenu
          label="Align"
          testid="drawing-align-submenu"
        >
          <MenuItem
            v-for="item in DRAWING_ALIGN_ITEMS"
            :key="item.id"
            :label="item.label"
            :testid="`drawing-${item.id}`"
            :disabled="!!drawingAlignReason"
            :reason="drawingAlignReason"
            @select="emit('command', `drawing.${item.id}`)"
          />
        </MenuSubmenu>
        <MenuSubmenu
          label="Distribute"
          testid="drawing-distribute-submenu"
        >
          <MenuItem
            v-for="item in DRAWING_DISTRIBUTE_ITEMS"
            :key="item.id"
            :label="item.label"
            :testid="`drawing-${item.id}`"
            :disabled="!!drawingDistributeReason"
            :reason="drawingDistributeReason"
            @select="emit('command', `drawing.${item.id}`)"
          />
        </MenuSubmenu>
        <MenuSubmenu
          label="Align to"
          testid="drawing-align-to-submenu"
        >
          <MenuRadioItem
            v-for="option in ALIGN_TO_REFERENCES"
            v-show="option.value !== 'column' || columnReferenceAvailable"
            :key="option.value"
            :label="option.label"
            :testid="`drawing-align-to-${option.value}`"
            :checked="alignToReference === option.value"
            @select="emit('drawing-align-to', option.value)"
          />
        </MenuSubmenu>
        <hr class="menu-separator">
        <MenuItem
          label="Position and size…"
          testid="drawing-position-size"
          :disabled="!editable"
          @select="emit('command', 'drawing.position-size')"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <div class="menu-section-label">
        Multilevel list
      </div>
      <MenuItem
        v-for="preset in MULTILEVEL_PRESETS"
        :key="preset.key"
        :label="preset.label"
        :testid="`multilevel-preset-${preset.key}`"
        :disabled="!editable"
        @select="emit('multilevel-preset', preset.key)"
      />
      <MenuItem
        label="Number headings 1–3"
        testid="multilevel-link-headings"
        :disabled="!editable"
        @select="emit('command', 'format.multilevel-link-headings')"
      />
      <MenuItem
        label="Define new list…"
        testid="multilevel-define-list"
        :disabled="!editable"
        @select="emit('command', 'format.multilevel-define-list')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Text colour"
        testid="format-menu-text-color"
        :disabled="!editable"
        @select="emit('command', 'format.text-color')"
      />
      <MenuItem
        label="Highlight colour"
        testid="format-menu-highlight"
        :disabled="!editable"
        @select="emit('command', 'format.highlight')"
      />
    </template>

    <template #tools>
      <MenuItem
        label="Search commands"
        testid="tools-menu-command-palette"
        @select="emit('command', 'app.commandPalette')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Word count"
        testid="tools-menu-word-count"
        @select="emit('command', 'tools.word-count')"
      />
      <MenuItem
        label="Custom dictionary"
        testid="tools-menu-dictionary"
        @select="emit('command', 'tools.dictionary')"
      />
      <MenuItem
        label="Update all fields"
        testid="tools-menu-update-all-fields"
        :disabled="!editable"
        @select="emit('command', 'tools.update-all-fields')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Check accessibility"
        testid="tools-menu-accessibility"
        @select="emit('command', 'tools.accessibility')"
      />
      <MenuItem
        label="Fonts"
        testid="tools-menu-fonts"
        @select="emit('command', 'tools.fonts')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Convert {{…}} text to merge fields"
        testid="tools-menu-convert-merge-fields"
        :disabled="!editable || !canConvertLiteralChips || compact"
        :reason="compactReason"
        @select="emit('command', 'tools.convert-merge-fields')"
      />
    </template>
  </MenuBar>
  <!-- The Insert menu's "Image…" command drives this hidden picker
       (SPEC-024). -->
  <input
    ref="menuImageInput"
    type="file"
    class="menu-image-input"
    data-testid="insert-image-file-input"
    :accept="ACCEPTED_IMAGE_ACCEPT_ATTR"
    @change="handleMenuImageChange"
  >
</template>

<style scoped>
/* The Insert menu's table-size quick-pick (SPEC-049), slotted into the menu
   popup. These rules reach it because slotted content keeps this component's
   scope id. */
.menu-image-input {
  display: none;
}

/* The hovered size reads out beside the grid rather than under it: the grid
   is far taller than it is wide, and the Insert menu needs every row of
   height it can give back to keep fitting a 720px viewport whole (SPEC-028
   AC-8/AC-9 asserts the break commands further down are inside it). */
.menu-table-grid-wrap {
  display: flex;
  align-items: flex-end;
  gap: var(--space-4);
  padding: var(--space-2) var(--space-5) var(--space-3);
}

.menu-table-grid {
  display: flex;
  flex-direction: column;
  gap: var(--border-width);
}

.menu-table-grid-row {
  display: flex;
  gap: var(--border-width);
}

.menu-table-grid-cell {
  width: 0.75rem;
  height: 0.75rem;
  padding: 0;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.menu-table-grid-cell.active {
  background: var(--accent-subtle);
  border-color: var(--accent);
}

.menu-table-grid-cell:disabled {
  cursor: default;
  opacity: 0.5;
}

.menu-table-grid-label {
  font-size: var(--text-sm);
  color: var(--text-secondary);
  white-space: nowrap;
}
</style>
