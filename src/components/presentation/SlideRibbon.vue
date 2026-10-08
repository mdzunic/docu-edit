<script setup lang="ts">
// The slides editor's ribbon (Toolbar Redesign, option 2c "Ribbon-lite").
//
// Before this, the slides editor stacked three separate rows of chrome under
// the menu bar — a deck header, the formatting toolbar and the insert bar —
// which is exactly what the handoff calls out: the word processor gives its
// commands one banded strip, and the deck gives them three competing ones.
//
// The deck line this component first drew above the tabs is gone too. It
// restated what the app bar already says — the deck title and its format —
// one row lower, so the title was on screen twice and the format three times
// ("PPTX" badge, "POWERPOINT (.PPTX)"). The handoff's 2c puts title, format,
// deck facts, save state and the deck actions on a single bar; the actions
// moved to AppShell's #app-bar-actions slot, which is where the word processor
// has always kept its own (EditorSaveStatus), so the two editors now wear the
// same top line.
//
// What is left is one chrome row and one strip. The row carries undo/redo —
// the deck's always-visible command pair, where DocumentToolbar's band 1 keeps
// the word processor's and where the handoff's own 2c strip begins — and two
// kinds of control that must NOT look alike:
//
//   · "Home" and "Insert" are tabs. They switch which slotted strip shows,
//     one at a time — a single-select tablist.
//   · Notes, Design, Text, Transitions, Animate, Layout, Table, Media and
//     Header & footer open side panels, any number of which can be open at
//     once (Text is SPEC-097's lists/spacing/autofit panel). Notes
//     joined them from the deck's action cluster, where it was the one entry
//     that opened a panel rather than running a command.
//
// They were all rendered as role="tab" with aria-selected, which put two
// "selected" tabs in one tablist (Insert *and* Table, say) — invalid for the
// WAI-ARIA tabs pattern, and on screen simply two accent-underlined tabs at
// once. So the panel toggles are now what they always were: aria-pressed
// buttons in their own trailing group, marked by a filled chip rather than the
// tab row's accent rule. Their testids are unchanged, so a panel still opens
// from the same place.
//
// The menu bar has since been merged into that row — the deck was showing two
// grey bars of chrome (File/Edit/View… on one, undo/redo + the tabs + the
// panel toggles on the next) where the word processor shows one, and the two
// even wore different grounds: the menu bar took --surface-raised and this row
// took whatever the canvas was painted, so a merge of the two also settles the
// ground. The bar comes in through the #menu slot, rendered `embedded` so it
// draws no bar of its own, and PresentationView slots this whole component
// into AppShell's #menubar region — above the canvas, which is a scroll
// container that would otherwise clip the menus' popups.
//
// The strip tabs could NOT stay in that merged row: they wear the same
// condensed uppercase face as the menu titles and mark themselves with the
// same accent underline the open menu uses, so "INSERT" the tab would have sat
// a thumb's width from "INSERT" the menu, identical in every respect but what
// clicking does. They moved to the head of the strip they switch instead,
// where the mark reads as "this is the strip you are looking at" and the row
// above is left with two unambiguous kinds of control: menus that open, and
// panel toggles that press.
//
// Runtime prop/emit declarations, not the generic forms — see AppShell.vue's
// note on the espree constraint every .vue script here shares.
import { ref } from "vue";
import AppIcon from "../AppIcon.vue";

const props = defineProps({
  canUndo: { type: Boolean, default: false },
  canRedo: { type: Boolean, default: false },
  /** Panel open states, rendered as each panel toggle's aria-pressed. */
  notesOpen: { type: Boolean, default: false },
  themeOpen: { type: Boolean, default: false },
  transitionOpen: { type: Boolean, default: false },
  animateOpen: { type: Boolean, default: false },
  layoutOpen: { type: Boolean, default: false },
  tableOpen: { type: Boolean, default: false },
  mediaOpen: { type: Boolean, default: false },
  headerFooterOpen: { type: Boolean, default: false },
  /** SPEC-097: the Text panel (lists, spacing, autofit, text body). */
  textOpen: { type: Boolean, default: false },
});

const emit = defineEmits([
  "undo",
  "redo",
  "notes-toggle",
  "theme-toggle",
  "transition-toggle",
  "animate-toggle",
  "layout-toggle",
  "table-toggle",
  "media-toggle",
  "header-footer-toggle",
  "text-toggle",
]);

/** Which slotted strip is showing. Both stay mounted (v-show, not v-if): the
 * strips own popovers and armed-insertion state that must survive a tab
 * change, exactly as the document toolbar's overflow keeps its controls
 * mounted rather than re-creating them. */
const tab = ref("home");

const STRIPS = [
  { id: "home", label: "Home" },
  { id: "insert", label: "Insert" },
];

// The panel toggles, in the handoff's order. `open` and `toggle` are functions,
// not values: `open` so the row re-reads the prop on every render rather than
// snapshotting it at setup, and `toggle` so each entry closes over its own
// literal event name — passing `emit(panel.event)` would widen the name to
// `string` and fail the emit's union type (and .vue scripts cannot use
// `as const` to narrow it back; espree rejects it).
const PANELS = [
  { id: "notes", label: "Notes", testid: "slides-notes-toggle", open: () => props.notesOpen, toggle: () => emit("notes-toggle") },
  { id: "layout", label: "Layout", testid: "slides-layout-toggle", open: () => props.layoutOpen, toggle: () => emit("layout-toggle") },
  { id: "design", label: "Design", testid: "slides-theme-toggle", open: () => props.themeOpen, toggle: () => emit("theme-toggle") },
  { id: "text", label: "Text", testid: "slides-text-toggle", open: () => props.textOpen, toggle: () => emit("text-toggle") },
  { id: "transition", label: "Transitions", testid: "slides-transition-toggle", open: () => props.transitionOpen, toggle: () => emit("transition-toggle") },
  { id: "animate", label: "Animate", testid: "slides-animate-toggle", open: () => props.animateOpen, toggle: () => emit("animate-toggle") },
  { id: "table", label: "Table", testid: "slides-table-toggle", open: () => props.tableOpen, toggle: () => emit("table-toggle") },
  { id: "media", label: "Media", testid: "slides-media-toggle", open: () => props.mediaOpen, toggle: () => emit("media-toggle") },
  { id: "header-footer", label: "Header & footer", testid: "slides-header-footer", open: () => props.headerFooterOpen, toggle: () => emit("header-footer-toggle") },
];

/** Roving arrow-key movement along the tab row (WAI-ARIA tabs pattern). Scoped
 * to the tablist, so it rings the two strip tabs only — the panel toggles are
 * ordinary buttons and stay in the natural tab order. */
function onTabKeydown(event = new KeyboardEvent("keydown")) {
  const step = event.key === "ArrowRight" ? 1 : event.key === "ArrowLeft" ? -1 : 0;
  if (step === 0) {
    return;
  }
  const row = event.currentTarget;
  if (!(row instanceof HTMLElement)) {
    return;
  }
  const tabs = Array.from(row.querySelectorAll('[role="tab"]'));
  const here = tabs.indexOf(document.activeElement instanceof Element ? document.activeElement : tabs[0]);
  const next = tabs[(here + step + tabs.length) % tabs.length];
  if (next instanceof HTMLElement) {
    event.preventDefault();
    next.focus();
  }
}
</script>

<template>
  <div class="slide-ribbon">
    <div class="ribbon-row">
      <!-- The menu bar, rendered `embedded` by its caller so the row's ground
           is the only one drawn. It leads: the menus are the map of the
           product, the way the word processor's bar leads its chrome. -->
      <div class="ribbon-menu">
        <slot name="menu" />
      </div>
      <!-- A rule before the tabs as well as after them. Menus and tabs are two
           kinds of control sharing one row of words: without the boundary,
           FILE EDIT VIEW SLIDE HOME INSERT reads as one list of six, four of
           which drop a popup and two of which swap the strip below. -->
      <span class="ribbon-rule" />
      <div
        class="ribbon-tabs"
        role="tablist"
        aria-label="Slide tools"
        data-testid="slides-ribbon-tabs"
        @keydown="onTabKeydown"
      >
        <button
          v-for="strip in STRIPS"
          :key="strip.id"
          type="button"
          role="tab"
          class="ribbon-tab"
          :class="{ 'is-selected': tab === strip.id }"
          :aria-selected="tab === strip.id ? 'true' : 'false'"
          :tabindex="tab === strip.id ? 0 : -1"
          :data-testid="`slides-tab-${strip.id}`"
          @click="tab = strip.id"
        >
          {{ strip.label }}
        </button>
      </div>
      <div
        class="ribbon-history"
        role="group"
        aria-label="History"
      >
        <span class="ribbon-rule" />
        <button
          type="button"
          class="ribbon-command"
          data-testid="presentation-undo"
          :disabled="!canUndo"
          aria-label="Undo"
          title="Undo"
          @click="emit('undo')"
        >
          <AppIcon name="undo" />
        </button>
        <button
          type="button"
          class="ribbon-command"
          data-testid="presentation-redo"
          :disabled="!canRedo"
          aria-label="Redo"
          title="Redo"
          @click="emit('redo')"
        >
          <AppIcon name="redo" />
        </button>
      </div>
      <!-- A flexible spacer rather than `margin-inline-start: auto` on the
           panel group: the auto margin also applies once the group has
           wrapped, which right-aligned it on a line of its own and left a
           ragged gap where the menus had been. A zero-basis spacer takes no
           part in the wrap decision, so it pushes the group to the trailing
           edge while everything fits on one line and simply collapses when it
           does not. -->
      <span class="ribbon-spacer" />
      <div
        class="ribbon-panels"
        role="group"
        aria-label="Panels"
        data-testid="slides-ribbon-panels"
      >
        <!-- The handoff names a zone with a kicker rather than boxing it (its
             "States" and "Fields" rows do the same). Without it the toggles
             read as more menu titles: same condensed uppercase face, same
             row. -->
        <span
          class="ribbon-panels-kicker"
          aria-hidden="true"
        >Panels</span>
        <button
          v-for="panel in PANELS"
          :key="panel.id"
          type="button"
          class="ribbon-panel-toggle"
          :class="{ 'is-open': panel.open() }"
          :aria-pressed="panel.open() ? 'true' : 'false'"
          :data-testid="panel.testid"
          @click="panel.toggle()"
        >
          {{ panel.label }}
        </button>
      </div>
    </div>
    <div class="ribbon-strips">
      <div
        v-show="tab === 'home'"
        class="ribbon-strip"
        role="tabpanel"
        aria-label="Home"
      >
        <slot name="home" />
      </div>
      <div
        v-show="tab === 'insert'"
        class="ribbon-strip"
        role="tabpanel"
        aria-label="Insert"
      >
        <slot name="insert" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.slide-ribbon {
  display: flex;
  flex-direction: column;
  border-bottom: var(--border-width) solid var(--border-subtle);
}

/* One line: the menus lead, then undo/redo, and the panel toggles sit at the
   trailing edge. It wraps rather than clipping — a narrow window drops the
   panel group to its own line instead of putting toggles out of reach, which
   is no worse than the two rows this row replaced. The ground is the chrome
   surface the strip below and the app's other bars already wear. */
.ribbon-row {
  display: flex;
  align-items: stretch;
  flex-wrap: wrap;
  gap: 0 var(--space-4);
  padding: 0 var(--space-3);
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
}

/* The embedded menu bar states its own height; this only stops the slot from
   stretching the row when the bar collapses to the compact tier's button. */
.ribbon-menu {
  display: flex;
  align-items: center;
}

.ribbon-tabs,
.ribbon-panels {
  display: flex;
  align-items: stretch;
  gap: 0;
}

.ribbon-history {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

/* Undo/redo are commands, not tabs and not toggles: the strip's plain
   transparent grammar, no selected mark of any kind. */
.ribbon-command {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: var(--height-control);
  height: var(--height-control);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.ribbon-command:hover:not(:disabled) {
  background: var(--state-hover);
}

.ribbon-command:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.ribbon-command:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.ribbon-rule {
  align-self: center;
  width: var(--border-width);
  height: var(--space-5);
  margin-inline: var(--space-2);
  background: var(--border-subtle);
}

.ribbon-spacer {
  flex: 1 1 0;
  min-width: 0;
}

.ribbon-panels {
  align-items: center;
  gap: var(--space-1);
}

/* The handoff's tab: condensed, uppercase and tracked, marked when selected by
   an accent rule along its bottom edge. The rule is a pseudo-element, not an
   inset box-shadow — tokens.spec allows only var(--shadow-*) shadows — and is
   out of flow so selecting a tab never shifts the row. The tabs stretch to the
   strip's full height, so that rule runs along the bottom of the strip they
   switch. */
.ribbon-tab,
.ribbon-panel-toggle {
  position: relative;
  display: inline-flex;
  align-items: center;
  height: var(--height-menu-bar);
  padding: 0 var(--space-4);
  border: var(--border-width) solid transparent;
  background: transparent;
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-md);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  white-space: nowrap;
  cursor: pointer;
}

/* The tabs stretch to the row's full height so the selected one's accent rule
   runs along the boundary with the strip it switches — the mark reads as
   "this strip", not as another open menu. The menu titles beside them are
   never persistently marked, are set a step smaller, and answer a click with
   a popup rather than a strip swap. */
.ribbon-tab {
  height: auto;
  align-self: stretch;
}

.ribbon-tab:hover,
.ribbon-panel-toggle:hover {
  background: var(--state-hover);
  color: var(--text-primary);
}

.ribbon-tab.is-selected {
  color: var(--text-accent);
}

.ribbon-tab.is-selected::after {
  content: "";
  position: absolute;
  inset-inline: 0;
  bottom: 0;
  height: var(--space-1);
  background: var(--accent);
}

/* An open panel is a pressed toggle, not a selected tab: the same filled chip
   UiToggleButton paints, so "which strip am I on" and "which panels are open"
   never read as the same mark. */
.ribbon-panel-toggle {
  height: var(--height-control);
  padding: 0 var(--space-3);
  font-size: var(--text-sm);
}

.ribbon-panels-kicker {
  padding-inline-end: var(--space-2);
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  color: var(--text-accent);
  white-space: nowrap;
}

.ribbon-panel-toggle.is-open {
  background: var(--surface-selected);
  border-color: var(--accent);
  color: var(--text-accent);
}

.ribbon-tab:focus-visible,
.ribbon-panel-toggle:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.ribbon-strips {
  display: flex;
  align-items: stretch;
  min-height: var(--height-toolbar);
  background: var(--surface-raised);
}

.ribbon-tabs {
  flex: none;
  margin-inline-start: var(--space-3);
}

.ribbon-strip {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  align-items: center;
  min-height: var(--height-toolbar);
  padding: 0 var(--space-3);
  --control-height: var(--height-toolbar-control);
}
</style>
