<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, useTemplateRef, watch } from "vue";
import { asMenuBarMenus } from "../lib/core/menuBar";
import UiMenuSurface from "./ui/UiMenuSurface.vue";
import { useTier } from "../composables/useResponsiveTier";

// The application menu bar (SPEC-049). This component is deliberately
// *structural only*: it renders the menu titles and their popups, and owns
// the menubar's keyboard model, single tab stop, pointer cross-open and
// open/close state. Each popup's box and its dismissal (outside pointerdown,
// Escape — even when a mouse open left focus on the page) are UiMenuSurface's
// (SPEC-058 AC-39); focus stays this component's, so the surface runs with
// returnFocus=false. The contents of each popup come from a slot named after
// the menu id, so the callers author their own role="menuitem" /
// menuitemcheckbox buttons (with their own shortcut labels, disabled and
// checked state). This component locates those items by role to drive
// Up/Down focus, so it never needs to know which commands exist.
//
// Runtime prop declaration (not defineProps<{...}>()) because ESLint parses
// .vue scripts with espree, which rejects TS-only syntax — see AppShell.vue.
const props = defineProps({
  /** The menus, in bar order. Each: { id: string, label: string }. */
  menus: { type: Array, default: () => [] },
  /** Accessible name for the menubar landmark. */
  ariaLabel: { type: String, default: "Application" },
  /** The bar is hosted inside another chrome row rather than owning a row of
   * its own (the slides ribbon merges it into the ribbon row), so it draws no
   * ground, no bottom border and no inline padding — the host supplies all
   * three, and the titles read as one zone of that row instead of a bar
   * stacked on it. */
  embedded: { type: Boolean, default: false },
});

const root = useTemplateRef("root");

// Each popup's UiMenuSurface anchor: the title button that opens it, so a
// pointerdown on the open menu's own title is left to the click toggle
// instead of dismiss-then-reopen. Collected per index by a template ref
// function; Array() (not []) so the array is not inferred as never[] — the
// same espree-constraint workaround family as UiSelect's Reflect.get.
const titleRefs = ref(Array());

// The untyped Array prop coerced into the typed list the template renders.
const menuList = computed(() => asMenuBarMenus(props.menus));

// Which menu is open (-1 = none) and which title holds the single tab stop.
const openIndex = ref(-1);
const activeIndex = ref(0);

function menuCount() {
  return menuList.value.length;
}

function titleEls() {
  return root.value ? Array.from(root.value.querySelectorAll("[data-menu-title]")) : [];
}

// Only the open popup carries data-menu-popup-open, so item queries never
// reach the (display:none) closed popups that keep-mounted surfaces hold in
// the DOM.
function itemEls() {
  const popup = root.value ? root.value.querySelector("[data-menu-popup-open]") : null;
  if (!popup) {
    return [];
  }
  return Array.from(
    popup.querySelectorAll('[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"]'),
  ).filter((el) => !el.hasAttribute("disabled") && el.getAttribute("aria-disabled") !== "true");
}

function focusTitle(index = 0) {
  activeIndex.value = index;
  void nextTick(() => {
    const el = titleEls()[index];
    if (el instanceof HTMLElement) {
      el.focus();
    }
  });
}

// `where` is "first", "last" or "" ("" — a pointer open — leaves focus on the
// title, matching how clicking a menu button behaves).
function openMenu(index = 0, where = "") {
  openIndex.value = index;
  activeIndex.value = index;
  if (!where) {
    return;
  }
  void nextTick(() => {
    const items = itemEls();
    items.forEach((el) => {
      if (el instanceof HTMLElement) {
        el.tabIndex = -1;
      }
    });
    const target = where === "last" ? items[items.length - 1] : items[0];
    if (target instanceof HTMLElement) {
      target.focus();
    }
  });
}

function closeMenu(returnFocus = true) {
  const wasOpen = openIndex.value;
  openIndex.value = -1;
  if (returnFocus && wasOpen >= 0) {
    focusTitle(wasOpen);
  }
}

function toggleMenu(index = 0) {
  if (openIndex.value === index) {
    closeMenu(false);
  } else {
    openMenu(index, "");
  }
  activeIndex.value = index;
}

function moveTitle(delta = 1) {
  const n = menuCount();
  if (n === 0) {
    return;
  }
  const next = (activeIndex.value + delta + n) % n;
  // With a menu open, Left/Right moves the open menu to the neighbour
  // (standard menu-bar behaviour); otherwise it only moves the tab stop.
  if (openIndex.value >= 0) {
    openMenu(next, "first");
  } else {
    focusTitle(next);
  }
}

function focusItemBy(delta = 1) {
  const items = itemEls();
  if (items.length === 0) {
    return;
  }
  const active = document.activeElement;
  const current = active ? items.indexOf(active) : -1;
  const next = current === -1 ? (delta > 0 ? 0 : items.length - 1) : (current + delta + items.length) % items.length;
  const target = items[next];
  if (target instanceof HTMLElement) {
    target.focus();
  }
}

function focusItemEdge(edge = "first") {
  const items = itemEls();
  if (items.length === 0) {
    return;
  }
  const target = edge === "first" ? items[0] : items[items.length - 1];
  if (target instanceof HTMLElement) {
    target.focus();
  }
}

function onTitleKeydown(index = 0, event = new KeyboardEvent("keydown")) {
  switch (event.key) {
    case "ArrowRight":
      event.preventDefault();
      moveTitle(1);
      break;
    case "ArrowLeft":
      event.preventDefault();
      moveTitle(-1);
      break;
    case "ArrowDown":
    case "Enter":
    case " ":
      event.preventDefault();
      openMenu(index, "first");
      break;
    case "ArrowUp":
      event.preventDefault();
      openMenu(index, "last");
      break;
    case "Escape":
      event.preventDefault();
      closeMenu(false);
      break;
    case "Home":
      event.preventDefault();
      focusTitle(0);
      break;
    case "End":
      event.preventDefault();
      focusTitle(menuCount() - 1);
      break;
    case "Tab":
      closeMenu(false);
      break;
    default:
      break;
  }
}

function onItemKeydown(event = new KeyboardEvent("keydown")) {
  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      focusItemBy(1);
      break;
    case "ArrowUp":
      event.preventDefault();
      focusItemBy(-1);
      break;
    case "Home":
      event.preventDefault();
      focusItemEdge("first");
      break;
    case "End":
      event.preventDefault();
      focusItemEdge("last");
      break;
    case "ArrowRight":
      event.preventDefault();
      moveTitle(1);
      break;
    case "ArrowLeft":
      event.preventDefault();
      moveTitle(-1);
      break;
    case "Escape":
      event.preventDefault();
      closeMenu(true);
      break;
    case "Tab":
      closeMenu(false);
      break;
    default:
      break;
  }
}

// A click on a real command closes the bar — unless that command opted out
// via data-menu-keep-open (the format painter, whose double-click-to-lock
// gesture needs the menu to stay open across the click sequence). A pointer
// click never restores focus to the title.
//
// A command is any menu-role item *or* any other button a caller put in the
// popup: a widget like the Insert menu's table-size grid can't give its cells
// a menuitem role without dragging all 100 of them into the Up/Down item ring,
// yet picking a size is still a command and must close the bar.
function onPopupClick(event = new MouseEvent("click")) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  const item = target.closest(
    '[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"],button',
  );
  if (!item || item.closest("[data-menu-keep-open]")) {
    return;
  }
  closeMenu(false);
}

function onTitleEnter(index = 0) {
  if (openIndex.value >= 0 && openIndex.value !== index) {
    openMenu(index, "");
  }
}

// Outside-click and Escape dismissal are owned by UiMenuSurface — one per
// popup — so this component no longer hand-rolls a document mousedown
// listener, and a mouse-opened menu now closes on Escape even though
// @mousedown.prevent left focus on the page (SPEC-058 AC-39).

// SPEC-068 AC-7/8 — at the compact tier the whole bar collapses into one
// "Menu" button that opens a full-height sheet listing the same menus in bar
// order, each command still authored by the caller's slot. The medium tier
// keeps the real menu bar (AC-6). The sheet is a single scrollable overlay
// rather than nested popups: at 390px a fly-out submenu has nowhere to fly, so
// every menu's items are stacked under its label and a tap on a command runs it
// and closes the sheet.
const tier = useTier();
const compact = computed(() => tier.value === "compact");
const sheetOpen = ref(false);

function openSheet() {
  sheetOpen.value = true;
}
function closeSheet() {
  sheetOpen.value = false;
}
function onSheetKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    event.preventDefault();
    closeSheet();
  }
}
// A tap on a real command inside the sheet runs it (the caller's @select fires
// on the same click) and then dismisses the sheet — mirroring onPopupClick for
// the docked bar. data-menu-keep-open opts a control out (the format painter).
function onSheetClick(event = new MouseEvent("click")) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  const item = target.closest(
    '[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"],button',
  );
  if (!item || item.closest("[data-menu-keep-open]") || item.closest("[data-menu-sheet-chrome]")) {
    return;
  }
  closeSheet();
}

// Leaving the compact tier (window widened, orientation change) must not strand
// an open sheet over a menu bar that is no longer collapsed.
watch(compact, (isCompact) => {
  if (!isCompact) {
    sheetOpen.value = false;
  }
});
onBeforeUnmount(() => {
  sheetOpen.value = false;
});
</script>

<template>
  <!-- Compact tier (SPEC-068 AC-7/8): one "Menu" button + a full-height sheet
       holding every menu's commands, stacked under their labels. -->
  <div
    v-if="compact"
    class="menu-bar menu-bar-compact"
    :class="{ 'menu-bar-embedded': embedded }"
    data-testid="menu-bar"
  >
    <button
      type="button"
      class="menu-title menu-button"
      data-testid="menu-button"
      aria-haspopup="menu"
      :aria-expanded="sheetOpen ? 'true' : 'false'"
      @click="openSheet"
    >
      Menu
    </button>
    <Teleport to="body">
      <div
        v-if="sheetOpen"
        class="menu-sheet-scrim"
        data-testid="menu-sheet-scrim"
        @click="closeSheet"
      />
      <div
        v-if="sheetOpen"
        class="menu-sheet"
        role="menu"
        :aria-label="ariaLabel"
        data-testid="menu-sheet"
        @keydown="onSheetKeydown"
        @click="onSheetClick"
      >
        <header
          class="menu-sheet-header"
          data-menu-sheet-chrome
        >
          <span class="menu-sheet-title">Menu</span>
          <button
            type="button"
            class="menu-sheet-close"
            data-testid="menu-sheet-close"
            aria-label="Close menu"
            @click="closeSheet"
          >
            ✕
          </button>
        </header>
        <section
          v-for="menu in menuList"
          :key="menu.id"
          class="menu-sheet-section"
          :data-testid="`menu-sheet-section-${menu.id}`"
        >
          <h2 class="menu-section-label">
            {{ menu.label }}
          </h2>
          <slot :name="menu.id" />
        </section>
      </div>
    </Teleport>
  </div>
  <div
    v-else
    ref="root"
    class="menu-bar"
    :class="{ 'menu-bar-embedded': embedded }"
    role="menubar"
    :aria-label="ariaLabel"
    data-testid="menu-bar"
  >
    <div
      v-for="(menu, index) in menuList"
      :key="menu.id"
      class="menu-bar-item"
      role="none"
    >
      <button
        :id="`menu-title-${menu.id}`"
        :ref="(el) => { titleRefs[index] = el }"
        type="button"
        class="menu-title"
        role="menuitem"
        data-menu-title
        :data-testid="`menu-title-${menu.id}`"
        aria-haspopup="menu"
        :aria-expanded="openIndex === index ? 'true' : 'false'"
        :aria-controls="`menu-${menu.id}`"
        :tabindex="activeIndex === index ? 0 : -1"
        @mousedown.prevent
        @click="toggleMenu(index)"
        @mouseenter="onTitleEnter(index)"
        @focus="activeIndex = index"
        @keydown="onTitleKeydown(index, $event)"
      >
        {{ menu.label }}
      </button>
      <UiMenuSurface
        :id="`menu-${menu.id}`"
        :open="openIndex === index"
        :anchor="titleRefs[index]"
        :return-focus="false"
        keep-mounted
        class="menu-popup"
        role="menu"
        :aria-label="menu.label"
        :data-testid="`menu-${menu.id}`"
        :data-menu-popup-open="openIndex === index ? '' : undefined"
        @dismiss="closeMenu(false)"
        @keydown="onItemKeydown($event)"
        @click="onPopupClick($event)"
      >
        <slot :name="menu.id" />
      </UiMenuSurface>
    </div>
  </div>
</template>

<style scoped>
.menu-bar {
  flex: none;
  display: flex;
  align-items: stretch;
  gap: 0;
  /* One stated height (AC-50): the bar never sizes to its content. */
  height: var(--height-menu-bar);
  padding: 0 var(--space-3);
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
}

/* Merged into a host row (the slides ribbon): the row owns the ground, the
   border and the inline padding, so the bar contributes only its titles. The
   stated height stays — it is what keeps the titles the same size as the
   controls beside them. */
.menu-bar-embedded {
  padding: 0;
  background: transparent;
  border-bottom: 0;
}

.menu-bar-item {
  position: relative;
  display: flex;
}

/* Toolbar Redesign — the bar's titles take the caption face: condensed,
   uppercase and tracked, so the menu bar reads as chrome rather than as body
   text sitting above the document. */
.menu-title {
  position: relative;
  border: var(--border-width) solid transparent;
  background: transparent;
  padding: var(--space-2) var(--space-4);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  color: var(--text-primary);
  cursor: pointer;
  white-space: nowrap;
}

.menu-title:hover,
.menu-title[aria-expanded="true"] {
  background: var(--state-hover);
}

/* The open menu is marked by an accent rule along the bottom of its title,
   drawn as a pseudo-element rather than an inset box-shadow (tokens.spec
   allows only var(--shadow-*) shadows) and out of flow so it cannot shift
   the title's own metrics. */
.menu-title[aria-expanded="true"] {
  color: var(--text-accent);
}

.menu-title[aria-expanded="true"]::after {
  content: "";
  position: absolute;
  inset-inline: 0;
  bottom: 0;
  height: var(--space-1);
  background: var(--accent);
}

/* Absolutely positioned so an open menu overlays the toolbar/canvas below
   rather than growing the shell. The menu bar is not an overflow-clipped
   container (unlike the toolbar strip), so no fixed-position escape is
   needed here.

   The box is UiMenuSurface's; the bar's own chrome — square top corners
   against the bar, the heavier shadow, the tighter vertical padding — must
   outrank the primitive's base box, whose selector has the same single-class
   specificity, so this one is qualified by the wrapper. */
.menu-bar-item .menu-popup {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 60;
  display: flex;
  flex-direction: column;
  min-width: 15rem;
  /* The handoff's .tb-menu: the box insets its rows rather than letting them
     run to the border, so a hovered row reads as a row and not as a band
     welded to the frame. Tighter vertically than horizontally — see the row
     rule below on what the Insert menu's height costs. */
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: 0 0 var(--radius-md) var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-3);
  /* A long menu (Insert) never runs off the bottom of the viewport: it caps
     its height and scrolls, so every command stays reachable. */
  max-height: calc(100vh - 6rem);
  overflow-y: auto;
}

/* Shared look for every command the callers place in a popup — authored as
   buttons with a menuitem role. Deep selectors because the items live in
   slotted content. */
.menu-popup :deep([role="menuitem"]),
.menu-popup :deep([role="menuitemcheckbox"]),
.menu-popup :deep([role="menuitemradio"]) {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  width: 100%;
  /* The handoff's .mi is a 30px row; this is 28. The Insert menu is 33 commands
     plus a table grid and two galleries, and SPEC-028 requires all four break
     items — which sit near its end — to be inside a 720px viewport without
     scrolling. At 30px they land 19px past it. The row is stated against the
     hit-target token rather than as a literal, so a coarse pointer still gets a
     finger-sized 38px row, and the vertical rhythm is still the handoff's:
     a fixed row height instead of the content-sized rows this replaced. */
  min-height: calc(var(--height-control) - var(--space-2));
  box-sizing: border-box;
  border: 0;
  background: transparent;
  padding: 0 var(--space-3);
  text-align: left;
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  cursor: pointer;
}

/* The active-row highlight is the background fill; focus keeps whatever ring
   the global :focus-visible rule paints on top of it — the outline is never
   removed (AC-46). */
.menu-popup :deep([role="menuitem"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitemcheckbox"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitemradio"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitem"]:focus),
.menu-popup :deep([role="menuitemcheckbox"]:focus),
.menu-popup :deep([role="menuitemradio"]:focus) {
  /* The handoff's .mi:hover is an accent wash, not the neutral --state-hover
     every other surface uses: a menu row is a command about to fire, and the
     accent is what marks intent everywhere else in the strip. Mixed from the
     palette's own --accent, so it follows every theme without a new token. */
  background: color-mix(in srgb, var(--accent) 12%, transparent);
}

.menu-popup :deep([role="menuitem"]:disabled),
.menu-popup :deep([role="menuitemcheckbox"]:disabled),
.menu-popup :deep([role="menuitemradio"]:disabled) {
  color: var(--text-disabled);
  cursor: default;
}

/* The leading gutter every command reserves, so labels line up in one column
   whatever the row carries: a check mark (MenuCheckItem), a glyph or nothing
   (MenuItem, MenuSubmenu). Before the handoff only toggles reserved it, so a
   menu of plain commands and toggles read as two ragged columns.

   The glyph is the shared 20px AppIcon scaled to the handoff's 16px — the
   menu's own measure, smaller than a toolbar button's. */
.menu-popup :deep(.menu-check),
.menu-popup :deep(.menu-icon) {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 1rem;
  flex: none;
}

.menu-popup :deep(.menu-check) {
  color: var(--accent);
}

.menu-popup :deep(.menu-icon .app-icon) {
  width: 1rem;
  height: 1rem;
}

/* The accelerator label, pushed to the trailing edge. */
.menu-popup :deep(.menu-shortcut) {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
}

.menu-popup :deep(.menu-separator) {
  margin: var(--space-1) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-subtle);
}

/* The handoff never draws a bare rule above a heading — the heading carries
   its own top rule (below). A separator that only exists to introduce the next
   section would double it, so it steps aside rather than every menu's markup
   being rewritten to drop it. */
.menu-popup :deep(.menu-separator:has(+ .menu-section-label)) {
  display: none;
}

/* The same kicker the toolbar's overflow sections use, so a section heading
   reads identically wherever a command list is grouped. Any kicker with
   something above it also carries the divider, which is how the handoff
   separates sections — one mark doing both jobs, and why no extra margin sits
   above it here. */
.menu-popup :deep(* + .menu-section-label) {
  border-top: var(--border-width) solid var(--border-subtle);
}

.menu-popup :deep(.menu-section-label) {
  padding: var(--space-2) var(--space-3) var(--space-1);
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  line-height: var(--line-tight);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  color: var(--text-accent);
}

/* ---- Compact tier: the "Menu" button + full-height sheet (SPEC-068) -------- */
.menu-bar-compact {
  align-items: center;
}

.menu-button {
  height: var(--height-control);
  padding: 0 var(--space-4);
  border-radius: var(--radius-sm);
  font-weight: var(--weight-medium);
}

/* A dim scrim behind the sheet; the sheet itself sits above it. Teleported to
   body so it covers the whole app, and pinned to the visual viewport edges so
   the on-screen keyboard cannot push it out of reach (SPEC-068). */
.menu-sheet-scrim {
  position: fixed;
  inset: 0;
  z-index: 90;
  background: var(--surface-overlay);
}

.menu-sheet {
  position: fixed;
  inset: 0;
  z-index: 91;
  display: flex;
  flex-direction: column;
  background: var(--surface-base);
  color: var(--text-primary);
  overflow-y: auto;
  overscroll-behavior: contain;
}

.menu-sheet-header {
  position: sticky;
  top: 0;
  z-index: 1;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: var(--space-3) var(--space-5);
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
}

.menu-sheet-title {
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
}

.menu-sheet-close {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: var(--height-control);
  height: var(--height-control);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: inherit;
  cursor: pointer;
  font-size: var(--text-lg);
}

.menu-sheet-close:hover {
  background: var(--state-hover);
}

.menu-sheet-section {
  padding: var(--space-2) 0;
  border-bottom: var(--border-width) solid var(--border-subtle);
}

/* The sheet reuses the popup command chrome so a caller's slotted items look
   and behave identically whether they render docked or in the sheet. */
.menu-sheet :deep([role="menuitem"]),
.menu-sheet :deep([role="menuitemcheckbox"]),
.menu-sheet :deep([role="menuitemradio"]) {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  width: 100%;
  box-sizing: border-box;
  min-height: var(--height-control);
  border: 0;
  background: transparent;
  padding: var(--space-2) var(--space-5);
  text-align: left;
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  cursor: pointer;
}

.menu-sheet :deep([role="menuitem"]:hover:not(:disabled)),
.menu-sheet :deep([role="menuitemcheckbox"]:hover:not(:disabled)),
.menu-sheet :deep([role="menuitemradio"]:hover:not(:disabled)),
.menu-sheet :deep([role="menuitem"]:focus),
.menu-sheet :deep([role="menuitemcheckbox"]:focus),
.menu-sheet :deep([role="menuitemradio"]:focus) {
  background: color-mix(in srgb, var(--accent) 12%, transparent);
}

.menu-sheet :deep([role="menuitem"]:disabled),
.menu-sheet :deep([role="menuitemcheckbox"]:disabled),
.menu-sheet :deep([role="menuitemradio"]:disabled) {
  color: var(--text-disabled);
  cursor: default;
}

.menu-sheet :deep(.menu-check),
.menu-sheet :deep(.menu-icon) {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 1rem;
  flex: none;
}

.menu-sheet :deep(.menu-check) {
  color: var(--accent);
}

.menu-sheet :deep(.menu-icon .app-icon) {
  width: 1rem;
  height: 1rem;
}

.menu-sheet :deep(.menu-shortcut) {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
}

.menu-sheet :deep(.menu-separator) {
  margin: var(--space-1) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-subtle);
}

.menu-sheet :deep(.menu-separator:has(+ .menu-section-label)) {
  display: none;
}

/* The sheet's headings wear the same kicker as the docked popup's — the two
   tiers are one menu shown two ways, so they must not read as two designs. */
.menu-sheet :deep(* + .menu-section-label) {
  margin-top: var(--space-1);
  border-top: var(--border-width) solid var(--border-subtle);
}

.menu-sheet :deep(.menu-section-label) {
  padding: var(--space-3) var(--space-5) var(--space-1);
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  line-height: var(--line-tight);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  color: var(--text-accent);
}
</style>
