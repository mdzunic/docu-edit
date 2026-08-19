<script setup lang="ts">
import { computed, nextTick, ref, useTemplateRef } from "vue";
import { asMenuBarMenus } from "../lib/core/menuBar";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

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
</script>

<template>
  <div
    ref="root"
    class="menu-bar"
    role="menubar"
    :aria-label="ariaLabel"
    data-testid="menu-bar"
  >
    <div
      v-for="(menu, index) in menuList"
      :key="menu.id"
      class="menu-bar-item"
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

.menu-bar-item {
  position: relative;
  display: flex;
}

.menu-title {
  border: var(--border-width) solid transparent;
  background: transparent;
  padding: var(--space-2) var(--space-4);
  font-size: var(--text-md);
  color: var(--text-primary);
  cursor: pointer;
  white-space: nowrap;
}

.menu-title:hover,
.menu-title[aria-expanded="true"] {
  background: var(--state-hover);
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
  padding: var(--space-1) 0;
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
  box-sizing: border-box;
  border: 0;
  background: transparent;
  padding: var(--space-2) var(--space-5);
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
  background: var(--state-hover);
}

.menu-popup :deep([role="menuitem"]:disabled),
.menu-popup :deep([role="menuitemcheckbox"]:disabled),
.menu-popup :deep([role="menuitemradio"]:disabled) {
  color: var(--text-disabled);
  cursor: default;
}

/* The leading gutter each command reserves for a check mark, so labels line
   up whether or not the row is a toggle. */
.menu-popup :deep(.menu-check) {
  display: inline-block;
  width: 1rem;
  flex: none;
  color: var(--accent);
}

/* The accelerator label, pushed to the trailing edge. */
.menu-popup :deep(.menu-shortcut) {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.menu-popup :deep(.menu-separator) {
  margin: var(--space-1) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-subtle);
}

.menu-popup :deep(.menu-section-label) {
  padding: var(--space-1) var(--space-5) var(--space-1);
  font-size: var(--text-xs);
  line-height: var(--line-tight);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}
</style>
