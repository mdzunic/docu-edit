<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, useTemplateRef } from "vue";
import { asMenuBarMenus } from "../lib/menuBar";

// The application menu bar (SPEC-049). This component is deliberately
// *structural only*: it renders the menu titles and their popups, and owns
// the menubar's keyboard model, single tab stop, pointer cross-open and
// open/close state. The contents of each popup come from a slot named after
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
// reach the (display:none) closed popups that v-show keeps in the DOM.
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
function onPopupClick(event = new MouseEvent("click")) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  const item = target.closest('[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"]');
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

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (
    openIndex.value >= 0 &&
    root.value &&
    event.target instanceof Node &&
    !root.value.contains(event.target)
  ) {
    closeMenu(false);
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
      <div
        v-show="openIndex === index"
        :id="`menu-${menu.id}`"
        class="menu-popup"
        role="menu"
        :aria-label="menu.label"
        :data-testid="`menu-${menu.id}`"
        :data-menu-popup-open="openIndex === index ? '' : undefined"
        @keydown="onItemKeydown($event)"
        @click="onPopupClick($event)"
      >
        <slot :name="menu.id" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.menu-bar {
  flex: none;
  display: flex;
  align-items: stretch;
  gap: 0;
  padding: 0 0.5rem;
  background: #fafafa;
  border-bottom: 1px solid #e0e0e0;
}

.menu-bar-item {
  position: relative;
  display: flex;
}

.menu-title {
  border: 1px solid transparent;
  background: transparent;
  padding: 0.3rem 0.75rem;
  font-size: 0.85rem;
  color: #3c4043;
  cursor: pointer;
  white-space: nowrap;
}

.menu-title:hover,
.menu-title[aria-expanded="true"] {
  background: #eceff1;
}

.menu-title:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: -2px;
}

/* Absolutely positioned so an open menu overlays the toolbar/canvas below
   rather than growing the shell. The menu bar is not an overflow-clipped
   container (unlike the toolbar strip), so no fixed-position escape is
   needed here. */
.menu-popup {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 60;
  display: flex;
  flex-direction: column;
  min-width: 15rem;
  padding: 0.25rem 0;
  border: 1px solid #dadce0;
  border-radius: 0 0 6px 6px;
  background: #fff;
  box-shadow: 0 4px 12px rgb(0 0 0 / 0.18);
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
  gap: 0.75rem;
  width: 100%;
  box-sizing: border-box;
  border: 0;
  background: transparent;
  padding: 0.4rem 1rem;
  text-align: left;
  font: inherit;
  font-size: 0.85rem;
  color: #202124;
  cursor: pointer;
}

.menu-popup :deep([role="menuitem"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitemcheckbox"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitemradio"]:hover:not(:disabled)),
.menu-popup :deep([role="menuitem"]:focus),
.menu-popup :deep([role="menuitemcheckbox"]:focus),
.menu-popup :deep([role="menuitemradio"]:focus) {
  background: #f1f3f4;
  outline: none;
}

.menu-popup :deep([role="menuitem"]:disabled),
.menu-popup :deep([role="menuitemcheckbox"]:disabled),
.menu-popup :deep([role="menuitemradio"]:disabled) {
  color: #9aa0a6;
  cursor: default;
}

/* The leading gutter each command reserves for a check mark, so labels line
   up whether or not the row is a toggle. */
.menu-popup :deep(.menu-check) {
  display: inline-block;
  width: 1rem;
  flex: none;
  color: #1a73e8;
}

/* The accelerator label, pushed to the trailing edge. */
.menu-popup :deep(.menu-shortcut) {
  margin-left: auto;
  padding-left: 1.5rem;
  color: #80868b;
  font-size: 0.78rem;
}

.menu-popup :deep(.menu-separator) {
  margin: 0.25rem 0;
  border: none;
  border-top: 1px solid #e8eaed;
}

.menu-popup :deep(.menu-section-label) {
  padding: 0.25rem 1rem 0.1rem;
  font-size: 0.7rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #80868b;
}
</style>
