<script setup lang="ts">
import { nextTick, ref, useTemplateRef } from "vue";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// A nested submenu inside a MenuBar popup (SPEC-059): a role="menuitem" that
// opens a flyout of its own menu items to the side — used for View →
// Appearance (AC-48/49). The flyout box, its outside-click and Escape
// dismissal are UiMenuSurface's; this component owns only the open state and
// the pointer/keyboard gestures that toggle it.
//
// The whole submenu carries data-menu-keep-open so MenuBar's popup-click
// handler leaves the bar open when the trigger — or a choice inside the
// flyout — is clicked (the parent bar closes on outside pointerdown instead).
// keydown for the keys handled here is stopped so MenuBar's own item-ring
// navigation (which reads ArrowLeft/Right as "move to the neighbouring menu")
// does not also fire.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
const props = defineProps({
  label: { type: String, required: true },
  testid: { type: String, default: "" },
});

const open = ref(false);
const trigger = useTemplateRef("trigger");
const flyout = useTemplateRef("flyout");

function focusFirstItem() {
  void nextTick(() => {
    const root = flyout.value && flyout.value.$el instanceof HTMLElement ? flyout.value.$el : null;
    const first = root
      ? root.querySelector(
          '[role="menuitemradio"]:not([disabled]),[role="menuitemcheckbox"]:not([disabled]),[role="menuitem"]:not([disabled])',
        )
      : null;
    if (first instanceof HTMLElement) {
      first.focus();
    }
  });
}

function openMenu(focus = false) {
  open.value = true;
  if (focus) {
    focusFirstItem();
  }
}

function closeMenu(returnFocus = true) {
  open.value = false;
  if (returnFocus && trigger.value instanceof HTMLElement) {
    trigger.value.focus();
  }
}

function onTriggerKeydown(event = new KeyboardEvent("keydown")) {
  switch (event.key) {
    case "ArrowRight":
    case "Enter":
    case " ":
      event.preventDefault();
      event.stopPropagation();
      openMenu(true);
      break;
    case "ArrowLeft":
    case "Escape":
      if (open.value) {
        event.preventDefault();
        event.stopPropagation();
        closeMenu();
      }
      break;
    default:
      break;
  }
}

function onFlyoutKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "ArrowLeft" || event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    closeMenu();
  }
}
</script>

<template>
  <div
    class="menu-submenu"
    data-menu-keep-open
  >
    <button
      ref="trigger"
      type="button"
      role="menuitem"
      class="menu-item menu-submenu-trigger"
      :data-testid="testid || undefined"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      tabindex="-1"
      @mousedown.prevent
      @click="open ? closeMenu(false) : openMenu(false)"
      @keydown="onTriggerKeydown"
    >
      <span
        class="menu-check"
        aria-hidden="true"
      />
      <span class="menu-item-label">{{ label }}</span>
      <span
        class="menu-submenu-arrow"
        aria-hidden="true"
      >▸</span>
    </button>
    <UiMenuSurface
      ref="flyout"
      :open="open"
      :anchor="trigger"
      class="menu-submenu-flyout"
      role="menu"
      :aria-label="label"
      :data-testid="testid ? `${testid}-flyout` : undefined"
      @dismiss="closeMenu(false)"
      @keydown="onFlyoutKeydown"
    >
      <slot />
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.menu-submenu {
  position: relative;
  display: flex;
}

.menu-submenu-trigger {
  width: 100%;
}

.menu-submenu-arrow {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

/* The flyout opens to the trailing edge of the trigger, overlapping the parent
   popup's right border by one step so the two read as one connected surface. */
.menu-submenu-flyout {
  position: absolute;
  top: 0;
  left: 100%;
  z-index: 61;
  display: flex;
  flex-direction: column;
  min-width: 14rem;
  padding: var(--space-1) 0;
}
</style>
