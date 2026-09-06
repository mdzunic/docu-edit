<script setup lang="ts">
import { nextTick, onBeforeUnmount, ref, useTemplateRef, watch } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import AppIcon from "./AppIcon.vue";
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
  /** The row's leading glyph — see MenuItem's note on the 16px gutter. */
  icon: { type: String, default: "" },
});

const open = ref(false);
const trigger = useTemplateRef("trigger");
const flyout = useTemplateRef("flyout");

// The parent popup caps its height and scrolls (overflow-y: auto), which also
// clips horizontal overflow — an absolutely-positioned flyout at left:100%
// lands entirely inside the clipped region and is invisible. The same
// fixed-position escape the toolbar pickers use frees it; "right" places it
// beside the trigger, top-aligned, flipping sides at the viewport edge.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => trigger.value,
  menu: () => (flyout.value && flyout.value.$el instanceof HTMLElement ? flyout.value.$el : null),
  placement: "right",
  offsetY: 0,
});

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

// The parent popup can close without the flyout ever being dismissed: a
// command elsewhere in it, an Escape, or a hover cross-open to a neighbouring
// menu all hide the popup (v-show) while this open state stays true — and the
// flyout, now unclipped, would be back on screen the next time the menu
// opens. So while open, watch the parent surface's inline style and close the
// moment it is hidden.
const root = useTemplateRef("root");

// One observer for the component's life (no TS-only nullable-let syntax — see
// the espree constraint above); observe/disconnect toggle with the open state.
const popupObserver =
  typeof MutationObserver === "undefined"
    ? null
    : new MutationObserver((mutations) => {
        const popup = mutations[0] ? mutations[0].target : null;
        if (popup instanceof HTMLElement && popup.style.display === "none") {
          closeMenu(false);
        }
      });

function observePopup() {
  const popup =
    root.value instanceof HTMLElement ? root.value.closest(".ui-menu-surface") : null;
  if (popupObserver && popup instanceof HTMLElement) {
    popupObserver.observe(popup, { attributes: true, attributeFilter: ["style"] });
  }
}

function unobservePopup() {
  if (popupObserver) {
    popupObserver.disconnect();
  }
}

watch(open, (isOpen) => {
  if (isOpen) {
    observePopup();
  } else {
    unobservePopup();
  }
});

onBeforeUnmount(unobservePopup);

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

// A click inside the flyout is a real command and closes it (and, bubbling on,
// the whole bar). But a nested submenu's own trigger carries
// data-menu-keep-open: this flyout can host another MenuSubmenu (Format →
// Drawing → Align), and clicking that child trigger must open its flyout, not
// collapse this one. Mirror MenuBar.onPopupClick and leave this open for any
// keep-open target — the child's flyout owns its own dismissal.
function onFlyoutClick(event = new MouseEvent("click")) {
  const target = event.target;
  if (target instanceof Element && target.closest("[data-menu-keep-open]")) {
    return;
  }
  open.value = false;
}
</script>

<template>
  <div
    ref="root"
    class="menu-submenu"
  >
    <!-- Only the trigger keeps the bar open (opening the flyout must not close
         it); a click on a choice inside the flyout is a real command and closes
         the whole bar, exactly like any other menu command. -->
    <button
      ref="trigger"
      type="button"
      role="menuitem"
      class="menu-item menu-submenu-trigger"
      :data-testid="testid || undefined"
      data-menu-keep-open
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      tabindex="-1"
      @mousedown.prevent
      @click="open ? closeMenu(false) : openMenu(false)"
      @keydown="onTriggerKeydown"
    >
      <span
        class="menu-icon"
        aria-hidden="true"
      ><AppIcon
        v-if="icon"
        :name="icon"
      /></span>
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
      :style="menuStyle"
      class="menu-submenu-flyout"
      role="menu"
      :aria-label="label"
      :data-testid="testid ? `${testid}-flyout` : undefined"
      @dismiss="closeMenu(false)"
      @keydown="onFlyoutKeydown"
      @click="onFlyoutClick"
    >
      <slot />
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.menu-submenu {
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

/* The flyout opens beside the trigger. Its position is the fixed-position
   escape useFloatingMenu computes: the parent popup is a scroll container
   (overflow-y: auto against tall menus), and anything absolutely positioned
   past its edge — the old left:100% — is clipped invisible. */
.menu-submenu-flyout {
  z-index: 61;
  display: flex;
  flex-direction: column;
  min-width: 14rem;
  padding: var(--space-1) 0;
}
</style>
