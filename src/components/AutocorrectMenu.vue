<script setup lang="ts">
import { onMounted, onUnmounted, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// The autocorrect settings menu (SPEC-026): toggles the smart-quote and dash
// substitutions independently. The parent owns the persisted settings and
// re-emits them here as props; this component only reflects and toggles them.
// Runtime prop/emit declarations per the espree constraint (see DownloadMenu).
const props = defineProps({
  smartQuotes: { type: Boolean, default: true },
  dashes: { type: Boolean, default: true },
});

const emit = defineEmits(["toggle-smart-quotes", "toggle-dashes"]);

const open = ref(false);
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

// A fixed-position popup, like the toolbar's other menus: .toolbar clips on
// overflow-y, so an absolutely-positioned list here is painted behind the
// strip's edge instead of over the canvas.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
});

function closeMenu() {
  open.value = false;
}

onMounted(() => document.addEventListener("click", closeMenu));
onUnmounted(() => document.removeEventListener("click", closeMenu));
</script>

<template>
  <div
    class="autocorrect"
    @click.stop
    @keydown.escape="closeMenu"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="autocorrect-menu"
      aria-haspopup="menu"
      :aria-expanded="open"
      title="Autocorrect"
      @mousedown.prevent
      @click="open = !open"
    >
      Autocorrect
    </button>
    <div
      v-if="open"
      ref="menu"
      class="autocorrect-list"
      role="menu"
      aria-label="Autocorrect settings"
      data-testid="autocorrect-menu-list"
      :style="menuStyle"
    >
      <button
        type="button"
        role="menuitemcheckbox"
        :aria-checked="props.smartQuotes"
        data-testid="autocorrect-smart-quotes"
        @click="emit('toggle-smart-quotes')"
      >
        <span class="check">{{ props.smartQuotes ? "✓" : "" }}</span>
        Smart quotes
      </button>
      <button
        type="button"
        role="menuitemcheckbox"
        :aria-checked="props.dashes"
        data-testid="autocorrect-dashes"
        @click="emit('toggle-dashes')"
      >
        <span class="check">{{ props.dashes ? "✓" : "" }}</span>
        Dash substitution
      </button>
    </div>
  </div>
</template>

<style scoped>
.autocorrect {
  display: inline-flex;
}

/* position/left/top come inline from useFloatingMenu's menuStyle. */
.autocorrect-list {
  z-index: 30;
  min-width: 12rem;
  display: flex;
  flex-direction: column;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  box-shadow: 0 2px 6px rgb(0 0 0 / 0.2);
  padding: 0.25rem 0;
}

.autocorrect-list button {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.autocorrect-list button:hover {
  background: #f1f3f4;
}

.check {
  display: inline-block;
  width: 1rem;
  color: #1a73e8;
}
</style>
