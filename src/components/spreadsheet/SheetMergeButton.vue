<script setup lang="ts">
// The toolbar's merge split button (SPEC-124 "Merges"), in SPEC-123's
// band-6 slot: the main half (`sheet-merge`) is Merge & Centre — pressed
// when the active cell is merged, its action then Unmerge — and the arrow
// (`sheet-merge-menu-button`) opens `sheet-merge-menu`: Merge & Centre
// (`sheet-merge-centre`), Merge Across (`sheet-merge-across`), Merge Cells
// (`sheet-merge-cells`) and Unmerge (`sheet-unmerge`). Everything runs
// through useMerges, and every part shows disabled with it (SPEC-138: while
// the workbook is read-only too). Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { asMerges } from "../../views/spreadsheet/useMerges";
import AppIcon from "../AppIcon.vue";
import SheetMenuButton from "./SheetMenuButton.vue";

const props = defineProps({
  /** useMerges' return. */
  merges: { type: Object, required: true },
});

const merges = computed(() => asMerges(props.merges));

const MODES = [
  { mode: "centre", label: "Merge & Centre", testid: "sheet-merge-centre" },
  { mode: "across", label: "Merge Across", testid: "sheet-merge-across" },
  { mode: "cells", label: "Merge Cells", testid: "sheet-merge-cells" },
];

function merge(mode = "centre") {
  merges.value.merge(mode === "across" ? "across" : mode === "cells" ? "cells" : "centre");
}
</script>

<template>
  <span class="split">
    <button
      type="button"
      class="main"
      data-testid="sheet-merge"
      :aria-label="merges.merged.value ? 'Unmerge cells' : 'Merge & Centre'"
      :title="merges.merged.value ? 'Unmerge cells' : 'Merge & Centre'"
      :aria-pressed="merges.merged.value ? 'true' : 'false'"
      :class="{ 'is-pressed': merges.merged.value }"
      :disabled="merges.disabled.value"
      @mousedown.prevent
      @click="merges.primary()"
    >
      <AppIcon name="merge-cells" />
    </button>
    <SheetMenuButton
      v-slot="{ close }"
      label="Merge options"
      testid="sheet-merge-menu-button"
      menu-testid="sheet-merge-menu"
      :disabled="merges.disabled.value"
    >
      <button
        v-for="item in MODES"
        :key="item.mode"
        type="button"
        role="menuitem"
        :data-testid="item.testid"
        :disabled="merges.disabled.value"
        @mousedown.prevent
        @click="merge(item.mode); close()"
      >
        {{ item.label }}
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="sheet-unmerge"
        :disabled="merges.disabled.value"
        @mousedown.prevent
        @click="merges.unmerge(); close()"
      >
        Unmerge
      </button>
    </SheetMenuButton>
  </span>
</template>

<style scoped>
.split {
  display: inline-flex;
  align-items: center;
}

.main {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: var(--control-height, var(--height-control));
  height: var(--control-height, var(--height-control));
  padding: 0;
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.main:not(:disabled):hover {
  background: var(--state-hover);
}

.main.is-pressed {
  background: var(--state-active);
  border-color: var(--border-default);
}

.main:disabled {
  opacity: 0.5;
  cursor: default;
}

.main:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}
</style>
