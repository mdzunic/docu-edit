<script setup lang="ts">
// SPEC-136's surfaces over the sheet panel, from the grid context's
// `printing`: the print preview (SheetPrintPages in the grid area beside the
// reused PrintPreviewDialog — its options slot the scope select, its actions
// slot "Page setup…" and "Download PDF"), the off-screen 1:1 pages the
// browser's print and the PDF export mount, the export's progress and Stop,
// a refusal's notice, and the Page Setup dialog. Screen-only affordances
// carry `.print-exclude`. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import PrintPreviewDialog from "../PrintPreviewDialog.vue";
import SheetPageSetupDialog from "./SheetPageSetupDialog.vue";
import SheetPrintPages from "./SheetPrintPages.vue";
import { asSheetPrinting, PRINT_SCOPES } from "../../views/spreadsheet/useSheetPrint";

const props = defineProps({
  /** The grid context's SheetPrinting. */
  printing: { type: Object, required: true },
  theme: { type: Object, default: null },
  imageSrc: { type: Function, default: undefined },
});

const p = computed(() => asSheetPrinting(props.printing));
const preview = computed(() => p.value.print.job.value.pages.map((page, index) => ({ page, index })));
const offscreen = computed(() => (p.value.print.printing.value ? p.value.print.printPages.value : p.value.pdf.sheetPages.value));
const notice = computed(() => p.value.print.notice.value || p.value.pdf.notice.value);

function scopeChanged(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    p.value.print.setScope(target.value);
  }
}
</script>

<template>
  <SheetPrintPages
    v-if="p.print.isOpen.value"
    class="print-exclude-scale"
    mode="preview"
    :pages="preview"
    :total="p.print.pageCount.value"
    :layout-for="p.print.layoutFor"
    :font-family="p.print.fontFamily.value"
    :theme="theme"
    :image-src="imageSrc"
  />
  <PrintPreviewDialog
    v-if="p.print.isOpen.value"
    :page-count="p.print.pageCount.value"
    :range-mode="p.print.rangeMode.value"
    :range-text="p.print.rangeText.value"
    :error="p.print.rangeError.value"
    :can-print="p.print.canPrint.value"
    :empty-message="p.print.emptyMessage"
    :show-background-option="false"
    @update:range-mode="(mode) => (p.print.rangeMode.value = mode === 'custom' ? 'custom' : 'all')"
    @update:range-text="(text) => (p.print.rangeText.value = String(text))"
    @print="p.print.print()"
    @close="p.print.close()"
  >
    <template #options>
      <label class="scope">
        Print
        <select
          data-testid="sheet-print-scope"
          :value="p.print.scope.value"
          @change="scopeChanged"
        >
          <option
            v-for="scope in PRINT_SCOPES"
            :key="scope.value"
            :value="scope.value"
          >{{ scope.label }}</option>
        </select>
      </label>
    </template>
    <template #actions>
      <button
        type="button"
        class="action"
        data-testid="sheet-print-page-setup"
        @click="p.print.openPageSetup()"
      >
        Page setup…
      </button>
      <button
        type="button"
        class="action"
        data-testid="sheet-print-download-pdf"
        :disabled="p.print.pageCount.value === 0 || p.pdf.busy.value"
        @click="p.pdf.exportPdf()"
      >
        Download PDF
      </button>
    </template>
  </PrintPreviewDialog>
  <div
    v-if="p.pdf.busy.value || notice"
    class="print-chip print-exclude"
  >
    <span
      v-if="p.pdf.busy.value"
      role="status"
      data-testid="sheet-print-progress"
    >Page {{ p.pdf.progress.value.done }} of {{ p.pdf.progress.value.total }}</span>
    <button
      v-if="p.pdf.busy.value"
      type="button"
      data-testid="sheet-print-abort"
      @click="p.pdf.abort()"
    >
      Stop
    </button>
    <span
      v-if="notice"
      role="alert"
      data-testid="sheet-print-notice"
    >{{ notice }}</span>
  </div>
  <SheetPrintPages
    v-if="offscreen.length > 0"
    mode="sheet"
    :pages="offscreen"
    :total="p.print.printing.value ? p.print.printPages.value.length : p.pdf.progress.value.total"
    :layout-for="p.print.layoutFor"
    :page-names="p.print.printing.value ? p.print.printPageNames.value : []"
    :font-family="p.print.fontFamily.value"
    :theme="theme"
    :image-src="imageSrc"
  />
  <SheetPageSetupDialog :dialog="p.pageSetup.dialog" />
</template>

<style scoped>
.scope {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-top: var(--space-4);
}

.action {
  font: inherit;
}

.print-chip {
  position: absolute;
  left: 50%;
  bottom: var(--space-5);
  transform: translateX(-50%);
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-5);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  z-index: 30;
}
</style>
