<script setup lang="ts">
// SPEC-133's table surfaces in the workbook body, bound to useSheetTables (the
// grid context's `tables`): the design bar (TableDesignBar) and the ODS notice
// (`sheet-table-ods-notice`, once per session, dismissable), both in flow
// BELOW the grid — WorkbookBody places this between the sheet and its tabs, so
// neither ever moves a cell under the pointer — and the modal Create Table
// (`sheet-create-table`) and Resize Table (`sheet-resize-table`) dialogs, kept
// mounted while their range is pointed at, and the convert confirmation
// (`sheet-table-convert-dialog`, OK `sheet-table-convert-confirm`). Extracted
// so WorkbookBody only places it. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { ODS_TABLE_NOTICE } from "../../views/spreadsheet/tableDesign";
import { asSheetTables } from "../../views/spreadsheet/useSheetTables";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import UiIconButton from "../ui/UiIconButton.vue";
import CreateTableDialog from "./CreateTableDialog.vue";
import ResizeTableDialog from "./ResizeTableDialog.vue";
import TableDesignBar from "./TableDesignBar.vue";

const props = defineProps({
  /** useSheetTables' return. */
  tables: { type: Object, required: true },
});

const t = computed(() => asSheetTables(props.tables));
const createDialog = computed(() => {
  const open = t.value.dialog.value;
  return open?.kind === "create" ? open : null;
});
const resizeDialog = computed(() => {
  const open = t.value.dialog.value;
  return open?.kind === "resize" ? open : null;
});

function create(text = "", headers = true) {
  t.value.confirmCreate(text, headers);
}

function resize(text = "") {
  t.value.confirmResize(text);
}
</script>

<template>
  <TableDesignBar :tables="props.tables" />
  <div
    v-if="t.odsNotice.value"
    class="table-ods-notice"
    role="status"
    data-testid="sheet-table-ods-notice"
  >
    <span>{{ ODS_TABLE_NOTICE }}</span>
    <UiIconButton
      class="table-ods-dismiss"
      icon="close"
      label="Dismiss"
      data-testid="sheet-table-ods-notice-dismiss"
      @click="t.dismissOdsNotice()"
    />
  </div>
  <CreateTableDialog
    v-if="createDialog"
    :range="createDialog.range"
    :headers="createDialog.headers"
    :error="t.dialogError.value"
    :picking="t.picking.value"
    :selection-text="t.selectionText.value"
    @ok="create"
    @cancel="t.closeDialog()"
    @pick="t.startPick()"
    @pick-done="t.endPick()"
  />
  <ResizeTableDialog
    v-if="resizeDialog"
    :range="resizeDialog.range"
    :error="t.dialogError.value"
    :picking="t.picking.value"
    :selection-text="t.selectionText.value"
    @ok="resize"
    @cancel="t.closeDialog()"
    @pick="t.startPick()"
    @pick-done="t.endPick()"
  />
  <BaseDialog
    v-if="t.converting.value"
    title="Convert to range"
    testid="sheet-table-convert-dialog"
    width="24rem"
    @dismiss="t.cancelConvert()"
  >
    <p class="convert-message">
      Convert the table "{{ t.converting.value.name }}" to a normal range? Its formatting stays, and formulas that refer to it use cell references instead.
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-table-convert-cancel"
        autofocus
        @click="t.cancelConvert()"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="sheet-table-convert-confirm"
        @click="t.confirmConvert()"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.table-ods-notice {
  flex: none;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  background: var(--surface-selected);
  border-top: var(--border-width) solid var(--accent-subtle);
  color: var(--text-accent);
  font-size: var(--text-sm);
}

.table-ods-dismiss {
  margin-left: auto;
}

.convert-message {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-md);
  overflow-wrap: anywhere;
}
</style>
