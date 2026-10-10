// The sheet tab context menu's item shape (SPEC-121 "UI"): SheetTabMenu takes
// its items as a runtime `Array` prop (the espree constraint on .vue scripts
// bars a typed PropType there — see AppShell.vue), so it coerces them through
// here. Framework-free. Item test ids are `sheet-tab-menu-<action>` —
// distinct from the menu bar's `sheet-menu-*` and SPEC-124's
// `sheet-delete-dialog`.

/** One command in the tab menu. An item that cannot apply is `disabled` with
 * the stated `reason` (shown as its title and inline) — never hidden. */
export interface SheetTabMenuItem {
  action: string;
  label: string;
  testid: string;
  disabled?: boolean;
  reason?: string;
  /** Destructive (Delete): drawn in the danger ink, as MenuItem does. */
  danger?: boolean;
  /** Draw a separator above this item, to group the commands. */
  separatorBefore?: boolean;
}

/** Coerces the untyped `items` prop into the typed list the menu renders. */
export function asSheetTabMenuItems(value: unknown): SheetTabMenuItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((entry) => {
    const item = (entry ?? {}) as Partial<SheetTabMenuItem>;
    return {
      action: String(item.action ?? ""),
      label: String(item.label ?? ""),
      testid: String(item.testid ?? ""),
      disabled: item.disabled === true,
      reason: typeof item.reason === "string" ? item.reason : "",
      danger: item.danger === true,
      separatorBefore: item.separatorBefore === true,
    };
  });
}
