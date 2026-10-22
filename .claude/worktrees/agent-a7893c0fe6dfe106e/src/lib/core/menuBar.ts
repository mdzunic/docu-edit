// Shape helpers for the application menu bar (SPEC-049). The MenuBar
// component takes its menu list as a runtime `Array` prop (the espree
// constraint on .vue scripts bars a typed PropType there — see AppShell.vue),
// so it coerces the untyped array through here, mirroring how StylePicker
// runs its `entries` prop through `asStyleEntries`.

/** One menu in the bar: a stable id (used for slot names and test ids) and
 * the visible title. */
export interface MenuBarMenu {
  id: string;
  label: string;
}

/** Coerces the untyped `menus` prop into the typed list the template needs. */
export function asMenuBarMenus(value: unknown[]): MenuBarMenu[] {
  return value.map((entry) => {
    const menu = (entry ?? {}) as Partial<MenuBarMenu>;
    return { id: String(menu.id ?? ""), label: String(menu.label ?? "") };
  });
}
