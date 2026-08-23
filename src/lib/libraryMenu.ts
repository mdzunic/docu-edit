// The shape of a RowMenu overflow-menu item (SPEC-067), plus a coercion for
// the runtime-typed (`Array`) prop the .vue component receives — the .vue
// ESLint parser bars the generic/PropType syntax that would type it in place,
// so the cast lives here in plain TypeScript (see CommandPalette's
// asCommandMetas for the same pattern).
export interface LibraryMenuItem {
  id: string;
  label: string;
  testid: string;
  danger?: boolean;
}

export function asMenuItems(items: unknown): LibraryMenuItem[] {
  return Array.isArray(items) ? (items as LibraryMenuItem[]) : [];
}
