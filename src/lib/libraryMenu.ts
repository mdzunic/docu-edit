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

/** The library formats that are images — the entries SPEC-108's "Convert
 * to PDF…" applies to. */
const IMAGE_FORMATS: ReadonlySet<string> = new Set(["png", "jpeg", "gif", "svg", "webp"]);

export function isImageFormat(format: string | undefined): boolean {
  return format !== undefined && IMAGE_FORMATS.has(format);
}

/** SPEC-108's row action, offered on image entries only. */
export const CONVERT_TO_PDF_ITEM: LibraryMenuItem = {
  id: "convert-pdf",
  label: "Convert to PDF…",
  testid: "menu-convert-pdf",
};

/** A library row/card's overflow menu (SPEC-067) — shared by the list view,
 * the grid cards and the keyboard paths so the actions stay identical. An
 * image entry (by its `format`) also gets "Convert to PDF…" (SPEC-108). */
export function libraryRowMenuItems(starred: boolean, format?: string): LibraryMenuItem[] {
  return [
    { id: "star", label: starred ? "Unstar" : "Star", testid: "menu-star" },
    { id: "rename", label: "Rename", testid: "menu-rename" },
    { id: "duplicate", label: "Duplicate", testid: "menu-duplicate" },
    ...(isImageFormat(format) ? [CONVERT_TO_PDF_ITEM] : []),
    { id: "delete", label: "Delete", testid: "menu-delete", danger: true },
  ];
}

/** A trashed row's overflow menu (SPEC-067). */
export const TRASH_MENU_ITEMS: LibraryMenuItem[] = [
  { id: "restore", label: "Restore", testid: "menu-restore" },
  { id: "delete-permanently", label: "Delete permanently", testid: "menu-delete-permanently", danger: true },
];
