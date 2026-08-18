---
id: SPEC-062
title: Command palette
status: draft
area: shell
depends_on: [SPEC-021, SPEC-049, SPEC-058]
---
## Why
SPEC-021 put a command palette out of scope when the product had a dozen
commands. It now has **82**, in the `editorCommands` Map in
`src/views/DocumentView.vue` — the single registry both the menu bar and
the keyboard path dispatch through — spread across six menus, several
submenus and seven panes. Finding "Check accessibility" means knowing it
is under Tools; finding "Convert {{…}} text to merge fields" means knowing
it exists at all. Every product at this feature depth has a search-the-
commands affordance for exactly this reason: Word's Search box, Google
Docs' Ctrl+/ menu search, LibreOffice's Tools ▸ Search Commands.

The registry that makes it cheap is already there, and already
disciplined: one id per command, dispatch through one Map, labels and
shortcut bindings in `src/lib/shortcuts.ts`. What is missing is a
user-facing name and a category for the menu-only ids — the palette needs
to *show* "Insert → Footnote", and today that string exists only as markup
inside `DocumentMenuBar.vue`. Fixing that is most of this spec's real
work, and it pays off past the palette: it is the same table SPEC-068's
small-screen layout and any future "what can this app do" surface need.

## Acceptance criteria
- [ ] Ctrl/Cmd+K opens the palette from anywhere in the editor, registered in `src/lib/shortcuts.ts` as `app.commandPalette` with category "View" and appearing in the shortcut reference dialog
- [ ] `Help → Search commands` — or, until a Help menu exists, `Tools → Search commands` — opens the same palette, so it is discoverable without knowing the shortcut
- [ ] The palette lists every command in the registry, with its menu path ("Insert → Footnote"), its icon where one exists, and its shortcut label read from `shortcutLabel(id)` — never restated in the palette
- [ ] Typing filters by subsequence match against the command's label, menu path and its aliases, ranked per "Matching and ranking", with matched characters highlighted in the result
- [ ] Enter runs the highlighted command; Up/Down move; Escape closes and returns focus to where it was, with the document selection intact
- [ ] Running a command from the palette is identical to running it from its menu item — same id, same handler, same undo entry — asserted by a test that drives one command both ways and compares the resulting model
- [ ] A command that does not apply to the current selection is listed and **disabled**, with a short reason ("needs a selection", "no table here"), rather than hidden — the palette is a map of the product, matching SPEC-049's rule for menus
- [ ] With an empty query the palette shows recent commands (most recent first, up to 8, persisted under `docu-edit:recent-commands`), then a short list of common ones
- [ ] The palette searches the document's **styles** as well as commands: typing a style name offers "Apply style: Heading 2", drawn live from the open document's style catalogue
- [ ] The palette opens over the editor without moving it, is dismissed by clicking outside, and never steals the document selection — reopening after a text selection still shows the selection-dependent commands as enabled
- [ ] The palette is `role="dialog"` with a `combobox`/`listbox` search pattern: `aria-expanded`, `aria-controls`, `aria-activedescendant` on the input, `role="option"` per row, and a live region announcing the result count
- [ ] Every command in the registry has a label and a menu path in the metadata table, asserted by a unit test that fails when an id is added to `editorCommands` without an entry — the palette cannot silently miss a command
- [ ] The palette is available in the library view too, limited to the commands that make sense there (new document, new from template, open, search library) — one component, two command sources
- [ ] Opening, typing and running are all keyboard-only paths; no step requires the pointer

## Out of scope
Running commands with arguments from the palette ("insert a 3×4 table") —
a command that needs a dialog opens its dialog. Natural-language or fuzzy
*semantic* search. Searching document *content* — that is SPEC-017's find
bar, and the palette links to it rather than duplicating it. Searching
help articles; there is no help content. User-defined command aliases or
macros. Customising shortcuts (a separate spec). Adding, removing or
changing any command's behaviour.

## Reference

### The registry this reads
`src/views/DocumentView.vue` holds:

```ts
const editorCommands = new Map([
  ["file.print", openPrintPreview],
  ["file.page-setup", openPageSetup],
  …                                    // 82 entries as of SPEC-057
]);
```

A `Map` rather than an object literal because espree cannot parse the index
signature a string-keyed object lookup needs in a `.vue` script — the
palette must keep using `get()` for the same reason.

`src/lib/shortcuts.ts` holds the binding side:

```ts
export interface ShortcutDefinition {
  id: string; label: string; category: string;
  bindings: ShortcutBinding[]; overridesBrowser?: boolean;
}
export const SHORTCUTS: ShortcutDefinition[];
export function shortcutLabel(id: string, isMac?: boolean): string;
```

Ids are shared where a command has a shortcut; menu-only commands use the
same dotted style (`file.page-setup`, `insert.footnote`,
`tools.accessibility`). The palette dispatches with
`editorCommands.get(id)?.()`, which is what `DocumentView`'s own menu
handler already does.

### Command metadata — the table this spec adds
`src/lib/commandCatalogue.ts`, framework-free, so it is unit-testable
without mounting the view and reusable by any later surface:

```ts
export interface CommandMeta {
  /** Matches an `editorCommands` key exactly. */
  id: string;
  /** User-facing name, e.g. "Footnote". */
  label: string;
  /** Menu path shown as context, e.g. ["Insert", "Notes"]. */
  path: string[];
  /** SPEC-058 icon name, when the command has one. */
  icon?: string;
  /** Extra search terms: synonyms and the words people actually type.
   *  "TOC" for table of contents, "spell" for check spelling. */
  aliases?: string[];
  /** Which surface offers it. */
  scope: "document" | "library";
}
export const COMMANDS: CommandMeta[];
export function searchCommands(query: string, all: CommandMeta[]): CommandMatch[];
```

`searchCommands` and its ranking are pure and unit-tested; the Vue layer
only renders and dispatches. The completeness test asserts
`new Set(editorCommands.keys())` equals
`new Set(COMMANDS.filter(c => c.scope === "document").map(c => c.id))` —
so the table cannot drift from the registry in either direction.

Labels and paths must match the menu bar's own text, which
`DocumentMenuBar.vue` currently writes inline as `label="…"` attributes.
Where the menu item's label and the catalogue's disagree, the catalogue is
wrong: a follow-up may have the menu read its labels from this table, but
this spec does not require that rewrite.

### Matching and ranking
Subsequence (not substring) matching, case- and accent-insensitive, over
the concatenation of label, path and aliases — so "ftnt" finds Footnote and
"instab" finds Insert → Table. Ranking, highest first:

1. exact label match;
2. label prefix match;
3. contiguous substring in the label;
4. subsequence in the label, scored by how early and how contiguous the
   matched characters are;
5. a match that lands only in the path or the aliases.

Ties break toward more-recently-used, then alphabetically. Disabled
commands always sort after enabled ones with the same score, so the
palette never puts an unusable row at the top.

### Enablement
The palette shows live enablement, which means it needs the same predicates
the menus use. `DocumentMenuBar.vue` receives those as props today
(`canUndo`, `hasSelection`, `inTable`, and so on). This spec introduces one
computed record in the view:

```ts
/** id -> "" when enabled, or a short reason when not. */
const commandAvailability: ComputedRef<Record<string, string>>;
```

fed to both the menu bar and the palette, so a command's enabled state is
computed in one place and cannot disagree between the two surfaces. The
reason strings are short, lower-case fragments ("needs a selection", "no
table here", "document is read-only"), rendered dimmed at the row's end.

### Style commands
The palette's second source is the open document's style catalogue —
`NamedStyle` entries with `type: "paragraph"` or `"character"`, already
resolved for the style picker by `src/lib/styles.ts`. Each becomes a
virtual row `{ label: "Apply style: <name>", run: applyStyle(id) }`,
ranked below exact command matches so typing "heading" still surfaces
"Insert → Header" alongside "Apply style: Heading 1". These rows are built
per open document, never persisted into `COMMANDS`.

### Recents
`docu-edit:recent-commands` stores an array of ids, most recent first,
capped at 8, following the existing `docu-edit:*` convention
(`docu-edit:recent-text-colors`, `docu-edit:recent-special-characters`) and
written through the storage guard so a blocked `localStorage` degrades to
an in-session list. An id that is no longer in the catalogue is dropped on
read.

### Placement in the shortcut registry
Ctrl/Cmd+K is free: `RESERVED_BROWSER_KEYS` in `src/lib/shortcuts.ts`
reserves only `t`, `w` and `n`, and no existing definition binds `k`. The
palette registers as:

```ts
{ id: "app.commandPalette", label: "Search commands",
  category: "View", bindings: [{ key: "k", mod: true }] }
```

It is handled at the document level alongside find and print — the two
commands `handleGlobalShortcut` already owns — so it opens even when the
canvas does not have focus and over a document too large to edit
(SPEC-047's read-only ceiling).

### Specs this changes
- `specs/021-keyboard-shortcuts-and-reference.md` — "A command palette" in
  its out-of-scope list is resolved here; the reference sheet gains
  Ctrl/Cmd+K under View. User-remappable shortcuts stay out of scope.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — one menu item is
  added, per that spec's rule that a new command names its menu. The
  palette does not replace the menus and does not change any placement.
