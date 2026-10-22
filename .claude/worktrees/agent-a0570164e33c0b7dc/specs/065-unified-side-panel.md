---
id: SPEC-065
title: The unified side panel
status: draft
area: shell
depends_on: [SPEC-037, SPEC-049, SPEC-058]
---
## Why
Seven specs each needed somewhere to put a pane, and each invented one.
`src/views/DocumentView.vue` now renders seven sibling docks —
`.outline-dock`, `.styles-dock`, `.comments-dock`, `.track-changes-dock`,
`.version-history-dock`, `.accessibility-dock`, `.fonts-dock` — at four
different widths (16rem, 18rem, 20rem), each reaching into its pane with a
`:deep()` rule to fix up styling the pane could not know about. The panes
themselves are 199 to 730 lines apiece and share nothing: no common header,
no close affordance convention, no resize, no scroll behaviour, no ARIA
shape. Three of the seven persist their open state
(`docu-edit:outline-pane`, `docu-edit:comments-pane`,
`docu-edit:version-history-pane`); four forget it.

What a user sees is worse than what a developer sees. Opening the outline
and the comments pane at once eats 34rem of a 1280px window, leaving the
page squeezed into what is left, with no way to narrow either. Nothing
tells you which panes exist — you find them by reading the View menu. And
because each dock is its own flex child, opening a fourth pane simply keeps
taking width from the document until the page has none.

Every editor with this many auxiliary views converges on the same answer:
one docked region, tabbed or stacked, resizable, remembered. This spec is
that region. It is a pure consolidation — no pane's *contents* change —
and it is the last chance to do it cheaply, because SPEC-072 and SPEC-076
both want panes of their own.

## Acceptance criteria
- [ ] One side-panel region hosts every auxiliary pane; the seven per-pane docks and their `:deep()` fix-ups are deleted
- [ ] The region docks to the right by default and can be moved to the left, with the choice persisted; the outline pane's historical left position becomes the default *tab* of a left-docked panel only if the user puts it there
- [ ] The panel shows one pane at a time, chosen from a tab rail of the open panes; opening a second pane adds a tab rather than a second dock
- [ ] Each tab shows its pane's icon (SPEC-058) and name, and a badge where the pane has a count to report — comment count, unresolved-change count, accessibility-issue count
- [ ] The panel has one header: the active pane's title, any pane-supplied header actions, and one close button that closes the whole panel
- [ ] The panel is resizable by dragging its inner edge, between a 240px minimum and half the window width, with the width persisted; the document area reflows and repaginates only if page geometry actually changed (it does not — the page width is fixed, so this is a scroll/zoom concern, per "Interaction with zoom")
- [ ] Double-clicking the resize handle collapses the panel to its tab rail and restores it, keeping the tabs reachable
- [ ] Panel state — docked side, width, open/collapsed, which tabs are open and which is active — persists under a single `docu-edit:side-panel` key, replacing the three per-pane keys, with a documented migration for the existing ones
- [ ] Every existing View-menu toggle keeps working and keeps its command id: toggling a pane that is closed opens the panel with that pane active; toggling the active pane closes it; toggling a pane that is open but inactive activates it
- [ ] The panel is `role="complementary"` with an `aria-label` naming the active pane; the tab rail is a `tablist`/`tab`/`tabpanel` set with arrow-key navigation, and the resize handle is a `separator` with `aria-orientation="vertical"` and arrow-key resizing
- [ ] Focus order runs document → panel tab rail → active pane contents, and closing the panel returns focus to the document at its prior selection
- [ ] Each pane keeps its own root class and every `data-testid` it ships with, so no existing e2e selector changes; the panes lose only their outer dock
- [ ] A pane that becomes irrelevant closes itself gracefully — closing version-history preview mode, resolving the last comment — without collapsing the panel if other tabs are open
- [ ] At viewport widths below 900px the panel overlays the document instead of displacing it, with a scrim, and closes on Escape
- [ ] `DocumentView.size.spec.ts`'s budget ratchets **down**: removing seven docks and their styles takes lines out of the view

## Out of scope
The contents, behaviour or model of any pane — SPEC-037 (outline),
SPEC-038 (comments), SPEC-039 (track changes), SPEC-040 (version history),
SPEC-042 (styles), SPEC-043 (accessibility) and SPEC-048 (fonts) each keep
their spec's behaviour verbatim. Floating or undockable panes. Two panels
at once, or splitting the panel horizontally. Drag-and-drop reordering of
tabs. A panel on the bottom edge. Adding any new pane. Mobile layout
(SPEC-068), beyond the overlay rule above.

## Reference

### What ships today
In `src/views/DocumentView.vue`'s template, inside `.editor-layout`:

```
.outline-dock          v-if="outlinePaneOpen"           16rem   OutlinePane        (428 lines)
.styles-dock           v-if="stylesPaneOpen"            18rem   StylesPane         (292)
.comments-dock         v-if="commentsPaneOpen"          18rem   CommentsPane       (730)
.track-changes-dock    v-if="trackChangesPaneOpen"      18rem   TrackChangesPane   (386)
.version-history-dock  v-if="versionHistoryPaneOpen"    20rem   VersionHistoryPane (375)
.accessibility-dock    v-if="accessibilityPaneOpen"     20rem   AccessibilityPane  (237)
.fonts-dock            v-if="fontsPaneOpen"             20rem   FontsPane          (199)
```

Each pane's root carries its own class (`.outline-pane`, `.styles-pane`,
…), which the dock then reaches into with `:deep()`. The View menu toggles
them through the `editorCommands` Map: `view.outline-pane`,
`view.styles-pane`, `view.comments-pane`, `view.track-changes-pane`,
`file.version-history`, plus the Tools entries for accessibility and fonts.

### The panel
`src/components/SidePanel.vue`, plus a composable that owns the state so
the view keeps only a wiring call, per CLAUDE.md's DocumentView rule:

```ts
// src/views/document/useSidePanel.ts
export type PaneId =
  | "outline" | "styles" | "comments" | "trackChanges"
  | "versionHistory" | "accessibility" | "fonts";

export interface SidePanelState {
  /** Tabs currently open, in the order they were opened. */
  openPanes: Ref<PaneId[]>;
  activePane: Ref<PaneId | null>;   // null means the panel is closed
  side: Ref<"left" | "right">;
  widthPx: Ref<number>;
  collapsed: Ref<boolean>;
}

export interface PaneDescriptor {
  id: PaneId;
  title: string;
  icon: string;                     // SPEC-058 icon name
  /** Live count for the tab badge, or 0 for none. */
  badge?: Readonly<Ref<number>>;
}

export function useSidePanel(deps: SidePanelDeps): { … };
```

`SidePanelDeps extends EditorContext` (`src/views/document/context.ts`),
following the established composable convention.

The panel renders: a tab rail (icon + title + badge per open pane, active
one selected), a header (active pane's title, its optional header actions
via a slot, a close button), and the active pane's component in the body.
Panes are kept alive when switched away from, so scroll position and
transient state survive tab switches — `VersionHistoryPane`'s preview mode
and `CommentsPane`'s draft reply both depend on that.

### Toggle semantics
Each existing command keeps its id and gains a three-way meaning:

| State when invoked | Result |
| --- | --- |
| Pane not open | Open it, make it active, open the panel |
| Pane open, not active | Make it active |
| Pane open and active | Close that pane; if it was the last, close the panel |

This preserves the "toggle" feel of every shipped View item while making a
tabbed panel behave the way tabbed panels do. The View menu's checked state
means "open", not "active" — so a `menuitemcheckbox` stays checked for a
background tab, which is what the user's mental model expects.

### Persistence and migration
One key replaces three:

```json
// docu-edit:side-panel
{ "side": "right", "widthPx": 320, "collapsed": false,
  "openPanes": ["outline", "comments"], "activePane": "comments" }
```

On first read, if the new key is absent, the three legacy keys
(`docu-edit:outline-pane`, `docu-edit:comments-pane`,
`docu-edit:version-history-pane`) are read once to seed `openPanes` and
then removed. A pane id in stored state that no longer exists is dropped.
Writes go through the existing storage-availability guard (SPEC-003), so a
blocked `localStorage` falls back to session-only state.

### Interaction with zoom
Resizing the panel changes the width available to the document, which is
what SPEC-063's Fit width reads. So a resize re-evaluates the fit zoom when
the mode is Fit width or Fit page, and otherwise leaves the zoom alone.
Page geometry never depends on panel width — the page is its section's
width in centimetres and always was — so a panel resize schedules **no
pagination work**, which is the invariant to assert rather than assume,
because it is the one that protects SPEC-047's baseline.

### Specs this changes
Seven specs' pane placement is superseded; none of their behaviour is.
Each keeps its component, its root class, its `data-testid`s and its
command id:

- `specs/037-document-outline-navigator.md` — its pane moves from a
  standalone left dock into the panel. Its "toggled from the View menu and
  the state persists" AC is satisfied by the panel's persistence.
- `specs/038-comments.md`, `specs/039-track-changes.md`,
  `specs/040-version-history-and-restore.md`,
  `specs/042-custom-styles-manager.md`,
  `specs/043-accessibility-checker.md`,
  `specs/048-font-fallback-and-embedding.md` — same: dock replaced, pane
  unchanged.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the View and Tools
  menus keep exactly the items they have; only what a toggle *does* when
  the pane is already open is refined, per the table above.
