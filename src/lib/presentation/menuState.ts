// Shape helper for the slides menu bar (SPEC-049's model, for the
// presentation product). PresentationMenuBar takes its state as one runtime
// `Object` prop — the espree constraint on .vue scripts bars a typed PropType
// there (see AppShell.vue) — so it coerces that prop through here, mirroring
// how MenuBar runs its `menus` prop through `asMenuBarMenus` and the Table
// menu runs its state through `asTableMenuState`.
//
// One object rather than a dozen props: the bar needs a checkmark or a
// disabled reason for most of its items, and threading that many bindings
// through the view's template would cost more than it explains.

/** What the bar needs in order to render: what is checked, what is disabled. */
export interface PresentationMenuState {
  canUndo: boolean;
  canRedo: boolean;
  canPresent: boolean;
  /** The deck's other export format, e.g. "OpenDocument (.odp)". "" hides the
   * item — a deck already in that format has no "other" to offer. */
  otherFormatLabel: string;
  designOpen: boolean;
  transitionOpen: boolean;
  animateOpen: boolean;
  layoutOpen: boolean;
  tableOpen: boolean;
  mediaOpen: boolean;
  headerFooterOpen: boolean;
  notesOpen: boolean;
  /** Whether anything on the stage is selected — the arrange and position
   * commands need a target. */
  hasSelection: boolean;
  canGroup: boolean;
  canUngroup: boolean;
  /** SPEC-115: whether the deck has slides for "Redesign…" to arrange (the
   * item itself renders only while the deck-designer experiment is on). */
  canRedesign: boolean;
}

/** Coerces the untyped `state` prop into the typed shape the template reads.
 * Every field defaults to the safe answer (nothing checked, nothing
 * available), so a bar mounted without state renders a stable, fully disabled
 * map of the product rather than throwing. */
export function asPresentationMenuState(value: unknown): PresentationMenuState {
  const state = (value ?? {}) as Partial<PresentationMenuState>;
  return {
    canUndo: !!state.canUndo,
    canRedo: !!state.canRedo,
    canPresent: !!state.canPresent,
    otherFormatLabel: String(state.otherFormatLabel ?? ""),
    designOpen: !!state.designOpen,
    transitionOpen: !!state.transitionOpen,
    animateOpen: !!state.animateOpen,
    layoutOpen: !!state.layoutOpen,
    tableOpen: !!state.tableOpen,
    mediaOpen: !!state.mediaOpen,
    headerFooterOpen: !!state.headerFooterOpen,
    notesOpen: !!state.notesOpen,
    hasSelection: !!state.hasSelection,
    canGroup: !!state.canGroup,
    canUngroup: !!state.canUngroup,
    canRedesign: !!state.canRedesign,
  };
}
