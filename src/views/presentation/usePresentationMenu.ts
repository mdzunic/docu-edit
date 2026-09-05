// The slides menu bar's command table (SPEC-049's model, for the other
// product). The word processor has had a menu bar since SPEC-049; the deck
// editor had only strips of buttons, so the same app offered two different
// ways to find a command and gave the same action two names ("Download" on one
// side, "Export" on the other). This is the deck's half of closing that.
//
// The shape follows DocumentMenuBar's: the bar emits a command id, and this
// maps ids to the handlers the view already owns. Nothing new is implemented
// here — every entry routes to a composable the strips already call — so the
// menu is a second, discoverable route to the same commands rather than a
// parallel implementation that can drift from them.
import { computed, type ComputedRef } from "vue";
import type { PresentationMenuState } from "../../lib/presentation/menuState";

/* eslint-disable @typescript-eslint/no-explicit-any */
/** The composables the view has already built. Typed loosely on purpose: this
 * is a routing table over them, and naming each composable's full shape here
 * would duplicate a dozen interfaces that already exist next door. */
export interface PresentationMenuDeps {
  undo: any;
  deckActions: any;
  deckSetup: any;
  notes: any;
  transitions: any;
  animations: any;
  layouts: any;
  tables: any;
  media: any;
  headerFooter: any;
  themePanel: any;
  stageEdit: any;
  charts: any;
  preferences: any;
  addSlide: () => void;
  toggleLayout: () => void;
}
/* eslint-enable @typescript-eslint/no-explicit-any */

export function usePresentationMenu(deps: PresentationMenuDeps): {
  state: ComputedRef<PresentationMenuState>;
  run: (command: string) => void;
} {
  const arrange = () => deps.stageEdit.arrange;
  const editing = () => deps.stageEdit.editing;

  const state = computed<PresentationMenuState>(() => ({
    canUndo: !!deps.undo.canUndo.value,
    canRedo: !!deps.undo.canRedo.value,
    canPresent: !!deps.deckActions.canPresent.value,
    otherFormatLabel: String(deps.deckActions.otherFormatLabel.value ?? ""),
    designOpen: !!deps.themePanel.open.value,
    transitionOpen: !!deps.transitions.open.value,
    animateOpen: !!deps.animations.open.value,
    layoutOpen: !!deps.layouts.open.value,
    tableOpen: !!deps.tables.open.value,
    mediaOpen: !!deps.media.open.value,
    headerFooterOpen: !!deps.headerFooter.open.value,
    notesOpen: !!deps.notes.open.value,
    hasSelection: arrange().selectedCount.value > 0,
    canGroup: arrange().selectedCount.value > 1,
    canUngroup: !!arrange().canUngroup?.value,
  }));

  const commands = new Map<string, () => void>([
    // File
    ["file.export", () => deps.deckActions.download()],
    ["file.export-copy", () => deps.deckActions.downloadCopy()],
    ["file.export-html", () => deps.deckActions.openHtmlExport()],
    ["file.deck-setup", () => deps.deckSetup.open()],
    ["app.preferences", () => deps.preferences.open()],
    ["file.close", () => deps.deckActions.close()],

    // Edit
    ["edit.undo", () => deps.undo.undo()],
    ["edit.redo", () => deps.undo.redo()],

    // View — the panel toggles the right-hand strip also drives, so a
    // checkmark here and a pressed button there report the same state.
    ["view.design", () => deps.themePanel.toggle()],
    ["view.transition", () => deps.transitions.toggle()],
    ["view.animate", () => deps.animations.toggle()],
    ["view.layout", () => deps.toggleLayout()],
    ["view.table", () => deps.tables.toggle()],
    ["view.media", () => deps.media.toggle()],
    ["view.header-footer", () => deps.headerFooter.show()],
    ["view.notes", () => deps.notes.toggle()],
    ["view.present", () => deps.deckActions.present()],

    // Insert
    ["insert.slide", () => deps.addSlide()],
    ["insert.text-box", () => deps.stageEdit.insertion.armTextBox()],

    // Format
    ["format.bold", () => editing().toggleBold()],
    ["format.italic", () => editing().toggleItalic()],
    ["format.underline", () => editing().toggleUnderline()],
    ["format.bullet", () => editing().toggleBulletCmd()],
    ["format.indent", () => editing().indent(1)],
    ["format.outdent", () => editing().indent(-1)],
    ["format.group", () => arrange().group()],
    ["format.ungroup", () => arrange().ungroup()],
    ["format.position", () => arrange().openPosition()],
  ]);

  /** Runs a command id, or a parameterised one: `insert.shape:<kind>`,
   * `insert.chart:<kind>`, `format.align:<value>`, `arrange:<value>`. The
   * suffix form keeps the bar emitting plain strings (as DocumentMenuBar
   * does) instead of structured payloads. */
  function run(command: string): void {
    const direct = commands.get(command);
    if (direct) {
      direct();
      return;
    }
    const [id, argument] = command.split(":");
    if (argument === undefined) {
      return;
    }
    switch (id) {
      case "insert.shape":
        deps.stageEdit.insertion.armShape(argument);
        break;
      case "insert.chart":
        deps.charts.pickChartType(argument);
        break;
      case "format.align":
        editing().setAlign(argument);
        break;
      case "arrange":
        deps.stageEdit.manipulation.arrange(argument);
        break;
      case "align":
        arrange().align(argument);
        break;
      default:
        break;
    }
  }

  return { state, run };
}
