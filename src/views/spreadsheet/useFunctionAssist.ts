// The formula-editing assists (SPEC-127 "The assists") over SPEC-122's
// editing session (useCellEditing): the function autocomplete, the argument
// hint and the Insert Function dialog. lib/spreadsheet/formula/functionHelp.ts
// does the reading; this owns the state, the keys and the insertions.
//
// - Autocomplete: while a formula is edited in the cell or the formula bar,
//   an identifier typed where a function may appear lists the registered
//   names starting with it, alphabetically (FunctionAutocomplete), with the
//   highlighted one's summary. Its keys come BEFORE the editor's (the key
//   filter useCellEditing runs first): Up/Down move the highlight, Tab or
//   Enter insert `NAME(` in place of the prefix (only the name when a "("
//   already follows, the caret then past it), Escape closes the list only —
//   so Enter does not commit and the arrows do not point. Typing refilters;
//   no match, or a string, closes it.
// - The hint (FunctionArgHint): the signature of the innermost enclosing
//   call of a registered function, its current argument bold; hidden while
//   the list is open.
// - The dialog (InsertFunctionDialog) opens from Shift+F3 (sheet.insertFunction
//   in SHEET_SHORTCUTS — in Ready from the window, while editing through the
//   key filter) and from the formula bar's fx button; on a cell protection
//   will not let be edited both do nothing and the button is disabled.
//   Insert while editing a formula puts `NAME(` at the caret; otherwise it
//   starts editing the active cell with `=NAME(`, as Excel does.
// - Accessibility: `editorAria` is what the editor hosts bind —
//   aria-autocomplete, aria-expanded, aria-controls and
//   aria-activedescendant while the list is open, aria-describedby the hint
//   while it shows. Every string is en-US whatever the locale.
import { computed, onScopeDispose, ref, watch, type Ref } from "vue";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { callAtCaret, completionAt, functionsStartingWith, signatureParts } from "../../lib/spreadsheet/formula/functionHelp";
import { getFunction } from "../../lib/spreadsheet/formula/registry";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";

export const AUTOCOMPLETE_ID = "function-autocomplete";
export const ARG_HINT_ID = "function-arg-hint";
export const PROTECTED_TITLE = "This cell is protected";

/** The editing session the assists ride on (useCellEditing's). */
export interface FunctionAssistDeps {
  editing: Readonly<Ref<boolean>>;
  text: Readonly<Ref<string>>;
  caret: Readonly<Ref<number>>;
  replaceRange(from: number, to: number, insert: string, caretAfter: number): void;
  /** Starts editing the active cell with a text; false when refused. */
  startEditing(text: string): boolean;
  canEditActiveCell(): boolean;
  /** Installs the assists' key handler ahead of the editor's own. */
  setKeyFilter(filter: ((event: KeyboardEvent) => boolean) | null): void;
  focusCell(): void;
}

/** Whether a keydown is Shift+F3, the Insert Function shortcut. */
function isInsertFunction(event: KeyboardEvent): boolean {
  return findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.insertFunction";
}

/** The focused editor host, for placing the list and hint below it. */
function focusedHost(): Element | null {
  if (typeof document === "undefined") {
    return null;
  }
  return document.activeElement?.closest('[data-testid="cell-editor"], [data-testid="formula-bar-content"]') ?? null;
}

export function useFunctionAssist(deps: FunctionAssistDeps) {
  const highlight = ref(0);
  /** The text the list was dismissed at (Escape); typing reopens it. */
  const dismissedAt = ref<string | null>(null);
  const dialogOpen = ref(false);

  const formula = computed(() => deps.editing.value && deps.text.value.startsWith("="));
  const completion = computed(() => (formula.value ? completionAt(deps.text.value, deps.caret.value) : null));
  const matches = computed(() => (completion.value ? functionsStartingWith(completion.value.prefix) : []));
  const listOpen = computed(() => matches.value.length > 0 && dismissedAt.value !== deps.text.value);
  const highlighted = computed(() => matches.value[Math.min(highlight.value, matches.value.length - 1)] ?? null);
  const summary = computed(() => (highlighted.value ? (getFunction(highlighted.value)?.summary ?? "") : ""));
  const hint = computed(() => {
    if (!formula.value || listOpen.value) {
      return null;
    }
    const call = callAtCaret(deps.text.value, deps.caret.value);
    const spec = call ? getFunction(call.name) : undefined;
    return call && spec ? { name: spec.name, parts: signatureParts(spec, call.argIndex) } : null;
  });

  watch(
    () => completion.value?.prefix,
    () => {
      highlight.value = 0;
    },
  );

  /** The list's prefix replaced by `NAME(` (see the module header). */
  function insert(name: string): void {
    const context = completion.value;
    if (!context) {
      return;
    }
    const caret = deps.caret.value;
    const parenFollows = deps.text.value[caret] === "(";
    deps.replaceRange(context.start, caret, parenFollows ? name : `${name}(`, context.start + name.length + 1);
  }

  function openDialog(): void {
    if (deps.canEditActiveCell()) {
      dialogOpen.value = true;
    }
  }

  /** The dialog's Insert (see the module header). */
  function insertFromDialog(name: string): void {
    dialogOpen.value = false;
    if (formula.value) {
      const caret = deps.caret.value;
      deps.replaceRange(caret, caret, `${name}(`, caret + name.length + 1);
      deps.focusCell();
      return;
    }
    deps.startEditing(`=${name}(`);
  }

  /** The editor's keys, before its own table: true when taken. */
  function onEditorKey(event: KeyboardEvent): boolean {
    if (isInsertFunction(event)) {
      event.preventDefault();
      openDialog();
      return true;
    }
    if (!listOpen.value || event.ctrlKey || event.metaKey || event.altKey) {
      return false;
    }
    const count = matches.value.length;
    switch (event.key) {
      case "ArrowDown":
        highlight.value = (Math.min(highlight.value, count - 1) + 1) % count;
        break;
      case "ArrowUp":
        highlight.value = (Math.min(highlight.value, count - 1) + count - 1) % count;
        break;
      case "Tab":
      case "Enter":
        if (highlighted.value) {
          insert(highlighted.value);
        }
        break;
      case "Escape":
        dismissedAt.value = deps.text.value;
        break;
      default:
        return false;
    }
    event.preventDefault();
    return true;
  }

  /** Shift+F3 in Ready, from anywhere but an input or a dialog. */
  function onWindowKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !isInsertFunction(event) || isOutsideDocumentInput(event.target)) {
      return;
    }
    event.preventDefault();
    openDialog();
  }

  deps.setKeyFilter(onEditorKey);
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onWindowKeydown);
  }
  onScopeDispose(() => {
    deps.setKeyFilter(null);
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onWindowKeydown);
    }
  });

  /** The list and hint's place: fixed, just below the focused editor host. */
  const anchorStyle = computed((): Record<string, string> => {
    void deps.text.value;
    void deps.caret.value;
    void listOpen.value;
    const rect = focusedHost()?.getBoundingClientRect();
    return rect ? { left: `${Math.round(rect.left)}px`, top: `${Math.round(rect.bottom + 2)}px` } : {};
  });

  const editorAria = computed((): Record<string, string> => {
    if (listOpen.value) {
      return {
        "aria-autocomplete": "list",
        "aria-expanded": "true",
        "aria-controls": AUTOCOMPLETE_ID,
        "aria-activedescendant": `${AUTOCOMPLETE_ID}-${highlighted.value ?? ""}`,
      };
    }
    return hint.value ? { "aria-autocomplete": "list", "aria-expanded": "false", "aria-describedby": ARG_HINT_ID } : {};
  });

  return {
    /** The autocomplete: open, its names, the highlighted one and its summary. */
    listOpen,
    matches,
    highlighted,
    summary,
    choose: insert,
    /** The argument hint, or null. */
    hint,
    dialogOpen,
    openDialog,
    closeDialog: (): void => {
      dialogOpen.value = false;
    },
    insertFromDialog,
    /** Whether the fx button is enabled (protection allows an edit). */
    canInsert: computed(() => {
      void deps.editing.value;
      return deps.canEditActiveCell();
    }),
    anchorStyle,
    editorAria,
    onEditorKey,
  };
}

export type FunctionAssist = ReturnType<typeof useFunctionAssist>;

/** Narrows a component's runtime-typed `assist` prop (the .vue parser bars
 * casts). */
export function asFunctionAssist(value: unknown): FunctionAssist {
  return value as FunctionAssist;
}
