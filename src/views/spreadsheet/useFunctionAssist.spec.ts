// SPEC-127 "The assists": useFunctionAssist over a stand-in editing session —
// the list opens on an identifier where a function may appear and never in a
// string, Up/Down/Tab/Enter/Escape are taken before the editor's keys while
// it is open, insertion replaces the prefix with NAME( (only the name before
// an existing "("), the hint follows the innermost call and hides while the
// list is open, the editor's ARIA, and the dialog: Shift+F3 and fx open it
// unless the cell is protected, Insert puts NAME( at the caret while a
// formula is edited and starts editing =NAME( otherwise.
import { afterEach, describe, expect, it, vi } from "vitest";
import { effectScope, nextTick, ref, type EffectScope } from "vue";
import { useFunctionAssist, type FunctionAssistDeps } from "./useFunctionAssist";

/** Every session's scope, stopped after each test (its window listener with it). */
const scopes: EffectScope[] = [];
afterEach(() => scopes.splice(0).forEach((scope) => scope.stop()));

function session(initial = "", editing = true, canEdit = true) {
  const text = ref(initial);
  const caret = ref(initial.length);
  let filter: ((event: KeyboardEvent) => boolean) | null = null;
  const deps: FunctionAssistDeps = {
    editing: ref(editing),
    text,
    caret,
    replaceRange(from, to, insert, caretAfter) {
      text.value = text.value.slice(0, from) + insert + text.value.slice(to);
      caret.value = caretAfter;
    },
    startEditing: vi.fn((value: string) => {
      text.value = value;
      caret.value = value.length;
      return true;
    }),
    canEditActiveCell: () => canEdit,
    setKeyFilter: (next) => {
      filter = next;
    },
    focusCell: vi.fn(),
  };
  const scope = effectScope();
  scopes.push(scope);
  const assist = scope.run(() => useFunctionAssist(deps))!;
  const type = (value: string) => {
    text.value = value;
    caret.value = value.length;
  };
  const key = (name: string, init: KeyboardEventInit = {}) => filter?.(new KeyboardEvent("keydown", { key: name, cancelable: true, ...init })) ?? false;
  return { assist, deps, text, caret, type, key, scope };
}

describe("the autocomplete", () => {
  it("lists the registered names starting with the typed identifier, alphabetically, with the highlighted one's summary", () => {
    const { assist, type } = session("=");
    type("=ROU");
    expect(assist.listOpen.value).toBe(true);
    expect(assist.matches.value).toEqual(["ROUND", "ROUNDDOWN", "ROUNDUP"]);
    expect(assist.highlighted.value).toBe("ROUND");
    expect(assist.summary.value.length).toBeGreaterThan(0);
  });

  it("never opens in a string, without a formula, or with no match", () => {
    const { assist, type } = session("=");
    type('="ROU');
    expect(assist.listOpen.value).toBe(false);
    type("ROU");
    expect(assist.listOpen.value).toBe(false);
    type("=SEQ");
    expect(assist.listOpen.value).toBe(false);
  });

  it("moves the highlight with Up/Down and inserts NAME( with Tab, replacing the prefix", async () => {
    const { assist, type, key, text, caret } = session("=");
    type("=1+rou");
    await nextTick();
    expect(key("ArrowDown")).toBe(true);
    expect(assist.highlighted.value).toBe("ROUNDDOWN");
    expect(key("ArrowUp")).toBe(true);
    expect(key("ArrowUp")).toBe(true);
    expect(assist.highlighted.value).toBe("ROUNDUP");
    expect(key("Tab")).toBe(true);
    expect(text.value).toBe("=1+ROUNDUP(");
    expect(caret.value).toBe(text.value.length);
    expect(assist.listOpen.value).toBe(false);
  });

  it("inserts with Enter, completing only the name before an existing '('", () => {
    const { key, text, caret } = session("=XL(1)");
    caret.value = 3;
    expect(key("Enter")).toBe(true);
    expect(text.value).toBe("=XLOOKUP(1)");
    expect(caret.value).toBe(9);
  });

  it("closes the list on Escape only, and reopens it as typing goes on", () => {
    const { assist, type, key } = session("=");
    type("=SU");
    expect(key("Escape")).toBe(true);
    expect(assist.listOpen.value).toBe(false);
    type("=SUM");
    expect(assist.listOpen.value).toBe(true);
  });

  it("leaves every key to the editor while the list is closed", () => {
    const { key } = session("=1+2");
    expect(key("Enter")).toBe(false);
    expect(key("ArrowUp")).toBe(false);
  });
});

describe("the argument hint", () => {
  it("shows the innermost call's signature with the current argument, hidden while the list is open", () => {
    const { assist, type } = session("=");
    type('=XLOOKUP("x",');
    expect(assist.hint.value?.name).toBe("XLOOKUP");
    expect(assist.hint.value?.parts.find((part) => part.current)?.text).toBe("lookup_array");
    type('=XLOOKUP("x",ROUND(');
    expect(assist.hint.value?.parts.find((part) => part.current)?.text).toBe("number");
    type('=XLOOKUP("x",RO');
    expect(assist.listOpen.value).toBe(true);
    expect(assist.hint.value).toBeNull();
  });

  it("gives the editor its ARIA: the list's while it is open, the hint's while it shows", () => {
    const { assist, type } = session("=");
    type("=XL");
    expect(assist.editorAria.value).toMatchObject({ "aria-autocomplete": "list", "aria-expanded": "true", "aria-controls": "function-autocomplete" });
    expect(assist.editorAria.value["aria-activedescendant"]).toBe("function-autocomplete-XLOOKUP");
    type("=XLOOKUP(");
    expect(assist.editorAria.value["aria-describedby"]).toBe("function-arg-hint");
    type("1");
    expect(assist.editorAria.value).toEqual({});
  });
});

describe("the Insert Function dialog", () => {
  it("opens from Shift+F3 while editing and from fx, but not on a protected cell", () => {
    const open = session("=");
    expect(open.key("F3", { shiftKey: true })).toBe(true);
    expect(open.assist.dialogOpen.value).toBe(true);
    const locked = session("", false, false);
    locked.assist.openDialog();
    expect(locked.assist.dialogOpen.value).toBe(false);
    expect(locked.assist.canInsert.value).toBe(false);
  });

  it("opens from Shift+F3 in Ready through the window", () => {
    const { assist } = session("", false);
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key: "F3", shiftKey: true, bubbles: true, cancelable: true }));
    expect(assist.dialogOpen.value).toBe(true);
  });

  it("puts NAME( at the caret while a formula is edited", () => {
    const { assist, text, deps } = session("=1+");
    assist.openDialog();
    assist.insertFromDialog("PMT");
    expect(text.value).toBe("=1+PMT(");
    expect(assist.dialogOpen.value).toBe(false);
    expect(deps.focusCell).toHaveBeenCalled();
  });

  it("starts editing the active cell with =NAME( otherwise", () => {
    const { assist, deps, text } = session("", false);
    assist.insertFromDialog("PMT");
    expect(deps.startEditing).toHaveBeenCalledWith("=PMT(");
    expect(text.value).toBe("=PMT(");
  });
});
