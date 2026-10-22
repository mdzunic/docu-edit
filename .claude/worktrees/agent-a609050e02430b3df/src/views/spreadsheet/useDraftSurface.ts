// One editable host of the cell editor's draft (SPEC-122 "Editor DOM"): the
// in-cell editor (CellEditor.vue) and the formula bar's content are both
// contenteditable hosts bound through this, so typing in either updates both
// and Enter/Tab/Escape behave identically in each.
//
// - Normalise-after-input: the browser applies an edit natively — the first
//   character typed in Ready included, so it is never lost or doubled — then
//   the host is read back to plain text (sanitizeXmlText'd), handed to
//   useCellEditing, and re-rendered as text plus one coloured span per
//   reference with the caret restored by offset (cellEditorDom.ts). Nothing
//   re-renders a host mid-composition: an IME composition is read back once
//   it ends (and from Ready it opens Enter mode when it starts).
// - Paste inserts plain text; drops and the browser's formatting and
//   history commands are refused; keys go to useCellEditing's key table
//   first and are stopped there when they are the editor's.
// - The formula bar shows the active cell's text in Ready; a click in it
//   opens Edit mode at the clicked character (keyboard focus: at the end),
//   and a commit from it returns focus to the in-cell editor.
// - SPEC-130: a selection the draft asks for (its `anchor` — a validation
//   alert's Retry selects the entered text) is rendered in the focused host;
//   once the host's own selection collapses the draft forgets it.
import { onBeforeUnmount, onMounted, watch, type Ref } from "vue";
import { sanitizeXmlText } from "../../lib/core/xmlText";
import { caretOffsetAtPoint, caretOffsetIn, placeCaretIn, readDraftText, renderDraft, selectionOffsetsIn } from "./cellEditorDom";
import type { CellEditing, SurfaceKind } from "./useCellEditing";

/** The browser's formatting and history commands, drops and paragraph
 * breaks never reach a host: the draft is plain text and its line breaks
 * come from the key table (Alt/Option+Enter). */
function onBeforeInput(event: InputEvent): void {
  const type = event.inputType ?? "";
  if (type.startsWith("format") || type.startsWith("history") || type === "insertFromDrop" || type === "insertParagraph" || type === "insertLineBreak") {
    event.preventDefault();
  }
}

export function useDraftSurface(
  host: Readonly<Ref<HTMLElement | null>>,
  editor: () => CellEditing | null,
  kind: SurfaceKind,
  /** What the host shows in Ready (the formula bar's text; the cell: ""). */
  readyText: () => string = () => "",
) {
  let composing = false;
  let refused = false;
  let pressed = false;
  let unregister: (() => void) | null = null;

  const focused = (): boolean => host.value !== null && host.value.ownerDocument.activeElement === host.value;

  /** Renders the draft (or the Ready text), restoring the caret when the
   * host has focus. */
  function render(): void {
    const el = host.value;
    const e = editor();
    if (!el || composing) {
      return;
    }
    const editing = e?.editing.value === true;
    renderDraft(el, editing ? e.text.value : readyText(), editing ? e.point.spans.value : []);
    if (editing && focused()) {
      const anchor = e.draft.anchor.value;
      const caret = e.caret.value;
      placeCaretIn(el, anchor === null ? caret : Math.min(anchor, caret), anchor === null ? caret : Math.max(anchor, caret));
    }
  }

  watch(
    () => [editor()?.revision.value, editor()?.mode.value, editor()?.point.spans.value, readyText()],
    () => render(),
    { flush: "post" },
  );

  /** Reads the host back into the draft, then renders the normalised form. */
  function absorb(): void {
    const el = host.value;
    const e = editor();
    if (!el || !e) {
      return;
    }
    const raw = readDraftText(el);
    const at = caretOffsetIn(el) ?? raw.length;
    const clean = sanitizeXmlText(raw);
    e.input(clean, clean === raw ? at : sanitizeXmlText(raw.slice(0, at)).length);
    render();
  }

  function onInput(event: Event): void {
    if (!composing && !(event as InputEvent).isComposing) {
      absorb();
    }
  }

  function onCompositionStart(): void {
    composing = true;
    refused = editor()?.compositionStart() === false;
  }

  function onCompositionEnd(): void {
    composing = false;
    editor()?.compositionEnd();
    if (refused) {
      refused = false;
      render();
    } else {
      absorb();
    }
  }



  function onPaste(event: ClipboardEvent): void {
    event.preventDefault();
    const e = editor();
    if (e?.editing.value) {
      e.insertText(sanitizeXmlText((event.clipboardData?.getData("text/plain") ?? "").replace(/\r\n?/g, "\n")));
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    const e = editor();
    if (e?.onKeydown(event)) {
      event.stopPropagation();
      if (kind === "bar" && !e.editing.value) {
        e.focusCell();
      }
    }
  }

  /** The formula bar: a click opens Edit mode at the clicked character;
   * keyboard focus opens it at the end; in Enter mode a click turns Edit. */
  function openFromBar(at: number | undefined): void {
    const e = editor();
    if (kind !== "bar" || !e) {
      return;
    }
    if (e.editing.value) {
      e.openEdit();
      return;
    }
    e.openEdit({ at });
  }

  function onPointerDown(): void {
    pressed = true;
  }

  function onFocus(): void {
    if (!pressed) {
      openFromBar(undefined);
    }
  }

  function onClick(): void {
    pressed = false;
    const el = host.value;
    openFromBar(el ? (caretOffsetIn(el) ?? undefined) : undefined);
  }

  function onSelectionChange(): void {
    const el = host.value;
    const e = editor();
    if (el && e?.editing.value && !composing && focused()) {
      e.caret.value = caretOffsetIn(el) ?? e.caret.value;
      const range = selectionOffsetsIn(el);
      if (e.draft.anchor.value !== null && (!range || range.start === range.end)) {
        e.draft.anchor.value = null;
      }
    }
  }

  let mounted = false;
  /** Registers this host with the editor on show (again when it changes). */
  function attach(): void {
    unregister?.();
    unregister =
      editor()?.register(kind, {
        selection: () => (host.value ? selectionOffsetsIn(host.value) : null),
        caretAt: (x, y) => (host.value ? caretOffsetAtPoint(host.value, x, y) : null),
        focus: () => host.value?.focus({ preventScroll: true }),
        focused,
      }) ?? null;
    render();
  }
  watch(editor, () => {
    if (mounted) {
      attach();
    }
  });

  onMounted(() => {
    mounted = true;
    attach();
    document.addEventListener("selectionchange", onSelectionChange);
  });
  onBeforeUnmount(() => {
    mounted = false;
    unregister?.();
    unregister = null;
    document.removeEventListener("selectionchange", onSelectionChange);
  });

  return {
    onInput,
    onBeforeInput,
    onCompositionStart,
    onCompositionEnd,
    onPaste,
    onKeydown,
    onPointerDown,
    onFocus,
    onClick,
    render,
  };
}
