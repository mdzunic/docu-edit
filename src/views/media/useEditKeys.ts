// The image editor's keyboard (SPEC-112), installed on the window while the
// editor is open: Escape leaves (Done), SPEC-015's Cmd/Ctrl+Z,
// Shift+Cmd/Ctrl+Z and Cmd/Ctrl+Y (from the shortcut registry) undo and
// redo — except in a text field, whose own undo they are — and Cmd/Ctrl+[ /
// Cmd/Ctrl+] rotate left and right. A dialog's keys are its own.
import { onScopeDispose, watch, type Ref } from "vue";
import { findShortcut } from "../../lib/shortcuts";

export interface EditKeysDeps {
  active: Ref<boolean>;
  /** Whether the tools may act (the session is ready and nothing is busy). */
  ready: () => boolean;
  done: () => void;
  rotate: (delta: 1 | -1) => void;
  undo: () => void;
  redo: () => void;
}

/** Whether a key event comes from a field where Cmd/Ctrl+Z edits text. */
function inTextField(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable || target.tagName === "TEXTAREA" || target.tagName === "SELECT") return true;
  return target instanceof HTMLInputElement && !["range", "checkbox", "radio", "button", "color"].includes(target.type);
}

export function useEditKeys(deps: EditKeysDeps): (event: KeyboardEvent) => boolean {
  function handleKeydown(event: KeyboardEvent): boolean {
    if (!deps.active.value || event.defaultPrevented) return false;
    const target = event.target;
    if (target instanceof Element && target.closest('[role="dialog"]')) return false;
    const mod = event.metaKey || event.ctrlKey;
    if (event.key === "Escape" && !mod) {
      deps.done();
      return true;
    }
    if (!deps.ready()) return false;
    const bracket = event.key === "[" || event.key === "]" || event.code === "BracketLeft" || event.code === "BracketRight";
    if (mod && !event.altKey && bracket) {
      deps.rotate(event.key === "]" || event.code === "BracketRight" ? 1 : -1);
      return true;
    }
    const shortcut = findShortcut(event);
    if ((shortcut?.id === "edit.undo" || shortcut?.id === "edit.redo") && !inTextField(target)) {
      if (shortcut.id === "edit.undo") deps.undo();
      else deps.redo();
      return true;
    }
    return false;
  }

  function onWindowKeydown(event: KeyboardEvent): void {
    if (handleKeydown(event)) event.preventDefault();
  }
  watch(deps.active, (on) => {
    if (typeof window === "undefined") return;
    if (on) window.addEventListener("keydown", onWindowKeydown);
    else window.removeEventListener("keydown", onWindowKeydown);
  });
  onScopeDispose(() => {
    if (typeof window !== "undefined") window.removeEventListener("keydown", onWindowKeydown);
  });
  return handleKeydown;
}
