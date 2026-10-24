// The open/closed state of the Preferences dialog (SPEC-066), as a module
// singleton so any view can open the one dialog AppShell hosts. The dialog
// itself lives in AppShell — the single component wrapping both the library
// (HomeView) and the editor (DocumentView) — so it is reachable from either
// (AC-39, AC-40); this shared flag is how the editor's File menu, the global
// Cmd+, shortcut and the library entry point all open that one instance.
import { ref } from "vue";

const isOpen = ref(false);

/** Set when the dialog should bring its Labs section into view (SPEC-113's
 * experiment badge links there); the Labs section clears it once shown. */
const labsRequested = ref(false);

export function usePreferencesDialog() {
  return {
    isOpen,
    labsRequested,
    open(): void {
      isOpen.value = true;
    },
    /** Opens the dialog at its Labs section. */
    openLabs(): void {
      labsRequested.value = true;
      isOpen.value = true;
    },
    close(): void {
      isOpen.value = false;
    },
  };
}
