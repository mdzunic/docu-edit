// The presentation view's window-level keydown wiring (SPEC-095 extraction),
// lifted out of PresentationView so the view stays a wiring hub within its size
// budget. The stage's global shortcuts, the notes shortcuts and the Present
// shortcut all listen on `window` for the editor's lifetime; this composable
// registers them on mount and tears them down on unmount, so the view gains one
// call instead of a paired onMounted/onBeforeUnmount block.
import { onBeforeUnmount, onMounted } from "vue";

/** Registers each handler as a `window` keydown listener for the calling
 * component's lifetime (added on mount, removed on unmount). */
export function useWindowKeys(handlers: Array<(event: KeyboardEvent) => void>): void {
  const listeners = handlers.slice();
  onMounted(() => {
    for (const handler of listeners) {
      window.addEventListener("keydown", handler);
    }
  });
  onBeforeUnmount(() => {
    for (const handler of listeners) {
      window.removeEventListener("keydown", handler);
    }
  });
}
