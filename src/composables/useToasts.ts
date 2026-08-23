import { ref } from "vue";

/** An optional action button on a toast — e.g. the library's Undo after a
 * move-to-trash (SPEC-067 AC-34). Running it dismisses the toast. */
export interface ToastAction {
  label: string;
  run: () => void;
}

export interface Toast {
  id: number;
  message: string;
  action?: ToastAction;
}

export const TOAST_DURATION_MS = 4000;

// Module-level state: toasts raised just before a navigation (e.g. "Deleted
// <title>" when deleting the open document) stay visible across the route
// change, because every ToastHost instance renders this same list.
const toasts = ref<Toast[]>([]);
let nextToastId = 1;

function addToast(message: string, action?: ToastAction): void {
  const id = nextToastId;
  nextToastId += 1;
  toasts.value = [...toasts.value, { id, message, action }];
  setTimeout(() => dismissToast(id), TOAST_DURATION_MS);
}

function dismissToast(id: number): void {
  toasts.value = toasts.value.filter((toast) => toast.id !== id);
}

/** Runs a toast's action (if any) then dismisses it — the wiring behind the
 * action button. */
function runToastAction(id: number): void {
  const toast = toasts.value.find((entry) => entry.id === id);
  toast?.action?.run();
  dismissToast(id);
}

export function useToasts() {
  return { toasts, addToast, dismissToast, runToastAction };
}
