import { ref } from "vue";

export interface Toast {
  id: number;
  message: string;
}

export const TOAST_DURATION_MS = 4000;

// Module-level state: toasts raised just before a navigation (e.g. "Deleted
// <title>" when deleting the open document) stay visible across the route
// change, because every ToastHost instance renders this same list.
const toasts = ref<Toast[]>([]);
let nextToastId = 1;

function addToast(message: string): void {
  const id = nextToastId;
  nextToastId += 1;
  toasts.value = [...toasts.value, { id, message }];
  setTimeout(() => dismissToast(id), TOAST_DURATION_MS);
}

function dismissToast(id: number): void {
  toasts.value = toasts.value.filter((toast) => toast.id !== id);
}

export function useToasts() {
  return { toasts, addToast, dismissToast };
}
