// The bridge between the model gateway's consent prompt and the shell's
// ModelConsentDialog (SPEC-113). The gateway (framework-free) asks through
// `askModelConsent`; AppShell hosts the one dialog and answers through
// `useModelConsentHost`. A module singleton, like usePreferencesDialog, so a
// request made from any view reaches the dialog of whichever shell is mounted.
//
// Fails closed: with no shell mounted to show the dialog, or a second request
// arriving while one is already on screen, the answer is "cancel" — a request
// nobody could see is a request nobody agreed to.
//
// Imports only the gateway's types (erased at build), so hosting the dialog
// loads none of the gateway's code.
import { computed, onMounted, onUnmounted, shallowRef } from "vue";
import type { ConsentChoice, ConsentRequest } from "../lib/ai/modelGateway";

interface PendingConsent {
  request: ConsentRequest;
  settle: (choice: ConsentChoice) => void;
}

const pending = shallowRef<PendingConsent | null>(null);
let hosts = 0;

/** The gateway's ConsentPrompt: shows `request` in the dialog and resolves
 * with the user's answer — or "cancel" when the caller's signal aborts while
 * the dialog is open. */
export function askModelConsent(request: ConsentRequest, signal?: AbortSignal): Promise<ConsentChoice> {
  if (hosts === 0 || pending.value !== null || signal?.aborted) {
    return Promise.resolve("cancel");
  }
  return new Promise((resolve) => {
    const entry: PendingConsent = {
      request,
      settle: (choice) => {
        signal?.removeEventListener("abort", onAbort);
        if (pending.value === entry) {
          pending.value = null;
        }
        resolve(choice);
      },
    };
    function onAbort(): void {
      entry.settle("cancel");
    }
    signal?.addEventListener("abort", onAbort, { once: true });
    pending.value = entry;
  });
}

/** For the shell that renders the dialog: the request on screen (or null)
 * and the answer handler. Registers the shell as a host while mounted; an
 * unmount with a request still open answers it "cancel". */
export function useModelConsentHost() {
  onMounted(() => {
    hosts += 1;
  });
  onUnmounted(() => {
    hosts -= 1;
    if (hosts === 0) {
      pending.value?.settle("cancel");
    }
  });
  return {
    request: computed(() => pending.value?.request ?? null),
    answer(choice: ConsentChoice): void {
      pending.value?.settle(choice);
    },
  };
}
