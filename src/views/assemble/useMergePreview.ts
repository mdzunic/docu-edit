// The page organiser's result preview (SPEC-111): the planned output, built
// into a blob: URL for the browser's own PDF viewer (SPEC-107's <object> —
// page cards render nothing of their own). It rebuilds a moment after the
// plan changes, at most once every PREVIEW_DEBOUNCE_MS; never by itself
// past PREVIEW_MAX_PAGES pages; and whenever Refresh is pressed (still
// within the debounce). A newer build supersedes an older one, and a URL is
// released when it is replaced and when the workspace closes. usePdfMerge
// supplies what to build.
import { computed, onScopeDispose, ref, watch, type ComputedRef, type Ref } from "vue";
import type { PdfDropRecord } from "../../lib/core/pdfDocument";

/** A change rebuilds the preview at most this often. */
export const PREVIEW_DEBOUNCE_MS = 750;
/** Past this many pages the preview is rebuilt only on request. */
export const PREVIEW_MAX_PAGES = 200;

export interface MergePreviewDeps {
  /** What the preview shows, as one comparable string ("" for nothing). */
  signature: ComputedRef<string>;
  pageCount: ComputedRef<number>;
  /** Builds the planned PDF; null when `stop` ended it. */
  build: (stop: () => boolean) => Promise<{ bytes: Uint8Array; dropped: PdfDropRecord[] } | null>;
  /** Each finished build's drop records ([] when the preview is cleared). */
  onDrops: (dropped: PdfDropRecord[]) => void;
  previewUrl: (bytes: Uint8Array) => string;
  revokeUrl: (url: string) => void;
  now: () => number;
  /** Turns an error into the words the panel shows. */
  describe: (cause: unknown) => string | null;
}

export interface PdfMergePreview {
  /** The planned output's blob: URL ("" when none is built). */
  url: Ref<string>;
  busy: Ref<boolean>;
  failure: Ref<string>;
  /** The plan changed since the preview shown was built. */
  stale: Ref<boolean>;
  /** Over PREVIEW_MAX_PAGES: rebuilt only on request. */
  manual: ComputedRef<boolean>;
  /** merge-refresh: rebuild now (or as soon as the debounce allows). */
  refresh: () => void;
}

export function useMergePreview(deps: MergePreviewDeps): PdfMergePreview {
  const url = ref("");
  const busy = ref(false);
  const failure = ref("");
  const stale = ref(false);
  const manual = computed(() => deps.pageCount.value > PREVIEW_MAX_PAGES);
  let timer: ReturnType<typeof setTimeout> | null = null;
  let lastBuiltAt: number | null = null;
  let token = 0;
  let version = 0;

  function clearTimer(): void {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  /** How long until the debounce allows the next build. */
  function untilAllowed(): number {
    return lastBuiltAt === null ? 0 : lastBuiltAt + PREVIEW_DEBOUNCE_MS - deps.now();
  }

  function show(next: string): void {
    if (url.value) {
      deps.revokeUrl(url.value);
    }
    url.value = next;
  }

  async function build(): Promise<void> {
    timer = null;
    const mine = ++token;
    const planned = version;
    lastBuiltAt = deps.now();
    busy.value = true;
    failure.value = "";
    const stop = () => mine !== token;
    try {
      const result = await deps.build(stop);
      if (!result || stop()) {
        return;
      }
      show(deps.previewUrl(result.bytes));
      deps.onDrops(result.dropped);
      stale.value = planned !== version;
    } catch (cause) {
      const reason = stop() ? null : deps.describe(cause);
      if (reason !== null) {
        failure.value = reason;
      }
    } finally {
      if (!stop()) {
        busy.value = false;
      }
    }
  }

  function clear(): void {
    clearTimer();
    token += 1;
    show("");
    busy.value = false;
    failure.value = "";
    stale.value = false;
    deps.onDrops([]);
  }

  watch(
    deps.signature,
    (signature) => {
      version += 1;
      if (!signature) {
        clear();
        return;
      }
      stale.value = true;
      clearTimer();
      // Never automatically past PREVIEW_MAX_PAGES: Refresh builds it.
      if (!manual.value) {
        timer = setTimeout(build, Math.max(PREVIEW_DEBOUNCE_MS, untilAllowed()));
      }
    },
    { immediate: true },
  );

  // Leaving the workspace frees the URL and drops a build in flight.
  onScopeDispose(clear);

  return {
    url,
    busy,
    failure,
    stale,
    manual,
    refresh: () => {
      clearTimer();
      timer = setTimeout(build, Math.max(0, untilAllowed()));
    },
  };
}
