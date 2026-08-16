// The backbone's feature-lifecycle hook registry. Extracted feature
// composables register callbacks the backbone runs at the points they care
// about (session close, unmount, scroll, post-edit find refresh), so the
// backbone need not import feature modules and there are no import cycles.
// Registration is idempotent per feature; the Sets are exposed so the
// backbone fires them at the right lifecycle moments.
export interface FeatureHooks {
  sessionCloseHooks: Set<() => void>;
  unmountHooks: Set<() => void>;
  scrollHooks: Set<() => void>;
  onSessionClose: (cb: () => void) => void;
  onUnmount: (cb: () => void) => void;
  onScroll: (cb: () => void) => void;
  registerFindRefresh: (fn: () => void) => void;
  /** Runs the registered post-edit find refresh, if any. */
  fireFindRefresh: () => void;
}

export function useFeatureHooks(): FeatureHooks {
  const sessionCloseHooks = new Set<() => void>();
  const unmountHooks = new Set<() => void>();
  const scrollHooks = new Set<() => void>();
  let findRefreshScheduler: () => void = () => {};

  return {
    sessionCloseHooks,
    unmountHooks,
    scrollHooks,
    onSessionClose: (cb) => {
      sessionCloseHooks.add(cb);
    },
    onUnmount: (cb) => {
      unmountHooks.add(cb);
    },
    onScroll: (cb) => {
      scrollHooks.add(cb);
    },
    registerFindRefresh: (fn) => {
      findRefreshScheduler = fn;
    },
    fireFindRefresh: () => findRefreshScheduler?.(),
  };
}
