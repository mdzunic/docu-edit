// The audience window, seen from the presenter window (SPEC-100): opening it,
// placing it, noticing when it goes, and closing it on exit. Everything the
// browser cannot promise degrades to a stated, visible fallback:
//
//   · Pop-ups. `window.open` returning null (blocked) puts the presenter in
//     single-window mode — status "blocked" — whose banner offers to try again
//     (a fresh click usually succeeds) or to present here. So does a browser
//     with no channel transport at all: a second window could not follow.
//   · Screens. The Window Management API is an enhancement only. Where
//     `window.getScreenDetails` exists AND its permission already resolves
//     granted, the window opens on the first non-primary screen through its
//     left/top/width/height features (and its URL says so, `placed=1`).
//     Otherwise it opens as an ordinary pop-up whose page shows the drag
//     hint, and — at most once per browser session, never blocking the open,
//     silent on refusal — the permission is asked for, to serve the next open.
//     `navigator.presentation` is not used at all.
//   · Loss. The audience's `bye` (closed or reloading) or the handle reading
//     `closed` marks it "lost"; its `hello` (a reload coming back, or a reopen)
//     marks it "connected" again — the session re-syncs it with one snapshot.
import { onBeforeUnmount, ref, type Ref } from "vue";
import type { WindowChannelTransport } from "../../lib/core/windowChannel";
import { audienceUrl, audienceWindowName, PRESENTER_HEARTBEAT_MS } from "../../lib/presentation/presenterSession";

export type AudienceStatus = "idle" | "opening" | "connected" | "lost" | "blocked";

/** Where a window goes on another screen (CSS px, screen coordinates). */
export interface AudiencePlacement {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** What the Window Management API says, reduced to what the open needs. */
export interface ScreenAccess {
  state: "granted" | "prompt" | "denied" | "unsupported";
  placement: AudiencePlacement | null;
}

export interface AudienceWindowDeps {
  deckId: Ref<string>;
  session: string;
  /** The 1-based slide the audience should open on. */
  slideNumber: () => number;
  transport: WindowChannelTransport;
  /** Test seams (default: the real window.open and Window Management API). */
  openWindow?: (url: string, name: string, features: string) => Window | null;
  screenAccess?: () => Promise<ScreenAccess>;
  askForScreens?: () => Promise<unknown>;
}

interface ScreenDetailed {
  isPrimary: boolean;
  availLeft: number;
  availTop: number;
  availWidth: number;
  availHeight: number;
}
type ScreenDetailsWindow = Window & { getScreenDetails?: () => Promise<{ screens: ScreenDetailed[] }> };

const SCREENS_ASKED_KEY = "docu-edit:presenter-screens-asked";
let screensAskedInMemory = false;

/** A permission query by a name TypeScript's DOM types may not know yet
 * ("window-placement" is the Window Management permission's name in Chromes
 * before 111). */
function queryPermission(name: string): Promise<PermissionStatus> {
  return navigator.permissions.query({ name: name as PermissionName });
}

/** The permission state and, when granted, the first non-primary screen.
 * Never throws: an absent API, an unknown permission name, a single screen
 * or any error all read as "no placement". */
export async function defaultScreenAccess(): Promise<ScreenAccess> {
  const win = window as ScreenDetailsWindow;
  if (typeof win.getScreenDetails !== "function") return { state: "unsupported", placement: null };
  try {
    const status = await queryPermission("window-management").catch(() => queryPermission("window-placement"));
    if (status.state !== "granted") return { state: status.state, placement: null };
    const details = await win.getScreenDetails();
    const screen = details.screens.find((candidate) => candidate.isPrimary === false);
    return {
      state: "granted",
      placement: screen
        ? { left: screen.availLeft, top: screen.availTop, width: screen.availWidth, height: screen.availHeight }
        : null,
    };
  } catch {
    return { state: "unsupported", placement: null };
  }
}

/** Whether the permission was already asked for in this browser session. */
function screensAsked(): boolean {
  if (screensAskedInMemory) return true;
  try {
    return sessionStorage.getItem(SCREENS_ASKED_KEY) === "1";
  } catch {
    return false;
  }
}

function markScreensAsked(): void {
  screensAskedInMemory = true;
  try {
    sessionStorage.setItem(SCREENS_ASKED_KEY, "1");
  } catch {
    // The in-memory flag still holds for this page.
  }
}

/** The ordinary pop-up's features: a 16:9 window offset from this one. */
function ordinaryFeatures(): string {
  const left = Math.round((window.screenX || 0) + 40);
  const top = Math.round((window.screenY || 0) + 40);
  return `popup=yes,width=960,height=540,left=${left},top=${top}`;
}

function placedFeatures(placement: AudiencePlacement): string {
  const { left, top, width, height } = placement;
  return `popup=yes,left=${Math.round(left)},top=${Math.round(top)},width=${Math.round(width)},height=${Math.round(height)}`;
}

export function useAudienceWindow(deps: AudienceWindowDeps) {
  const status = ref<AudienceStatus>("idle");
  let handle: Window | null = null;
  let poll = 0;
  let closing = false;

  function stopPolling(): void {
    if (poll) clearInterval(poll);
    poll = 0;
  }

  /** Opens (or re-opens, by name: the same window) the audience window on the
   * current slide. Called once on load and again from a click on
   * "Open audience window". */
  async function open(): Promise<void> {
    if (deps.transport === "none") {
      status.value = "blocked";
      return;
    }
    const access = await (deps.screenAccess ?? defaultScreenAccess)();
    const placement = access.placement;
    const url = audienceUrl(deps.deckId.value, deps.slideNumber(), deps.session, placement !== null);
    const features = placement ? placedFeatures(placement) : ordinaryFeatures();
    const opener = deps.openWindow ?? ((u: string, n: string, f: string) => window.open(u, n, f));
    let opened: Window | null = null;
    try {
      opened = opener(url, audienceWindowName(deps.deckId.value), features);
    } catch {
      opened = null;
    }
    if (!opened) {
      status.value = "blocked";
      return;
    }
    handle = opened;
    closing = false;
    if (status.value !== "connected") status.value = "opening";
    stopPolling();
    poll = window.setInterval(() => {
      if (handle?.closed && !closing) {
        status.value = "lost";
        stopPolling();
      }
    }, PRESENTER_HEARTBEAT_MS);
    if (access.state === "prompt" && !screensAsked()) {
      markScreensAsked();
      const ask = deps.askForScreens ?? (() => (window as ScreenDetailsWindow).getScreenDetails?.() ?? Promise.resolve());
      void Promise.resolve()
        .then(ask)
        .catch(() => {});
    }
  }

  /** The session heard from an audience window of this session. */
  function onAudience(event: "hello" | "bye"): void {
    if (closing) return;
    status.value = event === "hello" ? "connected" : "lost";
  }

  /** Closes the window this presenter opened (exit). */
  function close(): void {
    closing = true;
    stopPolling();
    try {
      if (handle && !handle.closed) handle.close();
    } catch {
      // A window that already went is fine.
    }
    handle = null;
  }

  onBeforeUnmount(stopPolling);

  return { status, open, close, onAudience };
}
