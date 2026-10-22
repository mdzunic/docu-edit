// Per-showing GIF URLs: a GIF gets a fresh object URL each time its slide is
// entered (so it restarts, as in PowerPoint), the previous showing's URLs stay
// live for a transition's outgoing stage, and older ones are revoked.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { MediaPart } from "../../lib/core/media";
import { useShowingMedia } from "./useShowingMedia";

const MEDIA: MediaPart[] = [
  { mediaId: "ppt/media/anim.gif", contentType: "image/gif", data: "R0lGODlh" },
  { mediaId: "ppt/media/still.png", contentType: "image/png", data: "iVBORw0K" },
];

let minted = 0;
const revoked: string[] = [];
beforeEach(() => {
  minted = 0;
  revoked.length = 0;
  vi.spyOn(URL, "createObjectURL").mockImplementation(() => `blob:showing/${++minted}`);
  vi.spyOn(URL, "revokeObjectURL").mockImplementation((url) => void revoked.push(url));
});
afterEach(() => vi.restoreAllMocks());

describe("useShowingMedia", () => {
  it("gives a GIF one object URL per showing and leaves other images on their data: URI", () => {
    const m = useShowingMedia();
    m.enter(MEDIA);
    const url = m.currentUrl()!;
    expect(url("ppt/media/anim.gif")).toBe("blob:showing/1");
    expect(url("ppt/media/anim.gif")).toBe("blob:showing/1"); // stable across re-renders
    expect(url("ppt/media/still.png")).toBeUndefined();
    expect(url("ppt/media/missing.gif")).toBeUndefined();
  });

  it("re-entering mints a fresh URL, keeps the previous showing live, and revokes the one before", () => {
    const m = useShowingMedia();
    m.enter(MEDIA);
    const first = m.currentUrl()!("ppt/media/anim.gif");
    m.enter(MEDIA);
    const second = m.currentUrl()!("ppt/media/anim.gif");
    expect(second).not.toBe(first);
    expect(m.previousUrl()!("ppt/media/anim.gif")).toBe(first); // the outgoing stage carries on
    expect(revoked).toEqual([]);
    m.enter(MEDIA);
    expect(revoked).toEqual([first]);
  });

  it("releaseAll revokes every live URL", () => {
    const m = useShowingMedia();
    m.enter(MEDIA);
    const first = m.currentUrl()!("ppt/media/anim.gif");
    m.enter(MEDIA);
    const second = m.currentUrl()!("ppt/media/anim.gif");
    m.releaseAll();
    expect(revoked.toSorted()).toEqual([first, second].toSorted());
    expect(m.currentUrl()).toBeUndefined();
  });

  it("a released showing mints nothing more", () => {
    const m = useShowingMedia();
    m.enter(MEDIA);
    const stale = m.currentUrl()!;
    m.enter(MEDIA);
    m.enter(MEDIA); // the first showing is now released
    expect(stale("ppt/media/anim.gif")).toBeUndefined();
  });
});
