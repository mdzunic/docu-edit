// The one browser-dependent decision about an audio/video clip, isolated
// behind an injected probe so it stays product- and DOM-neutral and
// unit-testable (SPEC-095). We never transcode: we play exactly what a plain
// `<video>`/`<audio>` element plays, so whether a clip is playable is the
// element's own `canPlayType` verdict — a clip we cannot play degrades to a
// labelled poster (`media-unsupported`), never an error.

/** Whether the browser reports it can play `contentType`, using an injected
 * `canPlayType`-shaped probe. `canPlayType` returns `"probably"`, `"maybe"`
 * or `""`; anything but `""` is a go. An absent/empty content type — or an
 * empty probe result — is "cannot play", so a linked or codec-less clip
 * resolves to `false` without touching the DOM here. */
export function canPlayMedia(
  contentType: string | undefined,
  probe: (type: string) => string,
): boolean {
  if (!contentType) {
    return false;
  }
  const verdict = probe(contentType);
  return verdict === "probably" || verdict === "maybe";
}
