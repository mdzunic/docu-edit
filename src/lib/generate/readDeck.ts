// Byte-based deck reading (SPEC-106), mirroring read.ts: generated `.pptx`/
// `.odp` bytes go straight back into a `PresentationModel` with no `File`, no
// disk and no browser. These wrap presentation/open.ts's
// `readPresentationBytes` — SPEC-080's byte readers, the same assembly the
// app's open path uses — behind nodeDom.ts's DOMParser polyfill, which the
// underlying XML readers need outside a browser (writing never does).
import { readPresentationBytes } from "../presentation/open";
import type { PresentationModel } from "../presentation/model";
import { ensureNodeXmlGlobals } from "./nodeDom";

/** Reads `.pptx` bytes into a `PresentationModel`, in the browser or in Node.
 * `title` defaults the way a blank deck's does. */
export async function readPptx(bytes: Uint8Array | ArrayBuffer, title?: string): Promise<PresentationModel> {
  await ensureNodeXmlGlobals();
  return readPresentationBytes(bytes, "pptx", title);
}

/** The `.odp` counterpart of `readPptx`. */
export async function readOdp(bytes: Uint8Array | ArrayBuffer, title?: string): Promise<PresentationModel> {
  await ensureNodeXmlGlobals();
  return readPresentationBytes(bytes, "odp", title);
}
