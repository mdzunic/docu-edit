// The pptx audio/video wire vocabulary (SPEC-095), shared by the reader, the
// serialiser and the golden so read and write stay symmetric and there is one
// source of truth for the relationship types and extension URIs. Framework-free
// (imports only core siblings); the DOM-side parse of the flags extension lives
// in reader.ts where the DOM helpers are.
import { escapeXmlAttribute } from "../../core/xml";
import type { MediaFrame } from "../model";

/** OOXML transitional media relationship types (ECMA-376 §15.2.x): a clip's
 * `a:videoFile`/`a:audioFile` points through one of these; the modern embedded
 * pointer under `p14:media` uses the Microsoft `media` type. */
export const VIDEO_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
export const AUDIO_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
export const MEDIA_RELATIONSHIP = "http://schemas.microsoft.com/office/2007/relationships/media";
/** The image relationship type, restated here so the golden (a poster is an
 * image rel) need not reach into reader.ts. */
export const IMAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** The PowerPoint 2010 media extension namespace + the `p:ext` GUID that carries
 * a `p14:media` under a media `p:pic`'s `p:nvPr` (ECMA-376 §19.5). */
export const P14_NAMESPACE = "http://schemas.microsoft.com/office/powerpoint/2010/main";
export const P14_MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/** A docu-edit `p:ext` carrying the SPEC-095 modelled playback flags on a
 * rebuilt media `p:pic`. PowerPoint stores these in the `p:timing` tree (which
 * SPEC-086 owns and keeps preserved); rather than reach into that tree, an
 * in-app edit records the flags in this app-private extension, which round-trips
 * losslessly through our own reader and is ignored by other apps (the `p:ext`
 * mechanism's purpose). A clip untouched since open re-emits its verbatim
 * `sourceXml` and never grows this extension. */
export const DE_MEDIA_NAMESPACE = "urn:docu-edit:pptx:media:2026";
export const DE_MEDIA_EXT_URI = "{7B2F1A6C-3E4D-4F5A-8B9C-0D1E2F3A4B5C}";

/** The playback-flags `p:ext` fragment for a media frame, or "" when it carries
 * no non-default flag (so the golden's "no flags" case is byte-clean). Attribute
 * order matches `parseMediaFlags`. */
export function mediaFlagsExtXml(frame: MediaFrame): string {
  const attrs: string[] = [];
  if (frame.autoStart) {
    attrs.push(`autoStart="1"`);
  }
  if (frame.loop) {
    attrs.push(`loop="1"`);
  }
  if (frame.muted) {
    attrs.push(`mute="1"`);
  }
  if (frame.volume !== undefined) {
    attrs.push(`vol="${frame.volume}"`);
  }
  if (frame.fullScreen) {
    attrs.push(`fullScreen="1"`);
  }
  if (attrs.length === 0) {
    return "";
  }
  return `<p:ext uri="${DE_MEDIA_EXT_URI}"><de:media xmlns:de="${escapeXmlAttribute(DE_MEDIA_NAMESPACE)}" ${attrs.join(" ")}/></p:ext>`;
}

/** The `p:nvPr` content of a rebuilt media `p:pic` (SPEC-095): the
 * `a:videoFile`/`a:audioFile` source ref, the `p14:media` embed pointer, and
 * the docu-edit playback-flags extension. `sourceRid` links the transitional
 * file element (embedded → the video/audio rel, linked → an external rel);
 * `mediaRid` is the modern `p14:media` embed (embedded clips only). */
export function mediaNvPrXml(frame: MediaFrame, rids: { sourceRid?: string; mediaRid?: string }): string {
  const fileTag = frame.media === "audio" ? "a:audioFile" : "a:videoFile";
  // The content type rides on the file element too, so a LINKED clip (no
  // embedded part to sniff it from) recovers it on read.
  const ct = frame.contentType ? ` contentType="${escapeXmlAttribute(frame.contentType)}"` : "";
  const sourceRef = rids.sourceRid
    ? `<${fileTag} r:link="${escapeXmlAttribute(rids.sourceRid)}"${ct}/>`
    : `<${fileTag}${ct}/>`;
  const flags = mediaFlagsExtXml(frame);
  const p14 = rids.mediaRid
    ? `<p:ext uri="${P14_MEDIA_EXT_URI}"><p14:media xmlns:p14="${escapeXmlAttribute(P14_NAMESPACE)}" r:embed="${escapeXmlAttribute(rids.mediaRid)}"/></p:ext>`
    : "";
  const extLst = p14 || flags ? `<p:extLst>${p14}${flags}</p:extLst>` : "";
  return `<p:nvPr>${sourceRef}${extLst}</p:nvPr>`;
}
