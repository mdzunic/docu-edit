// SPEC-095 golden fixtures for the pptx media wire, in SPEC-086's golden shape:
// a hand-authored model, the canonical bytes the serialiser must produce, and
// the slide rels the reader resolves against. Kept apart from the .spec so the
// odp golden can re-export GOLDEN_MEDIA_FRAMES and both formats prove they model
// the same clips. Authored against ECMA-376 §19.3/§19.5; the exact p14:media
// GUID and playback-flag encoding are this app's choice (see mediaWire.ts).
import type { MediaFrame } from "../model";
import type { MediaPart } from "../../core/media";
import type { Relationship } from "../../core/xml";
import {
  AUDIO_RELATIONSHIP,
  IMAGE_RELATIONSHIP_TYPE,
  MEDIA_RELATIONSHIP,
  VIDEO_RELATIONSHIP,
} from "./mediaWire";

/** The modelled clips both format goldens share (the odp golden re-exports
 * this). `id` is added per-frame at read; the goldens compare the rest. */
export const GOLDEN_MEDIA_VIDEO: Omit<MediaFrame, "id"> = {
  kind: "media",
  media: "video",
  xCm: 2,
  yCm: 2,
  widthCm: 16,
  heightCm: 9,
  mediaId: "ppt/media/media1.mp4",
  posterMediaId: "ppt/media/image1.png",
  contentType: "video/mp4",
  autoStart: true,
  loop: true,
  muted: true,
  volume: 80,
  fullScreen: true,
};

export const GOLDEN_MEDIA_AUDIO: Omit<MediaFrame, "id"> = {
  kind: "media",
  media: "audio",
  xCm: 1,
  yCm: 1,
  widthCm: 8,
  heightCm: 1.6,
  mediaId: "ppt/media/media2.wav",
  contentType: "audio/wav",
};

export const GOLDEN_MEDIA_FRAMES: Omit<MediaFrame, "id">[] = [GOLDEN_MEDIA_VIDEO, GOLDEN_MEDIA_AUDIO];

/** The media/poster parts the goldens reference (bytes irrelevant to the wire
 * golden, so empty). */
export const GOLDEN_MEDIA_PARTS: MediaPart[] = [
  { mediaId: "ppt/media/media1.mp4", contentType: "video/mp4", data: "" },
  { mediaId: "ppt/media/media2.wav", contentType: "audio/wav", data: "" },
  { mediaId: "ppt/media/image1.png", contentType: "image/png", data: "" },
];

/** The slide rels the video golden resolves against (the writer mints these). */
export const GOLDEN_VIDEO_RELS: Relationship[] = [
  { id: "rId2", type: IMAGE_RELATIONSHIP_TYPE, target: "../media/image1.png" },
  { id: "rId3", type: VIDEO_RELATIONSHIP, target: "../media/media1.mp4" },
  { id: "rId4", type: MEDIA_RELATIONSHIP, target: "../media/media1.mp4" },
];

export const GOLDEN_AUDIO_RELS: Relationship[] = [
  { id: "rId2", type: AUDIO_RELATIONSHIP, target: "../media/media2.wav" },
  { id: "rId3", type: MEDIA_RELATIONSHIP, target: "../media/media2.wav" },
];

/** The canonical rebuilt video `p:pic` (id=2, the first shape a slide mints). */
export const MEDIA_VIDEO_PIC_GOLDEN =
  `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Video 2"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
  `<p:nvPr><a:videoFile r:link="rId3"/>` +
  `<p:extLst>` +
  `<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId4"/></p:ext>` +
  `<p:ext uri="{7B2F1A6C-3E4D-4F5A-8B9C-0D1E2F3A4B5C}"><de:media xmlns:de="urn:docu-edit:pptx:media:2026" autoStart="1" loop="1" mute="1" vol="80" fullScreen="1"/></p:ext>` +
  `</p:extLst></p:nvPr></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="720000" y="720000"/><a:ext cx="5760000" cy="3240000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

/** The canonical rebuilt audio `p:pic` (id=2, no poster, no flags). */
export const MEDIA_AUDIO_PIC_GOLDEN =
  `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Audio 2"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
  `<p:nvPr><a:audioFile r:link="rId2"/>` +
  `<p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId3"/></p:ext></p:extLst>` +
  `</p:nvPr></p:nvPicPr>` +
  `<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="360000" y="360000"/><a:ext cx="2880000" cy="576000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

/** A PowerPoint-flavoured variant of the embedded video: the `p14:media` embed
 * only (no `a:videoFile`), an `a:hlinkClick` media action, and reordered
 * `cNvPr` attributes — the reader must still model the same clip. */
export const MEDIA_VIDEO_PIC_VARIANT =
  `<p:pic><p:nvPicPr><p:cNvPr name="Media 2" id="2"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="" action="ppaction://media"/></p:cNvPr>` +
  `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
  `<p:nvPr><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId4"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="720000" y="720000"/><a:ext cx="5760000" cy="3240000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

/** A `p:pic` that is NOT media (no video/audio/p14:media, no resolvable blip):
 * stays the SPEC-079 "Media" ForeignFrame, preserved verbatim. */
export const NON_MEDIA_PIC =
  `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="720000" y="720000"/><a:ext cx="1440000" cy="1440000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
