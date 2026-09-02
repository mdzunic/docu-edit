// SPEC-095 golden fixtures for the odp media wire (draw:plugin), in SPEC-086's
// golden shape. Re-exports the pptx GOLDEN_MEDIA_FRAMES so both formats trace to
// one logical set of clips; the odp goldens carry odp part paths (the reader
// keys media by href, not the pptx `ppt/media/` path). Authored against ODF 1.2
// §10.4.7/§14.9; the draw:param spellings are this app's canonical choice, with
// the reader also tolerant of Impress's VolumeDB (see reader.ts).
import type { MediaFrame } from "../model";

export { GOLDEN_MEDIA_FRAMES } from "../pptx/mediaGolden";

export const ODP_GOLDEN_MEDIA_VIDEO: Omit<MediaFrame, "id"> = {
  kind: "media",
  media: "video",
  xCm: 2,
  yCm: 2,
  widthCm: 16,
  heightCm: 9,
  mediaId: "media/media1.mp4",
  posterMediaId: "Pictures/image1.png",
  contentType: "video/mp4",
  autoStart: true,
  loop: true,
  muted: true,
  volume: 80,
  fullScreen: true,
};

export const ODP_GOLDEN_MEDIA_AUDIO: Omit<MediaFrame, "id"> = {
  kind: "media",
  media: "audio",
  xCm: 1,
  yCm: 1,
  widthCm: 8,
  heightCm: 1.6,
  mediaId: "media/media2.wav",
  contentType: "audio/wav",
};

/** The canonical rebuilt video `draw:frame` + `draw:plugin`. */
export const MEDIA_PLUGIN_GOLDEN =
  `<draw:frame svg:x="2cm" svg:y="2cm" svg:width="16cm" svg:height="9cm" presentation:play-full-screen="true">` +
  `<draw:plugin xlink:href="media/media1.mp4" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="video/mp4">` +
  `<draw:param draw:name="AutoPlay" draw:value="true"/>` +
  `<draw:param draw:name="Loop" draw:value="true"/>` +
  `<draw:param draw:name="Mute" draw:value="true"/>` +
  `<draw:param draw:name="Volume" draw:value="80"/>` +
  `</draw:plugin>` +
  `<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `</draw:frame>`;

/** The canonical rebuilt audio `draw:frame` (no poster, no flags). */
export const MEDIA_PLUGIN_AUDIO_GOLDEN =
  `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="8cm" svg:height="1.6cm">` +
  `<draw:plugin xlink:href="media/media2.wav" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="audio/wav">` +
  `</draw:plugin>` +
  `</draw:frame>`;

/** An Impress-flavoured variant: a media-typed `draw:object` (not `draw:plugin`),
 * reordered attributes, and a `VolumeDB` (decibel) param the reader converts. */
export const MEDIA_PLUGIN_VARIANT =
  `<draw:frame svg:width="16cm" svg:height="9cm" svg:x="2cm" svg:y="2cm">` +
  `<draw:object draw:mime-type="video/mp4" xlink:href="media/media1.mp4" xlink:type="simple">` +
  `<draw:param draw:value="true" draw:name="Loop"/>` +
  `<draw:param draw:name="VolumeDB" draw:value="-6"/>` +
  `</draw:object></draw:frame>`;

/** A `draw:frame` whose only child is an unresolvable `draw:image`: not media
 * (no plugin/media-object), so it stays the SPEC-079 "Media" ForeignFrame. */
export const NON_MEDIA_FRAME =
  `<draw:frame svg:x="2cm" svg:y="2cm" svg:width="4cm" svg:height="4cm">` +
  `<draw:image xlink:href="Pictures/missing.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `</draw:frame>`;
