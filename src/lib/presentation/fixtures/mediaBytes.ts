// Byte constants for the SPEC-095 media fixtures. DOM-free (fixtures/build.ts is
// Node-only): a minimal PCM WAV every browser decodes — the playable clip — and
// a ~512-byte synthetic MP4 (`ftyp`/`isom` header, no real codec) whose only job
// is to exercise the unsupported-codec degrade path (no encoder in-repo).
import { bytesToBase64 } from "../../core/packageParts";

function writeAscii(buffer: Uint8Array, offset: number, text: string): void {
  for (let i = 0; i < text.length; i++) {
    buffer[offset + i] = text.charCodeAt(i);
  }
}

/** A tiny mono 8-bit PCM WAV (8 samples): 44-byte RIFF/WAVE header + samples.
 * Every browser decodes PCM WAV, so this is the fixtures' playable clip. */
export function wavBytes(): Uint8Array {
  const samples = new Uint8Array([128, 160, 192, 160, 128, 96, 64, 96]);
  const dataSize = samples.length;
  const buffer = new Uint8Array(44 + dataSize);
  const view = new DataView(buffer.buffer);
  writeAscii(buffer, 0, "RIFF");
  view.setUint32(4, 36 + dataSize, true);
  writeAscii(buffer, 8, "WAVE");
  writeAscii(buffer, 12, "fmt ");
  view.setUint32(16, 16, true); // PCM fmt chunk size
  view.setUint16(20, 1, true); // audioFormat = PCM
  view.setUint16(22, 1, true); // channels = mono
  view.setUint32(24, 8000, true); // sample rate
  view.setUint32(28, 8000, true); // byte rate = rate * blockAlign
  view.setUint16(32, 1, true); // block align
  view.setUint16(34, 8, true); // bits per sample
  writeAscii(buffer, 36, "data");
  view.setUint32(40, dataSize, true);
  buffer.set(samples, 44);
  return buffer;
}

/** A ~512-byte structurally-plausible MP4: an `ftyp` box (major brand `isom`)
 * followed by a `free` padding box, and no `moov`/`mdat` codec data — a browser
 * cannot play it, so it drives the media-unsupported degrade (AC-33). */
export function syntheticMp4Bytes(): Uint8Array {
  const total = 512;
  const buffer = new Uint8Array(total);
  const view = new DataView(buffer.buffer);
  // ftyp box (24 bytes): size, 'ftyp', major 'isom', minor 512, brands isom/mp41
  view.setUint32(0, 24, false); // big-endian box size (ISO-BMFF)
  writeAscii(buffer, 4, "ftyp");
  writeAscii(buffer, 8, "isom");
  view.setUint32(12, 512, false);
  writeAscii(buffer, 16, "isom");
  writeAscii(buffer, 20, "mp41");
  // free box for the remaining padding.
  view.setUint32(24, total - 24, false);
  writeAscii(buffer, 28, "free");
  return buffer;
}

export const WAV_BASE64 = bytesToBase64(wavBytes());
export const SYNTHETIC_MP4_BASE64 = bytesToBase64(syntheticMp4Bytes());

/** A 2×2 solid PNG (75 bytes) reused as the audio clip's poster still. */
export const POSTER_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";
