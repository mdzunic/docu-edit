// Raw Snappy block decoder (SPEC-087). An IWA chunk's payload is one raw
// Snappy block — a varint uncompressed length followed by literal/copy tags —
// with no framing-format stream identifier and no CRC (verified format fact
// #3). Iterative by construction (AC-44) and DOM-free (AC-45): plain byte
// arithmetic over Uint8Array, runnable under Node vitest as-is.
import { KeynoteOpenError } from "./errors";

const TAG_LITERAL = 0;
const TAG_COPY_1 = 1; // 3-bit length, 11-bit offset
const TAG_COPY_2 = 2; // 16-bit LE offset
// tag 3: copy with 32-bit LE offset

/** Decodes one raw Snappy block. Truncated input — a tag, length byte,
 * offset or literal running past the end, or output not reaching the
 * declared uncompressed length — raises the typed error (AC-42). */
export function snappyDecode(bytes: Uint8Array): Uint8Array {
  let pos = 0;

  // Varint uncompressed length.
  let expected = 0;
  let shift = 0;
  for (;;) {
    if (pos >= bytes.length) {
      throw new KeynoteOpenError("Truncated Snappy block: unterminated length varint");
    }
    const byte = bytes[pos];
    pos += 1;
    expected += (byte & 0x7f) * 2 ** shift;
    if ((byte & 0x80) === 0) {
      break;
    }
    shift += 7;
    if (shift > 35) {
      throw new KeynoteOpenError("Truncated Snappy block: oversized length varint");
    }
  }

  const out = new Uint8Array(expected);
  let outPos = 0;

  while (pos < bytes.length) {
    const tag = bytes[pos];
    pos += 1;
    const kind = tag & 3;

    if (kind === TAG_LITERAL) {
      let length = (tag >> 2) + 1;
      if (length > 60) {
        // 61..64 encode the literal length in the next 1..4 bytes, LE.
        const lengthBytes = length - 60;
        if (pos + lengthBytes > bytes.length) {
          throw new KeynoteOpenError("Truncated Snappy block: literal length runs past the end");
        }
        length = 0;
        for (let i = 0; i < lengthBytes; i++) {
          length += bytes[pos + i] * 2 ** (8 * i);
        }
        length += 1;
        pos += lengthBytes;
      }
      if (pos + length > bytes.length || outPos + length > expected) {
        throw new KeynoteOpenError("Truncated Snappy block: literal runs past the end");
      }
      out.set(bytes.subarray(pos, pos + length), outPos);
      pos += length;
      outPos += length;
      continue;
    }

    let length: number;
    let offset: number;
    if (kind === TAG_COPY_1) {
      if (pos >= bytes.length) {
        throw new KeynoteOpenError("Truncated Snappy block: copy offset runs past the end");
      }
      length = ((tag >> 2) & 0x7) + 4;
      offset = ((tag >> 5) << 8) | bytes[pos];
      pos += 1;
    } else if (kind === TAG_COPY_2) {
      if (pos + 2 > bytes.length) {
        throw new KeynoteOpenError("Truncated Snappy block: copy offset runs past the end");
      }
      length = (tag >> 2) + 1;
      offset = bytes[pos] | (bytes[pos + 1] << 8);
      pos += 2;
    } else {
      if (pos + 4 > bytes.length) {
        throw new KeynoteOpenError("Truncated Snappy block: copy offset runs past the end");
      }
      length = (tag >> 2) + 1;
      offset = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16)) + bytes[pos + 3] * 0x1000000;
      pos += 4;
    }

    if (offset === 0 || offset > outPos) {
      throw new KeynoteOpenError("Corrupt Snappy block: copy offset before the start of output");
    }
    if (outPos + length > expected) {
      throw new KeynoteOpenError("Corrupt Snappy block: copy runs past the declared length");
    }
    // Byte-by-byte on purpose: offsets smaller than the length are legal and
    // repeat the just-written bytes (RLE-style), which a block memcpy breaks.
    for (let i = 0; i < length; i++) {
      out[outPos] = out[outPos - offset];
      outPos += 1;
    }
  }

  if (outPos !== expected) {
    throw new KeynoteOpenError("Truncated Snappy block: output shorter than the declared length");
  }
  return out;
}
