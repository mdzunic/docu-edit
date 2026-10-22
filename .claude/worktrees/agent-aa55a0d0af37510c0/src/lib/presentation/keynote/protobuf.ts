// A schema-less Protocol-Buffers wire reader (SPEC-087): just enough to walk
// an iWork Archive's messages field-by-field — no schema compiler, no
// codegen. Unknown fields are skipped by wire type (AC-40); truncation
// raises the typed error (AC-42). DOM-free and Node-runnable (AC-45).
import { KeynoteOpenError } from "./errors";

export const WIRE_VARINT = 0;
export const WIRE_FIXED64 = 1;
export const WIRE_LENGTH_DELIMITED = 2;
export const WIRE_FIXED32 = 5;

/** A mutable read position over a byte buffer. */
export interface ProtoCursor {
  bytes: Uint8Array;
  pos: number;
}

/** Reads one varint as an exact JS number. iWork's varints are lengths, ids
 * and type numbers — all far below 2^53 — so a value that cannot be
 * represented exactly is corrupt input, not a supported case. */
export function readVarint(cursor: ProtoCursor): number {
  let result = 0;
  let shift = 0;
  for (;;) {
    if (cursor.pos >= cursor.bytes.length) {
      throw new KeynoteOpenError("Truncated message: unterminated varint");
    }
    const byte = cursor.bytes[cursor.pos];
    cursor.pos += 1;
    result += (byte & 0x7f) * 2 ** shift;
    if ((byte & 0x80) === 0) {
      break;
    }
    shift += 7;
    if (shift > 63) {
      throw new KeynoteOpenError("Corrupt message: varint longer than 64 bits");
    }
  }
  if (!Number.isSafeInteger(result)) {
    throw new KeynoteOpenError("Corrupt message: varint beyond 2^53");
  }
  return result;
}

/** One decoded field. `value` is set for varints; `bytes` holds the raw
 * payload for fixed32 (4 bytes), fixed64 (8 bytes) and length-delimited
 * fields — the caller interprets them (float32, sub-message, string…). */
export interface ProtoField {
  num: number;
  wire: number;
  value?: number;
  bytes?: Uint8Array;
}

/** Decodes a whole message into its fields, in wire order. Groups (wire
 * types 3/4, long-deprecated) have no length and cannot be skipped, so they
 * make the message undecodable — the typed error, not a guess. */
export function fields(message: Uint8Array): ProtoField[] {
  const cursor: ProtoCursor = { bytes: message, pos: 0 };
  const result: ProtoField[] = [];
  while (cursor.pos < message.length) {
    const key = readVarint(cursor);
    const wire = key & 7;
    const num = Math.floor(key / 8);
    if (num === 0) {
      throw new KeynoteOpenError("Corrupt message: field number 0");
    }
    if (wire === WIRE_VARINT) {
      result.push({ num, wire, value: readVarint(cursor) });
    } else if (wire === WIRE_FIXED64) {
      if (cursor.pos + 8 > message.length) {
        throw new KeynoteOpenError("Truncated message: fixed64 runs past the end");
      }
      result.push({ num, wire, bytes: message.subarray(cursor.pos, cursor.pos + 8) });
      cursor.pos += 8;
    } else if (wire === WIRE_LENGTH_DELIMITED) {
      const length = readVarint(cursor);
      if (cursor.pos + length > message.length) {
        throw new KeynoteOpenError("Truncated message: length-delimited field runs past the end");
      }
      result.push({ num, wire, bytes: message.subarray(cursor.pos, cursor.pos + length) });
      cursor.pos += length;
    } else if (wire === WIRE_FIXED32) {
      if (cursor.pos + 4 > message.length) {
        throw new KeynoteOpenError("Truncated message: fixed32 runs past the end");
      }
      result.push({ num, wire, bytes: message.subarray(cursor.pos, cursor.pos + 4) });
      cursor.pos += 4;
    } else {
      throw new KeynoteOpenError(`Undecodable message: unsupported wire type ${wire}`);
    }
  }
  return result;
}

/** A fixed32 field's payload as a little-endian float32 — Keynote's point
 * coordinates and slide size (facts #6/#7 in the spec's Reference). */
export function float32Of(field: ProtoField): number {
  const bytes = field.bytes;
  if (field.wire !== WIRE_FIXED32 || !bytes) {
    throw new KeynoteOpenError("Corrupt message: expected a fixed32 field");
  }
  return new DataView(bytes.buffer, bytes.byteOffset, 4).getFloat32(0, true);
}
