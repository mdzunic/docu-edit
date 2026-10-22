// IWA chunk framing and the archive object stream (SPEC-087). An `.iwa` file
// is a sequence of chunks — 1 header byte 0x00, a 3-byte little-endian
// payload length, then a raw Snappy block — whose decoded payloads
// concatenate into one object stream of `[varint length][TSP.ArchiveInfo]
// [payload messages]` records (verified format facts #2 and #4). DOM-free
// and Node-runnable (AC-45); every walk is a loop, never recursion (AC-44).
import { KeynoteOpenError } from "./errors";
import { snappyDecode } from "./snappy";
import { fields, readVarint, WIRE_LENGTH_DELIMITED, WIRE_VARINT, type ProtoCursor } from "./protobuf";

/** Decodes an `.iwa` file's chunks into the single concatenated object
 * stream. A non-0x00 chunk header or a declared length running past the end
 * raises the typed error (AC-42). */
export function decodeIwa(bytes: Uint8Array): Uint8Array {
  const pieces: Uint8Array[] = [];
  let total = 0;
  let pos = 0;
  while (pos < bytes.length) {
    if (bytes[pos] !== 0x00) {
      throw new KeynoteOpenError(`Corrupt IWA file: unknown chunk type 0x${bytes[pos].toString(16)}`);
    }
    if (pos + 4 > bytes.length) {
      throw new KeynoteOpenError("Truncated IWA file: chunk header runs past the end");
    }
    const length = bytes[pos + 1] | (bytes[pos + 2] << 8) | (bytes[pos + 3] << 16);
    pos += 4;
    if (pos + length > bytes.length) {
      throw new KeynoteOpenError("Truncated IWA file: chunk payload runs past the end");
    }
    const piece = snappyDecode(bytes.subarray(pos, pos + length));
    pieces.push(piece);
    total += piece.length;
    pos += length;
  }
  const stream = new Uint8Array(total);
  let offset = 0;
  for (const piece of pieces) {
    stream.set(piece, offset);
    offset += piece.length;
  }
  return stream;
}

/** One payload message of an archived object, tagged with its iWork archive
 * type number (KN.DocumentArchive = 1, KN.ShowArchive = 2, …). */
export interface ArchiveMessage {
  type: number;
  payload: Uint8Array;
}

/** Parses an object stream into its archives, keyed by object identifier.
 * TSP.ArchiveInfo: field 1 = identifier, field 2 = repeated MessageInfo
 * (field 1 = type, field 3 = payload length — consumed consecutively after
 * the info block). Map insertion order is stream order, which callers rely
 * on to find a file's root archive. */
export function parseArchives(stream: Uint8Array): Map<bigint, ArchiveMessage[]> {
  const archives = new Map<bigint, ArchiveMessage[]>();
  const cursor: ProtoCursor = { bytes: stream, pos: 0 };
  while (cursor.pos < stream.length) {
    const infoLength = readVarint(cursor);
    if (cursor.pos + infoLength > stream.length) {
      throw new KeynoteOpenError("Truncated object stream: ArchiveInfo runs past the end");
    }
    const info = fields(stream.subarray(cursor.pos, cursor.pos + infoLength));
    cursor.pos += infoLength;

    let identifier: number | undefined;
    const messages: ArchiveMessage[] = [];
    for (const field of info) {
      if (field.num === 1 && field.wire === WIRE_VARINT && field.value !== undefined) {
        identifier = field.value;
      } else if (field.num === 2 && field.wire === WIRE_LENGTH_DELIMITED && field.bytes) {
        let type: number | undefined;
        let length: number | undefined;
        for (const messageField of fields(field.bytes)) {
          if (messageField.num === 1 && messageField.wire === WIRE_VARINT) {
            type = messageField.value;
          } else if (messageField.num === 3 && messageField.wire === WIRE_VARINT) {
            length = messageField.value;
          }
        }
        if (type === undefined || length === undefined) {
          throw new KeynoteOpenError("Corrupt object stream: MessageInfo without type or length");
        }
        if (cursor.pos + length > stream.length) {
          throw new KeynoteOpenError("Truncated object stream: payload runs past the end");
        }
        messages.push({ type, payload: stream.subarray(cursor.pos, cursor.pos + length) });
        cursor.pos += length;
      }
    }
    if (identifier === undefined) {
      throw new KeynoteOpenError("Corrupt object stream: ArchiveInfo without an identifier");
    }
    archives.set(BigInt(identifier), messages);
  }
  return archives;
}
