// Byte-builders for SPEC-087's unit tests: an independent, minimal encoder
// for protobuf messages, raw Snappy blocks, IWA chunks and whole synthetic
// iWork packages. Test-only (imported by the keynote specs alone), living
// beside its subject under a non-spec name like fixtures.captions.ts /
// fixtures.fields.ts do. Deliberately NOT the decoder run backwards — tags
// and varints are written by their own arithmetic so the specs cross-check
// the reader rather than mirror it. The real-file contract lives in
// readerFixture.spec.ts against e2e/fixtures/mini-deck.key.
import { zip, type ZipEntryInput } from "../../core/zipWriter";

const encoder = new TextEncoder();

export function concatBytes(...parts: (Uint8Array | number[])[]): Uint8Array {
  const arrays = parts.map((part) => (Array.isArray(part) ? Uint8Array.from(part) : part));
  const out = new Uint8Array(arrays.reduce((sum, arr) => sum + arr.length, 0));
  let offset = 0;
  for (const arr of arrays) {
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

export function varint(value: number): number[] {
  const bytes: number[] = [];
  let rest = value;
  for (;;) {
    const low = rest % 128;
    rest = Math.floor(rest / 128);
    if (rest === 0) {
      bytes.push(low);
      return bytes;
    }
    bytes.push(low | 0x80);
  }
}

export function varintField(num: number, value: number): Uint8Array {
  return concatBytes(varint(num * 8), varint(value));
}

export function lengthField(num: number, payload: Uint8Array | number[]): Uint8Array {
  const bytes = Array.isArray(payload) ? Uint8Array.from(payload) : payload;
  return concatBytes(varint(num * 8 + 2), varint(bytes.length), bytes);
}

export function stringField(num: number, text: string): Uint8Array {
  return lengthField(num, encoder.encode(text));
}

export function float32Field(num: number, value: number): Uint8Array {
  const view = new DataView(new ArrayBuffer(4));
  view.setFloat32(0, value, true);
  return concatBytes(varint(num * 8 + 5), new Uint8Array(view.buffer));
}

export function message(...fieldBytes: (Uint8Array | number[])[]): Uint8Array {
  return concatBytes(...fieldBytes);
}

/** A TSP.Reference-shaped sub-message at `num`: `{1: varint id}`. */
export function reference(num: number, id: number): Uint8Array {
  return lengthField(num, varintField(1, id));
}

/** A TSD.GeometryArchive: position {1: x, 2: y} + size {1: w, 2: h}, all
 * float32, in Keynote points. */
export function geometry(x: number, y: number, w: number, h: number): Uint8Array {
  return message(
    lengthField(1, message(float32Field(1, x), float32Field(2, y))),
    lengthField(2, message(float32Field(1, w), float32Field(2, h))),
  );
}

export interface ArchiveInput {
  id: number;
  type: number;
  payload: Uint8Array;
}

/** One object stream record: `[varint length][ArchiveInfo][payload]`.
 * ArchiveInfo: field 1 = id, field 2 = MessageInfo {1: type, 3: length}. */
export function archiveRecord(archive: ArchiveInput): Uint8Array {
  const info = message(
    varintField(1, archive.id),
    lengthField(2, message(varintField(1, archive.type), varintField(3, archive.payload.length))),
  );
  return concatBytes(varint(info.length), info, archive.payload);
}

/** A raw Snappy block that encodes `payload` as plain literals. */
export function snappyLiteralBlock(payload: Uint8Array): Uint8Array {
  const pieces: (Uint8Array | number[])[] = [varint(payload.length)];
  for (let pos = 0; pos < payload.length; pos += 60) {
    const slice = payload.subarray(pos, Math.min(pos + 60, payload.length));
    pieces.push([(slice.length - 1) << 2], slice);
  }
  return concatBytes(...pieces);
}

/** Wraps an object stream as an `.iwa` file: chunks of 0x00 + 3-byte LE
 * length + Snappy block. `chunkSplit` (byte offset into the stream) forces
 * a two-chunk file to exercise the cross-chunk path. */
export function iwaFile(stream: Uint8Array, chunkSplit?: number): Uint8Array {
  const chunks =
    chunkSplit !== undefined && chunkSplit > 0 && chunkSplit < stream.length
      ? [stream.subarray(0, chunkSplit), stream.subarray(chunkSplit)]
      : [stream];
  return concatBytes(
    ...chunks.map((piece) => {
      const block = snappyLiteralBlock(piece);
      return concatBytes([0x00, block.length & 0xff, (block.length >> 8) & 0xff, (block.length >> 16) & 0xff], block);
    }),
  );
}

export function iwaFileOf(archives: ArchiveInput[], chunkSplit?: number): Uint8Array {
  return iwaFile(concatBytes(...archives.map(archiveRecord)), chunkSplit);
}

/** An 8×8 PNG, byte-for-byte the checked-in e2e/fixtures/fixture-image.png
 * (the AppleScript generator writes the same base64). */
export const FIXTURE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGN4VqeEFTEMLQkASgJhgTHXKdIAAAAASUVORK5CYII=";

export function fixturePngBytes(): Uint8Array {
  const binary = atob(FIXTURE_PNG_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// ---------------------------------------------------------------------------
// A complete synthetic two-slide deck mirroring the fixture contract.

export const SYNTHETIC_IMAGE_GEOMETRY = { x: 1300, y: 620, w: 300, h: 300 };

export interface SyntheticDeckOptions {
  /** Nest the object files inside an `Index.zip` (the older layout, AC-6). */
  nested?: boolean;
  /** Give slide 1 a non-empty transition message (field 4) — a counted
   * dropped effect. */
  withTransition?: boolean;
  /** Add an extra drawable that is neither text nor image on slide 1 — a
   * counted dropped object. */
  withForeignDrawable?: boolean;
  /** Point the image's data reference at a data id with no Data/ entry. */
  unresolvableImage?: boolean;
  /** Make the root slide node's child list point back at itself too — the
   * visited-set must terminate (AC-43). */
  withCycle?: boolean;
  /** Override the show size (points). */
  size?: { w: number; h: number };
}

/** The object graph of a synthetic deck: two slides — slide 1 title
 * "Keynote fixture deck" + body "Imported by docu-edit"; slide 2 title
 * "Second slide", body "First bullet point\nSecond bullet point", one PNG
 * image drawable and a speaker note — matching SPEC-087's content contract
 * so the unit assertions and the real-fixture assertions read alike. */
export function syntheticDeckArchives(options: SyntheticDeckOptions = {}): {
  document: ArchiveInput[];
  metadata: ArchiveInput[];
} {
  const size = options.size ?? { w: 1920, h: 1080 };
  const dataId = 77;

  const placeholder = (geo: Uint8Array, storageId: number): Uint8Array =>
    // KN.PlaceholderArchive: field 1 super → {1: {1: geometry}, 2: storage ref}.
    message(lengthField(1, message(lengthField(1, message(lengthField(1, geo))), reference(2, storageId))));

  const rootChildren = [reference(1, 11), reference(1, 12)];
  if (options.withCycle) {
    rootChildren.push(reference(1, 10));
  }

  const slide1Fields: (Uint8Array | number[])[] = [
    reference(5, 31),
    reference(6, 32),
    lengthField(7, varintField(1, 31)),
    lengthField(7, varintField(1, 32)),
  ];
  if (options.withTransition) {
    slide1Fields.push(lengthField(4, message(varintField(1, 1), stringField(2, "apple:magic-move"))));
  }
  if (options.withForeignDrawable) {
    slide1Fields.push(lengthField(7, varintField(1, 36)));
  }

  const document: ArchiveInput[] = [
    { id: 1, type: 1, payload: message(reference(2, 3)) },
    {
      id: 3,
      type: 2,
      payload: message(
        lengthField(3, message(reference(2, 10))),
        lengthField(4, message(float32Field(1, size.w), float32Field(2, size.h))),
      ),
    },
    { id: 10, type: 4, payload: message(...rootChildren) },
    { id: 11, type: 4, payload: message(reference(2, 21)) },
    { id: 12, type: 4, payload: message(reference(2, 22)) },
    { id: 21, type: 5, payload: message(...slide1Fields) },
    {
      id: 22,
      type: 5,
      payload: message(
        reference(5, 33),
        reference(6, 34),
        lengthField(7, varintField(1, 33)),
        lengthField(7, varintField(1, 34)),
        lengthField(7, varintField(1, 35)),
        reference(27, 50),
      ),
    },
    { id: 31, type: 7, payload: placeholder(geometry(154, 87, 1612, 173), 41) },
    { id: 32, type: 7, payload: placeholder(geometry(154, 324, 1612, 648), 42) },
    { id: 33, type: 7, payload: placeholder(geometry(154, 87, 1612, 173), 43) },
    { id: 34, type: 7, payload: placeholder(geometry(154, 324, 1000, 648), 44) },
    {
      id: 35,
      // An image drawable: geometry on the super chain, plus a strict data
      // reference (field 3) resolved through the metadata index.
      type: 3005,
      payload: message(
        lengthField(
          1,
          message(
            lengthField(
              1,
              message(
                lengthField(
                  1,
                  geometry(
                    SYNTHETIC_IMAGE_GEOMETRY.x,
                    SYNTHETIC_IMAGE_GEOMETRY.y,
                    SYNTHETIC_IMAGE_GEOMETRY.w,
                    SYNTHETIC_IMAGE_GEOMETRY.h,
                  ),
                ),
              ),
            ),
          ),
        ),
        reference(3, options.unresolvableImage ? 99 : dataId),
      ),
    },
    // A drawable with neither storage nor data ref (e.g. a shape) — dropped.
    { id: 36, type: 3006, payload: message(varintField(9, 1)) },
    { id: 41, type: 2001, payload: message(stringField(3, "Keynote fixture deck")) },
    { id: 42, type: 2001, payload: message(stringField(3, "Imported by docu-edit")) },
    { id: 43, type: 2001, payload: message(stringField(3, "Second slide")) },
    { id: 44, type: 2001, payload: message(stringField(3, "First bullet point\nSecond bullet point")) },
    // The notes container: an unknown archive type whose graph reaches a
    // storage one reference hop away (AC-41's "unknown types ignored").
    { id: 50, type: 15, payload: message(lengthField(1, message(reference(2, 51)))) },
    { id: 51, type: 2001, payload: message(stringField(3, "A speaker note for the importer")) },
  ];

  const metadata: ArchiveInput[] = [
    {
      id: 2,
      type: 11006,
      payload: message(
        lengthField(3, message(varintField(1, 1), stringField(3, "Document"))),
        lengthField(4, message(varintField(1, dataId), stringField(3, "fixture-image.png"))),
      ),
    },
  ];

  return { document, metadata };
}

/** Zips a synthetic deck into a `.key`-shaped package (both layouts). */
export async function buildSyntheticKeynote(options: SyntheticDeckOptions = {}): Promise<Uint8Array> {
  const { document, metadata } = syntheticDeckArchives(options);
  const documentIwa = iwaFileOf(document, 40);
  const metadataIwa = iwaFileOf(metadata);
  const image = fixturePngBytes();

  const entries: ZipEntryInput[] = [];
  if (options.nested) {
    const nested = await zip([
      { name: "Document.iwa", data: documentIwa },
      { name: "Metadata.iwa", data: metadataIwa },
    ]);
    entries.push({ name: "Index.zip", data: nested, store: true });
  } else {
    entries.push({ name: "Index/Document.iwa", data: documentIwa });
    entries.push({ name: "Index/Metadata.iwa", data: metadataIwa });
  }
  entries.push({ name: `Data/fixture-image-77.png`, data: image });
  entries.push({ name: "Metadata/DocumentIdentifier", data: new TextEncoder().encode("synthetic") });
  return zip(entries);
}

/** Zips arbitrary archives into a package — for specs composing bespoke
 * graphs (role mapping, text cleaning, rejection cases) from the
 * primitives above. */
export async function buildKeynoteFromArchives(
  document: ArchiveInput[],
  metadata: ArchiveInput[] = [],
  extraEntries: ZipEntryInput[] = [],
): Promise<Uint8Array> {
  const entries: ZipEntryInput[] = [{ name: "Index/Document.iwa", data: iwaFileOf(document) }];
  if (metadata.length > 0) {
    entries.push({ name: "Index/Metadata.iwa", data: iwaFileOf(metadata) });
  }
  entries.push(...extraEntries);
  return zip(entries);
}

/** A stress deck: `slideCount` slides, each a node under the root with one
 * one-line body placeholder — the AC-44 iterative-decode case. */
export async function buildStressKeynote(slideCount: number): Promise<Uint8Array> {
  const archives: ArchiveInput[] = [
    { id: 1, type: 1, payload: message(reference(2, 3)) },
    {
      id: 3,
      type: 2,
      payload: message(
        lengthField(3, message(reference(2, 10))),
        lengthField(4, message(float32Field(1, 1920), float32Field(2, 1080))),
      ),
    },
  ];
  const rootChildren: Uint8Array[] = [];
  for (let i = 0; i < slideCount; i++) {
    const nodeId = 100 + i * 4;
    rootChildren.push(reference(1, nodeId));
    archives.push({ id: nodeId, type: 4, payload: message(reference(2, nodeId + 1)) });
    archives.push({
      id: nodeId + 1,
      type: 5,
      payload: message(reference(6, nodeId + 2), lengthField(7, varintField(1, nodeId + 2))),
    });
    archives.push({
      id: nodeId + 2,
      type: 7,
      payload: message(lengthField(1, message(lengthField(1, message(lengthField(1, geometry(10, 10, 500, 100)))), reference(2, nodeId + 3)))),
    });
    archives.push({ id: nodeId + 3, type: 2001, payload: message(stringField(3, `Slide ${i + 1}`)) });
  }
  archives.push({ id: 10, type: 4, payload: message(...rootChildren) });
  return zip([{ name: "Index/Document.iwa", data: iwaFileOf(archives) }]);
}
