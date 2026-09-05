// zlib-wrapped DEFLATE (RFC 1950) — what PDF's /FlateDecode and PNG's IDAT
// both read. Leans on the platform's CompressionStream rather than a
// dependency, per this project's protected package.json; core/zipWriter.ts
// holds the raw-DEFLATE (RFC 1951) counterpart the ZIP format wants.
//
// Product-neutral plumbing, so it lives here rather than beside either
// caller: lib/pdf's image encoder and its PDF writer both need it, and
// having the writer import it from the image encoder would make those two
// modules circular.

/** Compresses `data` into a zlib stream. */
export async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new CompressionStream("deflate"));

  const chunks: Uint8Array[] = [];
  let total = 0;
  const reader = stream.getReader();
  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    if (value) {
      chunks.push(value);
      total += value.length;
    }
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
