// The product-neutral half of SPEC-008's equality relations: byte identity,
// used for preserved opaque parts in any package format. The model-shaped
// relations (`equalStructural`, `equalSemantic`) stay in lib/equality.ts —
// they are typed against the document model; canonical XML comparison lives
// in core/xml.ts.

/** Byte identity between two buffers — the relation preserved opaque parts
 * are compared with before and after a save. */
export function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}
