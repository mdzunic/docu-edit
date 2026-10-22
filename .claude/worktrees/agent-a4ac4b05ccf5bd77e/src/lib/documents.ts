// Document lookup. SPEC-001 ships no persistence, so this is deliberately a
// stub: per the spec's Decisions, every id that can appear in a `/doc/:id`
// URL resolves to a document. Only the empty string — which the router can
// never produce — is rejected, so the not-found state stays real, reachable
// code until a later spec replaces this with an actual lookup.
export function documentExists(id: string): boolean {
  return id.length > 0;
}
