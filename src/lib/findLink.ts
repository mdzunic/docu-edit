// The `find` deep link SPEC-114 adds to the editor routes: /doc/:id,
// /slides/:id and /media/:id accept `?find=<text>` (and `&n=<k>`, the
// 1-based match to select) and pre-fill their existing find surface with it —
// no find behaviour changes. A library search result opens through it
// (lib/searchOpen.ts builds the link); the routes read it here. Pure: no
// imports, so each editor reads its query without pulling in anything else.

/** A route's `find` and `n` query parameters, read safely: the text (""
 * when absent) and the 1-based match to select (1 when absent or bad). */
export function findParams(query: Record<string, unknown>): { find: string; n: number } {
  const find = typeof query.find === "string" ? query.find.trim() : "";
  const n = typeof query.n === "string" ? Math.trunc(Number(query.n)) : 1;
  return { find, n: Number.isFinite(n) && n >= 1 ? n : 1 };
}
