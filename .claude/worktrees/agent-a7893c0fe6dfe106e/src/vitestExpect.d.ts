// The type-check's memory budget (SPEC-132). `npm run typecheck` runs
// vue-tsc in Node's default heap (≈2.2 GB here), and the project's type-check
// had grown to within a few dozen MB of it: vitest's `expect<T>(actual: T):
// Assertion<T>` instantiates the whole Assertion interface (every matcher,
// Chai's included) once per distinct `actual` type, and the unit suite passes
// thousands of distinct object-literal and model types. This overload is
// declared after vitest's own, so it is tried first (TypeScript orders a
// merged interface's later call signatures ahead of earlier ones): every
// `expect(actual)` is an `Assertion<unknown>`, instantiated once.
//
// No test is weakened: the matchers vitest declares are generic in their OWN
// type parameters (`toEqual<E>(expected: E)`, `toMatchObject<E>(…)`, …) and
// check nothing against `actual`'s type; what a test asserts is checked at
// run time, exactly as before. `expectTypeOf` (type-level assertions) is a
// separate function and is untouched.
import "@vitest/expect";

declare module "@vitest/expect" {
  interface ExpectStatic {
    (actual: unknown, message?: string): Assertion<unknown>;
  }
}
