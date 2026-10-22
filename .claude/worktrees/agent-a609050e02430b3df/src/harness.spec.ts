import { describe, it, expect } from "vitest";

const sum = (a: number, b: number): number => a + b;

// This suite exists so `npm test` has something to run and exits 0 from day
// one — the factory's regression gate refuses to start work on a repo whose
// test command fails. It asserts nothing about the product on purpose, so no
// feature spec can ever be blocked by it.
//
// Real unit tests live beside the code they cover (src/**/*.spec.ts) and are
// added by the spec that introduces the behaviour.
describe("test harness", () => {
  it("runs vitest and resolves TypeScript", () => {
    expect(sum(1, 1)).toBe(2);
  });
});
