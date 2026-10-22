// Shared shape validators for the generation API (SPEC-045 AC-6, AC-7,
// AC-9): every builder method and every JSON document-description field
// funnels through these before a single byte is written. Each throws
// `DocumentGenerationError` naming the offending path and the expected
// type/constraint — never a generic TypeError a caller has to decode.
import { DocumentGenerationError } from "./errors";

export function expectPlainObject(value: unknown, path: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new DocumentGenerationError(`${path}: expected an object`);
  }
  return value as Record<string, unknown>;
}

export function expectArray(value: unknown, path: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new DocumentGenerationError(`${path}: expected an array`);
  }
  return value;
}

export function expectString(value: unknown, path: string): string {
  if (typeof value !== "string") {
    throw new DocumentGenerationError(`${path}: expected a string`);
  }
  return value;
}

export function expectBoolean(value: unknown, path: string): boolean {
  if (typeof value !== "boolean") {
    throw new DocumentGenerationError(`${path}: expected a boolean`);
  }
  return value;
}

export function expectNumber(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new DocumentGenerationError(`${path}: expected a finite number`);
  }
  return value;
}

export function expectNonNegativeNumber(value: unknown, path: string): number {
  const n = expectNumber(value, path);
  if (n < 0) {
    throw new DocumentGenerationError(`${path}: expected a non-negative number`);
  }
  return n;
}

export function expectPositiveNumber(value: unknown, path: string): number {
  const n = expectNumber(value, path);
  if (n <= 0) {
    throw new DocumentGenerationError(`${path}: expected a positive number`);
  }
  return n;
}

/** An integer within `[min, max]` (inclusive) — used for heading/TOC levels. */
export function expectIntInRange(value: unknown, path: string, min: number, max: number): number {
  const n = expectNumber(value, path);
  if (!Number.isInteger(n) || n < min || n > max) {
    throw new DocumentGenerationError(`${path}: expected an integer between ${min} and ${max}`);
  }
  return n;
}

export function expectEnum<T extends string>(
  value: unknown,
  allowed: readonly T[],
  path: string,
): T {
  const str = expectString(value, path);
  if (!allowed.includes(str as T)) {
    throw new DocumentGenerationError(
      `${path}: expected one of ${allowed.map((v) => `"${v}"`).join(", ")}, got "${str}"`,
    );
  }
  return str as T;
}

/** A lowercase `#rrggbb` colour, the only colour shape the canonical model
 * stores (see documentModel.ts's `Hex`-style fields). */
export function expectHexColor(value: unknown, path: string): string {
  const str = expectString(value, path);
  if (!/^#[0-9a-f]{6}$/i.test(str)) {
    throw new DocumentGenerationError(`${path}: expected a "#rrggbb" colour, got "${str}"`);
  }
  return str.toLowerCase();
}

/** Rejects any key in `obj` not listed in `allowed` (AC-9): every object the
 * generation API accepts — a run, a block, a table spec, a whole
 * description — is validated closed, so a typo or a stray field is a
 * validation error rather than silently ignored data. */
export function rejectUnknownKeys(
  obj: Record<string, unknown>,
  allowed: readonly string[],
  path: string,
): void {
  for (const key of Object.keys(obj)) {
    if (!allowed.includes(key)) {
      throw new DocumentGenerationError(`${path}.${key}: unknown property`);
    }
  }
}
