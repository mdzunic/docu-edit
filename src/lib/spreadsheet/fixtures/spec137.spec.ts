// SPEC-137: proves the workbook performance benchmark's fixture generator is
// deterministic (AC-5 — "same script, same bytes") and shaped the way "The
// benchmark fixtures" describes, without paying for a full parse of the
// 500,000-cell fixture in the unit suite. The LIBRARY fixture (100,000
// cells) is parsed for shape; the LARGE fixture (500,000 cells) is only
// hashed/sized — see the fixture builders' own header for why a cheap
// determinism check does not need a parse either.
//
// NOTE: the xlsx reader is mid-refactor on this branch (open.ts's phases
// dispatch to xlsx/phases.ts, which is a stub until the READ-XLSX stream
// lands) — readSpreadsheetBytes may throw for xlsx in the meantime. This
// spec parses with xlsx/reader.ts's parseXlsx directly, which does not go
// through the phase split, so it is unaffected either way.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { parseXlsx } from "../xlsx/reader";
import { buildSpec137LargeXlsxBytes, buildSpec137LibraryXlsxBytes, SPEC137_LIBRARY_XLSX } from "./build";

describe("SPEC-137 fixtures", () => {
  it(
    "spec137-library.xlsx is deterministic and shaped per 'The benchmark fixtures' (AC-4, AC-5)",
    async () => {
      const first = await buildSpec137LibraryXlsxBytes();
      const parsed = await parseXlsx(first);
      const sheet = parsed.sheets[0];
      expect(parsed.sheets).toHaveLength(1);
      expect(sheet?.name).toBe(SPEC137_LIBRARY_XLSX.sheetName);
      expect(sheet?.rows).toHaveLength(SPEC137_LIBRARY_XLSX.rows);

      // Every row holds all 20 columns (A..T, 0-based col 0..19) — dense, not sparse.
      const lastRow = sheet!.rows[sheet!.rows.length - 1]!;
      expect(lastRow.index).toBe(SPEC137_LIBRARY_XLSX.rows - 1);
      expect(lastRow.cells).toHaveLength(SPEC137_LIBRARY_XLSX.cols);
      expect(Math.max(...lastRow.cells.map((cell) => cell.col)) + 1).toBe(SPEC137_LIBRARY_XLSX.cols);

      // Column A: shared strings (every value a string, sstIndex set).
      const firstRow = sheet!.rows[0]!;
      const aCell = firstRow.cells.find((cell) => cell.col === 0);
      expect(typeof aCell?.value).toBe("string");
      expect(aCell?.sstIndex).toBeDefined();

      // Column I (0-based col 8): dates — a numeric value under the
      // numFmt-14 built-in style.
      const iCell = firstRow.cells.find((cell) => cell.col === 8);
      expect(typeof iCell?.value).toBe("number");
      const iStyle = iCell?.styleId !== undefined ? parsed.styles[iCell.styleId] : undefined;
      expect(iStyle?.builtinFormatId).toBe(14);

      // Column J (0-based col 9): booleans.
      const jCell = firstRow.cells.find((cell) => cell.col === 9);
      expect(typeof jCell?.value).toBe("boolean");

      // Same script, same bytes (AC-5).
      const second = await buildSpec137LibraryXlsxBytes();
      expect(equalBytes(first, second)).toBe(true);
    },
    20_000,
  );

  it(
    "spec137-large.xlsx is deterministic and reasonably sized (AC-3, AC-5) — not parsed here (too slow for the unit suite)",
    async () => {
      const first = await buildSpec137LargeXlsxBytes();
      // Comfortably under SPEC-009's 50 MB open-path file-size ceiling
      // (specs/137 "Memory and the cell ceiling" sizes this fixture at a
      // quarter of the cell ceiling, not the byte one), and large enough
      // that an accidentally-empty generator would be caught.
      expect(first.length).toBeGreaterThan(200_000);
      expect(first.length).toBeLessThan(50_000_000);

      const second = await buildSpec137LargeXlsxBytes();
      expect(equalBytes(first, second)).toBe(true);
    },
    60_000,
  );
});
