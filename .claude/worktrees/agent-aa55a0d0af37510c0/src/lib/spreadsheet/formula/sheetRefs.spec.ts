// SPEC-121 "Sheet references in formulas": remapSheetReferences per token
// kind — references renamed (re-quoted as SPEC-119 prints) or turned into a
// prefix-less #REF!, 3-D endpoints moved inward along orderBefore (start,
// end, both, interior, collapse to one sheet, nothing left), sheet-scoped
// names and Sheet2!#REF! deleted or renamed, prefixed functions renamed only,
// external references never touched — in one simultaneous, case-insensitive
// pass that keeps whitespace byte-for-byte and returns an untouched formula
// as the identical string; for a single rename it equals SPEC-119's
// renameSheetInFormula over thousands of generated formulas. remapOdsAddresses
// rewrites only the address-list entries the remap touches, deleted sheets
// in LibreOffice's `$#REF!.` form.
import { describe, expect, it } from "vitest";
import { CANONICAL_CORPUS } from "../fixtures/formulaCorpus";
import { formatSheetPrefix } from "./print";
import { renameSheetInFormula } from "./references";
import { isEmptyRemap, remapOdsAddress, remapOdsAddresses, remapSheetReferences, type SheetRemap } from "./sheetRefs";

const ORDER = ["Jan", "Feb", "Mar", "Apr"];

function remap(renamed: Record<string, string>, deleted: string[] = [], orderBefore: string[] = ORDER): SheetRemap {
  return { renamed, deleted, orderBefore };
}

const renaming = (from: string, to: string): SheetRemap => remap({ [from]: to });
const deleting = (...names: string[]): SheetRemap => remap({}, names);

describe("isEmptyRemap", () => {
  it("is true exactly when nothing is renamed or deleted — an order alone is empty, an identity rename is not", () => {
    expect(isEmptyRemap(remap({}))).toBe(true);
    expect(isEmptyRemap(renaming("Data", "Facts"))).toBe(false);
    expect(isEmptyRemap(deleting("Data"))).toBe(false);
    // It still re-spells `data!A1` as `Data!A1`.
    expect(isEmptyRemap(renaming("Data", "Data"))).toBe(false);
  });

  it("an empty remap returns the text at once, whatever it is", () => {
    for (const text of ["Data!A1", "SUM(Jan:Mar!B2)", "'unterminated!A1", ""]) {
      expect(remapSheetReferences(text, remap({}))).toBe(text);
    }
  });
});

describe("remapSheetReferences — references", () => {
  it("renames every reference to the sheet: cells, ranges, whole columns and rows", () => {
    expect(remapSheetReferences("SUM(Data!A1:B2)+Data!$C$3*Data!A:A-Data!$1:$1", renaming("Data", "Facts"))).toBe(
      "SUM(Facts!A1:B2)+Facts!$C$3*Facts!A:A-Facts!$1:$1",
    );
  });

  it("turns a reference to a deleted sheet into #REF! as a whole — no prefix", () => {
    expect(remapSheetReferences("SUM(Lookup!A1:B5)+Lookup!C1*2+A1", deleting("Lookup"))).toBe("SUM(#REF!)+#REF!*2+A1");
    expect(remapSheetReferences("VLOOKUP(A1,Lookup!$A$1:$B$5,2,FALSE)", deleting("Lookup"))).toBe("VLOOKUP(A1,#REF!,2,FALSE)");
    expect(remapSheetReferences("'My Lookup'!A:A+Data!B1", deleting("My Lookup"))).toBe("#REF!+Data!B1");
  });

  it("re-quotes per SPEC-119's rule: to a name needing quotes and back", () => {
    const text = "Data!A1*2";
    expect(remapSheetReferences(text, renaming("Data", "Q1 Data"))).toBe("'Q1 Data'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "O'Brien"))).toBe("'O''Brien'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "2026"))).toBe("'2026'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "A1"))).toBe("'A1'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "TRUE"))).toBe("'TRUE'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "R1C1"))).toBe("'R1C1'!A1*2");
    expect(remapSheetReferences(text, renaming("Data", "Übersicht"))).toBe("Übersicht!A1*2");
    expect(remapSheetReferences("'Q1 Data'!A1+'O''Brien'!B2", remap({ "Q1 Data": "Summary", "O'Brien": "Smith" }))).toBe("Summary!A1+Smith!B2");
  });

  it("leaves references to other sheets, unqualified references and non-reference tokens verbatim", () => {
    const text = `SUM(A1,Other!b2)&"Data!A1"&Table1[Data]&#REF!A1&Rate&Data`;
    expect(remapSheetReferences(text, remap({ Data: "Facts" }, ["Lookup"]))).toBe(text);
  });
});

describe("remapSheetReferences — sheet-prefixed names, errors and functions", () => {
  it("a sheet-scoped name, and Sheet2!#REF!, become #REF! when their sheet is deleted", () => {
    expect(remapSheetReferences("Data!Rate*2+Data!#REF!+'Q1 Data'!Rate", deleting("Data", "Q1 Data"))).toBe("#REF!*2+#REF!+#REF!");
  });

  it("…and are renamed (prefix re-quoted, the rest as written) when it is renamed", () => {
    expect(remapSheetReferences("Data!Rate*2+Data!#REF!+Data!#ref!", renaming("Data", "Q1 Data"))).toBe("'Q1 Data'!Rate*2+'Q1 Data'!#REF!+'Q1 Data'!#ref!");
  });

  it("a 3-D #REF! or name prefix is #REF! when either endpoint is deleted, renamed per endpoint otherwise", () => {
    expect(remapSheetReferences("Jan:Mar!#REF!+1", deleting("Mar"))).toBe("#REF!+1");
    expect(remapSheetReferences("Jan:Mar!#REF!+1", deleting("Feb"))).toBe("Jan:Mar!#REF!+1");
    expect(remapSheetReferences("Jan:Mar!#REF!+1", renaming("Mar", "March 2026"))).toBe("'Jan:March 2026'!#REF!+1");
  });

  it("a prefixed function is renamed, never turned into #REF!", () => {
    expect(remapSheetReferences("Data!MYFN(1)+2", renaming("Data", "My Data"))).toBe("'My Data'!MYFN(1)+2");
    const text = "Data!MYFN(1)+2";
    expect(remapSheetReferences(text, deleting("Data"))).toBe(text);
  });

  it("never touches external references — they name another workbook's sheets", () => {
    const text = "[1]Data!A1+'[2]Data'!B2+[1]!Data+[1]Data:Lookup!C3";
    expect(remapSheetReferences(text, remap({ Data: "Facts" }, ["Lookup"]))).toBe(text);
  });
});

describe("remapSheetReferences — 3-D references", () => {
  it.each([
    [["Jan"], "SUM(Feb:Mar!B2)"],
    [["Mar"], "SUM(Jan:Feb!B2)"],
    [["Jan", "Mar"], "SUM(Feb!B2)"],
    [["Jan", "Feb"], "SUM(Mar!B2)"],
    [["Feb", "Mar"], "SUM(Jan!B2)"],
    [["Feb"], "SUM(Jan:Mar!B2)"],
    [["Apr"], "SUM(Jan:Mar!B2)"],
    [["Jan", "Feb", "Mar"], "SUM(#REF!)"],
  ])("SUM(Jan:Mar!B2) with %j deleted → %s — endpoints move inward, a one-sheet span is a plain reference, an empty one #REF!", (deleted, expected) => {
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", deleting(...deleted))).toBe(expected);
  });

  it("returns the identical string when only interior sheets are deleted", () => {
    const text = " SUM( Jan:Apr!B2 ) ";
    expect(remapSheetReferences(text, deleting("Feb", "Mar"))).toBe(text);
  });

  it("applies renames to the surviving endpoints — the one moved to included — and re-quotes the prefix whole", () => {
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", remap({ Feb: "Q1 Data" }, ["Jan"]))).toBe("SUM('Q1 Data:Mar'!B2)");
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", remap({ Jan: "January" }, ["Mar"]))).toBe("SUM(January:Feb!B2)");
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", remap({ Feb: "Only" }, ["Jan", "Mar"]))).toBe("SUM(Only!B2)");
    expect(remapSheetReferences("SUM('Q1 2026:Q4 2026'!C3)", remap({}, ["Q4 2026"], ["Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026"]))).toBe(
      "SUM('Q1 2026:Q3 2026'!C3)",
    );
  });

  it("renames either endpoint without deletions, and never collapses a span on a rename alone", () => {
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", renaming("Mar", "March 2026"))).toBe("SUM('Jan:March 2026'!B2)");
    expect(remapSheetReferences("SUM(Data:data!B2)", renaming("DATA", "Facts"))).toBe("SUM(Facts:Facts!B2)");
  });

  it("walks a reversed span the other way", () => {
    expect(remapSheetReferences("SUM(Mar:Jan!B2)", deleting("Mar"))).toBe("SUM(Feb:Jan!B2)");
    expect(remapSheetReferences("SUM(Mar:Jan!B2)", deleting("Jan"))).toBe("SUM(Mar:Feb!B2)");
    expect(remapSheetReferences("SUM(Mar:Jan!B2)", deleting("Mar", "Feb", "Jan"))).toBe("SUM(#REF!)");
  });

  it("is #REF! when a deleted endpoint's span cannot be walked: an endpoint orderBefore does not hold", () => {
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", remap({}, ["Jan"], []))).toBe("SUM(#REF!)");
    expect(remapSheetReferences("SUM(Jan:Mar!B2)", remap({}, ["Jan"], ["Jan", "Feb"]))).toBe("SUM(#REF!)");
    expect(remapSheetReferences("SUM(Jan:Jan!B2)", deleting("Jan"))).toBe("SUM(#REF!)");
  });

  it("matches orderBefore case-insensitively; an endpoint keeps its spelling, a moved one takes orderBefore's", () => {
    expect(remapSheetReferences("SUM(jan:MAR!b2)", remap({}, ["JAN"], ["Jan", "feb", "Mar"]))).toBe("SUM(feb:MAR!B2)");
  });

  it("moves whole-column and whole-row 3-D references alike", () => {
    expect(remapSheetReferences("SUM(Jan:Mar!A:A)+SUM(Jan:Mar!$2:$3)", deleting("Jan", "Mar"))).toBe("SUM(Feb!A:A)+SUM(Feb!$2:$3)");
  });
});

describe("remapSheetReferences — one simultaneous, case-insensitive pass", () => {
  it("swaps two names in one pass, and the swap round-trips", () => {
    const swap = remap({ Alpha: "Beta", Beta: "Alpha" }, [], ["Alpha", "Beta"]);
    const text = "Alpha!A1+Beta!B2+SUM(Alpha:Beta!C3)+Alpha!Rate-Beta!#REF!";
    const swapped = remapSheetReferences(text, swap);
    expect(swapped).toBe("Beta!A1+Alpha!B2+SUM(Beta:Alpha!C3)+Beta!Rate-Alpha!#REF!");
    expect(remapSheetReferences(swapped, swap)).toBe(text);
  });

  it("looks every reference up once: a chain of renames never hops twice", () => {
    expect(remapSheetReferences("First!A1+Second!A1+Third!A1", remap({ First: "Second", Second: "Third" }))).toBe("Second!A1+Third!A1+Third!A1");
  });

  it("renames and deletes together — a sheet renamed to a deleted one's old name included", () => {
    expect(remapSheetReferences("Data!A1+Lookup!A1", remap({ Data: "Facts" }, ["Lookup"]))).toBe("Facts!A1+#REF!");
    expect(remapSheetReferences("Lookup!A1+Archive!A1", remap({ Archive: "Lookup" }, ["Lookup"]))).toBe("#REF!+Lookup!A1");
  });

  it("treats a name both renamed and deleted as deleted", () => {
    expect(remapSheetReferences("Data!A1+Data!MYFN(1)", remap({ Data: "Facts" }, ["data"]))).toBe("#REF!+Data!MYFN(1)");
  });

  it("matches renamed and deleted names in any case, both sides upper-cased", () => {
    expect(remapSheetReferences("DATA!A1+data!B2+Data!C3+dAtA!Rate", renaming("dAtA", "Facts"))).toBe("Facts!A1+Facts!B2+Facts!C3+Facts!Rate");
    expect(remapSheetReferences("DATA!A1+data!B2+Data!#REF!", deleting("DaTa"))).toBe("#REF!+#REF!+#REF!");
    expect(remapSheetReferences("Übersicht!A1", deleting("ÜBERSICHT"))).toBe("#REF!");
  });

  it("re-spells a case-only rename, and keeps verbatim a reference already spelled as the new name", () => {
    expect(remapSheetReferences("data!a1+Data!a1", renaming("data", "Data"))).toBe("Data!A1+Data!a1");
  });
});

describe("remapSheetReferences — trivia and identity", () => {
  it("keeps every byte of whitespace and of every token it does not rewrite", () => {
    const text = "  SUM( Data!A1 ,\r\n\tLookup!B2 )  &  Other!c3 + Data!Rate\n";
    expect(remapSheetReferences(text, remap({ Data: "Q1 Data" }, ["Lookup"]))).toBe("  SUM( 'Q1 Data'!A1 ,\r\n\t#REF! )  &  Other!c3 + 'Q1 Data'!Rate\n");
  });

  it("keeps the intersection space when a reference in an intersection is rewritten", () => {
    expect(remapSheetReferences("SUM(Data!A1:C3 Data!B2:D4)", renaming("Data", "Facts"))).toBe("SUM(Facts!A1:C3 Facts!B2:D4)");
  });

  it("returns the identical string when nothing is touched", () => {
    for (const text of [" Feb!A1 +  'Other Sheet'!B2 ", "A1+B2*3", "Data2!A1+MyData!B2", "[1]Data!A1", "Data!", "'Data'", "\"Data!A1\""]) {
      expect(remapSheetReferences(text, remap({ Data: "Facts" }, ["Lookup"]))).toBe(text);
    }
  });
});

// --- the property: one rename is renameSheetInFormula ---------------------------------

/** mulberry32 — a small seeded PRNG, so the property runs the same every time. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Sheet names as formulas spell them — bare, needing quotes, case variants. */
const NAMES = ["Data", "data", "DATA", "Sheet1", "Q1 Data", "O'Brien", "2026", "A1", "Übersicht", "Jan", "Mar", "TRUE", "R1C1", "My.Sheet", "Facts"];
/** Names a rename gives — some needing quotes, some an existing name in another case. */
const TARGETS = ["Facts", "Q1 Data", "O'Brien", "2026", "A1", "TRUE", "Data", "data", "R2", "x.y", "Über Sicht", "Sheet1", "It's"];
const AREAS = ["A1", "$B$2", "a1:b2", "A:C", "$1:$3", "XFD1048576", "$A1:B$9"];
const LEADS = ["", "", "", " ", "  ", "\t", "\r\n"];
const OPERATORS = ["+", "-", "*", "&", ",", " ", ":", "="];

describe("property: a single rename is renameSheetInFormula", () => {
  it("equals renameSheetInFormula(t, from, to) for 4,000 generated formulas — every token kind, every spelling of a prefix", () => {
    const random = seeded(121);
    const pick = <T,>(items: readonly T[]): T => items[Math.floor(random() * items.length)]!;
    /** The sheet names the formula being generated spells. */
    let used: string[] = [];
    /** A prefix as written: as SPEC-119 prints it, or quoted where it need not be. */
    const prefix = (sheet: string, sheetEnd?: string): string => {
      used.push(sheet, ...(sheetEnd === undefined ? [] : [sheetEnd]));
      if (random() < 0.3) {
        const body = sheetEnd === undefined ? sheet : `${sheet}:${sheetEnd}`;
        return `'${body.replaceAll("'", "''")}'!`;
      }
      return formatSheetPrefix(sheet, sheetEnd);
    };
    const operand = (): string => {
      switch (Math.floor(random() * 11)) {
        case 0:
        case 1:
        case 2:
          return prefix(pick(NAMES)) + pick(AREAS);
        case 3:
          return prefix(pick(NAMES), pick(NAMES)) + pick(AREAS);
        case 4:
          return `${prefix(pick(NAMES))}Rate`;
        case 5:
          return random() < 0.5 ? `${prefix(pick(NAMES))}#REF!` : "#REF!A1";
        case 6:
          return `${prefix(pick(NAMES))}MYFN(${prefix(pick(NAMES))}A1)`;
        case 7:
          return random() < 0.5 ? `[1]${pick(["Data", "Jan", "Sheet1"])}!A1` : `'[2]${pick(NAMES).replaceAll("'", "''")}'!B2`;
        case 8:
          return pick(["Table1[Data]", `"Data!A1"`, "Rate", "B2", "1.5", "TRUE"]);
        case 9:
          return `SUM(${prefix(pick(NAMES))}${pick(AREAS)},${pick(LEADS)}${prefix(pick(NAMES))}${pick(AREAS)})`;
        default:
          return `(${prefix(pick(NAMES), pick(NAMES))}${pick(AREAS)})`;
      }
    };
    let changed = 0;
    for (let n = 0; n < 4_000; n += 1) {
      used = [];
      let text = pick(LEADS) + operand();
      const count = Math.floor(random() * 4);
      for (let i = 0; i < count; i += 1) {
        text += pick(LEADS) + pick(OPERATORS) + pick(LEADS) + operand();
      }
      text += pick(LEADS);
      // Mostly a sheet the formula names (in any case), sometimes any name.
      const named = used.length > 0 && random() < 0.8 ? pick(used) : pick(NAMES);
      const from = random() < 0.25 ? named.toUpperCase() : random() < 0.1 ? named.toLowerCase() : named;
      const to = pick(TARGETS);
      const expected = renameSheetInFormula(text, from, to);
      expect(remapSheetReferences(text, remap({ [from]: to }, [], NAMES)), JSON.stringify({ text, from, to })).toBe(expected);
      if (expected !== text) {
        changed += 1;
      }
    }
    // The property is not vacuous: most generated formulas name the sheet.
    expect(changed).toBeGreaterThan(1_000);
  });

  it("…and for every corpus formula, renaming each sheet it names", () => {
    const sheets = ["Sheet1", "sheet1", "Data", "Price List", "Jan", "Dec", "My Sheet", "O'Brien", "Übersicht", "Q1 2026", "Q4 2026", "Sheet3"];
    for (const text of CANONICAL_CORPUS) {
      for (const from of sheets) {
        for (const to of ["Facts", "Q1 Data", "2026", "O'Brien"]) {
          expect(remapSheetReferences(text, renaming(from, to)), JSON.stringify({ text, from, to })).toBe(renameSheetInFormula(text, from, to));
        }
      }
    }
  });

  it("…and for 5,000 seeded random strings over a reference-heavy alphabet, malformed text included", () => {
    const random = seeded(1210);
    const chunks = [
      "Ab!", "ab!", "'A b'!", "'a''b'!", "x:Ab!", "Ab!A1", "'A b'!b2:C3", "x:AB!$C$3", "ab!Rate", "X!#REF!",
      "AB", "A1", "b2:C3", "Rate", "#REF!", "!", "'", "''", ":", "$", " ", "\t", "[1]", "(", ")", ",", "+", "\"", ".", "1", "Ü",
    ];
    let changed = 0;
    for (let n = 0; n < 5_000; n += 1) {
      let text = "";
      const length = Math.floor(random() * 12);
      for (let i = 0; i < length; i += 1) {
        text += chunks[Math.floor(random() * chunks.length)];
      }
      const from = ["Ab", "A b", "x", "a'b"][Math.floor(random() * 4)]!;
      const to = ["Cd", "C d", "1", "x"][Math.floor(random() * 4)]!;
      const expected = renameSheetInFormula(text, from, to);
      expect(remapSheetReferences(text, renaming(from, to)), JSON.stringify({ text, from, to })).toBe(expected);
      if (expected !== text) {
        changed += 1;
      }
    }
    expect(changed).toBeGreaterThan(500);
  });
});

// --- ODS address lists -------------------------------------------------------------------

describe("remapOdsAddresses", () => {
  it("returns the list at once for an empty remap, and the identical string when no entry is touched", () => {
    const list = "$Data.A1 junk $#REF!.B2";
    expect(remapOdsAddresses(list, remap({}))).toBe(list);
    const untouched = "Other.a1:Other.b2   .C3 $'My Sheet'.D4 $Jan.A1:$Mar.B2";
    expect(remapOdsAddresses(untouched, remap({ Data: "Facts" }, ["Feb"]))).toBe(untouched);
  });

  it("rewrites only the renamed entries, every other entry and every separator byte-for-byte", () => {
    expect(remapOdsAddresses("$Data.$A$1:.$B$5 Other.c3   Data.A1:Data.B3  .D4", renaming("Data", "Facts"))).toBe(
      "$Facts.$A$1:.$B$5 Other.c3   $Facts.A1:.B3  .D4",
    );
  });

  it("quotes a new name as an address needs it, ' doubled, and drops quotes no longer needed", () => {
    expect(remapOdsAddresses("$Data.A1", renaming("Data", "Q1 Data"))).toBe("$'Q1 Data'.A1");
    expect(remapOdsAddresses("$Data.A1", renaming("Data", "O'Brien"))).toBe("$'O''Brien'.A1");
    expect(remapOdsAddresses("$Data.A1", renaming("Data", "x.y"))).toBe("$'x.y'.A1");
    expect(remapOdsAddresses("$'Q1 Data'.A1:.B2", renaming("Q1 Data", "Summary"))).toBe("$Summary.A1:.B2");
  });

  it("writes an entry on a deleted sheet in LibreOffice's $#REF!. form, keeping its cell parts", () => {
    expect(remapOdsAddresses("$Lookup.$A$1:.$B$5", deleting("Lookup"))).toBe("$#REF!.$A$1:.$B$5");
    expect(remapOdsAddresses("Lookup.A1", deleting("Lookup"))).toBe("$#REF!.A1");
    expect(remapOdsAddresses("Lookup.A1:Lookup.B2", deleting("Lookup"))).toBe("$#REF!.A1:.B2");
    expect(remapOdsAddresses("$Lookup.$A:.$C", deleting("Lookup"))).toBe("$#REF!.$A:.$C");
  });

  it("moves a 3-D entry's deleted endpoints inward first; only an entry left with no sheet becomes $#REF!.", () => {
    expect(remapOdsAddresses("$Jan.A1:$Mar.B2", deleting("Jan"))).toBe("$Feb.A1:$Mar.B2");
    expect(remapOdsAddresses("$Jan.A1:$Mar.B2", deleting("Jan", "Mar"))).toBe("$Feb.A1:.B2");
    expect(remapOdsAddresses("$Jan.A1:$Mar.B2", deleting("Feb"))).toBe("$Jan.A1:$Mar.B2");
    expect(remapOdsAddresses("$Jan.A1:$Mar.B2", deleting("Jan", "Feb", "Mar"))).toBe("$#REF!.A1:.B2");
    expect(remapOdsAddresses("$Jan.B5:$Mar.B5", deleting("Jan", "Mar"))).toBe("$Feb.B5");
    expect(remapOdsAddresses("$Jan.B5:$Mar.B5", deleting("Jan", "Feb", "Mar"))).toBe("$#REF!.B5");
  });

  it("handles a mixed list — renamed, deleted, other-sheet, sheet-less and unreadable entries — in one pass", () => {
    const list = "$Data.A1 $Lookup.B2:.C3 .D4  $Other.E5 $#REF!.F6 junk 'data'.G7";
    expect(remapOdsAddresses(list, remap({ Data: "Facts" }, ["Lookup"]))).toBe("$Facts.A1 $#REF!.B2:.C3 .D4  $Other.E5 $#REF!.F6 junk $Facts.G7");
  });

  it("is simultaneous and case-insensitive: a swap round-trips", () => {
    const swap = remap({ alpha: "Beta", BETA: "Alpha" });
    const swapped = remapOdsAddresses("$Alpha.A1 $Beta.B2:.C3", swap);
    expect(swapped).toBe("$Beta.A1 $Alpha.B2:.C3");
    expect(remapOdsAddresses(swapped, swap)).toBe("$Alpha.A1 $Beta.B2:.C3");
  });

  it("keeps an entry already spelled as the rename gives it", () => {
    expect(remapOdsAddresses("Data.a1", renaming("Data", "Data"))).toBe("Data.a1");
    expect(remapOdsAddresses("data.a1", renaming("Data", "Data"))).toBe("$Data.A1");
  });

  it("remapOdsAddress handles one entry the same way", () => {
    expect(remapOdsAddress("$Data.$A$1", renaming("Data", "Q1 Data"))).toBe("$'Q1 Data'.$A$1");
    expect(remapOdsAddress("$Lookup.$A$1", deleting("Lookup"))).toBe("$#REF!.$A$1");
    expect(remapOdsAddress("not an address", deleting("Lookup"))).toBe("not an address");
  });
});
