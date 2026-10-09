// The SPEC-119 formula corpus the formula specs share ("Tests"): canonical
// formulas (the tokenizer's round trip, the printer's oracle, the xlsx
// translator's canonical side), Excel-shaped xlsx wire formulas and
// LibreOffice-shaped OpenFormula ones — each list including every formula of
// the two SPEC-119 fixtures.
import { SPEC119_FORMULAS_ODS, SPEC119_FORMULAS_XLSX } from "./build";

/** Canonical formulas — every one parses. */
export const CANONICAL_CORPUS: readonly string[] = [
  `IF(A1="","",VLOOKUP(A1,'Price List'!$A$2:$D$500,4,FALSE))`,
  `SUMPRODUCT((Data!$B$2:$B$1000="West")*Data!$C$2:$C$1000)`,
  `"He said ""hi"""&A1`,
  `INDEX(A:A,MATCH(MAX(B:B),B:B,0))`,
  `SUM(Jan:Dec!B5)`,
  `SUM((A1:A3,C1:C3))`,
  `SUM(A1:C3 B2:D4)`,
  `  SUM( A1 , B1 ) `,
  `-2^2`,
  `1+-+-1`,
  `50%%`,
  `.5+1E3`,
  `LOG10(100)+ATAN2(1,1)`,
  `XFD1048576+XFE1`,
  `Sales[@Qty]*Sales[@Price]`,
  `Table1[[#Headers],[Col'#1]]`,
  `[1]!GlobalName`,
  `#REF!A1+Sheet1!#REF!`,
  `A1:INDEX(B:B,5)`,
  `'Übersicht'!A1`,
  `true+TRUE()`,
  "IF(A1>0,\r\n    \"positive\",\r\n    \"other\")",
  `'Q1 2026:Q4 2026'!C3+Sheet1:Sheet3!A1`,
  `SUM($A:$C)+SUM(1:1)+SUM($2:$4)`,
  `'My Sheet'!A1:B2*'O''Brien'!A1`,
  `[1]Sheet1!$A$1+'[2]Rates 2026'!B2`,
  `Table1[[#This Row],[Amount]]+[@Qty]`,
  `DAYS360(A1,B1)+R1C1+A0`,
  `SUM({1,2;3,4})+SUM({"a",TRUE;#N/A,-1.5})`,
  `IF(A1,,1)+F()+IFERROR(1/0,)`,
  `Sheet1!Rate*2+'My Sheet'!Rate`,
  `A1=B1=C1`,
  `2^3^2+2^3%+2^-2`,
  `-A1:A3`,
  `1-2-3`,
  `A1 &" "&  B1`,
  `CONCAT(A1,", ",B1)&TEXTJOIN(",",TRUE,A1:A3)`,
  `LET(rate,0.1,n,12,PMT(rate/n,n,1000))`,
  `#DIV/0!+#N/A+#NAME?+#NULL!+#NUM!+#VALUE!+#GETTING_DATA+#SPILL!+#CALC!+#FIELD!+#BLOCKED!+#CONNECT!+#BUSY!+#UNKNOWN!+#REF!`,
  `(A1:B2 B1:C3) (B2:D4)`,
  `$a$1+b2`,
  ...Object.values(SPEC119_FORMULAS_XLSX.formulas),
  ...Object.values(SPEC119_FORMULAS_ODS.formulas),
];

/** Excel-shaped `<f>` texts — toXlsxFormula(fromXlsxFormula(w)) === w. */
export const EXCEL_CORPUS: readonly string[] = [
  ...Object.values(SPEC119_FORMULAS_XLSX.wire),
  `_xlfn.CONCAT(A1,", ",B1)`,
  `_xlfn.TEXTJOIN(",",TRUE,A1:A3)`,
  `_xlfn.SWITCH(A1,1,"one",2,"two","many")`,
  `_xlfn.MINIFS(A1:A9,B1:B9,">0")+_xlfn.MAXIFS(A1:A9,B1:B9,">0")`,
  `_xlfn.XOR(A1,B1)+_xlfn.IFNA(C1,0)+_xlfn.DAYS(D1,E1)`,
  `_xlfn._xlws.FILTER(A1:A9,B1:B9>0)`,
  `_xlfn.STDEV.S(A1:A9)+_xlfn.CEILING.MATH(B1)+ISO.CEILING(C1)`,
  `_xlfn.LET(_xlpm.rate,0.1,_xlpm.n,12,PMT(_xlpm.rate/_xlpm.n,_xlpm.n,1000))`,
  `_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y))`,
  `_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)`,
  `Rate*_xlfn.LET(_xlpm.rate,2,_xlpm.rate)`,
  `VLOOKUP(A1,Prices!$A:$B,2,FALSE)`,
  `_xlfn.SINGLE(A1:A9)+_xlfn.ANCHORARRAY(B1)`,
];

/** LibreOffice-shaped `table:formula` values that survive the round trip
 * exactly — toOdsFormula(fromOdsFormula(w)) === w. */
export const LIBREOFFICE_CORPUS: readonly string[] = [
  ...Object.entries(SPEC119_FORMULAS_ODS.wire)
    .filter(([a1, wire]) => wire.startsWith("of:") && a1 !== "B6")
    .map(([, wire]) => wire),
  `of:=SUM(([.A1:.A3]~[.C1:.C3]))`,
  `of:=IF([.A1]="";"";VLOOKUP([.A1];[$'Price List'.$A$2:.$D$500];4;0))`,
  `of:=SUM([$Jan.B5:$Dec.B5])`,
  `of:=COM.MICROSOFT.STDEV.S([.A1:.A9])+LEGACY.CHIDIST([.A1];2)+FDIST([.A1];2;3;1)`,
  `of:=SUM([.A:.A])+SUM([.$1:.$1])`,
  `of:=[.A1] + [.B1] * 2`,
  `of:=-[.A1]^2&"x"`,
  `of:=['file:///home/user/rates.ods'#$Sheet1.A1]*2`,
  `of:=ORG.OPENOFFICE.ERRORTYPE([.A1])`,
  `of:=SUMPRODUCT(([$Data.$B$2:.$B$1000]="West")*[$Data.$C$2:.$C$1000])`,
  `of:=IFERROR(1/0;)`,
  `of:=[$'O''Brien'.A1]*2`,
];

/** The forms the ODS round trip changes on purpose (SPEC-119 "Known
 * non-identities"): wire → what toOdsFormula(fromOdsFormula(wire)) gives. */
export const ODS_NON_IDENTITIES: ReadonlyArray<readonly [string, string]> = [
  ["of:=[Sheet2.A1]", "of:=[$Sheet2.A1]"],
  ["of:=[$S.A1:$S.B2]", "of:=[$S.A1:.B2]"],
  [SPEC119_FORMULAS_ODS.wire.B6, "of:=SUM(([.B2:.B3]~[.B4]))"],
  ["of:=[.#REF!]+1", "of:=#REF!+1"],
  ["of:=SUM([$Sheet2.#REF!])", "of:=SUM(#REF!)"],
];
