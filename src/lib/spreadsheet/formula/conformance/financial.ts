// Financial conformance rows (SPEC-127 "Conformance tables"). Functions, in
// order: PMT, IPMT, PPMT, FV, PV, NPER, RATE, NPV, XNPV, IRR, XIRR, MIRR,
// SLN, SYD, DB, DDB, CUMIPMT, CUMPRINC, EFFECT, NOMINAL.
//
// Sources: each function's Microsoft reference page
// (support.microsoft.com/en-us/excel/functions/<name>-function) — its
// example table and Remarks bullets — for "Excel" rows; SPEC-127's Function
// pins and Numerical methods (RATE/IRR/XIRR convergence and tolerance, DB's
// three-decimal rate rounding, XNPV/XIRR date truncation) and hand
// arithmetic from the TVM identity pv*(1+r)^n + pmt*((1+r)^n-1)/r + fv = 0
// (fv=0,type=0 unless stated) for "pinned" rows. Dates are Excel 1900-system
// serial numbers (days since 1899-12-30, per SPEC-127's brief: 2008-10-30 =
// 39751); other dates below are computed from that anchor by day-of-year
// difference within the same or an adjacent (leap-aware) calendar year.
import type { ConformanceRow } from "./row";

export const FINANCIAL_ROWS: ConformanceRow[] = [
  // --- PMT: rate, nper, pv, [fv], [type] -------------------------------
  {
    formula: "PMT(A2/12,A3,A4)",
    cells: { A2: 0.08, A3: 10, A4: 10000 },
    expected: -1037.03,
    digits: 6,
    source: "Excel",
    note: "PMT function, example 1, row 1: 8% annual rate, 10 monthly payments, $10,000 loan; Microsoft prints ($1,037.03).",
  },
  {
    formula: "PMT(A2/12,A3,A4,,1)",
    cells: { A2: 0.08, A3: 10, A4: 10000 },
    expected: -1030.16,
    digits: 6,
    source: "Excel",
    note: "PMT function, example 1, row 2: same terms, type=1 (due at the start of the period, empty fv slot); Microsoft prints ($1,030.16).",
  },
  {
    formula: "PMT(A9/12,A10*12,0,A11)",
    cells: { A9: 0.06, A10: 18, A11: 50000 },
    expected: -129.08,
    digits: 5,
    source: "Excel",
    note: "PMT function, example 2: monthly deposit to reach $50,000 in 18 years at 6% annual, pv=0; Microsoft prints ($129.08).",
  },
  {
    formula: "PMT(1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (1, PMT needs rate, nper, pv); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- IPMT: rate, per, nper, pv, [fv], [type] --------------------------
  {
    formula: "IPMT(A2/12,A3,A4*12,A5)",
    cells: { A2: 0.1, A3: 1, A4: 3, A5: 8000 },
    expected: -66.67,
    digits: 4,
    source: "Excel",
    note: "IPMT function, example, row 1: interest due in month 1 of a 3-year loan at 10% annual on $8,000; Microsoft prints ($66.67).",
  },
  {
    formula: "IPMT(A2,3,A4,A5)",
    cells: { A2: 0.1, A4: 3, A5: 8000 },
    expected: -292.45,
    digits: 5,
    source: "Excel",
    note: "IPMT function, example, row 2: interest due in the last year with annual payments; Microsoft prints ($292.45).",
  },
  {
    formula: "IPMT(0.01,1,24,5000)",
    expected: -50,
    digits: 15,
    source: "pinned",
    note: "pinned: at per=1 (type=0) no principal has yet been repaid, so IPMT(rate,1,nper,pv) = -pv*rate = -5000*0.01 = -50 exactly.",
  },
  {
    formula: "IPMT(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (2, IPMT needs rate, per, nper, pv); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- PPMT: rate, per, nper, pv, [fv], [type] ---------------------------
  {
    formula: "PPMT(A2/12,1,A3*12,A4)",
    cells: { A2: 0.1, A3: 2, A4: 2000 },
    expected: -75.62,
    digits: 4,
    source: "Excel",
    note: "PPMT function, example 1: principal payment for month 1 of a 2-year loan at 10% annual on $2,000; Microsoft prints ($75.62).",
  },
  {
    formula: "PPMT(A8,A9,10,A10)",
    cells: { A8: 0.08, A9: 10, A10: 200000 },
    expected: -27598.05,
    digits: 7,
    source: "Excel",
    note: "PPMT function, example 2: principal payment for year 10 of a 10-year loan at 8% annual on $200,000; Microsoft prints ($27,598.05).",
  },
  {
    formula: "PPMT(0.1,1,1,1000)",
    expected: -1000,
    digits: 15,
    source: "pinned",
    note: "pinned: with nper=1 the single payment retires the entire principal (PMT=-pv*(1+rate), IPMT=-pv*rate, PPMT=PMT-IPMT=-pv), so PPMT(rate,1,1,1000) = -1000 exactly.",
  },
  {
    formula: "PPMT(1,2,3)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (3, PPMT needs rate, per, nper, pv); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- FV: rate, nper, pmt, [pv], [type] ----------------------------------
  {
    formula: "FV(A2/12,A3,A4,A5,A6)",
    cells: { A2: 0.06, A3: 10, A4: -200, A5: -500, A6: 1 },
    expected: 2581.4,
    digits: 6,
    source: "Excel",
    note: "FV function, example 1: future value of terms in A2:A5, payments due at period start; Microsoft prints $2,581.40.",
  },
  {
    formula: "FV(A2/12,A3,A4)",
    cells: { A2: 0.12, A3: 12, A4: -1000 },
    expected: 12682.5,
    digits: 7,
    source: "Excel",
    note: "FV function, example 2: future value of terms in A2:A4; Microsoft prints $12,682.50.",
  },
  {
    formula: "FV(A2/12,A3,A4,,A5)",
    cells: { A2: 0.11, A3: 35, A4: -2000, A5: 1 },
    expected: 82846.25,
    digits: 7,
    source: "Excel",
    note: "FV function, example 3: future value with pv omitted (empty slot), type=A5=1; Microsoft prints $82,846.25.",
  },
  {
    formula: "FV(A2/12,A3,A4,A5,A6)",
    cells: { A2: 0.06, A3: 12, A4: -100, A5: -1000, A6: 1 },
    expected: 2301.4,
    digits: 6,
    source: "Excel",
    note: "FV function, example 4: future value of terms in A2:A5; Microsoft prints $2,301.40.",
  },
  {
    formula: "FV(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (2, FV needs rate, nper, pmt); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- PV: rate, nper, pmt, [fv], [type] -----------------------------------
  {
    formula: "PV(A3/12,12*A4,A2,,0)",
    cells: { A2: 500, A3: 0.08, A4: 20 },
    expected: -59777.15,
    digits: 7,
    source: "Excel",
    note: "PV function, example: present value of a 20-year, 8% annuity paying $500/month; Microsoft prints ($59,777.15).",
  },
  {
    formula: "PV(0,12,-100)",
    expected: 1200,
    digits: 15,
    source: "pinned",
    note: "pinned: at rate=0 the TVM identity degenerates to pv + pmt*nper + fv = 0, so pv = -(-100*12) = 1200 exactly.",
  },
  {
    formula: "PV(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (2, PV needs rate, nper, pmt); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- NPER: rate, pmt, pv, [fv], [type] -----------------------------------
  {
    formula: "NPER(A2/12,A3,A4,A5,1)",
    cells: { A2: 0.12, A3: -100, A4: -1000, A5: 10000 },
    expected: 59.6738657,
    digits: 9,
    source: "Excel",
    note: "NPER function, example, row 1: periods to reach fv=10000 with payments at period start; Microsoft's live result 59.6738657.",
  },
  {
    formula: "NPER(A2/12,A3,A4,A5)",
    cells: { A2: 0.12, A3: -100, A4: -1000, A5: 10000 },
    expected: 60.0821229,
    digits: 9,
    source: "Excel",
    note: "NPER function, example, row 2: same terms, payments at period end; Microsoft's live result 60.0821229.",
  },
  {
    formula: "NPER(A2/12,A3,A4)",
    cells: { A2: 0.12, A3: -100, A4: -1000 },
    expected: -9.57859404,
    digits: 9,
    source: "Excel",
    note: "NPER function, example, row 3: fv omitted (0); Microsoft's live result -9.57859404.",
  },
  {
    formula: "NPER(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (2, NPER needs rate, pmt, pv); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- RATE: nper, pmt, pv, [fv], [type], [guess] --------------------------
  {
    formula: "RATE(A2*12,A3,A4)",
    cells: { A2: 4, A3: -200, A4: 8000 },
    expected: 0.0077,
    digits: 2,
    source: "Excel",
    note:
      "RATE function example: nper=4*12, pmt=-200, pv=8000. Microsoft's page shows this exact monthly rate as \"1%\" (0 decimal places) and the " +
      "annualized rate (the same monthly rate times 12, a separate formula/cell) as \"9.24%\" (2 decimal places); the 0-decimal cell is too " +
      "coarse to encode at any significant-digit count (0.77% rounds to 0.008 at 1 sig fig, not 0.01), so this row uses the more precise " +
      "figure: monthly = 9.24%/12 = 0.77% = 0.0077.",
  },
  {
    formula: "RATE(1,0,-100,110)",
    expected: 0.1,
    digits: 15,
    source: "pinned",
    note: "pinned: with pmt=0, nper=1, the TVM identity reduces to pv*(1+r)+fv=0 => -100*(1+r)+110=0 => r=0.1 exactly.",
  },
  {
    formula: "RATE(10,100,-1000,,,-1)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note:
      "pinned: guess=-1 (fv and type left empty) is itself <= -1; SPEC-127's Numerical methods pin returns #NUM! for a rate <= -1 without needing to iterate.",
  },

  // --- NPV: rate, value1, [value2], ... -------------------------------------
  {
    formula: "NPV(A2,A3,A4,A5,A6)",
    cells: { A2: 0.1, A3: -10000, A4: 3000, A5: 4200, A6: 6800 },
    expected: 1188.44,
    digits: 6,
    source: "Excel",
    note: "NPV function, example 1: net present value of a 3-year investment at 10% discount; Microsoft prints $1,188.44.",
  },
  {
    formula: "NPV(A2,A4:A8)",
    cells: { A2: 0.08, A4: 8000, A5: 9200, A6: 10000, A7: 12000, A8: 14500 },
    expected: 41922.06,
    digits: 7,
    source: "pinned",
    note:
      "pinned: NPV function, example 2 states NPV(A2,A4:A8)+A3 = $1,922.06 with A3=-40000 (the initial cost, added undiscounted outside NPV); " +
      "NPV(A2,A4:A8) alone = 1922.06 - (-40000) = 41922.06.",
  },
  {
    formula: "NPV(A2,A4:A8,-9000)",
    cells: { A2: 0.08, A4: 8000, A5: 9200, A6: 10000, A7: 12000, A8: 14500 },
    expected: 36250.53,
    digits: 7,
    source: "pinned",
    note:
      "pinned: the same example 2 states NPV(A2,A4:A8,-9000)+A3 = ($3,749.47) with A3=-40000, so NPV(A2,A4:A8,-9000) = -3749.47 - (-40000) = 36250.53.",
  },
  {
    formula: "NPV(0.1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (1, NPV needs rate, value1); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- XNPV: rate, values, dates ---------------------------------------------
  {
    formula: "XNPV(0.09,A2:A6,B2:B6)",
    cells: {
      A2: -10000, B2: 39448,
      A3: 2750, B3: 39508,
      A4: 4250, B4: 39751,
      A5: 3250, B5: 39859,
      A6: 2750, B6: 39904,
    },
    expected: 2086.65,
    digits: 6,
    source: "Excel",
    note:
      "XNPV function, example: 9% discount over payments dated 1/1/2008, 3/1/2008, 10/30/2008, 2/15/2009, 4/1/2009, converted to Excel serials " +
      "(days since 1899-12-30; 2008-10-30 = 39751 is SPEC-127's stated anchor, the rest computed by day-of-year difference, 2008 a leap year); " +
      "Microsoft prints $2,086.65.",
  },
  {
    formula: "XNPV(0.1,A1:A2,B1:B2)",
    cells: { A1: -1000, B1: 44000, A2: 1100, B2: 44365 },
    expected: 0,
    source: "pinned",
    note: "pinned: -1000 + 1100/(1.1)^((44365-44000)/365) = -1000 + 1100/1.1^1 = -1000 + 1000 = 0 (dates exactly 365 days apart).",
  },
  {
    formula: "XNPV(0.1,A1:A2,B1:B2)",
    cells: { A1: -1000, B1: 40000, A2: 600, B2: 39999 },
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "XNPV Remarks, verbatim: \"If any number in dates precedes the starting date, XNPV returns the #NUM! error value.\" (B2 < B1 here).",
  },
  {
    formula: "XNPV(0.1,A1:A3,B1:B2)",
    cells: { A1: -1000, A2: 500, A3: 600, B1: 40000, B2: 40100 },
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "XNPV Remarks, verbatim: \"If values and dates contain a different number of values, XNPV returns the #NUM! error value.\" (3 values, 2 dates).",
  },

  // --- IRR: values, [guess] ---------------------------------------------------
  {
    formula: "IRR(A2:A6)",
    cells: { A2: -70000, A3: 12000, A4: 15000, A5: 18000, A6: 21000 },
    expected: -0.021,
    digits: 2,
    source: "Excel",
    note: "IRR function, example, row 1: internal rate of return after four years; Microsoft prints -2.1%.",
  },
  {
    formula: "IRR(A2:A7)",
    cells: { A2: -70000, A3: 12000, A4: 15000, A5: 18000, A6: 21000, A7: 26000 },
    expected: 0.087,
    digits: 2,
    source: "Excel",
    note: "IRR function, example, row 2: internal rate of return after five years; Microsoft prints 8.7%.",
  },
  {
    formula: "IRR(A2:A4,-10%)",
    cells: { A2: -70000, A3: 12000, A4: 15000 },
    expected: -0.444,
    digits: 3,
    source: "Excel",
    note: "IRR function, example, row 3: internal rate of return after two years with guess=-10%; Microsoft prints -44.4%.",
  },
  {
    formula: "IRR({1,2,3})",
    expected: { error: "#NUM!" },
    source: "pinned",
    note:
      "pinned: SPEC-127's Numerical methods pin — a cash-flow series with no sign change has no rate zeroing its NPV, so IRR does not converge " +
      "and returns #NUM! (documented in the brief as IRR({1,2,3})).",
  },

  // --- XIRR: values, dates, [guess] --------------------------------------------
  {
    formula: "XIRR(A3:A7,B3:B7,0.1)",
    cells: {
      A3: -10000, B3: 39448,
      A4: 2750, B4: 39508,
      A5: 4250, B5: 39751,
      A6: 3250, B6: 39859,
      A7: 2750, B7: 39904,
    },
    expected: 0.373362535,
    digits: 8,
    source: "Excel",
    note:
      "XIRR function, example: the Description column states outright \"The internal rate of return (0.373362535 or 37.34%)\"; dates are the " +
      "same schedule as XNPV's example, converted to Excel serials from the stated 2008-10-30 = 39751 anchor. Compared at 8 digits: the " +
      "page's Remarks give XIRR's accuracy as 0.000001 percent (1e-8), so a ninth digit of 0.37… is below what Excel itself guarantees.",
  },
  {
    formula: "XIRR(A1:A2,B1:B2)",
    cells: { A1: -1000, B1: 44000, A2: 1100, B2: 44365 },
    expected: 0.1,
    digits: 15,
    source: "pinned",
    note: "pinned: 1000*(1+r) = 1100 with the two dates exactly 365 days apart (one period), so r = 0.1 exactly.",
  },
  {
    formula: "XIRR(A1:A3,B1:B3)",
    cells: { A1: -1000, B1: 40000, A2: 600, B2: 39999, A3: 600, B3: 40100 },
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — XNPV/XIRR truncate dates and a date before the first (B2=39999 < B1=40000) returns #NUM!.",
  },

  // --- MIRR: values, finance_rate, reinvest_rate --------------------------------
  {
    formula: "MIRR(A2:A7,A8,A9)",
    cells: { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A6: 37000, A7: 46000, A8: 0.1, A9: 0.12 },
    expected: 0.13,
    digits: 2,
    source: "Excel",
    note: "MIRR function, example, row 1: modified rate of return after five years; Microsoft prints 13%.",
  },
  {
    formula: "MIRR(A2:A5,A8,A9)",
    cells: { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A8: 0.1, A9: 0.12 },
    expected: -0.05,
    digits: 1,
    source: "Excel",
    note: "MIRR function, example, row 2: modified rate of return after three years; Microsoft prints -5%.",
  },
  {
    formula: "MIRR(A2:A7,A8,14%)",
    cells: { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A6: 37000, A7: 46000, A8: 0.1 },
    expected: 0.13,
    digits: 2,
    source: "Excel",
    note: "MIRR function, example, row 3: five-year modified rate with reinvest_rate=14% given as a literal; Microsoft prints 13%.",
  },
  {
    formula: "MIRR(A1:A2,0.1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (2, MIRR needs values, finance_rate, reinvest_rate); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- SLN: cost, salvage, life ---------------------------------------------------
  {
    formula: "SLN(A2,A3,A4)",
    cells: { A2: 30000, A3: 7500, A4: 10 },
    expected: 2250,
    digits: 4,
    source: "Excel",
    note: "SLN function, example: straight-line depreciation per year of a $30,000 asset with $7,500 salvage over 10 years; Microsoft prints $2,250.",
  },
  {
    formula: "SLN(1000,0,10)",
    expected: 100,
    digits: 15,
    source: "pinned",
    note: "pinned: SLN = (cost-salvage)/life = (1000-0)/10 = 100 exactly.",
  },
  {
    formula: "SLN(1000,0,0)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: SLN = (cost-salvage)/life divides by zero when life=0, so SLN(1000,0,0) is #DIV/0!.",
  },

  // --- SYD: cost, salvage, life, per -------------------------------------------------
  {
    formula: "SYD(A2,A3,A4,1)",
    cells: { A2: 30000, A3: 7500, A4: 10 },
    expected: 4090.91,
    digits: 6,
    source: "Excel",
    note: "SYD function, example, row 1: first-year depreciation allowance; Microsoft prints $4,090.91.",
  },
  {
    formula: "SYD(A2,A3,A4,10)",
    cells: { A2: 30000, A3: 7500, A4: 10 },
    expected: 409.09,
    digits: 5,
    source: "Excel",
    note: "SYD function, example, row 2: tenth-year depreciation allowance; Microsoft prints $409.09.",
  },
  {
    formula: "SYD(100,0,4,1)",
    expected: 40,
    digits: 15,
    source: "pinned",
    note: "pinned: SYD = (cost-salvage)*(life-per+1)*2/(life*(life+1)) = 100*4*2/(4*5) = 800/20 = 40 exactly.",
  },
  {
    formula: "SYD(1,2,3)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (3, SYD needs cost, salvage, life, per); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- DB: cost, salvage, life, period, [month] ---------------------------------------
  {
    formula: "DB(A2,A3,A4,1,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 186083.33,
    digits: 8,
    source: "Excel",
    note: "DB function, example, period 1 (7 months in the first year); Microsoft prints $186,083.33.",
  },
  {
    formula: "DB(A2,A3,A4,2,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 259639.42,
    digits: 8,
    source: "Excel",
    note: "DB function, example, period 2; Microsoft prints $259,639.42.",
  },
  {
    formula: "DB(A2,A3,A4,3,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 176814.44,
    digits: 8,
    source: "Excel",
    note: "DB function, example, period 3; Microsoft prints $176,814.44.",
  },
  {
    formula: "DB(A2,A3,A4,4,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 120410.64,
    digits: 8,
    source: "Excel",
    note: "DB function, example, period 4; Microsoft prints $120,410.64.",
  },
  {
    formula: "DB(A2,A3,A4,5,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 81999.64,
    digits: 7,
    source: "Excel",
    note: "DB function, example, period 5; Microsoft prints $81,999.64.",
  },
  {
    formula: "DB(A2,A3,A4,6,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 55841.76,
    digits: 7,
    source: "Excel",
    note: "DB function, example, period 6; Microsoft prints $55,841.76.",
  },
  {
    formula: "DB(A2,A3,A4,7,7)",
    cells: { A2: 1000000, A3: 100000, A4: 6 },
    expected: 15845.1,
    digits: 7,
    source: "Excel",
    note: "DB function, example, period 7 (only 5 months left); Microsoft prints $15,845.10.",
  },
  {
    formula: "DB(1,2,3)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (3, DB needs cost, salvage, life, period); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- DDB: cost, salvage, life, period, [factor] --------------------------------------
  {
    formula: "DDB(A2,A3,A4*365,1)",
    cells: { A2: 2400, A3: 300, A4: 10 },
    expected: 1.32,
    digits: 3,
    source: "Excel",
    note: "DDB function, example, row 1: first day's depreciation (life in days), default factor 2; Microsoft prints $1.32.",
  },
  {
    formula: "DDB(A2,A3,A4*12,1,2)",
    cells: { A2: 2400, A3: 300, A4: 10 },
    expected: 40,
    digits: 4,
    source: "Excel",
    note: "DDB function, example, row 2: first month's depreciation (life in months); Microsoft prints $40.00.",
  },
  {
    formula: "DDB(A2,A3,A4,1,2)",
    cells: { A2: 2400, A3: 300, A4: 10 },
    expected: 480,
    digits: 5,
    source: "Excel",
    note: "DDB function, example, row 3: first year's depreciation; Microsoft prints $480.00.",
  },
  {
    formula: "DDB(A2,A3,A4,2,1.5)",
    cells: { A2: 2400, A3: 300, A4: 10 },
    expected: 306,
    digits: 5,
    source: "Excel",
    note: "DDB function, example, row 4: second year's depreciation with factor 1.5; Microsoft prints $306.00.",
  },
  {
    formula: "DDB(A2,A3,A4,10)",
    cells: { A2: 2400, A3: 300, A4: 10 },
    expected: 22.12,
    digits: 4,
    source: "Excel",
    note: "DDB function, example, row 5: tenth year's depreciation, default factor 2; Microsoft prints $22.12.",
  },
  {
    formula: "DDB(1,2,3)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (3, DDB needs cost, salvage, life, period); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- CUMIPMT: rate, nper, pv, start_period, end_period, type -------------------------
  {
    formula: "CUMIPMT(A2/12,A3*12,A4,13,24,0)",
    cells: { A2: 0.09, A3: 30, A4: 125000 },
    expected: -11135.23213,
    digits: 10,
    source: "Excel",
    note: "CUMIPMT function, example, row 1: total interest in the second year (periods 13-24) of a 30-year, 9% loan on $125,000; Microsoft's live result -11135.23213.",
  },
  {
    formula: "CUMIPMT(A2/12,A3*12,A4,1,1,0)",
    cells: { A2: 0.09, A3: 30, A4: 125000 },
    expected: -937.5,
    digits: 4,
    source: "Excel",
    note: "CUMIPMT function, example, row 2: interest paid in the first month alone; Microsoft prints -937.5 ($125000*0.0075 exactly).",
  },
  {
    formula: "CUMIPMT(0,360,100000,1,12,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "CUMIPMT Remarks, verbatim: \"If rate <= 0, nper <= 0, or pv <= 0, CUMIPMT returns the #NUM! error value.\"",
  },

  // --- CUMPRINC: rate, nper, pv, start_period, end_period, type -------------------------
  {
    formula: "CUMPRINC(A2/12,A3*12,A4,13,24,0)",
    cells: { A2: 0.09, A3: 30, A4: 125000 },
    expected: -934.1071234,
    digits: 10,
    source: "Excel",
    note: "CUMPRINC function, example, row 1: total principal paid in the second year (periods 13-24); Microsoft's live result -934.1071234.",
  },
  {
    formula: "CUMPRINC(A2/12,A3*12,A4,1,1,0)",
    cells: { A2: 0.09, A3: 30, A4: 125000 },
    expected: -68.27827118,
    digits: 10,
    source: "Excel",
    note: "CUMPRINC function, example, row 2: principal paid in a single first-month payment; Microsoft's live result -68.27827118.",
  },
  {
    formula: "CUMPRINC(0,360,100000,1,12,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "CUMPRINC Remarks, verbatim: \"If rate <= 0, nper <= 0, or pv <= 0, CUMPRINC returns the #NUM! error value.\"",
  },

  // --- EFFECT: nominal_rate, npery -------------------------------------------------------
  {
    formula: "EFFECT(A2,A3)",
    cells: { A2: 0.0525, A3: 4 },
    expected: 0.0535427,
    digits: 6,
    source: "Excel",
    note: "EFFECT function, example: effective rate for a 5.25% nominal rate compounded quarterly; Microsoft prints 0.0535427.",
  },
  {
    formula: "EFFECT(0.12,12)",
    expected: 0.126825,
    digits: 6,
    source: "pinned",
    note: "pinned: EFFECT's stated formula effective = (1+nominal/npery)^npery - 1 = 1.01^12 - 1 ~= 0.126825 (via e^(12*ln 1.01) - 1).",
  },
  {
    formula: "EFFECT(0,4)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "EFFECT Remarks, verbatim: \"If nominal_rate <= 0 or if npery < 1, EFFECT returns the #NUM! error value.\"",
  },

  // --- NOMINAL: effect_rate, npery --------------------------------------------------------
  {
    formula: "NOMINAL(A2,A3)",
    cells: { A2: 0.053543, A3: 4 },
    expected: 0.05250032,
    digits: 7,
    source: "Excel",
    note: "NOMINAL function, example: nominal rate for a 5.3543% effective rate compounded quarterly; Microsoft prints 0.05250032.",
  },
  {
    formula: "NOMINAL(0.126825,12)",
    expected: 0.12,
    digits: 2,
    source: "pinned",
    note: "pinned: NOMINAL is EFFECT's inverse, nominal = npery*((1+effect)^(1/npery)-1); with effect=0.126825=1.01^12-1 (EFFECT(0.12,12) above), nominal = 12*(1.01-1) = 0.12.",
  },
  {
    formula: "NOMINAL(0,4)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "NOMINAL Remarks, verbatim: \"If effect_rate <= 0 or if npery < 1, NOMINAL returns the #NUM! error value.\"",
  },
];
