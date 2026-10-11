// SPEC-127 "Conformance tables": the Date & Time category's rows.
// Sources: Microsoft's function pages (support.microsoft.com/en-us/excel/functions/<name>-function)
// for ISOWEEKNUM, WORKDAY, WORKDAY.INTL, NETWORKDAYS, NETWORKDAYS.INTL, YEARFRAC, DATEDIF and
// DAYS360 (transcribed directly); every other row is "pinned" — hand-derived from this engine's
// stated date arithmetic (src/lib/spreadsheet/formula/functions/dateTime.ts's documented rules) or
// from SPEC-127's Function pins, never from running this engine or from memory of Microsoft's
// printed examples. Date serials are computed by hand: days since 1899-12-30 for dates on/after
// 1900-03-01 (the fictitious 1900-02-29 = serial 60), verified against the brief's anchors
// (1900-01-01 = 1, 2008-10-30 = 39751, 2026-10-08 = 46303).
import type { ConformanceRow } from "./row";

export const DATE_ROWS: ConformanceRow[] = [
  // DATE
  {
    formula: "DATE(2024,3,15)",
    expected: 45366,
    source: "pinned",
    note: "pinned: 2024-03-15 is 45366 days since 1899-12-30 (post-1900-03-01, fictitious 1900-02-29 included).",
  },
  {
    formula: "DATE(2026,10,8)",
    expected: 46303,
    source: "pinned",
    note: "pinned: matches the fixed conformance clock's TODAY() anchor, 2026-10-08 = 46303.",
  },
  {
    formula: "DATE(0,1,1)",
    expected: 1,
    source: "pinned",
    note: "pinned: dateTime.ts — a year 0-1899 has 1900 added, so DATE(0,1,1) is DATE(1900,1,1) = serial 1.",
  },
  {
    formula: "DATE(10000,1,1)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts — year outside 0-9999 (after the 0-1899 adjustment) is #NUM!.",
  },
  // TIME
  {
    formula: "TIME(14,30,0)",
    expected: 0.604166666666667,
    digits: 15,
    source: "pinned",
    note: "pinned: (14*3600+30*60)/86400 = 52200/86400 = 0.604166666666667 (2:30 PM as a day fraction).",
  },
  {
    formula: "TIME(0,0,0)",
    expected: 0,
    source: "pinned",
    note: "pinned: midnight is 0 seconds into the day.",
  },
  {
    formula: "TIME(24,0,0)",
    expected: 0,
    source: "pinned",
    note: "pinned: dateTime.ts — total seconds (24*3600=86400) taken modulo 86400 wraps to 0.",
  },
  {
    formula: "TIME(-1,0,0)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts — a negative total number of seconds is #NUM!.",
  },
  // YEAR
  {
    formula: "YEAR(45366)",
    expected: 2024,
    source: "pinned",
    note: "pinned: serial 45366 is 2024-03-15 (see DATE's derivation above).",
  },
  {
    formula: "YEAR(1)",
    expected: 1900,
    source: "pinned",
    note: "pinned: serial 1 is 1900-01-01.",
  },
  {
    formula: "YEAR(60)",
    expected: 1900,
    source: "pinned",
    note: "pinned: serial 60 is Excel's fictitious 1900-02-29.",
  },
  {
    formula: "YEAR(-1)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts's serialArg rejects a negative serial with #NUM!.",
  },
  // MONTH
  {
    formula: "MONTH(45366)",
    expected: 3,
    source: "pinned",
    note: "pinned: serial 45366 is 2024-03-15.",
  },
  {
    formula: "MONTH(60)",
    expected: 2,
    source: "pinned",
    note: "pinned: serial 60 is the fictitious 1900-02-29.",
  },
  {
    formula: "MONTH(1)",
    expected: 1,
    source: "pinned",
    note: "pinned: serial 1 is 1900-01-01.",
  },
  {
    formula: "MONTH(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate — an error argument returns itself.",
  },
  // DAY
  {
    formula: "DAY(45366)",
    expected: 15,
    source: "pinned",
    note: "pinned: serial 45366 is 2024-03-15.",
  },
  {
    formula: "DAY(60)",
    expected: 29,
    source: "pinned",
    note: "pinned: serial 60 is the fictitious 1900-02-29.",
  },
  {
    formula: "DAY(1)",
    expected: 1,
    source: "pinned",
    note: "pinned: serial 1 is 1900-01-01.",
  },
  {
    formula: "DAY()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a wrong argument count (0, needs 1) is #VALUE!.",
  },
  // HOUR
  {
    formula: "HOUR(46303.395833333333)",
    expected: 9,
    source: "pinned",
    note: "pinned: the fixed conformance clock, 09:30, is serial 46303 + 9.5/24.",
  },
  {
    formula: "HOUR(0.5)",
    expected: 12,
    source: "pinned",
    note: "pinned: 0.5 of a day is noon, hour 12.",
  },
  {
    formula: "HOUR(0)",
    expected: 0,
    source: "pinned",
    note: "pinned: serial with no fraction is midnight.",
  },
  {
    formula: "HOUR(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  // MINUTE
  {
    formula: "MINUTE(46303.395833333333)",
    expected: 30,
    source: "pinned",
    note: "pinned: the fixed conformance clock is 09:30.",
  },
  {
    formula: "MINUTE(0.5)",
    expected: 0,
    source: "pinned",
    note: "pinned: noon is exactly on the hour.",
  },
  {
    formula: "MINUTE(0.00694444444444444)",
    expected: 10,
    source: "pinned",
    note: "pinned: 600/86400 = 0:10:00, ten minutes past midnight.",
  },
  {
    formula: "MINUTE()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // SECOND
  {
    formula: "SECOND(46303.395833333333)",
    expected: 0,
    source: "pinned",
    note: "pinned: the fixed conformance clock is 09:30:00 exactly.",
  },
  {
    formula: "SECOND(1/86400)",
    expected: 1,
    source: "pinned",
    note: "pinned: 1/86400 of a day is one second past midnight.",
  },
  {
    formula: "SECOND(0)",
    expected: 0,
    source: "pinned",
    note: "pinned: no fractional part is zero seconds.",
  },
  {
    formula: "SECOND(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  // TODAY
  {
    formula: "TODAY()",
    expected: 46303,
    source: "pinned",
    note: "pinned: the fixed conformance clock is 2026-10-08 09:30, so TODAY() = 46303. Microsoft's own TODAY example is day-dependent and is skipped per the brief.",
  },
  {
    formula: "TODAY(1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — TODAY takes 0 arguments; 1 is a wrong argument count, #VALUE!.",
  },
  {
    formula: "TODAY(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a second wrong-argument-count variant, #VALUE! (a zero-arg function can only be varied this way).",
  },
  // NOW
  {
    formula: "NOW()",
    expected: 46303.395833333336,
    digits: 15,
    source: "pinned",
    note: "pinned: the fixed conformance clock is 2026-10-08 09:30 = 46303 + 9.5/24 = 46303.395833333333. Microsoft's own NOW example is moment-dependent and is skipped per the brief.",
  },
  {
    formula: "NOW(1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — NOW takes 0 arguments; 1 is a wrong argument count, #VALUE!.",
  },
  {
    formula: "NOW(1,2)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a second wrong-argument-count variant, #VALUE!.",
  },
  // WEEKDAY
  {
    formula: "WEEKDAY(45366)",
    expected: 6,
    source: "pinned",
    note: "pinned: serial 45366 is 2024-03-15, a Friday; default return_type 1 (Sunday=1) makes Friday 6. Excel's weekday arithmetic matches the real calendar from 1900-03-01 onward.",
  },
  {
    formula: "WEEKDAY(1)",
    expected: 1,
    source: "pinned",
    note: "pinned: dateTime.ts — serial 1 (1900-01-01) is defined as Sunday under Excel's (⌊serial⌋+6) mod 7 arithmetic, return_type 1 default.",
  },
  {
    formula: "WEEKDAY(45366,2)",
    expected: 5,
    source: "pinned",
    note: "pinned: return_type 2 (Monday=1); 2024-03-15 is a Friday, the 5th day when the week starts Monday.",
  },
  {
    formula: "WEEKDAY(45366,99)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts — a return_type outside {1,2,3,11..17} is #NUM!.",
  },
  // EDATE
  {
    formula: "EDATE(45322,1)",
    expected: 45351,
    source: "pinned",
    note: "pinned: 45322 = 2024-01-31; one month later clamps to the shorter month's end, 2024-02-29 (leap) = 45351.",
  },
  {
    formula: "EDATE(45306,2)",
    expected: 45366,
    source: "pinned",
    note: "pinned: 45306 = 2024-01-15; two months later, same day-of-month, is 2024-03-15 = 45366.",
  },
  {
    formula: "EDATE(45366,0)",
    expected: 45366,
    source: "pinned",
    note: "pinned: zero months later is the same date.",
  },
  {
    formula: "EDATE(45366,120000)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts — shifting the year outside 1900-9999 is #NUM!; 120000 months (10000 years) from 2024 overflows far past 9999.",
  },
  // EOMONTH
  {
    formula: "EOMONTH(45306,0)",
    expected: 45322,
    source: "pinned",
    note: "pinned: 45306 = 2024-01-15; the last day of January 2024 is 45322 (2024-01-31).",
  },
  {
    formula: "EOMONTH(45306,1)",
    expected: 45351,
    source: "pinned",
    note: "pinned: one month after January 2024 is February 2024 (leap); its last day is 45351 (2024-02-29).",
  },
  {
    formula: "EOMONTH(45322,0)",
    expected: 45322,
    source: "pinned",
    note: "pinned: 45322 (2024-01-31) is already the last day of its month.",
  },
  {
    formula: "EOMONTH(45366,120000)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: dateTime.ts — shifting the year outside 1900-9999 is #NUM!.",
  },
  // DAYS
  {
    formula: "DAYS(45352,45292)",
    expected: 60,
    source: "pinned",
    note: "pinned: DAYS(end_date,start_date) — end 45352 (2024-03-01) minus start 45292 (2024-01-01) is 60 days (31 Jan + 29 Feb, 2024 is a leap year).",
  },
  {
    formula: "DAYS(45292,45352)",
    expected: -60,
    source: "pinned",
    note: "pinned: arguments reversed from the row above, so the result's sign flips.",
  },
  {
    formula: "DAYS(1,1)",
    expected: 0,
    source: "pinned",
    note: "pinned: the same date both ends gives zero days.",
  },
  {
    formula: "DAYS(#N/A,1)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  // DATEVALUE
  {
    formula: 'DATEVALUE("3/15/2024")',
    expected: 45366,
    source: "pinned",
    note: "pinned: en-US month/day/year parse of 2024-03-15 is serial 45366.",
  },
  {
    formula: 'DATEVALUE("1/1/1900")',
    expected: 1,
    source: "pinned",
    note: "pinned: 1900-01-01 is serial 1.",
  },
  {
    formula: 'DATEVALUE("3/15/2024 2:30 PM")',
    expected: 45366,
    source: "pinned",
    note: 'pinned: per the Dates pin, DATEVALUE keeps only the integer (date) part even when the text also carries a time; 2024-03-15 is 45366.',
  },
  {
    formula: 'DATEVALUE("not a date")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: per the Dates pin, text that is not a recognisable date is #VALUE!.",
  },
  // TIMEVALUE
  {
    formula: 'TIMEVALUE("2:30 PM")',
    expected: 0.604166666666667,
    digits: 15,
    source: "pinned",
    note: "pinned: 2:30 PM is 52200/86400 = 0.604166666666667 of a day.",
  },
  {
    formula: 'TIMEVALUE("3/15/2024 2:30 PM")',
    expected: 0.604166666666667,
    digits: 15,
    source: "pinned",
    note: "pinned: per the Dates pin, TIMEVALUE keeps only the fractional (time) part even when the text also carries a date.",
  },
  {
    formula: 'TIMEVALUE("0:00")',
    expected: 0,
    source: "pinned",
    note: "pinned: midnight is a zero day-fraction.",
  },
  {
    formula: 'TIMEVALUE("not a time")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: per the Dates pin, text that is not a recognisable time is #VALUE!.",
  },
  // WEEKNUM
  {
    formula: "WEEKNUM(45292,1)",
    expected: 1,
    source: "pinned",
    note: "pinned: 45292 is 2024-01-01; the week containing January 1 is always week 1.",
  },
  {
    formula: "WEEKNUM(45366,1)",
    expected: 11,
    source: "pinned",
    note: "pinned: 45366 is 2024-03-15, day-of-year 75 (31+29+15, 2024 is a leap year); 2024-01-01 is a Monday (Sunday-indexed weekday 1); weeknum = floor((75+1-1)/7)+1 = 11 for return_type 1 (weeks start Sunday).",
  },
  {
    formula: "WEEKNUM(45366,2)",
    expected: 11,
    source: "pinned",
    note: "pinned: return_type 2 (weeks start Monday); since 2024-01-01 is itself a Monday (index 0), weeknum = floor((75+0-1)/7)+1 = 11, the same as return_type 1 here.",
  },
  {
    formula: "WEEKNUM(45366,99)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: an invalid return_type (not 1, 2, 11-17 or 21) is #NUM!.",
  },
  // ISOWEEKNUM
  {
    formula: "ISOWEEKNUM(A2)",
    cells: { A2: 40977 },
    expected: 10,
    source: "Excel",
    note: "ISOWEEKNUM function, Example: A2 = 3/9/2012 (2012-03-09, serial 40977). =ISOWEEKNUM(A2) is 10.",
  },
  {
    formula: "ISOWEEKNUM(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  {
    formula: "ISOWEEKNUM()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // WORKDAY
  {
    formula: "WORKDAY(A2,A3)",
    cells: { A2: 39722, A3: 151 },
    expected: 39933,
    source: "Excel",
    note: "WORKDAY function, Example row 1: A2 = 10/1/2008 (2008-10-01, serial 39722), A3 = 151. =WORKDAY(A2,A3) is 151 workdays later, 4/30/2009 (serial 39933).",
  },
  {
    formula: "WORKDAY(A2,A3,A4:A6)",
    cells: { A2: 39722, A3: 151, A4: 39778, A5: 39786, A6: 39834 },
    expected: 39938,
    source: "Excel",
    note: "WORKDAY function, Example row 2: same A2/A3, holidays A4:A6 = 11/26/2008, 12/4/2008, 1/21/2009 (serials 39778, 39786, 39834). Result 5/5/2009 (serial 39938).",
  },
  {
    formula: 'WORKDAY("abc",10)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'WORKDAY function, Remarks: "If any argument is not a valid date, WORKDAY returns the #VALUE! error value."',
  },
  {
    formula: "WORKDAY(1,-10)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'WORKDAY function, Remarks: "If start_date plus days yields an invalid date, WORKDAY returns the #NUM! error value" — serial 1 minus 10 workdays runs before the start of the calendar.',
  },
  // WORKDAY.INTL
  {
    formula: "WORKDAY.INTL(DATE(2012,1,1),30,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'WORKDAY.INTL function, Example row 1: "Using a 0 for the Weekend argument results in a #NUM! error."',
  },
  {
    formula: "WORKDAY.INTL(DATE(2012,1,1),90,11)",
    expected: 41013,
    source: "Excel",
    note: "WORKDAY.INTL function, Example row 2: 90 workdays from 2012-01-01, weekend code 11 (Sunday only). Result 41013 (2012-04-14).",
  },
  {
    formula: "WORKDAY.INTL(DATE(2012,1,1),30,17)",
    expected: 40944,
    source: "Excel",
    note: 'WORKDAY.INTL function, Example row 3: the page wraps this in TEXT to display "2/05/2012" but states the underlying "resulting serial number (40944)" outright; weekend code 17 (Saturday only).',
  },
  {
    formula: 'WORKDAY.INTL(DATE(2012,1,1),30,"1111111")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: SPEC-127\'s Function pins — a Monday-first weekend mask of all 1s leaves no workday in the week, "1111111" → #VALUE!.',
  },
  // NETWORKDAYS
  {
    formula: "NETWORKDAYS(A2,A3)",
    cells: { A2: 41183, A3: 41334 },
    expected: 110,
    source: "Excel",
    note: "NETWORKDAYS function, Example row 1: A2 = 10/1/2012 (serial 41183), A3 = 3/1/2013 (serial 41334). 110 workdays between them.",
  },
  {
    formula: "NETWORKDAYS(A2,A3,A4)",
    cells: { A2: 41183, A3: 41334, A4: 41235 },
    expected: 109,
    source: "Excel",
    note: "NETWORKDAYS function, Example row 2: A4 = 11/22/2012 (serial 41235) as a single holiday.",
  },
  {
    formula: "NETWORKDAYS(A2,A3,A4:A6)",
    cells: { A2: 41183, A3: 41334, A4: 41235, A5: 41247, A6: 41295 },
    expected: 107,
    source: "Excel",
    note: "NETWORKDAYS function, Example row 3: A4:A6 = 11/22/2012, 12/4/2012, 1/21/2013 (serials 41235, 41247, 41295) as three holidays.",
  },
  {
    formula: 'NETWORKDAYS("abc",A3)',
    cells: { A3: 41334 },
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'NETWORKDAYS function, Remarks: "If any argument is not a valid date, NETWORKDAYS returns the #VALUE! error value."',
  },
  // NETWORKDAYS.INTL
  {
    formula: "NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))",
    expected: 22,
    source: "Excel",
    note: "NETWORKDAYS.INTL function, Example row 1: default weekend (Saturday/Sunday); 31 total days less 9 weekend days is 22.",
  },
  {
    formula: "NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))",
    expected: -21,
    source: "Excel",
    note: "NETWORKDAYS.INTL function, Example row 2: start after end gives a negative result, -21.",
  },
  {
    formula: "NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{38719,38733})",
    expected: 22,
    source: "Excel",
    note: 'NETWORKDAYS.INTL function, Example row 3: weekend code 7 (Friday/Saturday), holidays 2006/1/2 and 2006/1/16 (serials 38719, 38733; Microsoft\'s page writes these as text dates in the array constant — given here as the equivalent serial numbers to avoid conflating with text-date parsing). Result 22.',
  },
  {
    formula: 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{38719,38733})',
    expected: 20,
    source: "Excel",
    note: "NETWORKDAYS.INTL function, Example row 4: same period and holidays as above, weekend mask \"0010001\" (Sunday and Wednesday off). Result 20.",
  },
  {
    formula: 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),"1111111")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: SPEC-127\'s Function pins — a Monday-first weekend mask of all 1s leaves no workday in the week, "1111111" → #VALUE!.',
  },
  // YEARFRAC
  {
    formula: "YEARFRAC(A2,A3)",
    cells: { A2: 40909, A3: 41120 },
    expected: 0.58055556,
    digits: 8,
    source: "Excel",
    note: "YEARFRAC function, Example row 1: A2 = 1/1/2012 (serial 40909), A3 = 7/30/2012 (serial 41120), basis omitted (0, US 30/360).",
  },
  {
    formula: "YEARFRAC(A2,A3,1)",
    cells: { A2: 40909, A3: 41120 },
    expected: 0.57650273,
    digits: 8,
    source: "Excel",
    note: "YEARFRAC function, Example row 2: basis 1 (actual/actual); 2012 is a leap year so the page uses a 366-day basis.",
  },
  {
    formula: "YEARFRAC(A2,A3,3)",
    cells: { A2: 40909, A3: 41120 },
    expected: 0.57808219,
    digits: 8,
    source: "Excel",
    note: "YEARFRAC function, Example row 3: basis 3 (actual/365).",
  },
  {
    formula: "YEARFRAC(A2,A3,2)",
    cells: { A2: 40909, A3: 41120 },
    expected: 0.586111111111111,
    digits: 15,
    source: "pinned",
    note: "pinned: basis 2 (actual/360) — 211 actual calendar days between 2012-01-01 and 2012-07-30 (31+29+31+30+31+30+29, 2012 is a leap year), divided by 360 = 0.586111111111111.",
  },
  {
    formula: "YEARFRAC(A2,A3,4)",
    cells: { A2: 40909, A3: 41120 },
    expected: 0.580555555555556,
    digits: 15,
    source: "pinned",
    note: "pinned: basis 4 (European 30/360) — neither 2012-01-01 (day 1) nor 2012-07-30 (day 30) is the 31st of its month, so the day components are unchanged: (0)*360+(7-1)*30+(30-1)=209 days / 360 = 0.580555555555556 (coincides with basis 0 for this pair since no month-end rule triggers).",
  },
  {
    formula: "YEARFRAC(A2,A3,5)",
    cells: { A2: 40909, A3: 41120 },
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "YEARFRAC function, Remarks: a basis outside 0-4 returns #NUM!.",
  },
  // DATEDIF
  {
    formula: 'DATEDIF(A2,B2,"Y")',
    cells: { A2: 36892, B2: 37622 },
    expected: 2,
    source: "Excel",
    note: 'DATEDIF function, Example: start 1/1/2001 (serial 36892), end 1/1/2003 (serial 37622) — the page writes this with named references Start_date/End_date; given here as plain cells A2/B2 holding the same values. =DATEDIF(A2,B2,"Y") is 2.',
  },
  {
    formula: 'DATEDIF(A3,B3,"D")',
    cells: { A3: 37043, B3: 37483 },
    expected: 440,
    source: "Excel",
    note: 'DATEDIF function, Example: start 6/1/2001 (serial 37043), end 8/15/2002 (serial 37483) — plain cells A3/B3 standing in for the page\'s named Start_date/End_date. =DATEDIF(A3,B3,"D") is 440.',
  },
  {
    formula: 'DATEDIF(A3,B3,"YD")',
    cells: { A3: 37043, B3: 37483 },
    expected: 75,
    source: "Excel",
    note: 'DATEDIF function, Example: same A3/B3 as above. =DATEDIF(A3,B3,"YD") is 75 (the day-of-year difference, ignoring years).',
  },
  {
    formula: 'DATEDIF(A3,B3,"M")',
    cells: { A3: 37043, B3: 37483 },
    expected: 14,
    source: "pinned",
    note: 'pinned: "M" is the number of complete months between 2001-06-01 and 2002-08-15 — (2002-2001)*12+(8-6)=14 complete months, since the end day (15) is not before the start day (1).',
  },
  {
    formula: 'DATEDIF(A3,B3,"MD")',
    cells: { A3: 37043, B3: 37483 },
    expected: 14,
    source: "pinned",
    note: 'pinned: per the Dates pin, "MD" (and DATEDIF generally) is computed as Microsoft documents it, hand-derived — the day components alone, ignoring months/years: 15 (end day) - 1 (start day) = 14.',
  },
  {
    formula: 'DATEDIF(A3,B3,"YM")',
    cells: { A3: 37043, B3: 37483 },
    expected: 2,
    source: "pinned",
    note: 'pinned: "YM" is the month components alone, ignoring days/years: 8 (end month) - 6 (start month) = 2.',
  },
  {
    formula: 'DATEDIF(B2,A2,"Y")',
    cells: { A2: 36892, B2: 37622 },
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "DATEDIF function, Remarks: if start_date exceeds end_date, DATEDIF returns #NUM! — arguments reversed from the Y example above (B2 2003 as start, A2 2001 as end).",
  },
  // DAYS360
  {
    formula: "DAYS360(A3,A4)",
    cells: { A2: 40544, A3: 40573, A4: 40575, A5: 40908 },
    expected: 1,
    source: "Excel",
    note: "DAYS360 function, Example row 1: A3 = 30-Jan-11 (serial 40573), A4 = 1-Feb-11 (serial 40575). US method (default): 1 day.",
  },
  {
    formula: "DAYS360(A2,A5)",
    cells: { A2: 40544, A3: 40573, A4: 40575, A5: 40908 },
    expected: 360,
    source: "Excel",
    note: "DAYS360 function, Example row 2: A2 = 1-Jan-11 (serial 40544), A5 = 31-Dec-11 (serial 40908). US method: 360 days (A5, the last day of December, with A2 earlier than the 30th, rolls A5 to the 1st of the next month).",
  },
  {
    formula: "DAYS360(A2,A4)",
    cells: { A2: 40544, A3: 40573, A4: 40575, A5: 40908 },
    expected: 30,
    source: "Excel",
    note: "DAYS360 function, Example row 3: A2 = 1-Jan-11, A4 = 1-Feb-11. US method: 30 days.",
  },
  {
    formula: "DAYS360(A2,A5,TRUE)",
    cells: { A2: 40544, A3: 40573, A4: 40575, A5: 40908 },
    expected: 359,
    source: "pinned",
    note: "pinned: European method only clamps a 31st-of-the-month date to the 30th. A2 (Jan 1) is unaffected; A5 (Dec 31) becomes Dec 30. (12-1)*30+(30-1) = 359 — one less than the US method's 360, since the European rule has no 'roll to next month' special case.",
  },
  {
    formula: "DAYS360(A2)",
    cells: { A2: 40544 },
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (1, needs at least 2) is #VALUE!.",
  },
];
