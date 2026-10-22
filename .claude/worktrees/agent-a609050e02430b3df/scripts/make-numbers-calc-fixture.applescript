-- Generates e2e/fixtures/spec120-numbers.xlsx for SPEC-120: the Numbers
-- oracle. One table whose column E holds formulas from SPEC-120's function
-- list that Numbers also implements; Numbers computes them and "Export To ▸
-- Excel" writes its results as the cached values. numbersOracle.spec.ts
-- recalculates the file and compares — a third-party engine's arithmetic,
-- which no builder can emulate.
--
-- Run from the repo root, on a Mac with Numbers installed and NOT running:
--   open -a Numbers --args -AppleLocale en_US    (first, or `activate`
--                                                 below fails with -600)
--   osascript scripts/make-numbers-calc-fixture.applescript
--
-- The locale argument makes Numbers parse and render formulas as en-US
-- whatever the Mac's region (a German region wants `ROUND(2,5;0)` and turns
-- 1/3 into "0,333…"), so its answers are comparable with the engine's
-- en-US defaults. If Numbers is already open, quit it first: the argument
-- only applies at launch.
--
-- Layout: row 1 is a header; rows 2–6 are a small data block (names,
-- amounts, a mixed column) that the aggregate, criteria and lookup rows
-- reference absolutely; every other row carries its own inputs in A–C.
-- (Numbers' rows: the export puts the table name in a row above them, so
-- every cell lands one row lower in the .xlsx — see spec120-numbers.md.)
-- Column D is the formula's text without "=", as SPEC-119 stores it, for
-- the report. Booleans never appear as inputs: Numbers exports a checkbox
-- as the number 1 (see spec116-numbers.md), so TRUE/FALSE are literals.
--
-- The list holds only functions Numbers 14.4 computes: it lacks IFNA, XOR,
-- DAYS, ISNA, ISERR, ISLOGICAL and TEXT. AppleScript cannot tell a function
-- Numbers lacks from a formula that evaluates to an error (both read as
-- missing value), so after the export this script checks the file itself:
-- Numbers drops some unknown functions from the export entirely, which is
-- a FAIL here, while an error result keeps its formula and has no cached
-- value (reported, not a failure — numbersOracle.spec.ts expects an error).
-- D must also still read as the text it was given. If any step reports
-- FAIL, fix and re-run until every row is ok, then commit the script and
-- the .xlsx together and update e2e/fixtures/spec120-numbers.md (the
-- Numbers version, and anything new about the export's shape).
-- Numbers may prompt for automation permission on first run; approve it
-- and re-run.

property mv : missing value

-- {A, B, C, formula}; a bare formula has no inputs. "{r}" is the row's own
-- number. Text inputs are stored as text (format set first), so "5" stays
-- the text "5".
property cases : {¬
  {"apple", 10, 1, "=SUM($B$2:$B$6)"}, ¬
  {"banana", 20, 2, "=AVERAGE($B$2:$B$6)"}, ¬
  {"cherry", 30, 3, "=COUNT($C$2:$C$6)"}, ¬
  {"date", 40, mv, "=COUNTA($C$2:$C$6)"}, ¬
  {"Banana", 50, "x", "=COUNTBLANK($C$2:$C$6)"}, ¬
  {1, 2, 3, "=A{r}+B{r}*C{r}"}, ¬
  {7, 2, mv, "=A{r}/B{r}"}, ¬
  {2, 3, 2, "=A{r}^B{r}^C{r}"}, ¬
  {2, mv, mv, "=-A{r}^2"}, ¬
  {2, mv, mv, "=0-A{r}^2"}, ¬
  {0.1, 0.2, 0.3, "=A{r}+B{r}=C{r}"}, ¬
  {0.1, 0.2, 0.3, "=A{r}+B{r}-C{r}"}, ¬
  {1, 3, mv, "=A{r}/B{r}"}, ¬
  {1, 0, mv, "=A{r}/B{r}"}, ¬
  {0, 0, mv, "=A{r}^B{r}"}, ¬
  {1, 3, mv, "=A{r}/B{r}&\"\""}, ¬
  {"5", mv, mv, "=A{r}+1"}, ¬
  {mv, mv, mv, "=A{r}"}, ¬
  {mv, mv, mv, "=A{r}=0"}, ¬
  {mv, mv, mv, "=A{r}=\"\""}, ¬
  {mv, mv, mv, "=A{r}=FALSE"}, ¬
  "=\"1\"+\"2\"", ¬
  "=\"1\"+\"$4.00\"", ¬
  "=\"6/1/2001\"-\"5/1/2001\"", ¬
  "=SQRT(\"8+1\")", ¬
  "=\"A\"&TRUE", ¬
  "=1+\"\"", ¬
  "=\"TRUE\"+0", ¬
  "=\"a\"*1", ¬
  "=TRUE+TRUE", ¬
  "=1=\"1\"", ¬
  "=1<\"1\"", ¬
  "=\"z\"<TRUE", ¬
  "=FALSE<TRUE", ¬
  "=\"A\"=\"a\"", ¬
  "=\"a\"<\"B\"", ¬
  "=NA()+1/0", ¬
  "=1/0+NA()", ¬
  "=IF(TRUE,1,1/0)", ¬
  "=ROUND(2.5,0)", ¬
  "=ROUND(-2.5,0)", ¬
  "=ROUND(1.005,2)", ¬
  "=ROUNDUP(3.14159,2)", ¬
  "=ROUNDDOWN(-3.14159,1)", ¬
  "=INT(-8.9)", ¬
  "=TRUNC(-8.9)", ¬
  "=MOD(-3,2)", ¬
  "=MOD(3,-2)", ¬
  "=CEILING(-2.5,2)", ¬
  "=CEILING(-2.5,-2)", ¬
  "=FLOOR(-2.5,-2)", ¬
  "=MROUND(10,3)", ¬
  "=FLOOR(0.3,0.1)", ¬
  "=ABS(-4.25)", ¬
  "=SIGN(-0.5)", ¬
  "=SQRT(16)", ¬
  "=POWER(2,10)", ¬
  "=EXP(1)", ¬
  "=LN(10)", ¬
  "=LOG(8,2)", ¬
  "=LOG10(1000)", ¬
  "=PI()", ¬
  {2, 3, 4, "=PRODUCT(A{r}:C{r})"}, ¬
  "=SUMPRODUCT($B$2:$B$4,$C$2:$C$4)", ¬
  "=SUMIF($A$2:$A$6,\"banana\",$B$2:$B$6)", ¬
  "=SUMIF($B$2:$B$6,\">25\")", ¬
  "=SUMIFS($B$2:$B$6,$A$2:$A$6,\"b*\",$C$2:$C$6,\">1\")", ¬
  "=AVERAGEIF($B$2:$B$6,\">=20\")", ¬
  "=AVERAGEIFS($B$2:$B$6,$C$2:$C$6,\"<3\")", ¬
  "=COUNTIF($A$2:$A$6,\"?a*\")", ¬
  "=COUNTIFS($B$2:$B$6,\">10\",$B$2:$B$6,\"<50\")", ¬
  "=MINIFS($B$2:$B$6,$A$2:$A$6,\"<>apple\")", ¬
  "=MAXIFS($B$2:$B$6,$C$2:$C$6,\">0\")", ¬
  "=MIN($B$2:$B$6)", ¬
  "=MAX($B$2:$B$6)", ¬
  "=SUM(\"3\")", ¬
  "=SUM(\"3\",TRUE)", ¬
  {"3", mv, mv, "=SUM(A{r}:B{r})"}, ¬
  {mv, mv, mv, "=AVERAGE(A{r}:C{r})"}, ¬
  {mv, mv, mv, "=MIN(A{r}:C{r})"}, ¬
  "=MAX(\"a\")", ¬
  {"1", 1, 2, "=COUNTIF(A{r}:C{r},1)"}, ¬
  {7, mv, mv, "=IF(A{r}>5,\"big\",\"small\")"}, ¬
  {3, mv, mv, "=IFS(A{r}<0,\"neg\",A{r}=0,\"zero\",TRUE,\"pos\")"}, ¬
  {0, mv, mv, "=IFERROR(1/A{r},\"none\")"}, ¬
  "=IFERROR(MATCH(\"zzz\",$A$2:$A$6,0),\"missing\")", ¬
  "=IFERROR(NA(),\"n/a\")", ¬
  "=AND(TRUE,FALSE)", ¬
  {0, 2, mv, "=OR(A{r}>1,B{r}>1)"}, ¬
  "=NOT(FALSE)", ¬
  {2, mv, mv, "=SWITCH(A{r},1,\"one\",2,\"two\",\"other\")"}, ¬
  "=CHOOSE(2.9,\"a\",\"b\",\"c\")", ¬
  {mv, mv, mv, "=ISBLANK(A{r})"}, ¬
  {"5", mv, mv, "=ISNUMBER(A{r})"}, ¬
  {"5", mv, mv, "=ISTEXT(A{r})"}, ¬
  "=ISERROR(1/0)", ¬
  "=ISERROR(NA())", ¬
  {"ab", 12, mv, "=CONCATENATE(A{r},\"-\",B{r})"}, ¬
  {"a", "b", "c", "=CONCAT(A{r}:C{r})"}, ¬
  {"x", mv, "z", "=TEXTJOIN(\", \",TRUE,A{r}:C{r})"}, ¬
  {"hello world", mv, mv, "=LEN(A{r})"}, ¬
  {"spreadsheet", mv, mv, "=LEFT(A{r},3)"}, ¬
  {"spreadsheet", mv, mv, "=RIGHT(A{r},5)"}, ¬
  {"spreadsheet", mv, mv, "=MID(A{r},3,4)"}, ¬
  "=UPPER(\"mixed Case\")", ¬
  "=LOWER(\"Mixed CASE\")", ¬
  "=PROPER(\"hello wide world\")", ¬
  "=PROPER(\"it's\")", ¬
  "=TRIM(\"  a   b  \")", ¬
  "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", ¬
  "=REPLACE(\"abcdef\",2,3,\"X\")", ¬
  "=FIND(\"b\",\"abcb\")", ¬
  "=FIND(\"B\",\"abcb\")", ¬
  "=SEARCH(\"B\",\"abcb\")", ¬
  "=SEARCH(\"c*b\",\"abcb\")", ¬
  "=VALUE(\"12.5\")", ¬
  "=REPT(\"ab\",3)", ¬
  "=EXACT(\"a\",\"A\")", ¬
  "=LEN(\"😀\")", ¬
  "=DATE(2026,10,8)", ¬
  "=DATE(2008,14,2)", ¬
  "=DATE(108,1,2)", ¬
  "=DATE(1900,2,29)", ¬
  "=YEAR(DATE(2026,10,8))", ¬
  "=MONTH(DATE(2026,10,8))", ¬
  "=DAY(DATE(2026,10,8))", ¬
  "=TIME(12,30,0)", ¬
  "=HOUR(TIME(13,45,30))", ¬
  "=MINUTE(TIME(13,45,30))", ¬
  "=SECOND(TIME(13,45,30))", ¬
  "=WEEKDAY(DATE(2026,10,8))", ¬
  "=WEEKDAY(DATE(2026,10,8),2)", ¬
  "=WEEKDAY(DATE(1900,1,1))", ¬
  "=EDATE(DATE(2026,1,31),1)", ¬
  "=EOMONTH(DATE(2026,2,10),0)", ¬
  "=DATE(2026,10,8)-DATE(2026,1,1)", ¬
  "=VLOOKUP(\"cherry\",$A$2:$B$6,2,FALSE)", ¬
  "=VLOOKUP(25,$B$2:$C$6,2,TRUE)", ¬
  {"a", "b", "c", "=HLOOKUP(\"b\",A{r}:C{r},1,FALSE)"}, ¬
  "=INDEX($B$2:$B$6,3)", ¬
  "=INDEX($A$2:$C$6,2,2)", ¬
  "=MATCH(\"date\",$A$2:$A$6,0)", ¬
  "=MATCH(35,$B$2:$B$6,1)", ¬
  "=MATCH(\"B*\",$A$2:$A$6,0)", ¬
  "=ROWS($A$2:$C$6)", ¬
  "=COLUMNS($A$2:$C$6)", ¬
  "=SUM(OFFSET($B$2,1,0,2,1))"}

on replaceText(theText, findText, replacement)
  set saved to AppleScript's text item delimiters
  set AppleScript's text item delimiters to findText
  set parts to text items of theText
  set AppleScript's text item delimiters to replacement
  set theText to parts as text
  set AppleScript's text item delimiters to saved
  return theText
end replaceText

on isText(v)
  return class of v is text
end isText

on withoutEquals(theFormula)
  return text 2 thru -1 of theFormula
end withoutEquals

on asText(v)
  return v as text
end asText

on run
  set repoRoot to do shell script "pwd"
  set fixturesDir to repoRoot & "/e2e/fixtures"
  do shell script "mkdir -p " & quoted form of fixturesDir
  set xlsxPosix to fixturesDir & "/spec120-numbers.xlsx"
  do shell script "rm -f " & quoted form of xlsxPosix

  set rowTotal to (count of cases) + 1
  set report to {}
  set failures to 0
  tell application "Numbers"
    activate
    set numbersVersion to version
    set theDoc to make new document
    tell theDoc
      tell sheet 1
        set name to "Calc"
        tell table 1
          try
            set name to "Oracle"
            set header column count to 0
            set header row count to 1
            set footer row count to 0
            set row count to rowTotal
            set column count to 5
            set end of report to "shape ok (" & rowTotal & " rows)"
          on error m
            set end of report to "shape FAIL " & m
            set failures to failures + 1
          end try
          set value of cell "A1" to "a"
          set value of cell "B1" to "b"
          set value of cell "C1" to "c"
          set value of cell "D1" to "formula"
          set value of cell "E1" to "Numbers"

          repeat with i from 1 to count of cases
            set r to i + 1
            set theCase to item i of cases
            if class of theCase is list then
              set inputs to items 1 thru 3 of theCase
              set formulaTemplate to item 4 of theCase
            else
              set inputs to {mv, mv, mv}
              set formulaTemplate to theCase
            end if
            set theFormula to my replaceText(formulaTemplate, "{r}", my asText(r))
            set formulaText to my withoutEquals(theFormula)
            try
              repeat with k from 1 to 3
                set inputValue to item k of inputs
                if inputValue is not mv then
                  set cellName to (item k of {"A", "B", "C"}) & r
                  if my isText(inputValue) then set format of cell cellName to text
                  set value of cell cellName to inputValue
                end if
              end repeat
              set format of cell ("D" & r) to text
              set value of cell ("D" & r) to formulaText
              set value of cell ("E" & r) to theFormula
              set problems to ""
              if formula of cell ("E" & r) is missing value then set problems to problems & " E is not a formula;"
              set shownText to my asText(value of cell ("D" & r))
              considering case
                if shownText is not formulaText then set problems to problems & " D reads " & shownText & ";"
              end considering
              if problems is not "" then
                set end of report to "row " & r & " FAIL " & formulaText & ":" & problems
                set failures to failures + 1
              end if
            on error m
              set end of report to "row " & r & " FAIL " & formulaText & ": " & m
              set failures to failures + 1
            end try
          end repeat
        end tell
      end tell
    end tell
    try
      export theDoc to (POSIX file xlsxPosix) as Microsoft Excel
      set end of report to "export ok"
    on error m
      set end of report to "export FAIL " & m
      set failures to failures + 1
    end try
    close theDoc saving no
  end tell

  -- Every case must reach the file as a formula; an error result is a
  -- formula with no <v>. Counted off the exported sheet itself.
  set sheetXml to "unzip -p " & quoted form of xlsxPosix & " xl/worksheets/sheet1.xml"
  set formulaCount to (do shell script sheetXml & " | perl -0ne 'print scalar(() = /<f[ >]/g)'") as integer
  set uncached to do shell script sheetXml & " | perl -0ne 'print join(\" \", /<c r=\"(E[0-9]+)\"[^>]*><f[^>]*>[^<]*<\\/f><\\/c>/g)'"
  if formulaCount is (count of cases) then
    set end of report to "export formulas ok (" & formulaCount & ")"
  else
    set end of report to "export formulas FAIL " & formulaCount & " of " & (count of cases) & " — a function Numbers lacks was dropped; see the E cells with no <f>"
    set failures to failures + 1
  end if
  set end of report to "no cached value (Numbers errors): " & uncached
  set end of report to "Numbers " & numbersVersion
  set end of report to ((count of cases) as text) & " formula rows, " & failures & " FAIL"
  set AppleScript's text item delimiters to linefeed
  return (report as text) & linefeed & "wrote " & xlsxPosix
end run
