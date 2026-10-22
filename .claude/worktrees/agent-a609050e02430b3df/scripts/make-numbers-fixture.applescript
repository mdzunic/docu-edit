-- Generates e2e/fixtures/spec116-numbers.xlsx for SPEC-116: a workbook
-- exported by Apple Numbers ("Export To ▸ Excel"), the one fixture no
-- builder can emulate because its shape is whatever Numbers writes — and
-- e2e/fixtures/spec116-numbers.numbers, the Numbers document it was
-- exported from, which SPEC-116's Numbers detection is tested against.
--
-- Run from the repo root, on a Mac with Numbers installed:
--   open -a Numbers    (first, or `activate` below fails with -600)
--   osascript scripts/make-numbers-fixture.applescript
--
-- CI containers cannot run Numbers, so the fixture is generated here and
-- committed. Its content is SPEC-116's fixture contract — if any step below
-- reports FAIL, fix and re-run until every step is ok, then commit both
-- files and update e2e/fixtures/spec116-numbers.md (the Numbers version and
-- the Index/ entries this script prints — SPEC-116 detects Numbers by them).
-- Numbers may prompt for automation permission on first run; approve it
-- and re-run.

on run
  set repoRoot to do shell script "pwd"
  set fixturesDir to repoRoot & "/e2e/fixtures"
  do shell script "mkdir -p " & quoted form of fixturesDir
  set xlsxPosix to fixturesDir & "/spec116-numbers.xlsx"
  set numbersPosix to fixturesDir & "/spec116-numbers.numbers"
  do shell script "rm -rf " & quoted form of xlsxPosix & " " & quoted form of numbersPosix

  -- A fixed date built field by field (date literals parse per locale).
  set theDate to current date
  set day of theDate to 1
  set year of theDate to 2024
  set month of theDate to March
  set day of theDate to 15
  set time of theDate to 0

  set report to {}
  tell application "Numbers"
    activate
    set numbersVersion to version
    set theDoc to make new document
    tell theDoc
      -- Sheet 1: one table holding text, numbers, a date and a formula.
      tell sheet 1
        try
          set name to "Summary"
          set end of report to "s1-name ok"
        on error m
          set end of report to "s1-name FAIL " & m
        end try
        tell table 1
          try
            set name to "Figures"
            set row count to 5
            set column count to 3
            set end of report to "t1-shape ok"
          on error m
            set end of report to "t1-shape FAIL " & m
          end try
          try
            set value of cell "A1" to "Item"
            set value of cell "B1" to "Amount"
            set value of cell "C1" to "Due"
            set value of cell "A2" to "Revenue"
            set value of cell "B2" to 1250.5
            set value of cell "C2" to theDate
            set value of cell "A3" to "Costs"
            set value of cell "B3" to 400
            set value of cell "A4" to "Margin"
            set value of cell "B4" to "=B2-B3"
            set value of cell "A5" to "Note"
            set value of cell "B5" to "Plain text, with a comma"
            set end of report to "t1-values ok"
          on error m
            set end of report to "t1-values FAIL " & m
          end try
        end tell
      end tell
      -- Sheet 2: a second, visible sheet with its own table.
      try
        set sheetTwo to make new sheet with properties {name:"Second"}
        set end of report to "s2-make ok"
      on error m
        set end of report to "s2-make FAIL " & m
      end try
      tell sheetTwo
        tell table 1
          try
            set name to "Notes"
            set row count to 3
            set column count to 2
            set value of cell "A1" to "Key"
            set value of cell "B1" to "Value"
            set value of cell "A2" to "count"
            set value of cell "B2" to 42
            set value of cell "A3" to "flag"
            set value of cell "B3" to true
            set end of report to "t2-values ok"
          on error m
            set end of report to "t2-values FAIL " & m
          end try
        end tell
      end tell
    end tell
    try
      export theDoc to (POSIX file xlsxPosix) as Microsoft Excel
      set end of report to "export ok"
    on error m
      set end of report to "export FAIL " & m
    end try
    save theDoc in (POSIX file numbersPosix)
    close theDoc saving no
  end tell

  -- The markers SPEC-116 detects Numbers by, read off the real source.
  -- (A package saved as a directory instead of one ZIP file is a FAIL:
  -- the library only ever receives single files.)
  set entries to do shell script "f=" & quoted form of numbersPosix & "; if [ -d \"$f\" ]; then echo 'source FAIL saved as a package directory'; else unzip -Z1 \"$f\" | grep '^Index/' | sed -E 's/-[0-9]+(-[0-9]+)?[.]iwa$/-<id>.iwa/' | sort | uniq -c; fi" without altering line endings
  set end of report to "Numbers " & numbersVersion
  set end of report to "source Index entries:" & linefeed & entries

  set AppleScript's text item delimiters to linefeed
  return (report as text) & "wrote " & xlsxPosix & linefeed & "wrote " & numbersPosix
end run
