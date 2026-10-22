-- Generates e2e/fixtures/mini-deck.key (and the embedded test image
-- e2e/fixtures/fixture-image.png) for SPEC-087's Keynote import.
--
-- Run from the repo root, on a Mac with Keynote installed:
--   osascript scripts/make-keynote-fixture.applescript
--
-- CI containers cannot run Keynote, so the fixture must be generated
-- here and committed. The deck's content is SPEC-087's fixture
-- contract — if any step below reports FAIL, fix and re-run until every
-- step is ok, then commit both files. Keynote may prompt for automation
-- permission on first run; approve it and re-run.

on run
  set repoRoot to do shell script "pwd"
  set fixturesDir to repoRoot & "/e2e/fixtures"
  do shell script "mkdir -p " & quoted form of fixturesDir
  set imgPosix to fixturesDir & "/fixture-image.png"
  set keyPosix to fixturesDir & "/mini-deck.key"
  -- 8x8 solid-colour PNG, fixed bytes (SPEC-087 asserts media equality).
  do shell script "printf %s 'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGN4VqeEFTEMLQkASgJhgTHXKdIAAAAASUVORK5CYII=' | base64 -D > " & quoted form of imgPosix
  do shell script "rm -f " & quoted form of keyPosix

  set report to {}
  tell application "Keynote"
    activate
    set theDoc to make new document with properties {width:1920, height:1080}
    tell theDoc
      tell slide 1
        try
          set object text of default title item to "Keynote fixture deck"
          set end of report to "s1-title ok"
        on error m
          set end of report to "s1-title FAIL " & m
        end try
        try
          set object text of default body item to "Imported by docu-edit"
          set end of report to "s1-body ok"
        on error m
          set end of report to "s1-body FAIL " & m
        end try
      end tell
      try
        set slideTwo to make new slide at end of slides with properties {base slide:master slide "Title & Bullets" of theDoc}
        set end of report to "s2-master ok"
      on error m
        set slideTwo to make new slide at end of slides
        set end of report to "s2-master fallback (" & m & ") — verify slide 2 still has title+body items"
      end try
      tell slideTwo
        try
          set object text of default title item to "Second slide"
          set end of report to "s2-title ok"
        on error m
          set end of report to "s2-title FAIL " & m
        end try
        try
          set object text of default body item to "First bullet point" & return & "Second bullet point"
          set end of report to "s2-body ok"
        on error m
          set end of report to "s2-body FAIL " & m
        end try
        try
          make new image with properties {file:(POSIX file imgPosix), position:{1300, 620}, width:300}
          set end of report to "s2-image ok"
        on error m
          set end of report to "s2-image FAIL " & m
        end try
        try
          set presenter notes to "A speaker note for the importer"
          set end of report to "s2-notes ok"
        on error m
          set end of report to "s2-notes FAIL " & m
        end try
      end tell
    end tell
    save theDoc in (POSIX file keyPosix)
    close theDoc saving no
  end tell

  set AppleScript's text item delimiters to linefeed
  return (report as text) & linefeed & "wrote " & keyPosix
end run
