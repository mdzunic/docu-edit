import { describe, it, expect } from "vitest";
import {
  SAVE_STATE_ERROR,
  SAVE_STATE_RECOVERED,
  SAVE_STATE_SAVED,
  SAVE_STATE_SAVING,
  saveStateText,
} from "./saveState";

describe("saveStateText (SPEC-010)", () => {
  it("names each state", () => {
    expect(saveStateText("saved")).toBe(SAVE_STATE_SAVED);
    expect(saveStateText("saving")).toBe(SAVE_STATE_SAVING);
    expect(saveStateText("error")).toBe(SAVE_STATE_ERROR);
    expect(saveStateText("recovered")).toBe(SAVE_STATE_RECOVERED);
  });

  it("does not claim the document is saved while a recovery decision is pending", () => {
    // The bug this state exists to stop: the bar read "All changes saved"
    // beside a banner announcing recovered UNSAVED changes, and the banner was
    // the truthful one — nothing is committed while it is up.
    expect(saveStateText("recovered")).not.toBe(SAVE_STATE_SAVED);
    expect(SAVE_STATE_RECOVERED).toBe("Not saved yet");
  });

  it("states the failure with its reason", () => {
    expect(SAVE_STATE_ERROR).toContain("browser storage is full");
  });
});
