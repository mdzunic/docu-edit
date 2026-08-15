import { afterEach, describe, expect, it } from "vitest";
import {
  COMMENT_AUTHOR_STORAGE_KEY,
  DEFAULT_COMMENT_AUTHOR,
  loadCommentAuthor,
  saveCommentAuthor,
} from "./commentIdentity";

afterEach(() => {
  localStorage.clear();
});

describe("comment author identity persistence (AC: settable identity, defaulting to 'You')", () => {
  it("defaults to 'You' for a new user (nothing stored)", () => {
    expect(DEFAULT_COMMENT_AUTHOR).toBe("You");
    expect(loadCommentAuthor()).toBe("You");
  });

  it("round-trips a stored name", () => {
    saveCommentAuthor("Ada Lovelace");
    expect(loadCommentAuthor()).toBe("Ada Lovelace");
  });

  it("trims whitespace on save", () => {
    saveCommentAuthor("  Bea  ");
    expect(loadCommentAuthor()).toBe("Bea");
  });

  it("falls back to the default for a blank name", () => {
    saveCommentAuthor("   ");
    expect(loadCommentAuthor()).toBe(DEFAULT_COMMENT_AUTHOR);
  });

  it("falls back to the default on corrupt storage", () => {
    localStorage.setItem(COMMENT_AUTHOR_STORAGE_KEY, "not json");
    expect(loadCommentAuthor()).toBe(DEFAULT_COMMENT_AUTHOR);
    localStorage.setItem(COMMENT_AUTHOR_STORAGE_KEY, JSON.stringify({ name: 42 }));
    expect(loadCommentAuthor()).toBe(DEFAULT_COMMENT_AUTHOR);
  });
});
