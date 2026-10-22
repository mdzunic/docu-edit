// SPEC-101 phase one (performed by SPEC-135): the comment helpers hoisted
// into core keep their SPEC-038 behaviour, the old module paths re-export
// the very same functions, and commentInitials follows the .docx writer's
// w:initials rule.
import { describe, expect, it } from "vitest";
import * as textComments from "../comments";
import { canModifyMessage, commentInitials, parseReplyParagraphText, replyParagraphText, type CommentMessage } from "./comments";

const message: CommentMessage = { messageId: "m1", author: "Ada Lovelace", text: "Where does this come from?", createdAt: "2026-10-01T09:30:00.000Z" };

describe("core/comments", () => {
  it("encodes a reply as one plain-text paragraph and parses it back", () => {
    const text = replyParagraphText(message);
    expect(text).toBe("Ada Lovelace (2026-10-01T09:30:00.000Z): Where does this come from?");
    expect(parseReplyParagraphText(text)).toEqual({ author: "Ada Lovelace", createdAt: "2026-10-01T09:30:00.000Z", text: "Where does this come from?" });
  });

  it("keeps a reply's colons and line breaks in its text", () => {
    const text = replyParagraphText({ ...message, text: "Note: see\nrow 4" });
    expect(parseReplyParagraphText(text)?.text).toBe("Note: see\nrow 4");
  });

  it("refuses paragraphs that are not the encoding", () => {
    expect(parseReplyParagraphText("Just a sentence.")).toBeNull();
    expect(parseReplyParagraphText("Bob (not a date): hi")).toBeNull();
  });

  it("lets only the message's own author modify it", () => {
    expect(canModifyMessage(message, "Ada Lovelace")).toBe(true);
    expect(canModifyMessage(message, "You")).toBe(false);
  });

  it("derives initials from each whitespace-separated word, at most three", () => {
    expect(commentInitials("Ada Lovelace")).toBe("AL");
    expect(commentInitials("  grace   brewster murray hopper ")).toBe("GBM");
    expect(commentInitials("You")).toBe("Y");
    expect(commentInitials("")).toBe("");
    expect(commentInitials("   ")).toBe("");
  });

  it("is what the text product's comments module re-exports", () => {
    expect(textComments.replyParagraphText).toBe(replyParagraphText);
    expect(textComments.parseReplyParagraphText).toBe(parseReplyParagraphText);
    expect(textComments.canModifyMessage).toBe(canModifyMessage);
  });
});
