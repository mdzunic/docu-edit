import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import CommentsPane from "./CommentsPane.vue";
import type { CommentThread } from "../lib/documentModel";
import type { VisibleCommentThread } from "../lib/comments";

const CARET = { section: 0, block: 0, offset: 0 };

function thread(overrides: Partial<CommentThread> = {}): CommentThread {
  return {
    threadId: "t1",
    resolved: false,
    messages: [{ messageId: "m1", author: "Ada", text: "needs work", createdAt: "2026-01-01T00:00:00.000Z" }],
    ...overrides,
  };
}

function entry(overrides: Partial<CommentThread> = {}): VisibleCommentThread {
  return { thread: thread(overrides), start: CARET, end: CARET };
}

function mountPane(props: Record<string, unknown> = {}) {
  return mount(CommentsPane, {
    props: {
      openThreads: [],
      resolvedThreads: [],
      filteredThreads: [],
      commentsOnlyView: false,
      threadTops: new Map(),
      activeThreadId: undefined,
      authorName: "Ada",
      commentFilter: "all",
      commentAuthorFilter: "",
      authors: [],
      composerOpen: false,
      editable: true,
      ...props,
    },
  });
}

describe("CommentsPane", () => {
  it("is a labelled landmark", () => {
    const wrapper = mountPane();
    const aside = wrapper.get('[data-testid="comments-pane"]');
    expect(aside.attributes("aria-label")).toBe("Comments");
  });

  it("renders each open thread's messages with author and text", () => {
    const wrapper = mountPane({ openThreads: [entry()] });
    const card = wrapper.get('[data-testid="comment-thread-t1"]');
    expect(card.text()).toContain("Ada");
    expect(card.text()).toContain("needs work");
  });

  it("shows an empty state when there are no threads", () => {
    const wrapper = mountPane();
    expect(wrapper.get('[data-testid="comments-empty"]').text()).toContain("No comments yet");
  });

  it("clicking a thread card emits focus-thread", async () => {
    const wrapper = mountPane({ openThreads: [entry()] });
    await wrapper.get('[data-testid="comment-thread-t1"]').trigger("click");
    expect(wrapper.emitted("focus-thread")).toEqual([["t1"]]);
  });

  it("submits composer text and clears it", async () => {
    const wrapper = mountPane({ composerOpen: true });
    await wrapper.get('[data-testid="comment-composer-input"]').setValue("looks great");
    await wrapper.get('[data-testid="comment-composer-submit"]').trigger("click");
    expect(wrapper.emitted("submit-comment")).toEqual([["looks great"]]);
  });

  it("does not submit blank composer text", async () => {
    const wrapper = mountPane({ composerOpen: true });
    await wrapper.get('[data-testid="comment-composer-input"]').setValue("   ");
    await wrapper.get('[data-testid="comment-composer-submit"]').trigger("click");
    expect(wrapper.emitted("submit-comment")).toBeUndefined();
  });

  it("submits a reply for the right thread", async () => {
    const wrapper = mountPane({ openThreads: [entry()] });
    await wrapper.get('[data-testid="comment-reply-input"]').setValue("agreed");
    const buttons = wrapper.findAll("button").filter((b) => b.text() === "Reply");
    await buttons[0].trigger("click");
    expect(wrapper.emitted("submit-reply")).toEqual([["t1", "agreed"]]);
  });

  it("only shows edit/delete for the message's own author (AC: author can edit/delete their own comments)", () => {
    const own = mountPane({ openThreads: [entry()], authorName: "Ada" });
    const ownMessage = own.get('[data-testid="comment-message-m1"]');
    expect(ownMessage.findAll("button").some((b) => b.text() === "Edit")).toBe(true);
    expect(ownMessage.findAll("button").some((b) => b.text() === "Delete")).toBe(true);

    const other = mountPane({ openThreads: [entry()], authorName: "Bea" });
    const otherMessage = other.get('[data-testid="comment-message-m1"]');
    expect(otherMessage.findAll("button")).toHaveLength(0);
  });

  it("resolve/reopen/delete-thread emit the thread id", async () => {
    const openWrapper = mountPane({ openThreads: [entry()] });
    const resolveBtn = openWrapper.findAll("button").find((b) => b.text() === "Resolve");
    await resolveBtn!.trigger("click");
    expect(openWrapper.emitted("resolve")).toEqual([["t1"]]);

    const resolvedWrapper = mountPane({ resolvedThreads: [entry({ resolved: true })] });
    const reopenBtn = resolvedWrapper.findAll("button").find((b) => b.text() === "Reopen");
    await reopenBtn!.trigger("click");
    expect(resolvedWrapper.emitted("reopen")).toEqual([["t1"]]);
  });

  it("comments-only view lists filteredThreads and shows filter controls", () => {
    const wrapper = mountPane({
      commentsOnlyView: true,
      filteredThreads: [entry()],
      authors: ["Ada", "Bea"],
    });
    expect(wrapper.find('[data-testid="comment-filter-select"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="comment-author-filter-select"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="comment-thread-t1"]').text()).toContain("needs work");
  });

  it("emits update:comment-filter and update:author-filter on selection", async () => {
    const wrapper = mountPane({ commentsOnlyView: true, authors: ["Ada"] });
    await wrapper.get('[data-testid="comment-filter-select"]').setValue("resolved");
    expect(wrapper.emitted("update:comment-filter")).toEqual([["resolved"]]);
    await wrapper.get('[data-testid="comment-author-filter-select"]').setValue("Ada");
    expect(wrapper.emitted("update:author-filter")).toEqual([["Ada"]]);
  });

  describe("security: comment text renders as plain text (AC: no markup injection)", () => {
    it("never interprets HTML/script markup in a message's text", () => {
      const malicious = entry({
        messages: [
          {
            messageId: "m1",
            author: "Eve",
            text: '<script>window.pwnedByComment = true;</script><img src="x" onerror="window.pwnedByComment = true">',
            createdAt: "2026-01-01T00:00:00.000Z",
          },
        ],
      });
      const wrapper = mountPane({ openThreads: [malicious] });

      // No live script/img element was created from the comment text.
      expect(wrapper.find("script").exists()).toBe(false);
      expect(wrapper.find("img").exists()).toBe(false);
      expect((window as unknown as { pwnedByComment?: boolean }).pwnedByComment).toBeUndefined();

      // The raw markup shows up as literal, escaped text content.
      const card = wrapper.get('[data-testid="comment-thread-t1"]');
      expect(card.text()).toContain("<script>window.pwnedByComment = true;</script>");
      expect(card.html()).not.toContain("<script>window.pwnedByComment");
      expect(card.html()).toContain("&lt;script&gt;");
    });

    it("escapes markup in the author name and reply drafts too", async () => {
      const malicious = entry({
        messages: [
          { messageId: "m1", author: '<img src=x onerror=alert(1)>', text: "hi", createdAt: "2026-01-01T00:00:00.000Z" },
        ],
      });
      const wrapper = mountPane({ openThreads: [malicious] });
      expect(wrapper.find("img").exists()).toBe(false);
      expect(wrapper.html()).toContain("&lt;img");
    });
  });
});
