// SPEC-113 — the destination pin: a typed baseUrl resolved to one origin, an
// http: origin refused unless it is loopback, and request URLs held to it.
import { describe, expect, it } from "vitest";
import { isLoopbackHostname, originOf, pinnedUrl, resolveEndpoint, type Endpoint } from "./origin";

function endpoint(baseUrl: string): Endpoint {
  const resolved = resolveEndpoint(baseUrl);
  if (!resolved.ok) {
    throw new Error(resolved.reason);
  }
  return resolved.endpoint;
}

describe("resolveEndpoint", () => {
  it("pins an https endpoint to its origin and normalises the base", () => {
    expect(endpoint("https://api.example.test/v1")).toEqual({
      origin: "https://api.example.test",
      base: "https://api.example.test/v1/",
      loopback: false,
    });
    expect(endpoint(" https://api.example.test/v1// ").base).toBe("https://api.example.test/v1/");
  });

  it("accepts plain http only for localhost, 127.0.0.1 and [::1]", () => {
    expect(endpoint("http://localhost:11434/v1")).toMatchObject({ origin: "http://localhost:11434", loopback: true });
    expect(endpoint("http://127.0.0.1:1234/v1").loopback).toBe(true);
    expect(endpoint("http://[::1]:8080/v1").loopback).toBe(true);
    for (const refused of ["http://api.example.test/v1", "http://192.168.1.4:11434/v1", "http://127.0.0.2/v1"]) {
      const result = resolveEndpoint(refused);
      expect(result.ok).toBe(false);
      expect(result.ok ? "" : result.reason).toContain("plain http: is only allowed for this computer");
    }
  });

  it("flags an https loopback endpoint as on this machine too", () => {
    expect(endpoint("https://localhost:8443/v1").loopback).toBe(true);
  });

  it("refuses an empty value, a non-URL, other schemes, credentials and a query or fragment", () => {
    for (const refused of [
      "",
      "localhost:11434",
      "ftp://files.example.test/v1",
      "javascript:alert(1)",
      "https://user:pass@api.example.test/v1",
      "https://api.example.test/v1?key=1",
      "https://api.example.test/v1#x",
    ]) {
      expect(resolveEndpoint(refused).ok, refused).toBe(false);
    }
  });

  it("originOf is the pinned origin, or null when refused", () => {
    expect(originOf("https://api.example.test/v1")).toBe("https://api.example.test");
    expect(originOf("http://api.example.test/v1")).toBeNull();
    expect(isLoopbackHostname("LOCALHOST")).toBe(true);
  });
});

describe("pinnedUrl", () => {
  const pinned = endpoint("https://api.example.test/v1");

  it("joins a path onto the base", () => {
    expect(pinnedUrl(pinned, "chat/completions")).toBe("https://api.example.test/v1/chat/completions");
    expect(pinnedUrl(pinned, "models")).toBe("https://api.example.test/v1/models");
  });

  it("refuses any URL that resolves off the pinned origin", () => {
    expect(pinnedUrl(pinned, "https://evil.example.test/v1/chat/completions")).toBeNull();
    expect(pinnedUrl(pinned, "//evil.example.test/x")).toBeNull();
    expect(pinnedUrl(pinned, "http://api.example.test/v1/models")).toBeNull();
    expect(pinnedUrl(pinned, "https://api.example.test:444/v1/models")).toBeNull();
  });
});
