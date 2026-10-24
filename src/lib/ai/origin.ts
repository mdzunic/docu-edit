// The destination pin (SPEC-113). A baseUrl the user typed is an input; this
// module turns it into the one origin a request may go to, once, and
// refuses everything else: a non-https: origin unless it is this machine's
// loopback, credentials smuggled into the URL, and any request URL that
// resolves outside the pinned origin. Framework-free and network-free —
// the gateway asks it before every request, the Labs form asks it to show
// the user where their text would go.

/** The three loopback hosts a plain-http: endpoint may use — exactly these,
 * as the WHATWG URL parser spells their hostnames. */
export const LOOPBACK_HOSTNAMES: readonly string[] = Object.freeze(["localhost", "127.0.0.1", "[::1]"]);

/** A resolved endpoint: the pinned origin, the base every path joins onto,
 * and whether the server is on this machine. */
export interface Endpoint {
  origin: string;
  /** The base URL normalised to end in exactly one "/". */
  base: string;
  loopback: boolean;
}

export type EndpointResolution = { ok: true; endpoint: Endpoint } | { ok: false; reason: string };

export function isLoopbackHostname(hostname: string): boolean {
  return LOOPBACK_HOSTNAMES.includes(hostname.toLowerCase());
}

/** Resolves a typed baseUrl to its endpoint, or the reason it is refused —
 * worded for the person who typed it. */
export function resolveEndpoint(baseUrl: string): EndpointResolution {
  const typed = baseUrl.trim();
  if (typed === "") {
    return { ok: false, reason: "No base URL is set. Enter the endpoint's OpenAI-compatible base URL in Labs." };
  }
  let url: URL;
  try {
    url = new URL(typed);
  } catch {
    return { ok: false, reason: `“${typed}” is not a URL. Enter one starting with https:// (or http:// for this computer).` };
  }
  if (url.protocol !== "https:" && url.protocol !== "http:") {
    return { ok: false, reason: `${url.protocol} is not an endpoint scheme. Use https:// (or http:// for this computer).` };
  }
  const loopback = isLoopbackHostname(url.hostname);
  if (url.protocol === "http:" && !loopback) {
    return {
      ok: false,
      reason:
        `Refused ${url.origin}: plain http: is only allowed for this computer (localhost, 127.0.0.1, [::1]). ` +
        "Use the endpoint's https:// address.",
    };
  }
  if (url.username !== "" || url.password !== "") {
    return { ok: false, reason: "Remove the user name and password from the base URL; a key belongs in the API key field." };
  }
  if (url.search !== "" || url.hash !== "") {
    return { ok: false, reason: "Remove the ?query or #fragment from the base URL; it should end at the API's version path, e.g. …/v1." };
  }
  const base = `${url.origin}${url.pathname.replace(/\/+$/, "")}/`;
  return { ok: true, endpoint: { origin: url.origin, base, loopback } };
}

/** The origin a baseUrl pins requests to, or null when it is refused. */
export function originOf(baseUrl: string): string | null {
  const resolved = resolveEndpoint(baseUrl);
  return resolved.ok ? resolved.endpoint.origin : null;
}

/** Resolves `path` against the endpoint's base and returns the URL only if
 * it stays on the pinned origin — an absolute or protocol-relative URL
 * pointing anywhere else returns null and is never requested. */
export function pinnedUrl(endpoint: Endpoint, path: string): string | null {
  let url: URL;
  try {
    url = new URL(path, endpoint.base);
  } catch {
    return null;
  }
  return url.origin === endpoint.origin ? url.href : null;
}
