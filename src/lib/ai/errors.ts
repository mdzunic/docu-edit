// The model gateway's failures (SPEC-113), typed and specific — never a
// spinner that stops. Each one's message says what happened and what to
// change, in words for the person using the app, and none ever carries the
// API key: the one field that echoes a server (ModelHttpError's body
// excerpt) is redacted by the gateway before it gets here.

/** The common base, so a caller can tell a gateway failure from its own. */
export class ModelGatewayError extends Error {
  override name = "ModelGatewayError";
}

/** The provider is off, or the config cannot be used (no base URL, no
 * model, a refused origin). Nothing was sent. */
export class ModelNotConfiguredError extends ModelGatewayError {
  override name = "ModelNotConfiguredError";
}

/** The user declined (or could not be asked) in the consent dialog. Nothing
 * was sent. */
export class ModelConsentDeniedError extends ModelGatewayError {
  override name = "ModelConsentDeniedError";

  constructor(readonly origin: string) {
    super(`You did not allow this request to ${origin}, so nothing was sent. Run it again to be asked again.`);
  }
}

/** No HTTP answer: nothing listening, DNS, CORS, a dropped connection — or a
 * redirect, which the gateway refuses. */
export class ModelUnreachableError extends ModelGatewayError {
  override name = "ModelUnreachableError";

  constructor(
    readonly origin: string,
    readonly loopback: boolean,
    readonly reason: "network" | "redirect" = "network",
  ) {
    super(
      reason === "redirect"
        ? `${origin} answered with a redirect to another address. Redirects are refused so a request — or a key — ` +
            "can never be moved to another host. Set the base URL to the address the endpoint really serves from."
        : loopback
          ? `Nothing is listening on ${origin}; is your local model running? If it is, check the port, and that ` +
            "the server accepts requests from this page (its CORS / allowed-origins setting)."
          : `Could not reach ${origin}. Check the base URL and your connection; the server may also be refusing ` +
            "requests from this page (CORS).",
    );
  }
}

/** The status-specific advice a ModelHttpError ends with. */
function httpAdvice(status: number): string {
  if (status === 401 || status === 403) {
    return "Check the API key in Labs, and that it is allowed to use this model.";
  }
  if (status === 404) {
    return "Check the base URL (it should end at the version path, e.g. …/v1) and the model name.";
  }
  if (status === 429) {
    return "The endpoint is rate limiting; wait a moment and try again.";
  }
  if (status >= 500) {
    return "The server failed; try again, or check the server's own logs.";
  }
  return "Check the base URL, the model name and the request settings in Labs.";
}

/** The endpoint answered with a non-2xx status. Carries the status and the
 * first 500 characters of the body. */
export class ModelHttpError extends ModelGatewayError {
  override name = "ModelHttpError";
  static readonly BODY_EXCERPT_LENGTH = 500;
  readonly body: string;

  constructor(
    readonly status: number,
    body: string,
  ) {
    const excerpt = body.slice(0, ModelHttpError.BODY_EXCERPT_LENGTH);
    super(`The endpoint answered HTTP ${status}${excerpt ? `: ${excerpt}` : ""}. ${httpAdvice(status)}`);
    this.body = excerpt;
  }
}

/** No complete answer within the configured timeout. */
export class ModelTimeoutError extends ModelGatewayError {
  override name = "ModelTimeoutError";

  constructor(readonly timeoutMs: number) {
    super(
      `No complete answer within ${timeoutMs / 1000} s, so the request was stopped. ` +
        "Raise the timeout in Labs, use a smaller model, or try again.",
    );
  }
}

/** The endpoint answered, but not with an OpenAI-compatible response the
 * gateway can read text from. */
export class ModelOutputError extends ModelGatewayError {
  override name = "ModelOutputError";

  constructor(readonly detail: string) {
    super(
      `The endpoint answered, but not with an OpenAI-compatible response (${detail}). ` +
        "Check that the base URL points at an OpenAI-compatible API.",
    );
  }
}

/** Whether `error` is a cancellation through the caller's AbortSignal. By
 * name, not class: a DOMException from another realm (a worker, a test
 * environment) is still a cancellation. */
export function isAbortError(error: unknown): boolean {
  return error !== null && typeof error === "object" && (error as { name?: unknown }).name === "AbortError";
}

/** The message to show for any failure out of complete/stream/testConnection:
 * the gateway's own wording, a plain line for a cancellation, and a generic
 * line for anything unexpected. */
export function describeModelError(error: unknown): string {
  if (error instanceof ModelGatewayError) {
    return error.message;
  }
  if (isAbortError(error)) {
    return "Cancelled. Nothing further was sent.";
  }
  return "The request failed unexpectedly. Nothing was kept; try again.";
}
