import { Client, type Context, type Middleware } from "@microsoft/microsoft-graph-client";

/**
 * A real SDK client whose terminal middleware records each request URL
 * instead of calling Graph.
 *
 * Chained `vi.fn()` mocks stop at `.query()` / `.filter()`, so they cannot
 * see what the SDK serializes. It joins query values raw, with no URL
 * encoding, so a test of what Graph receives has to go through this path.
 * Reading the result with `URL.searchParams` applies the same form-style
 * decoding that turned an unencoded `+` into a space on the Graph side.
 */
export function wireClient(body: unknown = { value: [] }): { client: Client; sent: URL[] } {
  const sent: URL[] = [];
  const middleware: Middleware = {
    execute: async (context: Context) => {
      sent.push(new URL(context.request as string));
      context.response = new Response(JSON.stringify(body), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    },
  };
  return { client: Client.initWithMiddleware({ middleware }), sent };
}
