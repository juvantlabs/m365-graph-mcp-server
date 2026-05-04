/**
 * Tool: m365-graph:cancel_event
 *
 * Cancel a meeting the user organizes. Sends a cancellation notice
 * to attendees. Two-phase spec/approval pattern (see delete_file
 * for the rationale): preview + confirmation_token, then execute.
 *
 * Note: this is for events the *user* organizes. For events where
 * the user is just an attendee, the right action is to *decline*
 * via Outlook (different Graph endpoint, not yet wrapped here).
 * Calling cancel_event on a non-organized event will produce a
 * clear Graph error.
 *
 * Required Graph scope: `Calendars.ReadWrite` (delegated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  consumeConfirmation,
  issueConfirmation,
} from "../auth/confirmation_tokens.js";
import { summarizeEvent } from "./list_events.js";
import {
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const TOOL_NAME = "m365-graph:cancel_event";

const definition: ToolDefinition = {
  name: TOOL_NAME,
  description:
    "Cancel a meeting the user organizes (sends cancellation to attendees). Two-phase: first call returns event preview + confirmation_token; second call executes. Use this only for events the user is the organizer of — for declining as an attendee, decline in Outlook (not yet wrapped as a tool here).",
  inputSchema: {
    type: "object",
    properties: {
      event_id: { type: "string", description: "Event to cancel." },
      comment: {
        type: "string",
        description: "Optional comment included in the cancellation notification sent to attendees.",
      },
      confirmation_token: {
        type: "string",
        description:
          "Omit on the first call (preview mode). Include the token returned by the preview call to execute the cancellation.",
      },
    },
    required: ["event_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const eventId = validateRequiredString(args.event_id, "event_id");
  const comment = validateOptionalString(args.comment, "comment");
  const confirmationToken = validateOptionalString(args.confirmation_token, "confirmation_token");

  // Spec for confirmation includes the comment too — if the agent
  // changes the comment between preview and execute, that's a
  // different action and the token must be re-issued.
  const spec: Record<string, unknown> = { event_id: eventId };
  if (comment !== undefined) spec.comment = comment;

  const apiPath = `/me/events/${encodeURIComponent(eventId)}`;

  // ---------- Phase 1: preview ----------
  if (!confirmationToken) {
    const event = (await graph.api(apiPath).get()) as Record<string, unknown>;
    const summary = summarizeEvent(event);
    const issued = issueConfirmation(TOOL_NAME, spec);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(
            {
              preview: {
                event: summary,
                cancellation_comment: comment ?? null,
                ...issued,
                instructions:
                  `Re-call ${TOOL_NAME} with the SAME args (event_id${comment ? ", comment" : ""}) ` +
                  `plus this confirmation_token to send the cancellation notice.`,
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  }

  // ---------- Phase 2: execute ----------
  const verdict = consumeConfirmation(confirmationToken, TOOL_NAME, spec);
  if (!verdict.ok) {
    throw new Error(
      `Refusing to cancel: confirmation_token ${verdict.error}. ` +
        `Re-run without confirmation_token to get a fresh preview + token.`,
    );
  }

  const cancelBody: Record<string, unknown> = {};
  if (comment !== undefined) cancelBody.Comment = comment;

  await graph.api(`${apiPath}/cancel`).post(cancelBody);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            cancelled: { event_id: eventId },
            note: "POST /events/{id}/cancel executed. Cancellation notice sent to attendees.",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const cancelEventTool: Tool = { category: "write_irreversible", definition, handler };
