/**
 * Tool: m365-graph:decline_event
 *
 * Decline an event the user is invited to (as an attendee — not as
 * organizer). Sends a decline notification to the organizer.
 *
 * For events the user *organizes*, use `cancel_event` instead.
 *
 * Two-phase spec/approval pattern (same as delete_file +
 * cancel_event): 1st call returns event preview + confirmation_token,
 * 2nd call (with same args + token) executes POST
 * `/me/events/{id}/decline`.
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

const TOOL_NAME = "m365-graph:decline_event";

const definition: ToolDefinition = {
  name: TOOL_NAME,
  description:
    "Decline an event the user is invited to (as attendee). Sends a decline RSVP to the organizer. Two-phase: first call returns event preview + confirmation_token, second call executes. For events the user organizes, use cancel_event instead.",
  inputSchema: {
    type: "object",
    properties: {
      event_id: { type: "string", description: "Event to decline." },
      comment: {
        type: "string",
        description: "Optional comment included in the decline notification sent to the organizer.",
      },
      send_response: {
        type: "boolean",
        description:
          "Whether to send a decline notice to the organizer (default true). If false, the decline is silent — organizer does not see a notification.",
      },
      confirmation_token: {
        type: "string",
        description:
          "Omit on the first call (preview mode). Include the token from the preview to execute the decline.",
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
  // sendResponse default true (Graph default + matches user expectation
  // of "I declined, the organizer should know"). To suppress: pass
  // explicit false.
  const sendResponse = args.send_response === false ? false : true;

  // Spec includes everything that affects what the action does.
  const spec: Record<string, unknown> = { event_id: eventId, send_response: sendResponse };
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
                decline_comment: comment ?? null,
                send_response: sendResponse,
                ...issued,
                instructions:
                  `Re-call ${TOOL_NAME} with the SAME args (event_id, send_response${comment ? ", comment" : ""}) ` +
                  `plus this confirmation_token to send the decline.`,
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
      `Refusing to decline: confirmation_token ${verdict.error}. ` +
        `Re-run without confirmation_token to get a fresh preview + token.`,
    );
  }

  const declineBody: Record<string, unknown> = { sendResponse };
  if (comment !== undefined) declineBody.comment = comment;

  await graph.api(`${apiPath}/decline`).post(declineBody);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            declined: { event_id: eventId, send_response: sendResponse },
            note:
              "POST /events/{id}/decline executed. " +
              (sendResponse
                ? "Decline notice sent to the organizer."
                : "Decline was silent — organizer not notified."),
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const declineEventTool: Tool = { category: "write_irreversible", definition, handler };
