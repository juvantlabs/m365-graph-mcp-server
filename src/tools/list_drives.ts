/**
 * Tool: m365-graph:list_drives
 *
 * Returns the drives the authenticated user has access to:
 *   - Their primary OneDrive (`/me/drive`)
 *   - Other drives accessible to them (`/me/drives`), which includes
 *     SharePoint document libraries the user is a member of.
 *
 * Required Graph scope: `Files.Read` (delegated). No write side-effects.
 *
 * This is the simplest read-only tool, used to validate the auth +
 * Graph plumbing end-to-end. Subsequent tools (search, list items,
 * download, etc.) follow the same pattern: typed input, typed output,
 * single Graph URL surface, no caller-supplied paths.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:list_drives",
  description:
    "List the drives the authenticated user has access to. Returns the user's primary OneDrive plus any shared SharePoint document libraries they're a member of. Read-only.",
  inputSchema: {
    type: "object",
    properties: {},
    required: [],
  },
};

interface DriveSummary {
  id: string;
  driveType: string;
  name: string;
  webUrl: string;
  owner: string | null;
}

export function summarizeDrive(drive: Record<string, unknown>): DriveSummary {
  const owner = drive.owner as Record<string, unknown> | undefined;
  const ownerUser = owner?.user as Record<string, unknown> | undefined;
  const ownerGroup = owner?.group as Record<string, unknown> | undefined;
  const ownerName =
    (ownerUser?.displayName as string | undefined) ??
    (ownerGroup?.displayName as string | undefined) ??
    null;
  return {
    id: String(drive.id ?? ""),
    driveType: String(drive.driveType ?? ""),
    name: String(drive.name ?? ""),
    webUrl: String(drive.webUrl ?? ""),
    owner: ownerName,
  };
}

const handler: ToolHandler = async (graph: Client): Promise<ToolResponse> => {
  const primary = await graph.api("/me/drive").get();
  const all = await graph.api("/me/drives").get();
  const driveList = Array.isArray(all?.value) ? all.value : [];

  const result = {
    primary: summarizeDrive(primary),
    accessible: driveList.map((d: Record<string, unknown>) => summarizeDrive(d)),
  };

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(result, null, 2),
      },
    ],
  };
};

export const listDrivesTool: Tool = { definition, handler };
