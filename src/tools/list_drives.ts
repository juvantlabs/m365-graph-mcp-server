/**
 * Tool: m365-graph:list_drives
 *
 * Two modes:
 *   - no `site_id`: the user's own drives: primary OneDrive
 *     (`/me/drive`) and `/me/drives`. In Microsoft 365 business tenants
 *     `/me/drives` returns only the user's OneDrive drives, NOT the
 *     SharePoint document libraries they can access.
 *   - `site_id` (from `list_sites`): the document libraries of that
 *     SharePoint site (`/sites/{site-id}/drives`). Their drive IDs work
 *     with every `drive_id` parameter of the file tools.
 *
 * Required Graph scopes: `Files.Read` (own drives); `Sites.Read.All`
 * for site libraries (covered by `Sites.ReadWrite.All`). Read-only.
 *
 * This is the simplest read-only tool, used to validate the auth +
 * Graph plumbing end-to-end. Subsequent tools (search, list items,
 * download, etc.) follow the same pattern: typed input, typed output,
 * single Graph URL surface, no caller-supplied paths.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";
import { validateOptionalString } from "../types/validators.js";

/**
 * Graph site ID: `{hostname},{site-collection-guid},{web-guid}`, exactly
 * as returned by `list_sites`. Validated strictly because it is placed
 * in the request path.
 */
export const SITE_ID_RE =
  /^[A-Za-z0-9.-]+,[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12},[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

const definition: ToolDefinition = {
  name: "m365-graph:list_drives",
  description:
    "List drives. Without site_id: the user's own OneDrive drives (primary + others). With site_id (from list_sites): the document libraries of that SharePoint site. SharePoint libraries are NOT returned without site_id: call list_sites first. Every returned drive id works as drive_id in the file tools. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      site_id: {
        type: "string",
        description:
          "Optional SharePoint site ID from list_sites ('{hostname},{guid},{guid}'). When set, returns that site's document libraries.",
      },
    },
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

function textResponse(result: unknown): ToolResponse {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(result, null, 2),
      },
    ],
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown> = {},
): Promise<ToolResponse> => {
  const siteId = validateOptionalString(args.site_id, "site_id");

  if (siteId !== undefined) {
    if (!SITE_ID_RE.test(siteId)) {
      throw new Error(
        "'site_id' must be a Graph site ID as returned by list_sites ('{hostname},{guid},{guid}')",
      );
    }
    const res = await graph.api(`/sites/${siteId}/drives`).get();
    const drives = (Array.isArray(res?.value) ? res.value : []).map(
      (d: Record<string, unknown>) => summarizeDrive(d),
    );
    return textResponse({ site_id: siteId, count: drives.length, drives });
  }

  const primary = await graph.api("/me/drive").get();
  const all = await graph.api("/me/drives").get();
  const driveList = Array.isArray(all?.value) ? all.value : [];

  return textResponse({
    primary: summarizeDrive(primary),
    accessible: driveList.map((d: Record<string, unknown>) => summarizeDrive(d)),
    note: "OneDrive drives only. For SharePoint document libraries call list_sites, then list_drives with site_id.",
  });
};

export const listDrivesTool: Tool = { category: "read", definition, handler };
