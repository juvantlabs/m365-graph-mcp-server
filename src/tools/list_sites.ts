/**
 * Tool: m365-graph:list_sites
 *
 * Finds the SharePoint sites the authenticated user can access, so the
 * agent can reach their document libraries:
 *
 *   list_sites  →  site id  →  list_drives({ site_id })  →  drive id
 *   →  list_items / search_files / download_file / upload_file /
 *      create_folder ({ drive_id })
 *
 * Wraps Graph `GET /sites?search={query}`. Microsoft Graph's
 * `/me/drives` does not return SharePoint document libraries in
 * Microsoft 365 business tenants, so site discovery is the only way
 * to enumerate them. Results are security-trimmed by SharePoint:
 * only sites the signed-in user can open are returned.
 *
 * Graph's site search is a keyword search: it can miss names such as
 * hyphenated site URLs ("team-finance"). When a specific query returns
 * nothing, the tool falls back to listing all accessible sites ('*')
 * and filtering them client-side by a case-insensitive substring match
 * on name, display name and URL (`match: "name_filter"` in the result).
 *
 * Required Graph scope: `Sites.Read.All` (covered by the
 * `Sites.ReadWrite.All` scope this server already requests). Read-only.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";
import { validateOptionalInteger, validateOptionalString } from "../types/validators.js";

const MAX_QUERY_LENGTH = 200;
/** Fallback scan: pages of 200 sites, at most 5 pages (1000 sites). */
const FALLBACK_PAGE_SIZE = 200;
const FALLBACK_MAX_PAGES = 5;
const SELECT = "id,name,displayName,webUrl";

const definition: ToolDefinition = {
  name: "m365-graph:list_sites",
  description:
    "Find SharePoint sites the user can access (search by keyword; default '*' lists all). If the keyword search finds nothing (Graph search can miss e.g. hyphenated names), the tool falls back to matching the query as a substring of each site's name, display name or URL. Returns each site's id, name and URL. Pass a site id to list_drives to get that site's document libraries (drive IDs), then use the drive ID with list_items, search_files, download_file, upload_file or create_folder. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description:
          "Optional keyword matched against site names and descriptions; if nothing matches, it is matched as a substring of name, display name or URL. Default '*' (all sites the user can access).",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        default: 50,
        description: "Maximum number of sites to return (1–100, default 50).",
      },
    },
    required: [],
  },
};

export interface SiteSummary {
  id: string;
  name: string;
  display_name: string;
  web_url: string;
}

export function summarizeSite(site: Record<string, unknown>): SiteSummary {
  return {
    id: String(site.id ?? ""),
    name: String(site.name ?? ""),
    display_name: String(site.displayName ?? ""),
    web_url: String(site.webUrl ?? ""),
  };
}

/** Case-insensitive substring match on name, display name and URL. */
export function matchesQuery(site: SiteSummary, query: string): boolean {
  const q = query.toLowerCase();
  return [site.name, site.display_name, site.web_url].some((v) => v.toLowerCase().includes(q));
}

/**
 * All sites the user can access ('*'), following `@odata.nextLink` for
 * at most FALLBACK_MAX_PAGES pages. Next links are only followed when
 * they point at Microsoft Graph over https.
 */
async function listAllSites(graph: Client): Promise<SiteSummary[]> {
  const out: SiteSummary[] = [];
  let page = await graph
    .api("/sites")
    .query({ search: "*" })
    .select(SELECT)
    .top(FALLBACK_PAGE_SIZE)
    .get();
  for (let i = 0; i < FALLBACK_MAX_PAGES; i++) {
    const value: Array<Record<string, unknown>> = Array.isArray(page?.value) ? page.value : [];
    out.push(...value.map(summarizeSite));
    const next = page?.["@odata.nextLink"];
    if (typeof next !== "string" || !next.startsWith("https://graph.microsoft.com/")) break;
    if (i === FALLBACK_MAX_PAGES - 1) break;
    page = await graph.api(next).get();
  }
  return out;
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const query = (validateOptionalString(args.query, "query") ?? "*").trim();
  if (query.length === 0 || query.length > MAX_QUERY_LENGTH) {
    throw new Error(`'query' must be 1–${MAX_QUERY_LENGTH} characters`);
  }
  const limit = validateOptionalInteger(args.limit, "limit", { min: 1, max: 100, default: 50 });

  // `.query()` URL-encodes the value; the search term never becomes
  // part of the path.
  const response = await graph
    .api("/sites")
    .query({ search: query })
    .select(SELECT)
    .top(limit)
    .get();
  const raw: Array<Record<string, unknown>> = Array.isArray(response?.value) ? response.value : [];
  let sites = raw.slice(0, limit).map(summarizeSite);
  let match: "search" | "name_filter" = "search";

  if (sites.length === 0 && query !== "*") {
    sites = (await listAllSites(graph)).filter((site) => matchesQuery(site, query)).slice(0, limit);
    match = "name_filter";
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ query, match, count: sites.length, sites }, null, 2),
      },
    ],
  };
};

export const listSitesTool: Tool = { category: "read", definition, handler };
