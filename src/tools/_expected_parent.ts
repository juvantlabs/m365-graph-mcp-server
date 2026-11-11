/**
 * Guard argument shared by delete_file and move_file: `expected_parent_id`.
 *
 * When the caller passes it, the tool reads the item first and refuses to
 * act unless the item's current parent folder has exactly that id. A caller
 * (or a policy hook in front of it) can then bound a destructive operation to
 * one folder: the argument says where the item is, and the server checks that
 * it is true. Omitted, the tools behave as before.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

export async function assertExpectedParent(
  graph: Client,
  itemApiPath: string,
  itemId: string,
  expectedParentId: string | undefined,
  action: string,
): Promise<Record<string, unknown> | undefined> {
  if (expectedParentId === undefined) return undefined;
  const item = (await graph.api(itemApiPath).get()) as Record<string, unknown>;
  const parent = item.parentReference as Record<string, unknown> | undefined;
  const actual = parent?.id !== undefined ? String(parent.id) : null;
  if (actual !== expectedParentId) {
    throw new Error(
      `Refusing to ${action}: item ${itemId} is in folder ${actual ?? "(unknown)"}, ` +
        `not in expected_parent_id ${expectedParentId}. Nothing was changed.`,
    );
  }
  return item;
}
