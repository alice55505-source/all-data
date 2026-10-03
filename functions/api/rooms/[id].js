import { jsonResponse, isValidRoomId, parseJsonColumn } from "../_lib.js";

export async function onRequestGet(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  // SELECT * so rooms in a database that hasn't gained periods_json yet
  // (it's added lazily on first save) still load.
  var row = await db.prepare("SELECT * FROM rooms WHERE id = ?").bind(id).first();
  if (!row) return jsonResponse({ error: "找不到這個房間" }, 404);

  return jsonResponse({
    id: id,
    name: row.name || "",
    groups: parseJsonColumn(row.groups_json, []),
    stats: parseJsonColumn(row.stats_json, {}),
    metrics: parseJsonColumn(row.metrics_json, { meetings: [], roles: [], extras: [] }),
    periods: parseJsonColumn(row.periods_json || "[]", [])
  });
}
