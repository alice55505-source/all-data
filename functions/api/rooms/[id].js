import { jsonResponse, isValidRoomId, parseJsonColumn } from "../_lib.js";

export async function onRequestGet(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  var row;
  try {
    row = await db.prepare("SELECT name, groups_json, stats_json, rollcall_json FROM rooms WHERE id = ?").bind(id).first();
  } catch (e) {
    // Database without migrations/0003 applied yet: keep the rest of the
    // app working, the 點名 tab just shows as empty until it is migrated.
    row = await db.prepare("SELECT name, groups_json, stats_json FROM rooms WHERE id = ?").bind(id).first();
  }
  if (!row) return jsonResponse({ error: "找不到這個房間" }, 404);

  return jsonResponse({
    id: id,
    name: row.name || "",
    groups: parseJsonColumn(row.groups_json, []),
    stats: parseJsonColumn(row.stats_json, {}),
    rollcall: parseJsonColumn(row.rollcall_json, {}) || {}
  });
}
