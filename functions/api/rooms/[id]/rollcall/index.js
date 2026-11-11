import { jsonResponse, isValidRoomId } from "../../../_lib.js";

// DELETE /api/rooms/:id/rollcall - clears every congregation's 點名 data.
export async function onRequestDelete(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  var existing = await db.prepare("SELECT id FROM rooms WHERE id = ?").bind(id).first();
  if (!existing) return jsonResponse({ error: "找不到這個房間" }, 404);

  var now = new Date().toISOString();
  await db.prepare("UPDATE rooms SET rollcall_json = '{}', updated_at = ? WHERE id = ?")
    .bind(now, id).run();

  return jsonResponse({ ok: true, updatedAt: now });
}
