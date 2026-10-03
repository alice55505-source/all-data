import { jsonResponse, isValidRoomId } from "../../_lib.js";

function normalizeStringArray(arr) {
  if (!Array.isArray(arr)) return [];
  var seen = {};
  var out = [];
  arr.forEach(function (v) {
    var s = typeof v === "string" ? v.trim() : "";
    if (!s || seen[s]) return;
    seen[s] = true;
    out.push(s);
  });
  return out;
}

export async function onRequestPut(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  var existing = await db.prepare("SELECT id FROM rooms WHERE id = ?").bind(id).first();
  if (!existing) return jsonResponse({ error: "找不到這個房間" }, 404);

  var body;
  try {
    body = await context.request.json();
  } catch (e) {
    return jsonResponse({ error: "請求格式錯誤" }, 400);
  }

  if (typeof body !== "object" || body === null) {
    return jsonResponse({ error: "資料格式錯誤" }, 400);
  }

  var metrics = {
    meetings: normalizeStringArray(body.meetings),
    roles: normalizeStringArray(body.roles),
    extras: normalizeStringArray(body.extras)
  };

  var now = new Date().toISOString();
  await db.prepare("UPDATE rooms SET metrics_json = ?, updated_at = ? WHERE id = ?")
    .bind(JSON.stringify(metrics), now, id).run();

  return jsonResponse({ ok: true, updatedAt: now });
}
