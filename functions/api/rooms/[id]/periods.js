import { jsonResponse, isValidRoomId } from "../../_lib.js";

var WEEK_PATTERN = /^\d+W\d+$/;

function normalizePeriods(raw) {
  if (!Array.isArray(raw)) return null;
  var out = [];
  raw.slice(0, 60).forEach(function (p) {
    if (!p || typeof p.name !== "string" || !Array.isArray(p.weeks)) return;
    var name = p.name.trim().slice(0, 30);
    var weeks = p.weeks.filter(function (w) { return typeof w === "string" && WEEK_PATTERN.test(w); }).slice(0, 120);
    if (name && weeks.length) out.push({ name: name, weeks: weeks });
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

  var periods = normalizePeriods(body && body.periods);
  if (!periods) return jsonResponse({ error: "資料格式錯誤" }, 400);

  var now = new Date().toISOString();
  var write = function () {
    return db.prepare("UPDATE rooms SET periods_json = ?, updated_at = ? WHERE id = ?")
      .bind(JSON.stringify(periods), now, id).run();
  };

  // Databases created before week periods existed lack this column; add it
  // on first use so nobody has to remember to run the migration by hand.
  try {
    await write();
  } catch (e) {
    if (!/no such column/i.test(String(e && e.message))) throw e;
    await db.prepare("ALTER TABLE rooms ADD COLUMN periods_json TEXT NOT NULL DEFAULT '[]'").run();
    await write();
  }

  return jsonResponse({ ok: true, updatedAt: now });
}
