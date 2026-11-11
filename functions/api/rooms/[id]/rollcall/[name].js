import { jsonResponse, isValidRoomId, parseJsonColumn } from "../../../_lib.js";

// Same read-modify-write-per-request approach as congregations/[name].js,
// but on the separate rollcall_json column so 點名報表 uploads never touch
// the 週報 (兒童／青職) data and vice versa.
async function withRollcall(db, id, mutate) {
  // SELECT * so a database without rollcall_json yet still works; the
  // column is added on first write (same approach as periods.js).
  var row = await db.prepare("SELECT * FROM rooms WHERE id = ?").bind(id).first();
  if (!row) return null;
  if (!("rollcall_json" in row)) {
    await db.prepare("ALTER TABLE rooms ADD COLUMN rollcall_json TEXT NOT NULL DEFAULT '{}'").run();
    row.rollcall_json = "{}";
  }

  var data = parseJsonColumn(row.rollcall_json, {}) || {};
  mutate(data);

  var now = new Date().toISOString();
  await db.prepare("UPDATE rooms SET rollcall_json = ?, updated_at = ? WHERE id = ?")
    .bind(JSON.stringify(data), now, id).run();

  return now;
}

function parseParams(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return { error: jsonResponse({ error: "房間代碼格式錯誤" }, 400) };
  var name = decodeURIComponent(context.params.name || "");
  if (!name) return { error: jsonResponse({ error: "缺少召會名稱" }, 400) };
  return { id: id, name: name };
}

export async function onRequestPut(context) {
  var p = parseParams(context);
  if (p.error) return p.error;

  var body;
  try {
    body = await context.request.json();
  } catch (e) {
    return jsonResponse({ error: "請求格式錯誤" }, 400);
  }
  if (typeof body !== "object" || body === null) {
    return jsonResponse({ error: "資料格式錯誤" }, 400);
  }

  var updatedAt = await withRollcall(context.env.DB, p.id, function (data) {
    data[p.name] = body;
  });
  if (updatedAt == null) return jsonResponse({ error: "找不到這個房間" }, 404);
  return jsonResponse({ ok: true, updatedAt: updatedAt });
}

export async function onRequestDelete(context) {
  var p = parseParams(context);
  if (p.error) return p.error;

  var updatedAt = await withRollcall(context.env.DB, p.id, function (data) {
    delete data[p.name];
  });
  if (updatedAt == null) return jsonResponse({ error: "找不到這個房間" }, 404);
  return jsonResponse({ ok: true, updatedAt: updatedAt });
}
