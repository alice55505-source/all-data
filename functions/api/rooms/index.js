import { jsonResponse, pruneEmptyRooms } from "../_lib.js";

const CODE_ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"; // no 0/O/1/I/L to avoid confusion

// Kept in sync with DEFAULT_GROUPS in js/app.js - only used to seed brand-new
// rooms so first-time users see a ready-to-edit example instead of a blank
// list. Once a room exists, an explicitly-cleared list stays empty (this
// seed is never re-applied), so it must not be derived at read time.
const DEFAULT_GROUPS = [
  { region: "雲東區", members: ["斗六", "古坑", "林內", "西螺", "莿桐", "斗南"] },
  { region: "雲西區", members: ["虎尾", "土庫", "崙背", "褒忠", "二崙", "麥寮", "北港", "口湖"] },
  { region: "嘉義區", members: ["嘉義市（梅山）", "中埔", "竹崎", "番路"] },
  { region: "民雄區", members: ["民雄", "溪口", "大林", "新港", "(六腳)"] },
  { region: "朴子區", members: ["朴子", "布袋", "鹿草", "太保", "水上"] }
];

function generateRoomId() {
  var out = "";
  var bytes = new Uint8Array(6);
  crypto.getRandomValues(bytes);
  for (var i = 0; i < bytes.length; i++) {
    out += CODE_ALPHABET[bytes[i] % CODE_ALPHABET.length];
  }
  return out;
}

export async function onRequestPost(context) {
  var db = context.env.DB;
  var now = new Date().toISOString();
  var groupsJson = JSON.stringify(DEFAULT_GROUPS);

  context.waitUntil(pruneEmptyRooms(db));

  for (var attempt = 0; attempt < 5; attempt++) {
    var id = generateRoomId();
    var existing = await db.prepare("SELECT id FROM rooms WHERE id = ?").bind(id).first();
    if (existing) continue;

    await db.prepare(
      "INSERT INTO rooms (id, created_at, groups_json, stats_json, updated_at) VALUES (?, ?, ?, '{}', ?)"
    ).bind(id, now, groupsJson, now).run();

    return jsonResponse({ id: id }, 201);
  }

  return jsonResponse({ error: "無法產生房間代碼，請重試" }, 500);
}
