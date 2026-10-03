export function jsonResponse(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { "Content-Type": "application/json" }
  });
}

export function isValidRoomId(id) {
  return typeof id === "string" && /^[A-Z0-9]{4,12}$/.test(id);
}

export function parseJsonColumn(text, fallback) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return fallback;
  }
}

// Cloudflare Pages Functions have no Cron Trigger support (that's a
// Workers-only feature), so instead of a real scheduled job this piggybacks
// on room creation - a low-frequency write endpoint - to opportunistically
// sweep out rooms that never got a single congregation upload. Best-effort:
// a failure here must never break the request it rides along on.
export async function pruneEmptyRooms(db) {
  var cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
  try {
    await db.prepare("DELETE FROM rooms WHERE stats_json = '{}' AND created_at < ?").bind(cutoff).run();
  } catch (e) {
    /* best-effort cleanup only */
  }
}
