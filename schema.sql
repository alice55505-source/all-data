CREATE TABLE IF NOT EXISTS rooms (
  id TEXT PRIMARY KEY,
  name TEXT,
  created_at TEXT NOT NULL,
  groups_json TEXT NOT NULL DEFAULT '[]',
  stats_json TEXT NOT NULL DEFAULT '{}',
  metrics_json TEXT NOT NULL DEFAULT '{"meetings":[],"roles":[],"extras":[]}',
  updated_at TEXT NOT NULL
);
