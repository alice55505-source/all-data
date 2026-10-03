CREATE TABLE IF NOT EXISTS rooms (
  id TEXT PRIMARY KEY,
  name TEXT,
  created_at TEXT NOT NULL,
  groups_json TEXT NOT NULL DEFAULT '[]',
  stats_json TEXT NOT NULL DEFAULT '{}',
  periods_json TEXT NOT NULL DEFAULT '[]',
  updated_at TEXT NOT NULL
);
