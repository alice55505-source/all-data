ALTER TABLE rooms ADD COLUMN metrics_json TEXT NOT NULL DEFAULT '{"meetings":[],"roles":[],"extras":[]}';
