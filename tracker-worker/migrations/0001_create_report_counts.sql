CREATE TABLE IF NOT EXISTS report_counts (
  day TEXT NOT NULL,
  form_version TEXT NOT NULL,
  language TEXT NOT NULL,
  count INTEGER NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (day, form_version, language)
);
