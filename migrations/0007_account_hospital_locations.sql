CREATE TABLE IF NOT EXISTS account_hospital_locations (
  email TEXT NOT NULL,
  source_type TEXT NOT NULL,
  location TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (email, source_type)
);

INSERT OR IGNORE INTO account_hospital_locations (email, source_type, location, updated_at)
SELECT email, 'mmc', COALESCE(NULLIF(json_extract(session_json, '$.settings.defaultLocationMmc'), ''), 'MMC Car Park, Tarella Road, Clayton VIC 3168, Australia'), updated_at
FROM account_states;

INSERT OR IGNORE INTO account_hospital_locations (email, source_type, location, updated_at)
SELECT email, 'ddh', COALESCE(NULLIF(json_extract(session_json, '$.settings.defaultLocationDdh'), ''), 'DDH Car Park, 135 David St, Dandenong VIC 3175, Australia'), updated_at
FROM account_states;

INSERT OR IGNORE INTO account_hospital_locations (email, source_type, location, updated_at)
SELECT email, 'casey', COALESCE(NULLIF(json_extract(session_json, '$.settings.defaultLocationCasey'), ''), 'Casey Hospital, 62-70 Kangan Drive, Berwick VIC 3806, Australia'), updated_at
FROM account_states;

INSERT OR IGNORE INTO account_hospital_locations (email, source_type, location, updated_at)
SELECT email, 'mch', COALESCE(NULLIF(json_extract(session_json, '$.settings.defaultLocationMch'), ''), 'Monash Children''s Hospital, 246 Clayton Road, Clayton VIC 3168, Australia'), updated_at
FROM account_states;
