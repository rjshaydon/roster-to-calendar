-- Private channels: a requester and an existing holder never share a thread.
CREATE TABLE support_cases (
 case_id TEXT PRIMARY KEY, requester_email TEXT NOT NULL,
 kind TEXT NOT NULL CHECK(kind IN ('support','identity')), subject TEXT NOT NULL CHECK(length(subject)<=120),
 person_id TEXT NOT NULL DEFAULT '', context_json TEXT NOT NULL DEFAULT '{}',
 status TEXT NOT NULL DEFAULT 'open' CHECK(status IN ('open','resolved')),
 version INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
 resolution_operation_id TEXT NOT NULL DEFAULT '', legacy_reference TEXT UNIQUE,
 admin_unread_count INTEGER NOT NULL DEFAULT 0 CHECK(admin_unread_count>=0),
 waiting_admin_count INTEGER NOT NULL DEFAULT 0 CHECK(waiting_admin_count>=0),
 has_admin_unread INTEGER GENERATED ALWAYS AS(admin_unread_count>0) STORED,
 waiting_admin INTEGER GENERATED ALWAYS AS(waiting_admin_count>0) STORED
);
CREATE INDEX idx_support_cases_status ON support_cases(status,updated_at DESC,case_id DESC);
CREATE INDEX idx_support_cases_unread ON support_cases(status,has_admin_unread,updated_at DESC,case_id DESC);
CREATE INDEX idx_support_cases_waiting ON support_cases(status,waiting_admin,updated_at DESC,case_id DESC);
CREATE INDEX idx_support_cases_requester ON support_cases(requester_email,case_id);
CREATE UNIQUE INDEX idx_support_identity_open ON support_cases(requester_email,person_id)
 WHERE kind='identity' AND status='open' AND person_id<>'';
CREATE TABLE support_channels (
 channel_id TEXT PRIMARY KEY, case_id TEXT NOT NULL REFERENCES support_cases(case_id) ON DELETE CASCADE,
 participant_email TEXT NOT NULL REFERENCES account_profiles(email) ON DELETE CASCADE,
 case_status TEXT NOT NULL DEFAULT 'open', waiting TEXT NOT NULL DEFAULT 'admin' CHECK(waiting IN ('admin','participant','none')),
 last_seq INTEGER NOT NULL DEFAULT 0, participant_read_seq INTEGER NOT NULL DEFAULT 0, admin_read_seq INTEGER NOT NULL DEFAULT 0,
 version INTEGER NOT NULL DEFAULT 1, updated_at TEXT NOT NULL,
 participant_unread INTEGER GENERATED ALWAYS AS(last_seq>participant_read_seq) STORED,
 admin_unread INTEGER GENERATED ALWAYS AS(last_seq>admin_read_seq) STORED,
 UNIQUE(case_id,participant_email)
);
CREATE INDEX idx_support_channels_inbox ON support_channels(participant_email,case_status,updated_at DESC,channel_id DESC);
CREATE INDEX idx_support_channels_unread ON support_channels(participant_email,case_status,participant_unread,updated_at DESC,channel_id DESC);
CREATE INDEX idx_support_channels_waiting ON support_channels(participant_email,case_status,waiting,updated_at DESC,channel_id DESC);
CREATE INDEX idx_support_channels_case ON support_channels(case_id,channel_id);
CREATE TABLE support_messages (
 channel_id TEXT NOT NULL REFERENCES support_channels(channel_id) ON DELETE CASCADE,
 sequence INTEGER NOT NULL, author_type TEXT NOT NULL CHECK(author_type IN ('admin','participant','system')),
 body TEXT NOT NULL CHECK(length(body)<=4000), created_at TEXT NOT NULL,
 PRIMARY KEY(channel_id,sequence)
);
CREATE TABLE support_inboxes (
 email TEXT PRIMARY KEY, version INTEGER NOT NULL DEFAULT 0,
 unread INTEGER NOT NULL DEFAULT 0 CHECK(unread>=0)
);
CREATE TABLE support_mutations (
 actor_email TEXT NOT NULL, mutation_id TEXT NOT NULL, digest TEXT NOT NULL, result_json TEXT NOT NULL,
 created_at TEXT NOT NULL, PRIMARY KEY(actor_email,mutation_id)
);
CREATE TABLE support_case_events (
 case_id TEXT NOT NULL REFERENCES support_cases(case_id) ON DELETE CASCADE,
 event_id TEXT NOT NULL, kind TEXT NOT NULL, actor TEXT NOT NULL,
 note TEXT NOT NULL DEFAULT '', operation_id TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL,
 PRIMARY KEY(case_id,event_id)
);
CREATE INDEX idx_support_case_events ON support_case_events(case_id,created_at DESC,event_id DESC);

CREATE TRIGGER support_channel_insert AFTER INSERT ON support_channels BEGIN
 INSERT INTO support_inboxes(email,version,unread) VALUES('rhaydon@gmail.com',1,NEW.admin_unread)
 ON CONFLICT(email) DO UPDATE SET version=version+1,unread=unread+NEW.admin_unread;
 INSERT INTO support_inboxes(email,version,unread) SELECT NEW.participant_email,1,NEW.participant_unread WHERE NEW.participant_email<>'rhaydon@gmail.com'
 ON CONFLICT(email) DO UPDATE SET version=version+1,unread=unread+NEW.participant_unread;
 UPDATE support_cases SET admin_unread_count=admin_unread_count+NEW.admin_unread,
  waiting_admin_count=waiting_admin_count+CASE WHEN NEW.waiting='admin' THEN 1 ELSE 0 END WHERE case_id=NEW.case_id;
END;
CREATE TRIGGER support_channel_update AFTER UPDATE ON support_channels BEGIN
 INSERT INTO support_inboxes(email,version,unread) VALUES('rhaydon@gmail.com',1,MAX(0,NEW.admin_unread-OLD.admin_unread))
 ON CONFLICT(email) DO UPDATE SET version=version+1,unread=unread+NEW.admin_unread-OLD.admin_unread;
 INSERT INTO support_inboxes(email,version,unread) SELECT NEW.participant_email,1,MAX(0,NEW.participant_unread-OLD.participant_unread) WHERE NEW.participant_email<>'rhaydon@gmail.com'
 ON CONFLICT(email) DO UPDATE SET version=version+1,unread=unread+NEW.participant_unread-OLD.participant_unread;
 UPDATE support_cases SET admin_unread_count=admin_unread_count+NEW.admin_unread-OLD.admin_unread,
  waiting_admin_count=waiting_admin_count+(CASE WHEN NEW.waiting='admin' THEN 1 ELSE 0 END)-(CASE WHEN OLD.waiting='admin' THEN 1 ELSE 0 END) WHERE case_id=NEW.case_id;
END;
CREATE TRIGGER support_channel_delete AFTER DELETE ON support_channels BEGIN
 UPDATE support_inboxes SET version=version+1,unread=unread-OLD.admin_unread WHERE email='rhaydon@gmail.com';
 UPDATE support_inboxes SET version=version+1,unread=unread-OLD.participant_unread WHERE email=OLD.participant_email AND email<>'rhaydon@gmail.com';
 UPDATE support_cases SET admin_unread_count=admin_unread_count-OLD.admin_unread,
  waiting_admin_count=waiting_admin_count-CASE WHEN OLD.waiting='admin' THEN 1 ELSE 0 END WHERE case_id=OLD.case_id;
END;
CREATE TRIGGER support_account_deleted BEFORE DELETE ON account_profiles BEGIN
 UPDATE support_case_events SET note='',actor=CASE WHEN actor=OLD.email THEN '' ELSE actor END
  WHERE case_id IN (SELECT case_id FROM support_channels WHERE participant_email=OLD.email);

 UPDATE support_cases SET requester_email='',subject='Account deleted',context_json='{}',person_id='',status='resolved',version=version+1
  WHERE requester_email=OLD.email;
 UPDATE support_channels SET case_status='resolved',waiting='none',version=version+1
  WHERE case_id IN (SELECT case_id FROM support_cases WHERE requester_email='' AND case_id IN
    (SELECT case_id FROM support_channels WHERE participant_email=OLD.email));
 DELETE FROM support_channels WHERE participant_email=OLD.email;
 DELETE FROM support_inboxes WHERE email=OLD.email;
 DELETE FROM support_mutations WHERE actor_email=OLD.email;
END;

-- Reusing the daily safety gate must not scan every settled receipt.
CREATE INDEX idx_maintenance_unfinished_day ON roster_maintenance_receipts(utc_day)
 WHERE reconciled_at='' AND (finished_at='' OR metadata_complete<>1);
