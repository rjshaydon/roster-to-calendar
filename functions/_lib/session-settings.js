import { applySessionPatch, sessionValueKey } from '../../public/static/session-settings-patch.js';
import { queryAccountCustomEvents, hospitalLocationsFromSession } from './d1-calendar.js';

export async function saveSessionSettings(db, { email, profile, changes, sanitizeCustomEvents }) {
  const id = profile?.profileId || email;
  const isProfile = Boolean(profile);
  const row = await db.prepare(isProfile ? 'SELECT state_json,settings_revision FROM doctor_profiles WHERE profile_id=?' : 'SELECT session_json,settings_revision FROM account_states WHERE email=?').bind(id).first();
  const state = isProfile ? JSON.parse(row?.state_json || '{}') : {};
  const current = isProfile ? state.session || {} : JSON.parse(row?.session_json || '{}');
  const previousEvents = isProfile ? current.customEvents || [] : await queryAccountCustomEvents(db, email, { limit: 201 });
  if (previousEvents.length > 200) throw new Error('Custom events exceed the settings save limit.');
  current.customEvents = sanitizeCustomEvents(previousEvents, isProfile ? '' : email);
  const next = applySessionPatch(current, changes);
  next.customEvents = sanitizeCustomEvents(next.customEvents, isProfile ? '' : email);
  if (next.customEvents.length > 200) throw new Error('Custom events exceed the settings save limit.');
  if (sessionValueKey(current) === sessionValueKey(next)) return { ok: true, unchanged: true };
  const revision = crypto.randomUUID(), now = new Date().toISOString();
  const durable = { ...next };
  if (!isProfile) delete durable.customEvents;
  const encoded = JSON.stringify(isProfile ? { ...state, session: durable } : durable);
  if (encoded.length > 256 * 1024) throw new Error('Saved settings exceed the size limit.');
  const statements = [];
  if (isProfile) {
    statements.push(row ? db.prepare('UPDATE doctor_profiles SET state_json=?,settings_revision=?,updated_at=? WHERE profile_id=? AND settings_revision=? AND state_json=?').bind(encoded, revision, now, id, row.settings_revision, row.state_json)
      : db.prepare("INSERT OR IGNORE INTO doctor_profiles(profile_id,doctor_key,display_name,source_types_json,state_json,created_at,updated_at,settings_revision) VALUES(?,?,?,?,?,?,?,?)").bind(id, profile.doctorKey, profile.displayName, JSON.stringify(profile.sourceTypes), encoded, now, now, revision));
  } else {
    statements.push(row ? db.prepare('UPDATE account_states SET session_json=?,settings_revision=?,updated_at=? WHERE email=? AND settings_revision=? AND session_json=?').bind(encoded, revision, now, email, row.settings_revision, row.session_json)
      : db.prepare('INSERT OR IGNORE INTO account_states(email,session_json,updated_at,settings_revision) VALUES(?,?,?,?)').bind(email, encoded, now, revision));
    const previous = new Map(current.customEvents.map(event => [event.id, event]));
    const incoming = new Map(next.customEvents.map(event => [event.id, event]));
    const eventChanges = [...new Set([...previous.keys(), ...incoming.keys()])].filter(key => sessionValueKey(previous.get(key)) !== sessionValueKey(incoming.get(key)));
    if (eventChanges.length > 40) throw new Error('Change at most 40 custom events in one save.');
    for (const key of eventChanges) {
      const event = incoming.get(key);
      if (!event) statements.push(db.prepare('DELETE FROM custom_events WHERE owner_email=? AND id=? AND EXISTS(SELECT 1 FROM account_states WHERE email=? AND settings_revision=?)').bind(email, key, email, revision));
      else statements.push(db.prepare(`INSERT INTO custom_events(owner_email,id,title,start_date,end_date,all_day,start_time,end_time,location,include,updated_at)
        SELECT ?,?,?,?,?,?,?,?,?,?,? WHERE EXISTS(SELECT 1 FROM account_states WHERE email=? AND settings_revision=?)
        ON CONFLICT(owner_email,id) DO UPDATE SET title=excluded.title,start_date=excluded.start_date,end_date=excluded.end_date,all_day=excluded.all_day,start_time=excluded.start_time,end_time=excluded.end_time,location=excluded.location,include=excluded.include,updated_at=excluded.updated_at`)
        .bind(email,event.id,event.title,event.startDate,event.endDate,event.allDay?1:0,event.startTime||'',event.endTime||'',event.location||'',event.include===false?0:1,now,email,revision));
    }
    const oldLocations = hospitalLocationsFromSession(current), newLocations = hospitalLocationsFromSession(next);
    for (const source of Object.keys(newLocations)) if (newLocations[source] !== oldLocations[source]) statements.push(db.prepare(`INSERT INTO account_hospital_locations(email,source_type,location,updated_at)
      SELECT ?,?,?,? WHERE EXISTS(SELECT 1 FROM account_states WHERE email=? AND settings_revision=?)
      ON CONFLICT(email,source_type) DO UPDATE SET location=excluded.location,updated_at=excluded.updated_at`).bind(email,source,newLocations[source],now,email,revision));
  }
  const result = await db.batch(statements);
  if (Number(result[0]?.meta?.changes || 0) !== 1) {
    const error = new Error('Settings changed during this save. Reload before saving; your local edit has been retained.');
    error.code = 'SESSION_SETTINGS_CONFLICT'; throw error;
  }
  return { ok: true };
}
