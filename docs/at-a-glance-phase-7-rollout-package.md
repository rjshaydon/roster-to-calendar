# At a glance Phase 7 rollout package

This is an operational worksheet, not permission to deploy. Phase 7 must be
performed gradually, with an explicit approval before each remote step.

## Release identity and protected state

- Release branch: `codex/at-a-glance-d1-optimization`
- Protected production base: `aa9eed8`
- Release candidate: the Phase 7A commit containing this worksheet
- Doctor identity branch: `codex/durable-doctor-identity-aliases` (untouched)
- Current repository automation defaults: roster writes `false`; watchdog
  automation `false`
- Production D1, R2, Pages, secrets, metrics and migration history have not
  been accessed while preparing this package.

Before any deployment, record the exact release commit, current production
commit, current migration ledger, D1 daily read/write usage and UTC reset time.
Stop if the deployed commit or migration ledger differs from the recorded
state.

## 7A safety controls

The release candidate adds independent, fail-closed controls:

| Purpose | Setting | Inert value |
| --- | --- | --- |
| Immediate shared-path stop | `FACILITY_SHARED_EMERGENCY_PAUSED` | `true` |
| EDs allowed to publish | `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST` | empty |
| EDs allowed to use shared readers | `FACILITY_SHARED_READER_SOURCE_ALLOWLIST` | empty |
| Users allowed to use shared readers | `FACILITY_SHARED_READER_COHORT` | empty |
| Automated roster writes | `ROSTER_AUTOMATION_WRITES_ENABLED` | `false` |
| Automated source IDs allowed | `ROSTER_AUTOMATION_SOURCE_ALLOWLIST` | empty |
| Queue dispatch/polling | `ROSTER_AUTOMATION_QUEUE_ENABLED` | `false` |
| Manual advanced maintenance | `ROSTER_ADVANCED_MAINTENANCE_ENABLED` | `false` |
| External watchdog | `ROSTER_AUTOMATION_ENABLED` | `false` |

All existing phase controls must also remain false at inert deployment:

- `FACILITY_ACCESS_MATERIALIZATION_ENABLED`
- `FACILITY_SHARED_METADATA_BUILD_ENABLED`
- `FACILITY_SHARED_METADATA_ENABLED`
- `FACILITY_SHARED_DAYS_BUILD_ENABLED`
- `FACILITY_SHARED_DAYS_ENABLED`
- `FACILITY_SHARED_CONTACTS_BUILD_ENABLED`
- `FACILITY_SHARED_CONTACTS_ENABLED`

Missing allowlists and cohorts deny access. The Creator cohort applies only to
the Creator's own view; entering another person's profile does not join that
person to the canary. An ED must be present in the applicable allowlist even
when a global phase setting is accidentally enabled.

## Migration worksheet

Apply no migration until the production ledger and quota headroom have been
reviewed. Apply only missing migrations, one at a time, in this order:

1. `0025_runtime_schema_parity.sql` — explicit tables historically created by
   runtime repair. Confirm existing table/index compatibility first.
2. `0026_facility_overview_materialisation.sql` — compact coverage, term-staff
   contribution and visibility tables.
3. `0027_facility_access_materialisation.sql` — one short-lived access row per
   subject.
4. `0028_facility_stream_catalog_materialisation.sql` — compact stream facts.
5. `0029_facility_day_publications.sql` — one publication/fencing row per ED.

These migrations create empty structures and indexes; none contains a
backfill. After each migration, record success and current D1 read/write usage.
Do not retry an unclear result until the migration ledger has been inspected.

## Initial publication budget

Initial publication uses the token-protected
`POST /api/automation/facility-materialize` endpoint. It requires all of:

- roster writes enabled;
- advanced maintenance enabled;
- the requested ED in `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST`;
- the emergency pause disabled.

The endpoint is a dry run unless the JSON body contains `"execute": true`.
Use one ED and one term only. Start with `mmc`, unless the release review names
a different ED. Example dry-run body:

```json
{
  "sourceType": "mmc",
  "termStart": "2026-08-03",
  "maximumDates": 120
}
```

The response lists every planned date and these upper bounds:

- one indexed day query per planned date;
- at most four publication-state writes;
- at most `planned dates + affected months + 3` R2 writes;
- zero broad `roster_events` scans.

Reject the run if it exceeds 120 dates, includes an unexpected date, reports a
broad roster scan, or its predicted cost is not comfortably inside the day's
remaining quota. Execution requires a second, separately approved request with
the identical inputs plus `"execute": true`. Never initialise multiple EDs in
parallel.

Medical term ends are derived from the next actual first-Monday term boundary,
not a fixed 90-day approximation. The 14-day pre-term visibility rule remains
unchanged.

## Gradual canary sequence

Each numbered step is a separate checkpoint. Do not combine them into one
deployment action.

1. Deploy code with the emergency pause true, empty allowlists/cohort, all
   phase controls false, roster automation false and watchdog false. Confirm
   ordinary existing behaviour only; do not publish data.
2. With adequate quota headroom, apply the reviewed empty-schema migrations,
   one at a time. Keep every new reader and writer disabled.
3. Permit only the chosen ED in `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST`.
   Temporarily enable roster writes and advanced maintenance solely for the
   approved dry run. Review its exact dates and upper-bound estimate.
4. If separately approved, execute that one bounded initial publication.
   Disable advanced maintenance immediately afterward and inspect D1/R2 usage.
5. Set reader ED allowlist to the same single ED and cohort to `creator`.
   Enable access, metadata, day and contact readers in dependency order, only
   after their corresponding objects exist. Test the Creator's own view; do
   not test impersonation as the canary path deliberately excludes it.
6. Observe real Cloudflare request and row metrics through an agreed quiet
   period. Compare examined rows, returned rows and writes with the local
   estimates. Any unexplained growth stops the rollout.
7. Add EDs one at a time. Publish and verify each ED before adding it to the
   reader allowlist. Expand `FACILITY_SHARED_READER_COHORT` from `creator` to
   `all` only after every intended ED has passed its own checkpoint.
8. Restore automation separately, source by source. A source must be present in
   `ROSTER_AUTOMATION_SOURCE_ALLOWLIST`; queue and watchdog controls remain
   independent. Do not restore them merely because readers are healthy.

## Capacity and acceptance

The planning envelope is 50 simultaneously visible On shift pages, viewed for
12 hours, refreshing every 60 seconds: 36,000 refresh requests per day. (The
earlier 8-hour example would be 24,000.) Unchanged contact refreshes have a
local acceptance target of zero D1 contact reads/writes and three bounded R2
reads. Authentication and a single indexed access-session read are accounted
separately; their actual production row cost must be measured during canary.

Local scale estimates use the deterministic 109,200-event fixture. They are
not Cloudflare billing measurements. Production acceptance requires:

- no reader-time `roster_events` or `roster_daily_presence` scan;
- unchanged contact refresh: zero D1 contact reads and zero writes;
- unchanged ingestion: zero event, membership, presence and cache rewrites;
- all queries bounded by ED/date, subject key or another indexed compact key;
- actual usage consistent with the dry-run estimate and canary request count;
- at least 50% of the daily D1 quota remaining after the approved test window.

Pause and investigate at 50% daily usage. Stop all optional writes and shared
rollout activity no later than 70%, or immediately on any unexplained spike.
These are operational ceilings, not targets.

## Rollback

Rollback is configuration-first and does not delete data:

1. Set `FACILITY_SHARED_EMERGENCY_PAUSED=true`.
2. Set all shared reader/build controls false and empty all source allowlists
   and the reader cohort.
3. Set roster writes, queue, advanced maintenance and watchdog controls false.
4. Confirm the old application path is serving ordinary requests and inspect
   usage before considering a code rollback.
5. Retain immutable R2 snapshots and compact D1 rows for diagnosis. Do not run
   cleanup, backfill, repair or repeat publication during an incident.

The shared reader deliberately returns to the pre-rollout path when disabled.
If that legacy path itself causes unsafe use, keep the affected feature paused
rather than repeatedly retrying it.

## Evidence required before wider release

Record for every checkpoint: UTC time, commit, settings changed, ED/cohort,
dry-run output, requests made, D1 rows read/written before and after, R2
operations, response correctness and rollback result. Production numbers must
come from Cloudflare metrics; local estimates must remain labelled as such.
