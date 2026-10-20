# All EDs staff-history search — implementation checkpoint

Merged PR #34; production activation on 20 October. Based on the agreed 19 October plan.

## Completed on the branch

- Existing All EDs / Find staff controls select history mode for eligible
  SMS/CMO accounts and Creator troubleshooting. Single-site and blank searches
  retain ordinary selected-term staff lists. Repeated per-site partial-coverage
  badges are removed; real missing-source messages remain.
- A compact name catalogue covers visible retained published terms. Approved
  aliases share history; unapproved same-name people remain separate. Name
  normalization happens once per catalogue page in the browser, then typing
  filters locally. Twelve matching histories are requested at a time, with
  explicit Show more. No event SQL, calendar hydration or contact reads occur
  in the history reader.
- Results show exact term ranges, recorded grades, roster spellings and upcoming
  allocations. Adjacent same-site/same-grade terms combine in the browser.
  Grade changes, breaks and simultaneous sites remain distinct. Wording is
  “Recorded on roster”, never proof of attendance. Synthetic senior continuity,
  provider-only directory rows and admin designations are not work evidence.
- Background preparation uses versioned per-term fragments. Older cached staff
  lists without contribution counts receive a bounded, indexed contribution
  read under existing admission/priority controls; retired files are excluded.
  Insufficient evidence is not silently treated as absence. Oversized evidence
  stops preparation rather than scanning event history.
- Catalogue assembly runs in GitHub, with immutable name pages/history shards
  and an atomic current pointer. Interrupted preparation reuses stored fragments
  and objects. Unchanged preparation performs no D1 work or R2 writes. Aliases
  are keyset-paged against a pinned identity revision; a concurrent roster or
  identity change prevents activation of a mixed projection.
- Successful roster/grade publication, identity publication and registration/audit
  completion request a coalesced build. Existing roster maintenance also requests
  a catch-up; there is no new schedule or browser polling loop. Existing visible
  roster revision refresh invalidates the browser's history cache.
- Server-side authorization precedes history reads. Normal clinical accounts need
  At a glance and current all-site SMS/CMO eligibility. The existing Creator/owner
  exception is retained for their own troubleshooting context; entered accounts
  retain their own entitlements. Nonclinical director access is not expanded.
  Caches are memory-only and cleared on logout/account change/access loss/expiry;
  late responses cannot repopulate a cleared cache or overwrite a newer search.

## Validation

`test:staff-history` exercises the actual background endpoint/Node runner with
SQLite/R2 fixtures, the cache reader, browser cache and actual state-action gate.
It covers historical-only names, approved alternative spellings, distinct same
names, term grades/overrides, continuity exclusion, older contribution evidence,
future release dates, zero-D1 history reads, unchanged no-op, missing objects,
identity invalidation, interrupted page loads and denial before cache reads.

Relevant existing suites passed: facility access, cached insights, publication
and delivery, source isolation, authority handoff, manual import caches,
restoration mutations, maintenance account budgets, cold route startup, external
identity audit and identity maintenance runner. `npm run check`, diff whitespace
checks and the Pages Functions build passed. Wrangler could not write its normal
log in the sandbox; compilation itself completed successfully.

A disposable Worker with **no production D1/R2 bindings** ran 75 successful
synthetic requests. Observability aggregate samples recorded 1–3 ms CPU for name
pages, twelve histories from a sixteen-person shard (32 term records per person),
one-source terms, a 512-member fragment and artifact storage. See the aggregate
JSON in `docs/evidence/staff-history-native-cpu-2026-10-20.json`. This measures the
new cache operations, not a complete authenticated production request. The test
Worker was deleted afterward.

Browser checks used the actual history renderer, actual cache client and app CSS
with synthetic data. They confirmed surname/approved-alias search, retained input
focus, a single-column history card, coalesced adjacent terms, upcoming labels,
explicit continuation and zero extra calls for already-cached matches. They do
not replace the authenticated real-data acceptance below.

## Production activation — 20 October

Initial preparation run 39347779262 completed successfully with 723 people.
Replay 39348176983 returned `unchanged`; unchanged preparation performs no
clinical D1 writes or catalogue rewrites. All four normal source import workflows
completed successfully after midnight UTC (11 am AEDT). No source queue was
manually bypassed. Preview deployments retain isolated databases and their
maintenance gate; production is the acceptance environment.

Both production flags are enabled. Initial acceptance used the Creator-only
cohort before release to existing eligible SMS/CMO accounts with At a glance.
No access entitlement or source roster was changed. Rollback is
`STAFF_HISTORY_ENABLED="false"`; the build flag can remain enabled.

Authenticated Safari acceptance confirmed:

- Ordinary blank All EDs lists and single-site lists remain selected-term views;
  repeated partial-coverage badges are absent. Genuine unavailable Casey
  coverage remains reported.
- Surname search finds Richard Haydon's DDH/MMC records and upcoming Term 4.
  Old MMC grades absent from retained evidence remain Unknown.
- Jayantha finds the approved Jay Weeraratne identity with DDH/VHH records and
  source spelling retained.
- Aeshan's two unmerged spellings appear separately, including MCH Term 1 and
  DDH Term 2 historical records. Search does not auto-merge names.
- Lee returns separate Jeremy/Jun identities and explicit continuation.

The complete authenticated native handler was exercised with Creator, SMS,
trainee, disabled and entered-account cases using disposable synthetic data.
48 authorized reads succeeded and four prohibited cases returned 403. Final
aggregate sampled CPU: names median 3 ms, p95 7 ms, maximum 9 ms; histories
median 3 ms, p95 5 ms, maximum 6 ms. Creator steady reads use four D1 rows and
zero writes; SMS steady reads use nine rows and zero writes. The first SMS
scope cache miss uses 17 reads and three existing permission-cache writes.
Names envelopes stream already-compressed R2 objects; no names parsing or
serialization occurs in the Worker. Typing filters cached names locally.
See the two `staff-history-authenticated-native-*-2026-10-20.json` evidence files.
The disposable Worker and database were deleted; the database ID remains in
both account inventories for accurate current-day accounting.

Account budget check at 00:31 UTC was GO: settled sample through 00:16 UTC
showed 25,244 reads and 556 writes. This is lagged observed usage, not the final
daily total. Existing forecasts, reservations and publication holds remain
in force; no D1 safeguards or Cloudflare plan were relaxed.

## Reset-day release review

The 11:08 am check correctly withheld admission while analytics' 15-minute
settlement window still pointed into the previous UTC day. No allowance was
inferred from an empty new-day sample.

Before rollout, semantic history checkpoints were added: if publication tags
change but recorded names/terms/grades and identity evidence do not, only the
source watermark is acknowledged. Identity export and catalogue rewrites are
skipped. Cache renewal no longer mistakes a renewed access expiry for an account
switch, and Show more requests made during a pending fetch are continued.

Initial production acceptance uses `STAFF_HISTORY_ENABLED="creator"`; ordinary
accounts remain excluded until the live gates pass. Creator context does not
carry into an entered ordinary account. The normal global setting remains
`"true"` once released. BUILD is enabled for initial preparation.

The targeted history/access/publication/cold-start suites pass. The separate
older clinician-access suite currently cannot run cleanly against migrations
alone: it lacks the legacy runtime index `idx_roster_events_file_doctor`; a
throwaway local schema compatibility run then reaches an older assertion about
legacy event lookups. No production code or stored data was changed to satisfy
that legacy fixture. The new endpoint's authorization gates are covered by the
focused tests; authenticated live acceptance remains required.
