# All EDs staff-history search — implementation checkpoint

Branch: `codex/ed-staff-history-search`. Based on the agreed 19 October plan.

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

## Activation intentionally pending

Both `STAFF_HISTORY_BUILD_ENABLED` and `STAFF_HISTORY_ENABLED` are false. No new
production deployment or clinical roster replacement was performed for this
feature. No schema migration or credential change is required.

At the latest read-only check, D1 had observed 95,475 writes in the UTC day and
only 907 advisory bulk writes after protecting pending publication capacity.
Do not start the optional identity projection export with that remaining margin.
This is Cloudflare's daily D1 allowance, independent of the assistant's five-hour
usage window. The UTC reset is 11 am Melbourne during daylight saving.

After the reset, complete these together:

1. Review/merge the branch, deploy with BUILD true and the user-facing flag false,
   then dispatch `staff-history.yml`. Confirm its initial result is complete.
   Preparation logs must contain aggregate counts/status only, never names,
   aliases, identifiers, tokens or response bodies. Keep the normal admission
   and unfinished-work safeguards; stop on any deferral.
2. Check published term coverage against retained source manifests, including
   genuinely unavailable Casey/older terms. Verify a historical-only example,
   an approved multi-site alias, a real grade transition and an upcoming term
   through authenticated Safari. Do not export private roster rows for testing.
3. Enable the user-facing flag only after the complete initial catalogue exists.
   Verify the entire authenticated history action's CPU and D1 cost, normal blank
   and single-site browsing, unchanged replay and entered-account restrictions.
4. Record release/deployment and aggregate acceptance evidence here. If any gate
   fails, leave the feature disabled and preserve existing staff/calendar views.

Rollback: set `STAFF_HISTORY_ENABLED=false`; stop optional preparation with BUILD
false if needed. Do not delete existing roster facts or published calendars.
The prior authority handoff/checker/manual-import remediation is already deployed
on main. This branch does not start restoration Batches 4/5/8 or deferred 3/7.
