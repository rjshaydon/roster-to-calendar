# Manual roster import optimisation — 20 October 2026

User approved implementation on Medium, before the historical staff-search work.
Scope: remove the duplicate daily-presence index from new manual imports when
all relevant live readers use published artifacts. No clinical data repair,
source cadence, historical grade or source-authority rule changes are included.

## Implementation

The existing manual begin phase negotiates a mode without writes. New browsers
advertise support; the server chooses cached mode only for a supported site with
all-cohort published metadata/day readers, colleague readers, subscription
artifacts and automatic publication enabled, legacy readers paused, no emergency
pause and the required R2 binding. Initialization pins the mode in the existing
immutable manifest/hash. This is server-gated, not a browser-selected shortcut.

Cached presence preparation uses the existing single guarded cursor advance.
The browser stops after that advance, instead of making an empty request for
every event batch. Facts, compact metadata, authority activation, publication,
receipts and account-wide reservations remain intact. Current imports retain
their original cached/legacy mode and hashes across retries; old browsers remain
supported. If the required rollout is withdrawn during a cached import, further
mutation/activation fails explicitly while retaining staging for retry.

The browser entrypoint version is advanced with this release. No new deployment
flag, schema migration or change to normal calendar/contact polling is required.

## Cost evidence

Actual browser import orchestration and production manual handler are exercised
locally against all migrations, using existing MMC/MCH full-workbook fixtures.
These are **conservative D1 write reservations**, not native Cloudflare billing:

| Full fixture | Events | Legacy presence rows | Legacy staging reservation | Cached staging reservation |
| --- | ---: | ---: | ---: | ---: |
| MMC Adult | 4,295 | 8,547 | 88,699 | 44,844 |
| MCH | 2,363 | 4,288 | 48,638 | 26,638 |

Savings are approximately 49.4% and 45.2% respectively. Publication and raw-file
retention/admission costs are separate; actual unused reservations settle/refund
through the existing meter. Local logical SQL changes fell from 14,010 to 5,413
for MMC and 7,045 to 2,732 for MCH, but those figures do not include Cloudflare's
index-write billing and must not be presented as billed rows.

The comparison is repeatable with `npm run test:manual-roster-cache`. Neither
variant's full fixture is uploaded to production to manufacture a benchmark.

## Correctness and regression evidence

The new test executes the actual browser function with real manual-handler calls.
It proves complete event counts, zero new presence rows, one presence request,
zero-write completed replay, interrupted import recovery, and pre-rollout legacy
job compatibility. Synthetic examples at MMC/MCH/DDH/VHH publish their roster
facts and verify overnight start/end preservation, leave, term grades, personal
calendar results, subscriptions and cached colleague matches without D1 reads.
Forged cached opt-in is rejected before writes when the required readers are off;
withdrawn rollout blocks activation. Existing authority, source-isolation,
publication-delivery, cached-insight, subscription/doctor-artifact, cold-start,
maintenance-budget and restoration-mutation suites are also checked.

This reuses the existing publication/metadata builders. It does not claim a new
native CPU benchmark, eliminate the complete-term metadata pass, implement
cross-file event diffs, or deduplicate a fresh upload under a new manual file ID.
Those latter optimisations need additional snapshot/identity/authority work;
expanding into them would add risk to this focused reduction. The first genuine
new manual upload remains the native usage/end-to-end acceptance opportunity.

## Release and rollback

Commit/deployment and live read-only checks are recorded below after release.
Rollback may restore the older upload client for future legacy-mode jobs while
keeping shared readers enabled. Do not re-enable SQL presence-based readers for
files whose manifests record cached mode without a separately budgeted index
rebuild. Preserve in-progress plans and published calendars; do not delete
clinical data or disable account safeguards. No staff-search work is included.

Released as commit `91001086` and pushed to `main`. Production Pages deployment:
`https://6c2cce6d.roster-to-calendar.pages.dev`.
The custom domain serves `app.js?v=20261020-manual-cache` and the negotiated
initialization code. Authenticated Safari reload completed and displayed the
existing calendar, including MMC Term 4 shifts. No production roster was replaced
for testing and no migration was needed.

The live read-only checker returned GO with a valid account-wide sample:
2,103,826 reads and 82,838 writes observed in the current UTC accounting day;
advisory bulk headroom was 15,606 writes after forecast/uncertainty. This is not
enough for a full MMC manual upload today, and deployment does not bypass that
limit. Native import billing remains to be observed on the next genuine upload.
The five-hour assistant allowance had 78% remaining at release verification.
