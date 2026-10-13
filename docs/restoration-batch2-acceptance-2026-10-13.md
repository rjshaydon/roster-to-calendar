# Batch 2 production acceptance — 13 October 2026 (Melbourne)

## Current status

The prepared app-calendar reader is enabled in production. Commit `11580ca1`
was deployed successfully as Pages deployment
`c670fd45-2e57-4cec-921c-dc2d49d91425`; its effective
`SUBSCRIPTION_ARTIFACT_MODE` is `ready`. Preview remains `off`.

The calendar CPU implementation and rollout are complete. Final additional
site-specific live comparisons remain subject to approval, and a subsequent
natural MCH changed-file delivery has not been observed in this pass. These are
verification items, not evidence of another known calendar CPU defect.
No clinical roster is changed merely to manufacture a test.

This record supersedes earlier incomplete audit/prime and universal 10-ms
acceptance assertions in this document and the 12 October verification notes.

## Changes and safeguards

- GitHub/Node prepares immutable per-doctor/month and complete per-doctor/source
  artifacts. Sixteen small identity index groups replace whole-site catalogue
  decoding in interactive calendar requests. Only changed content is uploaded.
- Publication requires complete derivatives and a matching source-manifest
  generation. Interrupted uploads, stale parents and missing groups defer
  delivery rather than publishing partial calendars. The existing cache remains
  available while a replacement is prepared.
- Current aliases, account scope, historical grades, locations, custom events,
  overrides, ownership and revocation remain enforced. Identity changes invalidate
  the correct calendar; a colleague-only edit preserves an unaffected doctor's
  revision.
- The actual browser automatic refresh now sends `cachedRevision` for personal,
  admin and profile views. An unchanged response updates the poll acknowledgement
  without rebuilding, persisting or repainting the snapshot. Late responses
  cannot replace a different person's current view.
- Historical caches are pinned by key/revision. Small externally prepared date
  summaries allow skipping their compressed bodies only where all relevant
  history is covered by published terms or outside the requested range. Any gap,
  invalid summary or changed revision retains the historical read.
- The annual window now covers January through the following January (bounded
  to 400 days). A two-candidate primary-key registry lookup retains history from
  the former 370-day cache while preferring an exact-range replacement.
- D1 caps, source priority, protected import headroom, reservations and unknown
  outcome holds remain unchanged. No paid Cloudflare upgrade was made.

## Completed production preparation

Doctor-source run `38357051324` completed at 06:40 UTC. Read-only readiness checks
again confirmed matching parent generations and grouped catalogues for
MMC/MCH/DDH/VHH. Casey has no shared manifest in this pipeline and remains in its
later restoration batch.

Historical coverage run `38381623053` completed at 10:20:48 UTC with 556 summaries.
Continuation/check run `38382709040` completed with zero further preparation:
unchanged original snapshots were not downloaded again. Old revision labels can
exceed 2 KB; summaries store a compact SHA-256 scope, retaining generation fences
without repeating those labels in routine metadata.

External production identity audit `38326766417` completed at 00:24:16 UTC:
1,013 examined, 97 proposals, 119 checkpoints. The earlier assertion that no
production audit had completed is obsolete. Subscription workflow `38378555581`
completed at 09:45:48 UTC with one changed feed built; completed priming and the
readiness certificate were established in the preceding acceptance pass.

## Exact delivery and performance evidence

The user approved the temporary read-only comparison of their own live accounts
in their Cloudflare account. Its D1/R2 wrappers reject mutations; temporary
authentication hashes exist only in memory. Private local responses have restricted
permissions. No real account password or roster content was changed. The temporary Worker
and its secret were deleted after the completed tests; the local access key was
removed, and both native/production tails were stopped.

The final annual-window comparisons returned HTTP 200 and matched every event,
review field and ordering against the existing reader: 128 standard-user events
and 155 Creator events. Each account had eight full loads and eight conditional
checks; all conditional responses acknowledged the existing snapshot.

Available native tail samples recorded 15 optimized full loads at 6–10 ms CPU
and seven conditional checks at 4–5 ms. All sampled outcomes were `ok`, and none
downloaded compressed legacy history. Full wall times were 697–927 ms;
conditional checks 363–516 ms. The two old-reader baseline samples used 52 and
103 ms CPU. These limited samples are not a universal Pages performance guarantee.

Actual production Safari reloads displayed the owner's calendar successfully.
Pages samples include initial/repeat calendar requests at 24 and 11 ms CPU and
an ensuing calendar check at 26 ms; all returned HTTP 200 with 32–43 D1 rows read,
zero writes and complete accounting metadata. Background samples also include
occasional higher CPU values, all with successful outcomes. Do not misreport
private warmed timings as universal production or cold-start timings.

Cloudflare documents CPU rollover for infrequent excess. Acceptance considers
actual resource-exhaustion outcomes, cold/warm distributions and sustained use;
an isolated over-10-ms successful sample is not automatically a failed Free-tier
invocation. The optimization removes broad roster and historical decoding from
routine checks; it does not guarantee every invocation is below 10 ms.

Fresh production subscription comparisons returned HTTP 200 and exactly matched
all current fields against read-only rendering (apart from generated timestamps):
128 staff events and 155 owner events, with no prior event IDs removed. Actual
production feed samples used 4–5 ms CPU and zero D1 writes.

Settled account analytics at the evening check: approximately 1,290,681 rows read
and 29,276 written in the current UTC day. This diagnostic sample does not override
normal admission; its stand-alone report still lacks billing/second-sample
clearance. D1 resets at 11 am Melbourne during daylight saving.

## Validation

Passing targeted suites: doctor-calendar artifacts, calendar refresh,
subscriptions, facility materialization, request attribution, client request
budgets, clinician switching, cold route startup, roster-delivery reliability,
external identity audit and audit checkpoints.

Tests cover removals/corrections, grade history, current authorization, alias
changes, source/range isolation, colleague-only edits, conditional zero-body
responses, interrupted/resumed publication, conservative history coverage, long
legacy revision labels, late-January delivery and old-cache migration.

The earlier broad `test:fixtures` run stopped at an existing source-text assertion
around line 922 expecting the former active-file-selection implementation in
unmodified `derived.js`. It is separate from runtime parity and remains recorded;
the assertion has not been weakened to claim a clean broad run.

## Remaining observation and deferred work

- Additional site-specific live comparison needs explicit approval for sending
  another clinician's name/identifier to the private temporary test Worker.
  Automatic approval review rejected that extension; existing account comparisons
  and workflow logs remain usable. Do not bypass the rejection.
- Observe the next genuine MCH changed-file import and resulting calendar delivery.
  Existing MCH import and prepared source generation are present, but a newly
  edited clinical file was not manufactured for acceptance tonight.
- Further cold-start/background CPU tuning belongs to measured follow-up where
  necessary; no sampled resource-exhaustion failure was observed in this rollout.
- Investigate the iPhone bottom bar stranded midway up the page after keyboard
  use/app resume, after Batch 2 acceptance. Reported 13 October; deferred by user.
- Then follow the agreed batch order: 6, 7, 3, 4, 5, 8.
