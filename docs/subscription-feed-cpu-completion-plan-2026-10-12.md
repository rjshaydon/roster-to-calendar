# Complete subscription CPU reliability

Status: preparation only. No additional endpoint, cache or workflow is enabled by
this document. Cloudflare Free and existing D1 safeguards stay unchanged.

## Evidence and remaining acceptance

CPU changes through `e510289b` / Pages `798eb35a` preserve the two real subscription
UID sets (128 and 154 events). Final sampled feeds use six statements, 285/372
reads and zero writes, with complete metadata. Captured CPU is 5–15 ms; earlier
samples reached 30 ms. Cold/repeated successes do not guarantee continued
reliability. Cloudflare documents a 10 ms Free HTTP limit with per-isolate burst
flexibility: https://developers.cloudflare.com/workers/platform/limits/#cpu-time

The read-only small Worker comparison still used 8–17 ms and has been deleted.
Moving the existing route to a smaller Worker alone is not sufficient. All local
regressions for grades, merge/reversal, permissions, aliases, custom events,
locations, ICU/DST behavior, admission and source priority pass.

## Next coherent implementation batch

1. Profile the remaining synchronous feed stages against captured private inputs
   locally, avoiding another large synthetic D1 write exercise. Keep clinical
   names, tokens, event payloads and credentials out of public artifacts/logs.
   Determine whether serialization/filtering or binding response processing
   dominates. Local elapsed time does not establish native CPU acceptance.
2. Prepare bounded external subscription assembly in Node/GitHub if profiling
   confirms dynamic assembly cannot reliably fit the Free allowance. Reuse the
   existing roster/custom-event/override and ICS semantics rather than inventing
   a parallel renderer. Export only each account's authorized aliases and selected
   Creator identity; page event data with stable evidence and no silent truncation.
   Do not load every clinician's event history in one request.
3. Publish private ready-to-serve ICS artifacts with conditional, atomic promotion.
   Pin account settings, token ownership, identity links/selected person, active
   roster generations and custom-event revisions. Validate them again before
   activation. Roster edits, identity changes, custom events and account settings
   must invalidate the relevant artifact. Revocation/reassignment must never serve
   a previous person's calendar. Cover full and selected-range subscriptions.
4. Have the request path authenticate the token, verify the relevant evidence and
   return the matching artifact. Define missing/stale behavior explicitly; do not
   disguise missing publication with an empty calendar, drop older shifts, or
   silently serve a former identity. Preserve existing URLs and UID generation.
   Coalesce rebuilds, run them after material changes, and avoid periodic full
   database scans or unchanged-roster downloads.
5. Meter every bounded export/callback, keep source imports ahead of optional work,
   and retain unknown receipts. Use read-only/pure preparation until safe write
   headroom exists. R2 output does not justify bypassing D1 admission. Test atomic
   interrupted publishing, stale fences, owner isolation, aliases/merges/reversal,
   custom events/locations/ranges, unchanged no-ops and bounded cold/warm CPU.
6. Enable only after native verification; compare exact live event IDs and fields,
   then complete registration and the production identity audit when admitted.
   Verify actual changed-roster delivery for MMC/MCH/VHH as well as the already
   observed DDH next-term delivery before accepting Batch 2.

Dedicated account-budget/registration CPU profiling remains part of the same
acceptance work. The observed one-query budget/early priority checks still have
samples above 10 ms. Do not assume a subscription artifact fixes those endpoints.
No paid plan, limit increase or manual reservation refund is proposed.

## Current stop conditions

The last sampled grant has about 2.18M read headroom, zero optional write headroom
and no sticky stop reason. No external production audit is complete yet. Observed
account analytics total roughly 749K reads / 47K writes; the comprehensive CLI
report still says STOP for missing billing evidence and a second sample, so it
is not a general clearance. Automatic maintenance continues to enforce its own
fail-closed analytics grant, import priority and reservations.

Implementation changes are committed/deployed. The remaining work is not cosmetic.
