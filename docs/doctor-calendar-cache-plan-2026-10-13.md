# Batch 2 revision: fast, selectively updated app calendars

Front-end acceptance comes first: complete calendars must become visible at least
as quickly as the current working path. Unchanged checks must not download shifts
or redraw calendars. Preserve personal settings, custom events, location overrides,
approved aliases, historical grades, ranges and removals. Keep Cloudflare Free and
all D1 safeguards. Subscription ICS and app JSON are distinct acceptance paths.

## Implementation batch

1. Reuse immutable monthly publications as external inputs. Node/GitHub splits
   them into content-addressed doctor-month objects. Identical doctor content
   retains its object key across roster generations. Compare prior completion
   indexes so unchanged doctors perform no upload requests. Only changed monthly
   inputs are downloaded; no repeated D1 event scans.
2. Complete every doctor object before publishing a versioned month completion
   index. Prepare these indexes before the source manifest becomes visible.
   Interrupted/replayed work reuses immutable objects, never activates a prefix.
3. For an authorized user's selected aliases, read compact indexes in parallel.
   Revision includes only selected doctor content, relevant grades/term visibility,
   range and personal inputs. Unchanged requests return a current envelope without
   shift downloads. Changed requests read only their small doctor objects.
4. Backfill retained months externally, verify preview/native parity and latency,
   then enable reads. Existing production reader stays in place until acceptance.
5. Narrow subscription invalidation without sacrificing immediate ownership and
   changed-roster fences. Finish the production audit/warm-up and real request cost
   verification. Record CPU exceptions rather than treating HTTP success as proof.

## Verification gates

- Personal/profile calendars: exact event and review parity with old reader.
- No-change: no shift data read, no upload/object writes, no browser redraw.
- A colleague's changed shift does not alter this doctor's revision.
- A removed shift/alias cannot survive through an older cache; grades stay historical.
- Range/settings/custom-event changes invalidate the relevant result.
- Interrupted upload, concurrent source update and CAS replay cannot publish partial
  or misattributed data. New source pointer waits for doctor completion indexes.
- Native changed and unchanged cold/warm CPU and response latency, including the
  complete state route; isolated helper timings alone are insufficient acceptance.
- Backfill is bounded/resumable, D1-event-scan-free and keeps live source work safe.

Rollback keeps immutable artifacts and switches the reader back to its existing
mode. Do not delete history or undo additive cache preparation to roll back reads.

## Implementation evidence and outstanding acceptance

The tested implementation now publishes content-addressed complete doctor/source
artifacts and compact source catalogues outside Cloudflare. Versioned month
objects are reusable build inputs. A colleague edit preserves this doctor's
content key; only changed doctors upload. Metadata-only checks read no shifts.
Publication waits for complete derivatives when the new reader is enabled.
Interrupted source builds retain per-doctor completion receipts. CAS updates
protect current pointers and permit safe rebasing after concurrent preparation.

Local integration regressions pass for current ownership, ranges, historical
grades, custom events, locations, removals, previous-month nights, partial-upload
failure, unchanged input, publication ordering and indexed personal inputs.
Sources remain in build mode, with the new interactive reader disabled. Existing
source delivery does not wait for this experimental preparation in build mode.

Native isolated tests preserve exact events and review details: 128 staff events
and 149 owner events within 2026-01-01 to 2027-01-06. Compact reader responses
sampled 323/399 ms versus 1,628/3,243 ms on the previous reader. Metadata-only
responses sampled 203–301 ms. Initial helper CPU is 8 ms for one site and 14 ms
for the five-site owner; metadata-only samples are 3–11 ms. The complete
synthetic-owner state route still exhibits cold/occasional CPU excess (up to
38 ms in the captured optimized samples). HTTP success and lower latency are
not full Free-tier CPU acceptance. The implementation is not a declaration of
Batch 2 completion. Final authenticated cold/warm CPU work remains required
before enabling the reader.

All canary artifacts are private and isolated; production bindings are read-only.
One isolated synthetic account costs four native rows written at setup. No
clinical roster, production account, identity or subscription is changed by QA.
Production external identity audit has meanwhile completed (20 candidate batches,
64 feature batches); recent subscription warm-up reports completed builds.
