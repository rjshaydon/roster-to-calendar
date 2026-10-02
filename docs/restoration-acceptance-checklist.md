# Restoration acceptance — 2 October 2026

Implementation and release evidence lives in
[bounded-roster-import-checkpoint.md](./bounded-roster-import-checkpoint.md).
This checklist separates remaining user checks from already deployed features.
Use ordinary app actions; do not edit clinical rosters merely to create test data.

| Journey | Evidence available | Remaining acceptance |
| --- | --- | --- |
| Ordinary login and personal calendar | Local authenticated tests: bounded published shifts, no history discovery or automatic claim writes; existing live calendar checks | Log in, confirm expected shifts and reopen from cache |
| Account linking | Published identity directory verified for all sites; transactional ownership/race/conflict and no-op tests pass | Check suggestions and an intended real link when needed; do not create a synthetic clinical claim |
| Creator directory and switching | Existing Safari switching/picker check; bounded directory tests pass; live 45-account directory and SMS filter (14 accounts) verified | Confirm grades and switch between two intended doctor profiles |
| At a glance | All-site current objects verified; user explored views without errors this morning | Check representative current shifts and site access after this release |
| Contacts | All three five-minute flows previously verified; DDH/MMC/MCH user-confirmed; VHH fresh and unchanged runs verified | Compare current handset holders with the sheet; stale worksheet names need not receive a handset |
| Colleague tools | Cached tests and previous live checks pass | Open both tools on a representative shift |
| Manual imports/replacement/removal | Actual SQLite interruption, rollback, no-op, concurrency and future-term-preservation tests pass; feature gates deployed | Use the next genuinely needed manual change; avoid replacing automated clinical terms merely to test |
| Session settings | Single-record optimistic patch tests and previous unchanged live save pass | Change a harmless display preference and verify on a second device, then restore it |
| Subscriptions | SQLite ICS tests prove immediate active replacement, inactive exclusion, indexed queries and zero writes | Confirm next real change in a subscribed calendar after its client refresh interval |
| Automatic current rosters | MMC/MCH changed deliveries and R2 publication succeeded; VHH real automatic import succeeded; DDH prior import succeeded | All-site maintenance and MMC latest-provider recovery succeeded; review the next unattended source change |
| Upcoming terms | VHH current/next objects verified with 19 October visibility; all-source full-term staging tests pass | Verify real MMC/MCH/DDH upcoming-term deliveries when providers publish them |
| D1 safety | Account-wide admission, atomic reservations, bounded batches and unchanged-input no-ops deployed; latest checker GO | Post-recovery sample GO at 06:12 UTC: 191,377 reads / 18,967 writes; continue ordinary-use observation |
| Deployment hygiene | Approved cleanup deleted 103 deployments with zero failures; independent inventory confirms current Production plus verified rollback and zero Preview | Complete for this release |

The bulk restoration code is deployed. The Mac is unlocked; live directory/filter and personal calendar checks passed.
Remaining identity and cross-device journeys are still acceptance items. Provider publication and subscribed-client refresh are external
dependencies. Global repair/bootstrap, login repair, Creator startup fan-out and
whole-account snapshot warm-up remain closed by design. Durable Doctor Names and
identity merging are separate future product work, not prerequisites for baseline
restoration.
