# Clinician overview follow-ups — 4 October 2026

The Creator can use the doctor switcher inside an entered account's At a glance
header. Switching continues to resolve the entered account's own permissions.

Enabled clinicians can browse the ED Staff names/designations directory across
enabled hospitals and published terms. Restricted clinicians receive no shift
events or individual coverage dates from that directory. On shift, By stream,
individual roster APIs and Working together remain subject to roster permissions.
Working together now admits verified future-term hospital memberships, alongside
historical memberships, rather than truncating them at the current term.

Term menus are populated from committed publication manifests. Staff directory
terms span the available hospitals; restricted Working together terms are matched
to that clinician's verified memberships. The existing publication release dates
still apply: an ingested but unreleased term is not yet available for review.
No imaginary next-year terms are generated. Date-range searches remain available.

Casey is a supported roster identity. A missing Casey publication no longer
rejects an otherwise valid multi-site personal calendar. Available hospitals are
shown, missing publication coverage is reported, and retained unpublished-source
shifts are preserved. Missing month/staff objects still prevent destructive
replacement of a calendar. Identity count and hospital allowlist bounds remain.

Working together defaults only to the entered profile when it occurs in the
selected period's authorised directory. Clearing all selections stays empty;
no alphabetical or all-staff roster is rendered or requested by that UI.

Validation includes real API tests for cross-hospital directory names without
shift access, previous/future memberships and unrelated-site denial, dynamic
term discovery, valid Casey identities and retained shifts, and browser
regressions for profile selection and cleared selections. Switching, snapshot,
rollout, access, cached-insights, request-budget and restoration-mutation suites
pass; the Worker compiles. No migration or account edits are needed.
