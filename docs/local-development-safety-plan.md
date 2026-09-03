# Production-isolated local development safety plan

## Purpose

Create a repeatable local development environment for the complete roster
application so feature work can be built, exercised, and performance-tested
without reading from or writing to Cloudflare D1, R2, Pages, or the production
automation services.

This is not a second version of the application. The local environment must run
the same static application and Pages Functions used in production, with
Wrangler's local D1 and R2 implementations replacing the remote bindings.

The immediate goals are:

1. make local login reliable;
2. make local data disposable and reproducible;
3. make accidental remote access during development difficult;
4. expose database work clearly enough to catch an expensive query before it
   reaches Cloudflare; and
5. make successful local safety checks a prerequisite for any remote release.

## Why this is now required

On 3 September 2026, production exceeded both Cloudflare D1 free-tier row-read
and row-write allowances. The write incident was caused by an independently
deployed roster watchdog initiating two complete imports of a 4,173-event
roster. The import path inserted derived events, rebuilt daily-presence rows,
and removed superseded rows. Cloudflare recorded more than 272,000 rows written
in 24 hours.

Reverting the doctor-identity branch did not stop that worker because the
worker and Pages application have separate deployment lifecycles. This exposed
three process failures:

- local and remote validation were not separated strongly enough;
- database cost was not a tested property of a change; and
- reverting application code was mistaken for reverting all active production
  processes.

The production containment is on `origin/main`:

- `cff238d Pause quota-heavy roster automation`
- `8d0cf60 Pause manual quota-heavy roster writes`

Those commits pause the watchdog, fail closed at the roster automation
endpoints, guard large Creator roster mutations, and remove repeated invite
schema DDL from ordinary requests.

## Current repository state

This safety harness was adapted on `codex/at-a-glance-d1-optimization` from
the protected production baseline `aa9eed8`.

- The baseline includes `cff238d`, `8d0cf60`, and `aa9eed8`.
- The separate `codex/durable-doctor-identity-aliases` branch remains untouched.
- The repository already has persistent Wrangler state under `.wrangler/state`.
  That state contains a local Creator account with a password hash from a prior
  run. It explains the current combination of “account already exists” and
  “incorrect password” messages.
- `.wrangler/` is ignored by Git, so local database and object-store data will
  not be committed.
- `npm run dev` currently invokes `wrangler pages dev public`. Wrangler uses
  local bindings in this mode, but the project has no deterministic reset,
  migration, seed, credentials, or remote-access preflight around it.
- The local lifecycle exercises the production Pages Functions against local
  Wrangler D1 and R2 bindings without contacting Cloudflare.

Implementation must preserve the existing uncommitted work and unrelated
untracked directories. It must not deploy, push to `main`, run a remote D1
command, or modify a Cloudflare binding.

## Implementation status

The original lifecycle was developed locally on 3 September 2026 and is being
revalidated on the production baseline for the At a glance optimisation:

- all production quota-containment commits remain in the branch history;
- `local:check`, `local:reset`, `local:migrate`, `local:seed`, `local:setup`,
  and `dev:local` now share one fixed `.wrangler/local-safe` environment;
- the synthetic seed provides deterministic Creator and user credentials,
  all five roster sources, representative events, and name variations;
- the lifecycle test proves reset, migration, idempotent seeding, correct and
  incorrect password handling, persistence, and restart behaviour; and
- an actual local Creator login is exercised against bindings that Wrangler
  reports as `Mode local`.

Phase 2 was completed locally on 3 September 2026:

- every server-side network integration now passes through one environment-aware
  outbound guard;
- `LOCAL_ONLY=true` blocks non-loopback requests before `fetch` is called;
- GitHub workflow dispatch returns before querying D1 when running locally;
- FindMyShift, Postmark email, and internal automation request paths fail closed
  with a clear `local-disabled` response;
- automation endpoints are tested with throwing D1 and R2 bindings, proving the
  local-disabled path touches neither service; and
- `npm run test:local-isolation` uses traps and static configuration checks and
  does not reset the developer's current local data.

No Preview or Production resource was used or changed while completing either
phase. Phases 3–5 remain outstanding.

## Non-negotiable safety rules

These rules are release blockers:

- Local development must not use `wrangler dev --remote`, a binding with
  `remote = true`, `wrangler d1 ... --remote`, or `--preview`.
- Local scripts must not contain a production or Preview D1 database ID, R2
  bucket name, Pages URL, watchdog URL, provider URL, or Cloudflare API token.
- The local application must run with roster automation, provider polling,
  GitHub workflow dispatch, email delivery, and snapshot warm-up dispatch
  disabled or replaced with local fakes.
- A test must fail if a local-only process tries to call a non-loopback URL.
- Local reset tooling may delete only one exact repository-owned directory,
  `.wrangler/local-safe`. It must refuse an empty, root, home, repository-root,
  `.wrangler`, or caller-supplied path.
- Seed data must be synthetic or deliberately sanitised. Production account
  password hashes, API secrets, subscription tokens, patient information, and
  raw private rosters must never be copied into the local database.
- The documented local Creator password must be a test-only value and must
  never be accepted as a production secret.
- Opening, searching, or regenerating doctor-name suggestions must perform no
  D1 query. Suggestions use the already-loaded doctor list.
- No feature may be remotely enabled merely because its local interface works.
  It must also pass the database budgets and regression gates in this plan.
- Production automation remains paused until a separate incremental-import
  design is implemented and approved. The local harness must not silently
  re-enable it.

## Target local architecture

```text
Browser at 127.0.0.1
        |
        v
Wrangler Pages local runtime
        |
        +--> local D1 files in .wrangler/local-safe
        +--> local R2 files in .wrangler/local-safe
        +--> blocked/fake automation and external services

No path to Cloudflare D1, Cloudflare R2, Pages deployment,
the production watchdog, GitHub Actions, FindMyShift, or email delivery
```

Cloudflare documents that Pages local development runs the Pages application
and Functions locally, and that local D1 is a standalone environment. Local D1
commands must use `--local`; omitting it can address a remote database:

- <https://developers.cloudflare.com/pages/functions/local-development/>
- <https://developers.cloudflare.com/d1/best-practices/local-development/>
- <https://developers.cloudflare.com/workers/local-development/>

## Required developer commands

Add a small set of commands with narrow, documented meanings. Names may change
during implementation only if Wrangler requires it, but their safety contracts
must not.

### `npm run local:check`

Perform a read-only preflight and fail unless all of the following are true:

- the requested host is `127.0.0.1` or `localhost`;
- the local persistence target resolves exactly to
  `<repository>/.wrangler/local-safe`;
- no D1 or R2 binding has `remote = true`;
- local automation and outbound-network guards are enabled;
- no production secret is loaded from `.dev.vars`, an environment-specific
  file, or the command environment; and
- the configured port is not being served by a different project.

This command must not log environment variable values.

### `npm run local:reset`

Stop with a clear error if the local Pages process is running. Otherwise:

1. resolve and validate the exact local-safe path;
2. remove only `.wrangler/local-safe`;
3. recreate the directory; and
4. report that only disposable local state was removed.

Implement the path validation in a reviewed JavaScript script. Do not expose a
general recursive-delete command in `package.json`.

### `npm run local:migrate`

Apply repository migrations to the local database only, using both:

```text
--local
--persist-to .wrangler/local-safe
```

The wrapper must reject `--remote`, `--preview`, arbitrary database names, and
arbitrary persistence paths. It must print the local database path before
applying anything.

Because the remote migration ledger has previously diverged from the repository
files, this local command must not be reused as a production migration wrapper.

### `npm run local:seed`

Seed a known, minimal but realistic local dataset:

- Creator: `rhaydon@gmail.com`
- a prominently documented test-only password;
- at least two ordinary users, including one unclaimed account;
- doctors and events from every supported source type;
- the agreed name-variation examples;
- approved, rejected, and deferred identity examples;
- a modest set of custom events, claims, facility permissions, and retained
  roster metadata; and
- no real secrets or production password material.

Seeding must be idempotent. Running it twice produces the same logical state
without multiplying events, aliases, decisions, or users.

Use the application's real password hashing and repository helpers where
practical. If seed-only SQL is needed, generate fixed test data through a
reviewed script rather than maintaining opaque password hashes by hand.

### `npm run dev:local`

Run the Pages application on the loopback interface using the dedicated
`.wrangler/local-safe` persistence directory. It must run `local:check` first
and must set explicit local-only flags for:

- external network blocking;
- roster automation disabled;
- quota-heavy roster writes disabled unless a specific local test enables
  them;
- doctor-name review enabled for local testing; and
- fake email and workflow delivery.

It must not accept or pass a `--remote` argument. The command should print:

- the local URL;
- the test Creator email;
- where to find the test password;
- the local data directory; and
- “Cloudflare D1/R2: not in use”.

### `npm run test:local`

Run the local schema, repository, API, and browser smoke tests without requiring
a Cloudflare login or network connection. It should start its own isolated
runtime on a non-production port and clean it up when complete.

## Reliable local login

Do not depend on the password that happens to exist in `.wrangler/state`.

The supported workflow becomes:

```text
npm run local:reset
npm run local:migrate
npm run local:seed
npm run dev:local
```

The login screen should then accept the documented local Creator credentials.
Seed verification must perform one real login request and fail setup if the
response is not successful.

For convenience, `local:setup` may compose reset, migration, seed, and
verification, but the individual commands must remain available for diagnosis.

Do not add a universal login bypass to application code. A development-only
bypass can conceal authentication defects and risks leaking into production.

## Fixture strategy

Maintain two fixture sizes.

### Small functional fixture

Fast enough for every test run. It covers:

- Creator and normal-user login;
- calendar loading and doctor switching;
- one roster file from each source;
- claimed and unclaimed doctors;
- apostrophe, whitespace, initial, title, case, and misspelling variations;
- one safe identity combination;
- one existing-person conflict;
- different-doctor and not-sure decisions; and
- undo/idempotent retry where supported.

### Large synthetic fixture

Used for performance and query-plan gates. It should be at least ten times the
current doctor count and at least twice the largest expected active-event set.
Generate it deterministically from synthetic names and events; do not duplicate
a private production export.

The fixture must include deliberately difficult surname buckets so the
in-browser matcher's comparison cap is exercised.

## Database observability and budgets

Local success must include cost evidence, not only correct output.

### Operation accounting

Add test-only instrumentation around the local D1 binding or repository layer
that records, per request:

- prepared statement count;
- executed statement count;
- rows read;
- rows written;
- returned row count;
- elapsed time; and
- normalized SQL/query label without bound values.

Never log passwords, tokens, raw roster bytes, or SQL parameter values.

Store each budget report as a disposable test artifact. Tests should compare
against ceilings and fail loudly when a change increases database work.

### Query-plan gates

Every new or changed identity query must have an `EXPLAIN QUERY PLAN` test
against the large local fixture.

- Point and pair lookups must use indexed `SEARCH` operations.
- A plan containing `SCAN roster_events`, `SCAN roster_doctors`, or another
  roster-sized table fails.
- An `OR` over separately indexed identity features fails unless its measured
  plan is demonstrably bounded.
- A server loop that issues one query per doctor fails.
- Returning few rows is not evidence that few rows were scanned.

### Identity feature ceilings

Retain the budgets from the doctor-name plan:

| Operation | Additional rows read | Rows written |
| --- | ---: | ---: |
| Open Doctor names | 0 for names/suggestions; under 500 only when loading bounded prior decisions | 0 |
| Search or regenerate suggestions | 0 | 0 |
| Open one doctor's history | under 100 | 0 |
| Save different/not sure | under 20 | at most 1 |
| Confirm a safe same-doctor pair | under 100 | under 20 |

The large fixture must not materially increase the rows touched by a point
mutation. If work grows in proportion to the total roster size, the
implementation is rejected.

### Roster-import budgets

Do not use the identity work as a reason to re-enable roster imports. In the
local environment, separately measure a complete import so its amplification
is visible. Before automation can ever return, it needs its own approved plan
with:

- provider-version and content-hash deduplication proven before writes;
- no second daily-presence rebuild;
- changed-row/delta processing rather than whole-roster replacement where
  feasible;
- a predicted write count checked before mutation;
- a hard per-run write ceiling and abort path; and
- a daily circuit breaker well below the Cloudflare allowance.

## Outbound-network protection

Local D1/R2 simulation is necessary but not sufficient because the application
also knows about provider and workflow endpoints.

Introduce a local-only runtime flag such as `LOCAL_ONLY=true`. When set, one
shared outbound helper must reject every non-loopback request before `fetch`.
Automation dispatch, provider polling, email delivery, and webhook paths must
use that helper or have an earlier local-only guard.

Tests must replace global `fetch` with a trap and prove that:

- loading and using the application performs no external request;
- doctor-name detection performs no request at all;
- automation endpoints return a local-disabled response without accessing D1,
  R2, GitHub, FindMyShift, Pages, or the watchdog; and
- seed and reset scripts reject a non-loopback base URL.

## Automated test matrix

`npm run test:local` is complete only when it covers:

| Area | Required proof |
| --- | --- |
| Isolation | Network trap records zero non-loopback calls |
| Authentication | Seeded Creator and user can log in; wrong password fails |
| Persistence | Restarting local Pages retains seeded data |
| Reset | Reset removes local-safe data only and produces a clean database |
| Idempotency | A second seed and repeated identity decision do not add duplicates |
| Doctor discovery | Suggestions are generated from the loaded client list with zero D1 calls |
| Identity mutation | Correct result, conflict handling, atomic failure, and row budgets |
| Existing application | Calendar, switcher, account, facility, contact, feed, and roster fixtures pass |
| Large fixture | Runtime remains bounded and query plans remain indexed |
| Automation | All production automation remains disabled in local normal use |
| Build | Pages Functions compile from the same source used locally |

Add a single `test:safe` aggregate command that runs the existing regression
suite, local isolation tests, query-plan checks, and function build. This is the
minimum pre-release gate.

## Git and deployment barriers

Local testing does not by itself prevent an accidental production push. Add
separate guardrails:

1. Keep feature development on a feature branch. Do not push it to `main`
   during local iteration.
2. Add a checked-in pre-push hook installer that blocks direct pushes to
   `main` unless an explicit one-use release confirmation is supplied.
3. Replace the generic `npm run deploy` entry with a production-named wrapper
   that requires:
   - the expected branch and commit;
   - a clean working tree;
   - a fresh passing `test:safe` report;
   - explicit typed confirmation; and
   - confirmation that D1 metrics have been recorded.
4. Keep remote migration commands out of normal development scripts.
5. Where repository settings permit, protect `main` and require the safety test
   before merge.
6. Treat Preview as remote infrastructure. Do not use Preview merely as an
   extension of local development.

These barriers supplement human review; they do not grant permission to deploy.

## Safe branch integration sequence

Before implementing the harness:

1. Review the current uncommitted diff and identify only the intended
   doctor-identity files.
2. Preserve those files in a named feature-branch checkpoint. Exclude unrelated
   `.cursor`, `backups`, and other untracked material unless explicitly wanted.
3. Merge `origin/main` into the feature branch so `cff238d` and `8d0cf60` are
   genuinely present.
4. Resolve shared-file conflicts in favour of retaining both the emergency
   guards and the intended bounded identity implementation.
5. Run the existing tests before adding local tooling.
6. Confirm `ROSTER_AUTOMATION_WRITES_ENABLED=false` remains the default after
   the merge.

Do not cherry-pick only part of the containment or manually re-create it from
memory.

## Implementation sequence

### Phase 1: local lifecycle

- Add the path-safe reset script.
- Add local-only migration and preflight scripts.
- Add the deterministic small seed.
- Add `local:setup` and `dev:local`.
- Prove login works after reset and after restart.

Stop here if any command contacts Cloudflare or an external provider.

### Phase 2: isolation tests

- Add the non-loopback network trap.
- Add tests for D1, R2, workflow, email, and provider isolation.
- Test the failure messages so a developer can distinguish a deliberate local
  block from an application defect.

### Phase 3: database cost gates

- Add per-request D1 accounting.
- Add the large synthetic fixture generator.
- Add query-plan assertions and operation budgets.
- Produce a human-readable local cost report.

### Phase 4: local browser workflow

- Automate Creator login.
- Exercise Doctor names in normal language.
- Verify suggestions, confirmation, conflict, rejection, defer, and reload.
- Verify ordinary user workflows and calendar rendering.

### Phase 5: release tooling

- Add `test:safe`.
- Add branch/deploy barriers.
- Update README instructions so `npm run dev` no longer directs contributors
  into an ambiguous persistent database.
- Update the doctor-name plan so this plan replaces its ordinary Preview test
  step.

## Remote validation and eventual release

Completion of this plan does not authorize a remote test or deployment.

After all local gates pass, a separate release decision must specify:

- the exact commit;
- the exact migration, if any;
- predicted rows read and written;
- a feature flag that defaults off and is stored outside D1;
- how the feature can be stopped without D1 access;
- the pre-test D1 account metrics;
- one bounded canary action;
- an immediate post-action metric check; and
- a rollback covering Pages, Workers, schedules, and configuration—not just
  Git history.

Do not apply a migration or open the remote Doctor names workspace merely to
“see whether it works.” Remote validation exists only to confirm a locally
proven, bounded operation.

## Acceptance criteria

The local safety environment is complete only when:

- a new developer can run one documented setup command and log in locally with
  known test credentials;
- reset, migration, seed, restart, and login are deterministic;
- D1 and R2 data live only under `.wrangler/local-safe`;
- normal local use and the complete local test suite generate zero Cloudflare
  D1/R2 operations;
- no local path can dispatch a workflow, poll a provider, send email, or call
  the production watchdog;
- a network trap and static configuration checks enforce those claims;
- the large synthetic fixture exercises more data than production without
  remote infrastructure;
- every new identity query has an indexed local query plan and passes its row
  budget;
- opening and searching Doctor names perform no D1 discovery work;
- all existing regression tests and the Pages Functions build pass;
- the feature branch contains both production quota-containment commits;
- the README clearly distinguishes local, Preview, and Production commands;
- Preview is treated as quota-consuming remote infrastructure; and
- no production or Preview change occurs without a separate explicit release
  instruction.

## Out of scope

This plan does not:

- re-enable roster automation;
- redesign incremental roster ingestion;
- copy production data locally;
- repair or apply remote migrations;
- deploy the doctor-identity feature;
- enable the feature in Preview or Production;
- reset Cloudflare quotas; or
- replace production monitoring and backups.

Those require separate plans or explicit release approval after the local
safety foundation is complete.
