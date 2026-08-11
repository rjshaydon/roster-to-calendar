# At a glance — By stream polish handover for Terra

## Status

Completed on `codex/implement-by-stream`:

1. `76a38f7 Standardize At a glance seniority order`
2. `8b9407e Synchronize By stream banner dates`
3. `153c027 Polish responsive By stream comparison`

## Purpose

Polish the existing By stream implementation on branch `codex/implement-by-stream` without redesigning its data API. This pass has three product outcomes:

1. By stream uses the same banner composition as On shift, with usable From, To, and Today controls synchronised with the lower By stream options.
2. Additional stream results appear beside the first when space permits: three to four on wide desktop and two on normal mobile, with selectors above the results on mobile.
3. Every At a glance view and every hospital uses the same seniority hierarchy: **SMS → SR → CMO → TR → JR → HMO → NP → Physio → Intern → Unknown**.

The existing By stream API, preferred-ED resolver, stream catalogue, coverage states, selector validation, and maximum of six selections remain in scope and should be preserved.

## Current implementation facts

Terra should verify these locations before editing because line numbers will move:

- `facilityOverviewState` in `public/static/app.js` already contains `date`, `byStreamFrom`, `byStreamTo`, and the By stream request/data fields.
- `renderFacilityOverviewHeader()` currently shows full term controls for ED staff, a text note for Working together, a `Compare stream coverage` note for By stream, and the On shift-style From/To/Today banner for the remaining view.
- The existing On shift banner range buttons are disabled and read from the broader calendar preview range. Do not bind By stream to `latestPreview.previewStart`, `latestPreview.previewEnd`, `settings.dateFrom`, or `settings.dateTo`.
- `renderFacilityOverview()` currently renders a second, functional By stream From/To/Today set in `facilityOverviewControls`.
- Both lower By stream date inputs already use `data-facility-overview-by-stream-date`.
- The delegated Today handler already branches on the active tab and sets both By stream bounds to today. Refactor it to use the shared range mutation helper described below.
- `.facility-overview-by-stream-results` currently stacks full-width lanes.
- `FACILITY_OVERVIEW_SENIORITY_ORDER` currently places CMO ahead of Senior Registrar and Intern ahead of NP/Physio.
- `renderFacilityOverviewOnShiftNames()` currently sorts deduplicated people only by display name, which is why DDH stream tiles can appear in the wrong hierarchy.
- The By stream All team ordering currently prioritises shift start before seniority.

## Product contract

### Banner and date synchronisation

When the active tab is By stream, the banner must retain the same outer composition as On shift:

- Doctor identity at the left.
- A From control.
- A To control.
- A Today button.
- Log out at the right.

Replace the By stream `Compare stream coverage` placeholder with functional native date inputs styled consistently with the existing banner controls. Retain the lower By stream From, To, and Today options; the user explicitly wants both surfaces.

Both surfaces are projections of the same state:

```js
facilityOverviewState.byStreamFrom
facilityOverviewState.byStreamTo
```

Do not introduce `byStreamHeaderFrom`, `byStreamHeaderTo`, DOM-to-DOM copying, or any other second source of truth.

Create one mutation path, for example:

```js
setFacilityOverviewByStreamRange({ from, to, changedField, refresh: true })
```

The exact name is not important. Its contract is:

1. Accept a proposed From and/or To value from either surface.
2. Keep ISO `YYYY-MM-DD` strings only.
3. If From moves after To, move To forward to From.
4. If To moves before From, move From backward to To.
5. Commit both state fields together.
6. Rerender so banner and lower controls immediately show the same values.
7. Trigger exactly one `loadFacilityOverviewByStream()` call for a valid completed range.
8. Preserve the existing request-id protection so a stale response cannot replace newer results.

Both Today buttons must call that same path with Melbourne today for both bounds. A single click must cause one state transition and one By stream request, not one request per rendered control.

Use the existing delegated change listener for both banner and lower inputs. The simplest durable contract is to give both sets the existing `data-facility-overview-by-stream-date="from|to"` attributes. Multiple matching controls are expected; event delegation acts only on the changed target.

Set an accurate accessible label such as `By stream date range`. Do not retain `Calendar range; unavailable in At a glance` on an enabled By stream range. Preserve visible `From` and `To` labels and native keyboard/date-picker behaviour.

On shift remains a single-date query using `facilityOverviewState.date`. This pass does not merge On shift date state with the By stream range. Switching tabs should preserve each tab’s deliberate date selection for the lifetime of the open workspace.

### Result layout

Interpret “put another stream beside the first” as the result lanes. Selector forms remain in the right rail on desktop. At the existing narrow-layout breakpoint, the whole selector panel moves above the results.

Use two result modes:

#### One selected stream

- Retain the approved Sol composition.
- The single lane uses the available result width.
- Dates run horizontally within that lane.

#### Two or more selected streams

- Each selector row produces one result lane in selector order.
- Result lanes form an auto-fitting grid.
- Dates run vertically inside each lane.
- The same visible date occupies the same row in every lane.
- Calculate the visible-date list once for the entire comparison before rendering lanes. Do not filter dates independently per lane, because that destroys row alignment.
- If `Hide dates without assignments` is on, remove a date only according to the shared comparison rule. Preserve explicit uncovered-roster dates so missing coverage is not mistaken for no assignment.
- Fifth and sixth selections wrap onto another grid row on desktop.

Target sizing:

- Wide desktop: three to four lanes across when the actual result container has room.
- Smaller desktop/tablet: two to three lanes.
- Normal mobile: two lanes across, with the selector panel above them.
- Exceptionally narrow mobile: one lane; readability takes priority over forcing two columns.

Prefer container-responsive `repeat(auto-fit, minmax(...))` behaviour or an equivalent results modifier over viewport-only assumptions. A lane minimum in the approximate 150–220 px range may be needed across the mobile and desktop breakpoints; tune against real names and times. Do not add one horizontal scrollbar per lane and do not permit page-level horizontal overflow.

Keep lane headers concise and unambiguous:

- Stream and seniority, using compact labels SR/TR/JR where currently appropriate.
- ED and selected range.
- Date heading in every date cell.

The selector rail remains sticky only on desktop. On mobile it must be non-sticky and precede the result grid in reading and focus order.

### Canonical seniority hierarchy

Replace the current order with canonical internal values in this exact rank:

```js
[
  "SMS",
  "Senior Registrar",                  // SR
  "CMO",
  "Transitional/Intermediate Registrar", // TR
  "Junior Registrar",                  // JR
  "HMO",
  "NP",
  "Physio",
  "Intern",
  "Unknown",
]
```

Use the existing normaliser for aliases and extend it only where fixtures reveal a missing alias. The following values must occupy the same rank:

- `SR` and `Senior Registrar`.
- `TR`, `IR`, Transitional Registrar, and Intermediate Registrar.
- `JR` and `Junior Registrar`.
- `NP`, `ENP`, and Nurse Practitioner.
- `Physio`, `AMP`, Physiotherapist.
- `Intern` and `I`.

Blank or genuinely unrecognised values display/sort as Unknown for these At a glance comparisons. Do not rewrite the source roster or persisted parser data as part of display sorting.

Provide shared comparators rather than repeating rank arithmetic. At minimum, cover:

- Seniority labels: rank, then a deterministic label tie-break.
- People in On shift stream tiles: effective seniority rank, then display name.
- By stream `All team`: seniority rank, then shift start, then display name.
- By stream single-seniority results: shift start, then display name.

Apply the hierarchy consistently to:

- On shift stream tiles at MMC, DDH, Casey, and MCH.
- DDH’s specially placed Orange, Silver, Fast Track, AVAO, SSU, night SR, main-team, and SSU-team cards.
- Generic On shift stream and unstreamed cards.
- By stream result assignments.
- By stream seniority selector options and stream-catalogue seniorities.
- ED staff seniority sections.

Hospital-specific tile placement remains unchanged. This work changes the ordering of people/seniority within those tiles, not the established hospital stream layout.

## Implementation map

Primary files:

- `public/static/app.js`
  - Correct `FACILITY_OVERVIEW_SENIORITY_ORDER`.
  - Add/reuse shared seniority comparators.
  - Change `renderFacilityOverviewOnShiftNames()` ordering.
  - Change By stream assignment ordering for All team.
  - Add the shared By stream range mutation helper.
  - Render enabled banner From/To/Today controls for the By stream branch of `renderFacilityOverviewHeader()`.
  - Refactor result rendering into one-lane and aligned multi-lane modes.
- `public/static/styles.css`
  - Style native banner date inputs consistently with the existing range buttons.
  - Add responsive multi-lane result grid rules.
  - Preserve selector-above-results behaviour below the current breakpoint.
  - Prevent narrow-name/time overflow.
- `scripts/test-fixtures.mjs`
  - Add regression assertions for the shared banner controls, range mutation path, responsive lane classes, and canonical seniority order.
  - Prefer behavioural fixture assertions where the current harness permits them; static source assertions alone are insufficient for ordering regressions.
- `docs/by-stream-implementation-plan.md`
  - Already updated with the approved polish contract.

No D1 migration or API contract change is expected. Do not change the server query unless a client-side test exposes missing data needed for the agreed UI.

## Required tests

### Date controls

1. Open By stream: banner and lower controls show the same default From and To.
2. Change banner From to a date after To: both To controls advance to match; one request is sent.
3. Change lower To to a date before From: both From controls move back to match; one request is sent.
4. Change either bound normally: the corresponding banner and lower inputs agree after render.
5. Press banner Today after creating a range: all four date inputs become today; one request is sent.
6. Press lower Today after creating a range: same result and request count.
7. Rapid successive edits: only the newest response is rendered.
8. Switch to On shift and back: On shift retains its single date and By stream retains its range.

### Responsive comparison

Test at minimum with one, two, four, and six stream selections:

- One stream: full-width Sol lane with horizontal dates.
- Four streams at a wide desktop viewport: three or four appear across depending on actual container width, with readable names and times.
- Six streams: remaining lanes wrap in selector order.
- Two streams at a representative mobile viewport around 375–430 CSS px: selectors are above and both result lanes appear beside one another without page overflow.
- Very narrow viewport around 320 CSS px: safe one-column fallback is acceptable.
- Multi-day, multi-stream comparison: date rows align across lanes.
- Hide-empty on/off: alignment remains intact and uncovered-roster dates remain distinguishable.

### Seniority

Create deliberately shuffled representative assignments at each of MMC, DDH, Casey, and MCH. Verify rendered order:

`SMS, SR, CMO, TR, JR, HMO, NP, Physio, Intern, Unknown`

Also verify:

- Names within the same seniority are alphabetical.
- All team uses hierarchy before time.
- Single-seniority By stream results use time before name.
- Aliases rank with their canonical grade.
- Unknown is last.
- DDH special card placement is unchanged.

### Regression commands

Run at minimum:

```sh
npm run check
node --check functions/api/state.js
node --check functions/_lib/d1-calendar.js
npm run test:fixtures
git diff --check
```

Then smoke-test authenticated representative data for MMC, DDH, Casey, and MCH on the branch preview.

## Delivery sequence

Keep the work on `codex/implement-by-stream`. Preserve unrelated untracked files and do not stage them.

Produce independently testable rounds, committing, pushing, and deploying each round before continuing:

1. **Canonical seniority order** — comparator changes plus multi-hospital fixtures.
2. **Shared banner range** — enabled banner inputs, single range mutation path, Today parity, synchronisation and request-count tests.
3. **Responsive stream comparison** — adjacent/aligned result lanes, desktop/mobile wrapping and overflow tests.

Suggested commit messages:

- `Standardize At a glance seniority order`
- `Synchronize By stream banner dates`
- `Polish responsive By stream comparison`

Use the existing branch preview alias after every push:

`https://codex-implement-by-stream.roster-to-calendar.pages.dev`

## Acceptance checklist

- [ ] By stream banner visually matches the On shift banner composition.
- [ ] Banner From/To are enabled and use By stream state.
- [ ] Banner and lower From/To remain identical after every edit.
- [ ] Both Today buttons set both bounds to Melbourne today.
- [ ] Each user date action starts at most one By stream request.
- [ ] One selected stream retains the Sol horizontal-date layout.
- [ ] Multiple streams appear beside one another with aligned date rows.
- [ ] Wide desktop supports three to four readable lanes.
- [ ] Normal mobile supports two readable lanes with selectors above.
- [ ] No page-level horizontal overflow is introduced.
- [ ] All four hospitals use the exact canonical seniority hierarchy.
- [ ] DDH stream-tile staff are hierarchy-first rather than alphabetic-only.
- [ ] Unknown always appears last.
- [ ] Existing coverage, duplicate-selection, default-ED, and stale-response behaviour still passes.
- [ ] Every implementation round is committed, pushed, deployed, and ready for user testing.

## Non-goals

- Sharing date state between On shift and By stream.
- Removing the lower By stream date controls.
- Changing the preferred-ED resolver.
- Changing hospital-specific stream placement.
- Adding drag-and-drop selector ordering.
- Persisting named comparisons.
- Adding a database migration without evidence that the UI work requires one.
