# At a glance — By stream implementation handoff

## Status

Planning is complete for implementation on `codex/by-stream-at-a-glance`.

This handoff adds a fourth At a glance view, **By stream**, between **On shift** and **ED staff**. It also replaces the current first-available/last-used ED behaviour with a shared, deterministic preferred-ED resolver when At a glance is entered from **My calendar**.

No feature code is included in this planning change.

## Confirmed product decisions

- Tab order: **On shift**, **By stream**, **ED staff**, **Working together**.
- By stream defaults to a one-day range where both bounds are today in `Australia/Melbourne`.
- The **Today** button sets both range bounds to today in one action.
- Date controls sit across the top of the view.
- Stream-selection rows are vertically stacked in a right-hand rail on desktop.
- On narrow layouts, the selector rail moves above the results.
- A selector row contains **ED**, **Stream**, and **Seniority**.
- Seniority defaults to **SMS** and offers **All team** plus the canonical seniorities observed for that ED/stream.
- Users can add, remove, and compare multiple selector rows.
- New user-facing code should say **ED**, following the existing At a glance terminology, even where the product discussion used “hospital”.
- Stream and seniority choices must be derived from active roster data and the same normalisation rules used by **On shift**. Do not maintain a separate hard-coded universal stream list.

## Entry and state rules

There are two kinds of navigation into At a glance:

1. **Normal entry from My calendar**
   - Recompute the preferred ED for the active doctor/account.
   - On shift opens on today at the preferred ED.
   - ED staff opens on the current Australian medical term at the preferred ED.
   - The first By stream selector row uses the preferred ED.
   - By stream initially uses today for both date bounds.

2. **Explicit internal navigation**
   - A drill-through that supplies a tab, ED, date, stream, or person takes priority over defaults.
   - Examples include opening On shift from a Working together result or a future deep link.
   - Do not run the preferred-ED resolver over an explicit destination.

While the workspace remains open, preserve deliberate user selections when switching tabs. Recompute defaults only on normal entry from My calendar, when changing the active doctor/account, or when the active roster revision invalidates the selected ED.

## Preferred ED resolver

### Purpose

The resolver supplies one preferred ED for the active doctor. It must use active roster evidence, not the order in which facilities happen to appear in a client-side array.

### Inputs

- All canonical and alias `doctorKey` values for the active doctor.
- The current instant and date in `Australia/Melbourne`.
- Active roster files only.
- Recognised working events, including Clinical Support for the purpose of establishing where a person works.
- Current-term membership as a fallback when no recognised working event exists.

Leave, PHNW, public-holiday, hidden, unresolved, and other non-working records do not establish a preferred ED.

### Definitions

- **Current term:** the existing 13-week Australian medical term containing today.
- **Current week:** Monday through Sunday containing today, in `Australia/Melbourne`.
- **Current shift:** a timed working event where `start <= now < end`; this includes an overnight shift that began yesterday.
- **Today’s shift:** an event assigned to today by the existing roster-date semantics.
- **Next shift:** the earliest recognised working event whose start is after now.

### Resolution order

1. Find distinct EDs with recognised working events for the doctor during the current term.
2. If there is exactly one current-term ED, return it.
3. If there are multiple current-term EDs, find the distinct EDs worked during the current week.
4. If there is exactly one current-week ED, return it.
5. If there are multiple current-week EDs:
   1. Prefer the ED of a shift active at the current instant.
   2. Otherwise prefer the ED of today’s next not-yet-started shift.
   3. Otherwise, if all of today’s shifts are complete, prefer the ED of today’s most recently completed shift.
6. If the preceding steps do not decide, return the ED of the next future shift in the current term.
7. If there is no future shift, return the ED of the most recent past working shift in the current term.
8. If there are no recognised current-term working events:
   1. Return the sole current-term membership ED when membership identifies exactly one.
   2. Otherwise return the sole linked ED for the active doctor/account.
   3. Otherwise reuse the last manually selected At a glance ED only when it remains linked and has current-term coverage.
   4. Finally, use the first linked ED with current-term coverage in the established ED display order.

If data contains simultaneous shifts at different EDs, choose deterministically by earliest start, then the existing ED display order. This is a data-quality edge case; it must not make the default unstable.

### Result shape

Return enough information for diagnostics without exposing it in normal UI:

```js
{
  facilityKey: "mmc",
  reason: "sole-current-week-facility",
  evidenceDate: "2026-08-11"
}
```

Suggested reason values:

- `sole-current-term-facility`
- `sole-current-week-facility`
- `active-shift`
- `today-next-shift`
- `today-last-shift`
- `next-term-shift`
- `last-term-shift`
- `sole-term-membership`
- `sole-linked-facility`
- `last-valid-selection`
- `first-covered-linked-facility`

### Implementation placement

Implement the resolver against D1 data rather than the filtered calendar preview. The preview can omit data because of UI filters and should not define a product default.

Add a small authorised API action such as `queryFacilityOverviewDefaults`, or include this result in a new shared At a glance metadata response. Cache it for the active doctor, Melbourne date, and roster revision. Invalidate it after roster activation/replacement, doctor/account switching, or a date rollover.

## By stream controls and state

Extend `facilityOverviewState` with isolated fields rather than reusing the single-date On shift fields:

```js
byStreamFrom: today,
byStreamTo: today,
byStreamRows: [
  {
    id: stableClientId,
    facilityKey: preferredFacilityKey,
    streamKey: preferredStreamKey,
    seniority: "SMS"
  }
],
byStreamCatalog: null,
byStreamContent: "",
byStreamData: null,
byStreamLoading: false,
byStreamRequestId: 0,
byStreamHideEmptyDates: true
```

Do not share the general `requestId` if doing so would let an unrelated staff/menu refresh cancel a By stream request. A dedicated request sequence or `AbortController` is preferred.

### Initial stream

Choose an initial stream at the preferred ED in this order:

1. The active doctor’s current or next meaningful stream at that ED, using the same assignment classifier as On shift.
2. The first observed meaningful stream in the hospital-specific `whoTeamRank` order.
3. If no meaningful stream is available, leave the stream field unselected and show `Choose stream…`; do not invent a stream.

Seniority remains SMS even when there is no SMS match in the initial date range. The empty state should then explain that no SMS assignment was found rather than silently changing the filter.

### Selector row behaviour

- Present **ED** before **Stream**, because stream options depend on ED.
- Changing ED clears an invalid stream and selects the first valid stream using the initial-stream rule.
- Changing ED or stream refreshes the available seniorities.
- Seniority options are based on active observed coverage for the ED/stream, not only the selected dates.
- Always include `All team` and `SMS`; then add observed canonical grades in `FACILITY_OVERVIEW_SENIORITY_ORDER`.
- Permit the same ED/stream more than once when seniority differs.
- Block exact duplicate ED/stream/seniority rows and announce the validation error.
- Every row after the first has a labelled Remove button. If only one row remains, it cannot be removed.
- Initially allow up to six rows. Disable **Add another stream** at the limit with a short explanation.

### Date controls

- Use two native date inputs labelled **From** and **To**.
- Both default to the Melbourne date for today.
- **Today** sets both bounds to today and triggers one refresh.
- Editing From past To moves To forward to match From.
- Editing To before From moves From backward to match To.
- Submit only complete valid ranges.
- Use an inclusive date range and existing roster-date semantics, so a preceding day’s PM shift ending at midnight is not attributed to the next day.
- Match the existing Working together API bound of at most 370 inclusive days unless measured performance requires a smaller documented limit.

## Results layout

### Desktop

- Top toolbar: From, To, Today.
- Main area: `minmax(0, 1fr)` results plus a roughly 260–320 px selector rail.
- Keep the selector rail visible with sticky positioning only when the containing workspace can do so without nested-scroll or clipping problems.
- Results are date-first. Each date section contains one result card per selector row in selector order.
- With one selector, use the available width rather than rendering a narrow column.
- With multiple selectors, use a responsive grid; do not force horizontal timeline scrolling for long date ranges.

### Narrow screens

- Move selectors above results.
- Keep date controls above selectors.
- Stack result cards within each date.
- Avoid a permanently sticky selector region on mobile.

### Assignment content

Each result card shows:

- ED and stream
- Active seniority filter
- Staff display name
- Canonical seniority when `All team` is selected
- Exact start/end time
- Relevant in-charge/IC marker already derived by the assignment classifier
- Multiple non-duplicate assignments when responsibility changes during the day

Sort people by shift start, seniority rank, then display name. Deduplicate equivalent active-roster events before rendering.

`Hide dates without assignments` defaults on for multi-day ranges. When off, retain dates with explicit empty or uncovered states. Even when hidden, show a summary such as `8 dates shown · 5 dates without a match hidden`.

## Coverage and empty states

The UI must distinguish:

1. No active roster covers the ED/date.
2. Coverage exists, but the stream was not observed.
3. The stream exists, but nobody matched the selected seniority.
4. Matching assignments exist.
5. The service is unavailable.

Coverage is calculated per ED and date, not inferred from the presence of a match. A missing result must never imply that nobody was rostered when source coverage is absent.

## Stream catalogue and canonicalisation

The current On shift UI derives assignments with `buildWhoAssignment`, `whoTeamLabel`, `whoDisplayTeamLabel`, parser-rule overrides, and `facilityOverviewIsMeaningfulStream`. By stream must produce the same stream label and key for the same event.

Before building the range query, extract or otherwise centralise the canonical pieces required by both views. Do not create a second set of regexes that can drift from On shift.

A catalogue entry should contain:

```js
{
  facilityKey: "mmc",
  streamKey: "resus",
  label: "Resus",
  seniorities: ["SMS", "Senior Registrar", "HMO"],
  firstSeenDate: "2026-02-02",
  lastSeenDate: "2026-11-01"
}
```

The catalogue should include meaningful streams observed in active roster events. Do not include `Other`, generic `AM/PM/Night`, grade names, Float/Rover, or Clinical Support unless the existing meaningful-stream contract is deliberately changed for both On shift and By stream.

## API and query contract

For minimal architectural disruption, add actions to the existing authenticated `/api/state` flow used by the other At a glance views. A separate overview endpoint can be a later refactor.

### Metadata/defaults action

Request:

```js
{
  action: "queryFacilityOverviewMetadata",
  doctorKeys: ["..."],
  today: "2026-08-11"
}
```

Response:

```js
{
  ok: true,
  preferredFacility: { facilityKey, reason, evidenceDate },
  facilities: [{ facilityKey, label, coverage }],
  streams: [{ facilityKey, streamKey, label, seniorities, firstSeenDate, lastSeenDate }]
}
```

The server should derive today using `Australia/Melbourne`; the client date is advisory only.

### By stream range action

Request:

```js
{
  action: "queryFacilityOverviewByStream",
  startDate: "2026-08-11",
  endDate: "2026-08-13",
  selections: [
    { id: "row-1", facilityKey: "mmc", streamKey: "resus", seniority: "SMS" }
  ]
}
```

Response:

```js
{
  ok: true,
  startDate,
  endDate,
  coverage: [{ facilityKey, date, covered, partial }],
  selections: [
    {
      id: "row-1",
      facilityKey: "mmc",
      streamKey: "resus",
      seniority: "SMS",
      assignments: [{ date, doctorKey, displayName, seniority, start, end, roleNote }]
    }
  ],
  queryMs
}
```

Return explicit DTOs rather than raw `event_json` where practical.

### Validation and query strategy

- Require At a glance permission exactly as the existing actions do.
- Accept 1–6 UI selections; reject more than 8 server-side.
- Require canonical facilities, stream keys, and seniorities.
- Reject invalid or reversed ranges and ranges longer than 370 inclusive days.
- Deduplicate requested selections server-side.
- Query all required facilities/range in one bounded D1 operation, then classify and match selections once. Do not issue one database/API request per row per day.
- Query only `roster_files.active = 1`.
- Use `roster_events.start_date` for date attribution, consistent with On shift.
- Reuse seniority overrides for the applicable term. A range can cross term boundaries, so resolve the override against each assignment’s term rather than only the range start.
- Filter with the existing recognised-working-event contract before stream matching.
- Record `queryMs` and log slow queries without including names or credentials.

The existing `(source_type, start_date, end_date)` index is a suitable starting point. Measure representative term and one-year ranges before adding a new migration.

## Frontend integration points

Primary files:

- `public/index.html`: insert the tab between On shift and ED staff.
- `public/static/app.js`: state, event handling, default resolution, loading, rendering, and extracted shared classification helpers.
- `public/static/styles.css`: responsive toolbar, selector rail, selector rows, date sections, and cards.
- `functions/api/state.js`: authorisation, validation, metadata/defaults action, and range action.
- `functions/_lib/d1-calendar.js`: preferred-ED and range queries.

Update `openFacilityOverview` so tab dispatch is explicit:

```js
if (tab === "staff") loadFacilityOverviewStaff();
else if (tab === "by-stream") loadFacilityOverviewByStream();
else if (tab === "together") loadFacilityOverviewTogether();
else loadFacilityOverviewOnShift();
```

Avoid the current pattern where every non-staff tab can fall through to On shift.

## Accessibility

- Keep the existing `role="tablist"`, `role="tab"`, `aria-selected` contract.
- Add a matching tab panel relationship if practical while touching the tabs.
- Every selector row should be a fieldset or labelled group such as `Stream selection 2`.
- Remove buttons need row-specific accessible names.
- Announce loading, duplicate-row errors, hidden empty-date counts, and result changes through the existing live region.
- Preserve native keyboard behaviour for date inputs and selects.
- Do not encode ED, stream, seniority, or coverage solely by colour.

## Test plan

### Preferred ED resolver

- One ED in current term -> that ED.
- Multiple term EDs, one ED this week -> this week’s ED.
- Multiple EDs this week, active shift now -> active shift ED.
- Multiple EDs this week, no active shift, shift later today -> later-today ED.
- Multiple completed shifts today -> most recently completed ED.
- No shift today -> next future shift ED.
- No future term shift -> most recent past term shift ED.
- Overnight shift begun yesterday -> active overnight ED.
- No working events, one membership ED -> membership ED.
- No events/membership, one linked ED -> linked ED.
- Simultaneous conflicting ED shifts -> deterministic tie-break.
- Normal entry recomputes; tab switching preserves manual choices; explicit drill-through wins.

### By stream

- Tab appears in the required order and remains permission-gated.
- Initial dates are both today; Today resets both after either is changed.
- Initial ED comes from the shared resolver.
- Initial stream follows current/next assignment, then canonical first-stream fallback.
- Seniority defaults to SMS.
- ED changes update streams and seniorities without stale responses winning.
- Exact duplicates are rejected; same stream with another seniority is allowed.
- Add/remove works up to the documented limit.
- One request covers multiple rows and dates.
- Term-crossing ranges apply the correct seniority override to each date.
- Midnight-ending PM shifts remain on their roster date.
- Empty match and missing coverage render differently.
- Hide-empty-dates summary is accurate.
- MMC, DDH, Casey, and MCH stream labels match On shift for the same fixtures.
- Mobile layout places selectors above results with no horizontal page overflow.

Run at minimum:

```sh
npm run check
npm run test:fixtures
```

Then smoke-test the feature against active representative data for all four ED sources.

## Suggested implementation sequence for Terra

1. Extract/test shared working-event, seniority, and stream canonicalisation without changing existing output.
2. Implement the preferred-ED resolver and apply it on normal At a glance entry.
3. Add the By stream tab, state, and static responsive shell.
4. Add metadata/catalogue loading and initial selector resolution.
5. Add the bounded range action/query and render real results.
6. Add coverage/empty/error states and hide-empty-dates behaviour.
7. Complete accessibility, stale-request protection, validation, and mobile refinement.
8. Run syntax/fixture tests and perform four-source smoke testing.

## Non-goals for this change

- Exporting or sharing By stream results.
- Persisting named comparison presets.
- Reordering selector rows by drag and drop.
- Introducing a new client router.
- Moving all existing At a glance actions into a new endpoint.
- Adding a new roster-events index before query timing demonstrates a need.
