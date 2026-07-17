# Roster Converter

Cloudflare Pages application for converting MMC, DDH, and Casey roster exports into an Apple/Google Calendar-compatible `.ics` file.

## Architecture

- `public/`: static frontend served by Cloudflare Pages
- `public/static/roster.js`: shared browser-side roster parsing and calendar generation logic
- `functions/api/`: Cloudflare Pages Functions for account, repository, and fallback calendar endpoints
- `functions/_lib/roster.js`: server fallback copy of roster parsing and `.ics` generation logic
- `fixtures/`: sample MMC/DDH/Casey spreadsheets and MMC PDF exports for parser regression checks
- `scripts/test-fixtures.mjs`: fixture smoke test
- `docs/PRD.md`: consolidated product rules

## Local Development

Install dependencies:

```bash
npm install
```

Run the parser smoke test:

```bash
npm run test:fixtures
```

Run the Pages app locally:

```bash
npm run dev
```

Wrangler will print the local URL, usually `http://127.0.0.1:8788`.

## Cloudflare Deployment

1. Push this repo to GitHub.
2. In Cloudflare, create a new Pages project and connect the GitHub repo.
3. Use these build settings:

```text
Framework preset: None
Build command: exit 0
Build output directory: public
Root directory: /
```

4. Cloudflare Pages will pick up the `/functions` directory automatically.
5. Create or attach the D1 database used by the app.
6. In the Pages project, add a D1 binding:

```text
Variable name: ROSTER_DB
D1 database: roster-converter-calendar
```

7. Redeploy the Pages project after adding the binding.
8. Add your custom domain in Cloudflare Pages once the deploy succeeds.

Without `ROSTER_DB`, account login, calendar loading, roster-derived events, coworker lookups, and subscription feeds will not run. KV is no longer a runtime store; after the D1-only cutover, creator roster files must be re-uploaded once so parsed rows can be persisted to D1.

## CLI Deploy

After authenticating Wrangler:

```bash
npm run deploy
```

## Notes

- The browser auto-detects MMC Excel/PDF uploads, DDH FindMyShift spreadsheet exports, and Casey weekly roster workbooks, then saves parsed roster metadata and events to D1.
- If only one consultant is detected, the UI shows the doctor name directly.
- Preview renders a Monday-start weekly grid before export.
- Users log in with an email address.
- `rhaydon@gmail.com` is the Creator account and is bootstrapped on first D1-backed login.
- Creator storage is unrestricted.
- Cross-device persistence requires the `ROSTER_DB` D1 binding above.
- Accounts with saved roster data can expose a tokenized subscription feed at `/api/feed?token=...`, which Apple Calendar or Google Calendar can subscribe to as a read-only `.ics` URL.
- The **When am I working with…?** insight excludes Clinical Support overlaps by default; tick **Include CS** in the modal Options to show them.
