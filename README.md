# Roster Converter

Cloudflare Pages application for converting MMC and DDH roster exports into an Apple/Google Calendar-compatible `.ics` file.

## Architecture

- `public/`: static frontend served by Cloudflare Pages
- `public/static/roster.js`: shared browser-side roster parsing and calendar generation logic
- `functions/api/`: Cloudflare Pages Functions for account, repository, and fallback calendar endpoints
- `functions/_lib/roster.js`: server fallback copy of roster parsing and `.ics` generation logic
- `fixtures/`: sample MMC/DDH spreadsheets and MMC PDF exports for parser regression checks
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
5. Create a Workers KV namespace, for example `roster-converter-state`.
6. In the Pages project, add a KV namespace binding:

```text
Variable name: ROSTER_STORE
KV namespace: roster-converter-state
```

7. Redeploy the Pages project after adding the binding.
8. Add your custom domain in Cloudflare Pages once the deploy succeeds.

Without `ROSTER_STORE`, the account and shared repository features will not run. Roster repository data, user accounts, and persistence must be server-side for the deployed app.

## CLI Deploy

After authenticating Wrangler:

```bash
npm run deploy
```

## Notes

- The browser auto-detects MMC Excel/PDF uploads and DDH FindMyShift spreadsheet exports, then saves source files and parsed metadata to Cloudflare storage.
- If only one consultant is detected, the UI shows the doctor name directly.
- Preview renders a Monday-start weekly grid before export.
- Users log in with an email address.
- `rhaydon@gmail.com` is the Creator account and is bootstrapped on first server-backed login if the KV store is empty.
- Creator storage is unrestricted.
- Standard accounts are prompted to keep only the latest 6 months active.
- Cross-device persistence requires the `ROSTER_STORE` KV binding above.
- Claimed accounts can expose a tokenized subscription feed at `/api/feed?token=...`, which Apple Calendar or Google Calendar can subscribe to as a read-only `.ics` URL.
