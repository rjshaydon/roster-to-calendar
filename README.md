# Roster Converter

Cloudflare Pages application for converting MMC and DDH roster exports into an Apple/Google Calendar-compatible `.ics` file.

## Architecture

- `public/`: static frontend served by Cloudflare Pages
- `functions/api/`: Cloudflare Pages Functions for analyze, preview, and export
- `functions/_lib/roster.js`: shared roster parsing and `.ics` generation logic
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

Without `ROSTER_STORE`, the app can still run a local browser-only fallback for development, but it is not a shared account system: roster repository data, user accounts, and persistence will not be available across browsers, devices, or other users.

## CLI Deploy

After authenticating Wrangler:

```bash
npm run deploy
```

## Notes

- The backend auto-detects MMC Excel/PDF uploads and DDH FindMyShift spreadsheet exports.
- If only one consultant is detected, the UI shows the doctor name directly.
- Preview renders a Monday-start weekly grid before export.
- Users log in with an email address.
- `rhaydon@gmail.com` is the Creator account.
- Creator storage is unrestricted.
- Standard accounts are prompted to keep only the latest 6 months active.
- Cross-device persistence requires the `ROSTER_STORE` KV binding above.
