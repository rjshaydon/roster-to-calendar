# Roster Converter

Cloudflare Pages application for converting MMC and DDH roster exports into an Apple Calendar-compatible `.ics` file.

## Architecture

- `public/`: static frontend served by Cloudflare Pages
- `functions/api/`: Cloudflare Pages Functions for analyze, preview, and export
- `functions/_lib/roster.js`: shared roster parsing and `.ics` generation logic
- `docs/PRD.md`: consolidated product rules

The older Python prototype remains in the repo as a reference, but the long-term deployment target is Cloudflare Pages.

## Local Development

Install dependencies:

```bash
npm install
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
Build command: npm install
Build output directory: public
Root directory: /
```

4. Cloudflare Pages will pick up the `/functions` directory automatically.
5. Add your custom domain in Cloudflare Pages once the first deploy succeeds.

## CLI Deploy

After authenticating Wrangler:

```bash
npm run deploy
```

## Notes

- The frontend accepts up to two roster files.
- The backend auto-detects MMC vs DDH uploads.
- If only one consultant is detected, the UI shows the doctor name directly.
- Preview renders a Monday-start weekly grid before export.
