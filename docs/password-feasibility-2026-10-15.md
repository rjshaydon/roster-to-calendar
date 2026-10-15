# Native password feasibility — 15 October 2026

The approved feasibility gate found no acceptable local Free-tier implementation among the tested settings. Product/session choices are approved, but authentication activation needs a hosting decision. Production code and data are unchanged.

## Password history

Git commit `a9ba6e67` (23 April 2026, “Add password account isolation”) introduced `SHA-256(salt + ':' + password)` in `functions/api/state.js`. The same algorithm remains at this audit. It predates the September D1 incident; the CPU restoration work did not replace a stronger password algorithm with SHA-256.

## Native experiment

Temporary Worker `roster-password-benchmark-oct15`, version `15cb1d32-98b6-4d51-92f3-7d8a866d104f`, had no D1, R2 or production bindings. It accepted only authenticated fixed synthetic tests; no user passwords, account emails or rosters were supplied. It was successfully deleted after testing. No app dependency was added.

| Test | Parameters | Result |
| --- | --- | --- |
| Argon2id, hash-wasm 4.12.0 | 19,456 KiB, two passes, parallelism one, 32-byte output | All three responses matched the local reference; two captured native traces used **142 ms and 124 ms CPU**. Above the required 10 ms budget. |
| Native PBKDF2 / SHA-256 | 600,000 iterations, 32-byte output | All three rejected: “iteration counts above 100000 are not supported”. |
| Native PBKDF2 / SHA-512 | 220,000 iterations, 32-byte output | All three rejected by the same runtime limit. |
| Session cryptographic control | Random 32-byte token and SHA-256 digest | Three correct responses; captured CPU **1, 0 and 0 ms**. Excludes D1/account/authorization/startup work. |

Workers disallow compiling arbitrary WASM during a request. The benchmark adapted only hash-wasm's module loading to static deployment imports of its existing Argon2 and BLAKE2b binaries. Hashing code, parameters and memory work were unchanged. The package license was retained in the temporary build.

Limitations: the tail captured 11/12 invocations, including only two Argon2id CPU timings; successful responses cannot stand in for missing CPU traces. In-Worker `Date.now()` stayed unchanged during CPU-only work, so native trace `cpuTime` is authoritative here. HTTP 200 does not prove Free-tier CPU reliability. These results reject the tested paths, not every possible implementation, and do not establish that the whole eventual sign-in route fits any budget.

Sanitized response/trace metrics and provenance: `password-feasibility-native-2026-10-15.json`. Raw trace headers/IPs and the temporary bearer key are not included. Reproducer: `scripts/security/prepare-password-feasibility.mjs`; it builds temporary artifacts only and never deploys, reads app accounts or changes dependencies. It requires an unpacked `hash-wasm@4.12.0` package, obtained with lifecycle scripts disabled. Capture native traces during a manual isolated deployment and remove that Worker afterwards. Do not use real passwords in this experiment.

## Decision options

1. **Remain on Cloudflare Free; evaluate external password authentication.** Supabase currently advertises 50,000 monthly active users on Free and documents bcrypt password storage. Roster/identity/calendar data can remain in D1/R2; the provider would hold authentication emails, password hashes and associated authentication metadata. Login network/provider availability, recovery/email configuration, rate limiting, migration and revocation integration require verification before activation. Its Free projects may pause after a week without provider activity. Long-lived app sessions can mean few provider calls even while calendars are used, so this is a real availability consideration. Do not manufacture keepalive traffic to evade an inactivity policy. No account or real credentials have been sent to Supabase.
2. **Remain on Cloudflare Free; investigate passkeys.** Public-key authentication avoids the expensive password KDF on this server. Requires a real enrollment/recovery design and Safari/iPhone/Citrix checks. A retained password fallback would still need appropriate storage and would not solve the tested KDF constraint. This changes login and is a larger product choice.
3. **Explicitly reconsider Workers Paid.** The documented minimum is US$5/month, with usage charges beyond included allowances. The tested Argon2id cost has ample space compared with paid per-request CPU allowances, although full-route testing and abuse controls remain necessary. This keeps authentication local and avoids a new auth provider; existing D1 guards and changed-only/cache work should stay. The user's existing Free preference remains in force unless explicitly changed.

Clerk was also reviewed: its Free plan currently fixes session lifetime at seven days. That conflicts with the approved remembered-device experience. Do not bypass a provider's session restrictions with an unreviewed custom flow or silently purchase its paid plan.

An external-provider migration must specify what happens to inactive legacy password hashes as well as active users. Hashes cannot be converted back into passwords. Successful reauthentication can upgrade credentials; any migration deadline, old-hash retirement or recovery requirement needs a concrete reviewed design. Do not invent cryptography, bulk-export plaintext passwords, send password work to GitHub Actions or reduce parameters to fit Free CPU.

## Next action

The user has been asked to choose the next architecture direction. Continue native auth/session/request integration once the selected path is reviewable and authorized; the rest of the approved Batch 6 product plan remains in place. Do not deploy a partial authentication migration with weak new password storage merely to make progress.

## Primary references

- [OWASP password storage](https://cheatsheetseries.owasp.org/cheatsheets/Password_Storage_Cheat_Sheet.html)
- [hash-wasm maintained source](https://github.com/Daninet/hash-wasm)
- [Cloudflare Workers limits](https://developers.cloudflare.com/workers/platform/limits/)
- [Cloudflare Workers pricing](https://developers.cloudflare.com/workers/platform/pricing/)
- [Supabase pricing and inactivity policy](https://supabase.com/pricing)
- [Supabase password storage](https://supabase.com/docs/guides/auth/password-security)
- [Clerk pricing and session restrictions](https://clerk.com/pricing)
