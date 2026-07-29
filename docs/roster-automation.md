# Roster automation setup

The app now has a token-protected automation ingress at:

```text
POST https://<your-pages-domain>/api/automation/ingest
```

It accepts the existing roster spreadsheets, retains the original file in R2, and queues it for the background processor. The scheduled GitHub Action parses the workbook with the same code as a Creator upload and sends the derived rows back to D1 in small batches. This keeps spreadsheet parsing outside the Cloudflare Pages Function CPU limit. The SharePoint filename and provider version are checked before hashing or storing another copy, so an unchanged file is not reparsed even if Power Automate sends it again. Content hashes provide a fallback when a connector does not supply a provider version.

Admin → Files shows the last provider modification time supplied by the connector and the corresponding successful import time. An automated source is never labelled current before the first successful source update; it remains **Not connected** or **Waiting for first source update**.

Supported `sourceId` values are:

| Source | `sourceId` | Parsed calendar source |
| --- | --- | --- |
| Monash Adults | `monash-adults` | MMC |
| Monash Paediatrics | `monash-paeds` | MCH |
| Dandenong Findmyshift | `dandenong-findmyshift` | DDH |

Casey deliberately has no automated source. It remains a Creator-only manual upload.

## Secure Cloudflare configuration

Create a long random value and save it as a Pages secret named `ROSTER_AUTOMATION_TOKEN`. Do not save SharePoint or Findmyshift passwords in D1, R2, the browser, or this repository.

```bash
npx wrangler pages secret put ROSTER_AUTOMATION_TOKEN
```

Save the same value as the encrypted GitHub Actions repository secret `ROSTER_AUTOMATION_TOKEN`. The `Process Monash roster queue` workflow checks at minutes 7, 22, 37 and 52 and can also be dispatched manually. The offset avoids GitHub's start-of-hour scheduling peak. When no changed file is queued, the job exits before checking out the repository or installing dependencies. The Pages Function only authenticates and stores changed incoming files; opening the website is not required.

Apply the new D1 migration before deploying the Pages Function:

```bash
npx wrangler d1 migrations apply roster-converter-calendar --remote
```

The token is sent only as:

```text
Authorization: Bearer <ROSTER_AUTOMATION_TOKEN>
```

## Monash SharePoint / Power Automate

Create a Power Automate flow in the Monash tenant. The signed-in Monash account and its MFA remain in Microsoft 365; the app receives only the roster file.

1. Use **SharePoint – When a file is created or modified (properties only)** for the `Medical Roster` folder.
2. Filter to the two approved filenames/prefixes and spreadsheet extensions. Use separate branches:
   - Adult roster: `sourceId` = `monash-adults`
   - Paediatric roster: `sourceId` = `monash-paeds`
3. Use **Get file content** for the triggering file.
4. Use the tenant-approved HTTP action to POST JSON to the ingress URL. A successful new upload returns HTTP `202` with `status: queued`. This connector is often Premium and may be blocked by DLP; confirm that with the Monash Power Platform administrator before enabling it.

Use this JSON body (replace the dynamic-content fields with the corresponding Power Automate values):

```json
{
  "sourceId": "monash-adults",
  "fileName": "@{triggerOutputs()?['body/{FilenameWithExtension}']}",
  "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "contentBase64": "@{body('Get_file_content')?['$content']}",
  "providerModifiedAt": "@{triggerOutputs()?['body/Modified']}",
  "providerVersion": "@{triggerOutputs()?['body/ETag']}"
}
```

Set the request header `Authorization` to `Bearer <the Cloudflare secret>`. Limit the flow to roster-folder maintainers and turn on failure notifications. Test by updating a copy first; the ingress rejects a file whose detected roster type does not match its source id.

The `providerVersion` value must be the SharePoint file ETag/version, not the flow run time. It must remain unchanged until that specific file changes. Filename plus provider version is the primary change identity, allowing different term files to have the same version number without being confused with one another.

## Advanced roster recovery

Normal source updates never require a full rebuild. In Admin → System, **Advanced recovery** is reserved for a confirmed corruption of the derived roster database while every retained source file is known to be correct. Recovery is blocked while an automated update is queued or processing and requires an explicit `REBUILD` confirmation. Prefer the per-file reparse control when only one roster is affected.

## Dandenong / Findmyshift

The ingress is ready for a full Findmyshift spreadsheet export under `dandenong-findmyshift`; it uses the existing DDH parser and does not truncate the roster. The remaining connection step is tenant-specific: configure a scheduled Findmyshift API/export job to retrieve the full roster and POST it to this endpoint as multipart (`rosterFile`) or the JSON format above.

Store the Findmyshift API key or OAuth refresh token as a secret in that scheduler (or as a Cloudflare Worker secret if a Worker is used), never as a user-visible app setting. Before enabling it, verify the export/API includes every stream, staff member, leave entry, and the required date horizon; the currently available browser roster is full-DDH scope, but no API credential was exposed during implementation.

## Manual access policy

The roster import button, drag-and-drop import, and the two persistent import actions are Creator-only. This is enforced in both the browser and `/api/state`; ordinary users cannot bypass it by calling the API directly.
