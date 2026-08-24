# MMC live contact allocations

The On shift view accepts a short-lived, doctors-only extract from `SHIFT ALLOCATIONS.xlsx`.
It does not use the workbook as a staff directory and will only display an extract for its
matching roster date. The extract expires at 10:00 Melbourne time on the following day.

## Excel Office Script

Create or replace the Office Script named `Extract MMC doctor shift contacts` with
[`scripts/mmc-contact-allocations-office-script.ts`](../scripts/mmc-contact-allocations-office-script.ts).
It returns Adult and Paediatric AM/PM/Night doctor rows, excludes all NIC and nursing rows,
and treats a row with no name as unallocated even if its extension remains present.

## Power Automate request

After `Run script from SharePoint library`, send an HTTP `POST` to:

```
https://roster-to-calendar.pages.dev/api/automation/contact-list-extract
```

Headers:

```
Authorization: Bearer <ROSTER_AUTOMATION_TOKEN>
Content-Type: application/json
```

Body shape:

```json
{
  "sourceId": "mmc-shift-allocations",
  "sourceDate": "<script result sourceDate>",
  "providerModifiedAt": "<SharePoint Modified value>",
  "providerVersion": "<SharePoint version or ETag>",
  "contacts": "<script result contacts>"
}
```

The endpoint retains only one contact extract. A successful JSON import deletes the legacy
workbook object for this source.
