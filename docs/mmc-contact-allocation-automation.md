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

## Inserted SSU Intern row — 9 October 2026

The current Flow uses Recurrence → Run script from SharePoint library → HTTP
(JSON extract). Its existing extraction script was updated in place, retaining
its identity and the existing five-minute schedule. No whole-workbook upload
or additional database request is introduced.

The script reads only A1:I64 and discovers the PAEDIATRIC EMERGENCY and subsequent
ADULTS headings. It includes inserted adult rows (SSU Intern is now row 28),
keeps shifted Paediatric doctor rows in MCH, and stops before nursing/service
tables. Missing section boundaries fail explicitly. ROLE/NAME/PHONE and shift
headings are excluded. Phone cells retain slash-separated numbers and instructions
as entered; the app treats each number separately for handset conflicts.

At 10:49 the successful Flow extract already contained Maria alone and
25144/25187, but omitted Mary because the old adult range ended at row 27.
The older Maria/Mary combined entry was therefore not the newest successful
extract. UI refresh and publication timing can temporarily retain an older view;
this observation does not prove a Flow rollback.
