// Mirrors the verified account inventory; the maintenance test detects drift.
export default {
  "accountId": "653047f50279c40888ef9715aa7c4c5f",
  "complete": true,
  "verifiedAt": "2026-10-18T05:55:05.195532Z",
  "evidence": "Wrangler D1 inventory and creation results on 5 October 2026; isolated Batch 2 Preview added without production data bindings. Exact synthetic-only authentication/session test databases created and deleted on 17 October 2026 are retained for current-day analytics attribution; no active caller remains. Verified 18 October synthetic passkey and request test database IDs are retained for same-day analytics attribution; their reads/writes remain included in account totals.",
  "databases": [
    {
      "id": "237d0d52-3a7c-4e02-8648-9f4dedbc1cb0",
      "name": "roster-converter-calendar",
      "environment": "production",
      "knownCallers": [
        "Pages Production and retained Production deployments: roster-to-calendar"
      ]
    },
    {
      "id": "541d7779-a43a-4603-a4ab-c95e9b4b015d",
      "name": "roster-converter-calendar-preview",
      "environment": "preview",
      "knownCallers": [
        "Retained Pages Preview deployments: roster-to-calendar"
      ]
    },
    {
      "id": "46829fee-f518-4d23-a992-d050f53c4c41",
      "name": "acem-exam-tutor-db",
      "environment": "production",
      "knownCallers": [
        "Worker: acem-exam-tutor (env.DB)"
      ]
    },
    {
      "id": "741ce534-eafe-4aa7-9305-7b7ed5661979",
      "name": "roster-identity-batch2-preview",
      "environment": "preview",
      "knownCallers": [
        "Isolated Batch 2 synthetic acceptance Preview: identity-batch2-acceptance"
      ]
    },
    {
      "id": "387e37d7-f4d3-4db6-80ff-f140f8fb4bc6",
      "name": "roster-session-native-validation",
      "environment": "temporary-deleted",
      "knownCallers": [
        "Gated synthetic session-cost probe; Worker and database deleted 17 October 2026"
      ]
    },
    {
      "id": "5c3d71d4-52c1-4eb0-9b32-0a5f2c3d9f02",
      "name": "roster-auth-live-acceptance-20261017",
      "environment": "temporary-deleted",
      "knownCallers": [
        "Gated synthetic Supabase lifecycle probe; Worker and database deleted 17 October 2026"
      ]
    },
    {
      "id": "a19cebcf-1ccd-43d5-84cd-cd105333cbdb",
      "name": "roster-passkey-native-20261018",
      "environment": "temporary-deleted",
      "knownCallers": [
        "Gated synthetic passkey proof-cost test; database and Worker deleted 18 October 2026"
      ]
    },
    {
      "id": "7b87d1be-2602-4a3e-8841-f6d525ad323d",
      "name": "roster-requests-native-20261018",
      "environment": "temporary-deleted",
      "knownCallers": [
        "Gated synthetic requests test only; Worker and database deleted 18 October 2026; no production, provider or mail bindings"
      ]
    }
  ]
};
