import assert from "node:assert/strict";
import { existsSync } from "node:fs";

import {
  LOCAL_CREATOR_EMAIL,
  LOCAL_CREATOR_PASSWORD,
  LOCAL_STATE_DIRECTORY,
  LOCAL_USER_EMAIL,
  LOCAL_USER_PASSWORD,
  localLogin,
  migrateLocalDatabase,
  resetLocalState,
  seedLocalDatabase,
  startLocalServer,
  stopLocalServer,
  verifyLocalLogin,
  waitForLocalServer,
} from "./local-dev.mjs";

const port = 8799;

async function withLocalServer(callback) {
  const child = await startLocalServer({ port, capture: true });
  try {
    await waitForLocalServer(child);
    await callback();
  } finally {
    await stopLocalServer(child);
  }
}

await resetLocalState({ port });
assert.equal(existsSync(LOCAL_STATE_DIRECTORY), true, "reset should recreate only the local-safe directory");
await migrateLocalDatabase();
await seedLocalDatabase();
await seedLocalDatabase();

await withLocalServer(async () => {
  const creator = await verifyLocalLogin(port);
  assert.equal(creator.viewedAccountId, LOCAL_CREATOR_EMAIL);
  assert.equal(creator.role, "creator");

  const user = await localLogin(port, { email: LOCAL_USER_EMAIL, password: LOCAL_USER_PASSWORD });
  assert.equal(user.response.status, 200, user.payload?.error || "seeded user should log in");
  assert.equal(user.payload.viewedAccountId, LOCAL_USER_EMAIL);

  const wrongPassword = await localLogin(port, { email: LOCAL_CREATOR_EMAIL, password: `${LOCAL_CREATOR_PASSWORD}-wrong` });
  assert.equal(wrongPassword.response.status, 401, "wrong local password should be rejected");
});

await withLocalServer(async () => {
  const restarted = await verifyLocalLogin(port);
  assert.equal(restarted.viewedAccountId, LOCAL_CREATOR_EMAIL, "Creator login should survive a local runtime restart");
});

console.log("Production-isolated local lifecycle passed reset, migration, idempotent seed, login, and restart checks.");
