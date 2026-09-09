const ENABLED_VALUES = new Set(["1", "true", "yes", "on"]);

export function creatorStartupHydrationEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.CREATOR_STARTUP_HYDRATION_ENABLED || "").trim().toLowerCase());
}

export function creatorDirectoryEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.CREATOR_DIRECTORY_ENABLED || "").trim().toLowerCase());
}
