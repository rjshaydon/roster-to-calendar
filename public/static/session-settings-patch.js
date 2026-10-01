const ROOTS = new Set(['doctorKey', 'settings', 'exportRange', 'overrides', 'customEvents', 'conflictSelections', 'hadPreview']);
const UNSAFE = new Set(['__proto__', 'constructor', 'prototype']);
export function sessionValueKey(value) {
  if (Array.isArray(value)) return JSON.stringify(value.map(item => JSON.parse(sessionValueKey(item))));
  if (value && typeof value === 'object') return JSON.stringify(Object.fromEntries(Object.keys(value).sort().map(key => [key, JSON.parse(sessionValueKey(value[key]))])));
  return JSON.stringify(value ?? null);
}
const valueKey = (value, path) => sessionValueKey(path[0] === 'customEvents' && Array.isArray(value) ? [...value].sort((a,b) => String(a.id).localeCompare(String(b.id))) : value);
const object = value => value && typeof value === 'object' && !Array.isArray(value);
export function makeSessionPatch(before = {}, after = {}) {
  before = { ...before, customEvents: before.customEvents || [] };
  after = { ...after, customEvents: after.customEvents || [] };
  const changes = [];
  const visit = (left, right, path, leftExists, rightExists) => {
    if (leftExists === rightExists && valueKey(left, path) === valueKey(right, path)) return;
    if (object(left) && object(right) && path.length < 4) {
      for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) visit(left[key], right[key], [...path, key], Object.hasOwn(left, key), Object.hasOwn(right, key));
    } else changes.push({ path, before: left ?? null, after: right ?? null, beforeExists: leftExists, afterExists: rightExists });
  };
  for (const key of ROOTS) visit(before[key], after[key], [key], Object.hasOwn(before, key), Object.hasOwn(after, key));
  validateSessionPatch(changes);
  return changes;
}
export function validateSessionPatch(changes) {
  if (!Array.isArray(changes) || changes.length > 256 || new TextEncoder().encode(JSON.stringify(changes)).length > 256 * 1024) throw new Error('Settings changes exceed the save limit.');
  const inspect = (value, depth = 0) => {
    if (!value || typeof value !== 'object') return;
    if (depth > 16 || Object.keys(value).some(key => UNSAFE.has(key))) throw new Error('Invalid settings change.');
    for (const item of Object.values(value)) inspect(item, depth + 1);
  };
  for (const change of changes) {
    inspect(change.before); inspect(change.after);
    if (!Array.isArray(change.path) || change.path.length < 1 || change.path.length > 4 || !ROOTS.has(change.path[0]) || change.path.some(key => typeof key !== 'string' || !key || key.length > 300 || UNSAFE.has(key)) || typeof change.beforeExists !== 'boolean' || typeof change.afterExists !== 'boolean') throw new Error('Invalid settings change.');
  }
}
export function applySessionPatch(current, changes) {
  validateSessionPatch(changes);
  const result = JSON.parse(JSON.stringify(current || {}));
  for (const change of changes) {
    let node = result;
    for (const key of change.path.slice(0, -1)) {
      if (Object.hasOwn(node, key) && !object(node[key])) {
        const error = new Error('Settings changed on another device. Reload before saving.');
        error.code = 'SESSION_SETTINGS_CONFLICT'; throw error;
      }
      if (!Object.hasOwn(node, key)) node[key] = {};
      node = node[key];
    }
    const key = change.path.at(-1);
    const exists = Object.hasOwn(node, key);
    const same = (expected, expectedExists) => exists === expectedExists && valueKey(node[key], change.path) === valueKey(expected, change.path);
    // A replay after a lost successful response is already applied.
    if (same(change.after, change.afterExists)) continue;
    if (!same(change.before, change.beforeExists)) {
      const error = new Error('These settings changed on another device. Reload before saving this change; your local edit has been retained.');
      error.code = 'SESSION_SETTINGS_CONFLICT';
      throw error;
    }
    if (change.afterExists) node[key] = JSON.parse(JSON.stringify(change.after));
    else delete node[key];
  }
  return result;
}
