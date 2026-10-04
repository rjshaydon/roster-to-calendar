// One in-flight refresh per visible view. A context change invalidates old work;
// failures retain the last good fingerprint so the next tick can retry.
export function createCalendarRevisionPoller({ context, eligible, read, refresh, onError = () => {}, schedule = setTimeout, cancel = clearTimeout, interval = 60000 }) {
  let timer = 0, running = false;
  const revisions = new Map();
  const update = () => {
    cancel(timer); timer = 0;
    if (context()) timer = schedule(tick, interval);
  };
  const tick = async () => {
    cancel(timer); timer = 0;
    const view = context();
    if (!view || running || !eligible()) { update(); return; }
    running = true;
    try {
      const revision = await read(view.sources);
      if (!revision || context()?.key !== view.key || !eligible()) return;
      if (revisions.get(view.key) !== revision) {
        const applied = await refresh(view);
        if (applied !== false && context()?.key === view.key) {
          revisions.set(view.key, revision);
          if (revisions.size > 32) revisions.delete(revisions.keys().next().value);
        }
      }
    } catch (error) { onError(error); }
    finally { running = false; update(); }
  };
  return { update, tick };
}
