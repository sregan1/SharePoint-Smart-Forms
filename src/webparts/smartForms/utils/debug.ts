/**
 * Console diagnostics, matching the convention used across the Smart* web
 * parts (see SharePointSmartPermissionsWebPart's userAccess.ts / App.tsx).
 *
 * `logError` always prints — it is the one thing a form owner needs to see
 * when something silently fails (a load, a publish, a submit) instead of just
 * a generic on-screen message. `debugLog` is opt-in and verbose: it traces
 * routine lifecycle events (list load, publish, provisioning) that would spam
 * every user's console if left on by default, so it only fires once
 * `localStorage.setItem('smartFormsDebug', '1')` has been set — handy when
 * asking someone to reproduce an issue and send back their console output.
 */

const PREFIX = '[SmartForms]';

const isDebugEnabled = (): boolean => {
  try {
    return window.localStorage && window.localStorage.getItem('smartFormsDebug') === '1';
  } catch {
    // localStorage unavailable (e.g. private browsing, or a locked-down host)
    return false;
  }
};

/** Verbose tracing — silent unless smartFormsDebug is enabled in localStorage. */
export function debugLog(...args: unknown[]): void {
  try {
    if (isDebugEnabled()) {
      // eslint-disable-next-line no-console
      console.debug(PREFIX, ...args);
    }
  } catch {
    // logging must never be the thing that throws
  }
}

/** Always-on error reporting, so a swallowed exception still leaves a trace. */
export function logError(context: string, error: unknown): void {
  // eslint-disable-next-line no-console
  console.error(PREFIX, context + ' failed:', error);
}

/** Always-on warning for a non-fatal problem worth knowing about. */
export function logWarning(context: string, detail: unknown): void {
  // eslint-disable-next-line no-console
  console.warn(PREFIX, context + ':', detail);
}
