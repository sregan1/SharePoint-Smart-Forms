/**
 * Minimal test harness.
 *
 * Deliberately dependency-free: the project has no Jest rig, and wiring one into
 * an SPFx build is more moving parts than a suite of pure-function tests needs.
 * Tests compile with the rest of the sources via config/tsconfig.test.json and
 * run on plain Node.
 */

// The project has no @types/node (it targets the browser), and adding it just to
// set an exit code isn't worth the dependency — declare the one member used.
declare const process: { exit: (code: number) => void };

interface IFailure {
  suite: string;
  name: string;
  message: string;
}

const failures: IFailure[] = [];
let passed = 0;
let currentSuite = '';

export const suite = (name: string, body: () => void): void => {
  currentSuite = name;
  // eslint-disable-next-line no-console
  console.log('\n  ' + name);
  body();
  currentSuite = '';
};

export const test = (name: string, body: () => void): void => {
  try {
    body();
    passed++;
    // eslint-disable-next-line no-console
    console.log('    ok   ' + name);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    failures.push({ suite: currentSuite, name: name, message: message });
    // eslint-disable-next-line no-console
    console.log('    FAIL ' + name + '\n         ' + message);
  }
};

const show = (value: unknown): string => {
  if (typeof value === 'string') {
    return JSON.stringify(value);
  }
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
};

export const assertEqual = (actual: unknown, expected: unknown, note?: string): void => {
  if (actual !== expected) {
    throw new Error(
      (note ? note + ': ' : '') + 'expected ' + show(expected) + ' but got ' + show(actual)
    );
  }
};

export const assertDeepEqual = (actual: unknown, expected: unknown, note?: string): void => {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) {
    throw new Error((note ? note + ': ' : '') + 'expected ' + b + ' but got ' + a);
  }
};

export const assertTrue = (value: boolean, note?: string): void => {
  if (value !== true) {
    throw new Error(note || 'expected true');
  }
};

export const assertFalse = (value: boolean, note?: string): void => {
  if (value !== false) {
    throw new Error(note || 'expected false');
  }
};

export const assertUndefined = (value: unknown, note?: string): void => {
  if (value !== undefined) {
    throw new Error((note ? note + ': ' : '') + 'expected undefined but got ' + show(value));
  }
};

export const assertDefined = (value: unknown, note?: string): void => {
  if (value === undefined || value === null) {
    throw new Error(note || 'expected a value');
  }
};

export const assertClose = (actual: number, expected: number, tolerance: number, note?: string): void => {
  if (typeof actual !== 'number' || isNaN(actual) || Math.abs(actual - expected) > tolerance) {
    throw new Error(
      (note ? note + ': ' : '') + 'expected ' + expected + ' ± ' + tolerance + ' but got ' + actual
    );
  }
};

export const assertContains = (haystack: string, needle: string, note?: string): void => {
  if (String(haystack).indexOf(needle) === -1) {
    throw new Error(
      (note ? note + ': ' : '') + 'expected ' + show(haystack) + ' to contain ' + show(needle)
    );
  }
};

export const assertNotContains = (haystack: string, needle: string, note?: string): void => {
  if (String(haystack).indexOf(needle) !== -1) {
    throw new Error(
      (note ? note + ': ' : '') + 'expected ' + show(haystack) + ' NOT to contain ' + show(needle)
    );
  }
};

export const report = (): void => {
  // eslint-disable-next-line no-console
  console.log('\n' + passed + ' passed, ' + failures.length + ' failed\n');
  if (failures.length > 0) {
    failures.forEach((failure) => {
      // eslint-disable-next-line no-console
      console.log('  FAIL ' + failure.suite + ' › ' + failure.name);
      // eslint-disable-next-line no-console
      console.log('       ' + failure.message);
    });
    process.exit(1);
  }
};
