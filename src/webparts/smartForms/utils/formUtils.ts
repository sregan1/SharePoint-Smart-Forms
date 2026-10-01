import {
  ChartKind,
  ConditionOperator,
  CURRENT_SCHEMA_VERSION,
  FieldType,
  IAddressValue,
  ICondition,
  IConditionGroup,
  IFormDefinition,
  IFormField,
  IFormFile,
  IFormSection,
  IFormValues,
  IHyperlinkValue,
  ILikertValue,
  IPersonInfo,
  IMessageBag,
  isInputType,
  migrateDefinition,
  msg,
  RANGE_OPERATORS,
  UNARY_OPERATORS
} from '../models';
import { evaluateFormula } from './formula';
import { stripHtml } from './sanitizeHtml';

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const PHONE_REGEX = /^[+]?[\d\s().-]{7,25}$/;
const URL_REGEX = /^(https?:\/\/)[^\s]+\.[^\s]{2,}/i;

/** Separator used for list-ish values stored in a single text column. */
export const MULTI_SEPARATOR = '; ';

// ---------------------------------------------------------------------------
// field traversal
// ---------------------------------------------------------------------------

/** Every field of a definition, including layout-only blocks, in display order. */
export const allFields = (definition: IFormDefinition): IFormField[] => {
  const result: IFormField[] = [];
  (definition.sections || []).forEach((s) => (s.fields || []).forEach((f) => result.push(f)));
  return result;
};

/** Only fields that collect an answer — these are the ones with list columns. */
export const inputFields = (definition: IFormDefinition): IFormField[] =>
  allFields(definition).filter((f) => isInputType(f.type));

export const findField = (definition: IFormDefinition, fieldId: string): IFormField | undefined =>
  allFields(definition).filter((f) => f.id === fieldId)[0];

export const findFieldByTitle = (definition: IFormDefinition, title: string): IFormField | undefined => {
  const wanted = (title || '').trim().toLowerCase();
  return allFields(definition).filter((f) => (f.title || '').trim().toLowerCase() === wanted)[0];
};

/**
 * Map of field id to its 1-based display number, built once per render.
 * Callers used to derive this inside a per-field loop, which made numbering
 * O(n^2) in the number of questions.
 */
export const buildNumberMap = (fields: IFormField[]): { [fieldId: string]: number } => {
  const map: { [fieldId: string]: number } = {};
  let n = 0;
  fields.forEach((field) => {
    if (isInputType(field.type)) {
      n++;
      map[field.id] = n;
    }
  });
  return map;
};

/**
 * Generate a SharePoint internal column name from a display title.
 * Letters and digits only (no underscore escaping headaches), prefixed so
 * Smart Forms columns are easy to recognize, unique within the definition.
 *
 * `retired` lists names of columns whose questions were deleted after being
 * published: the column still exists in the list, so reusing its name would
 * make a new question silently adopt the old column (wrong type, old data).
 */
export const generateInternalName = (title: string, existing: string[], retired?: string[]): string => {
  let base = 'SF';
  const words = (title || 'Field').split(/[^a-zA-Z0-9]+/).filter((w) => w.length > 0);
  words.forEach((w) => {
    base += w.charAt(0).toUpperCase() + w.slice(1);
  });
  base = base.replace(/[^a-zA-Z0-9]/g, '').slice(0, 28);
  if (base === 'SF') {
    base = 'SFField';
  }
  let candidate = base;
  let suffix = 2;
  const taken = (existing || []).concat(retired || []).map((e) => (e || '').toLowerCase());
  while (taken.indexOf(candidate.toLowerCase()) !== -1) {
    candidate = base + suffix;
    suffix++;
  }
  return candidate;
};

/** Every internal name that must not be handed out again: live fields plus retired columns. */
export const reservedInternalNames = (definition: IFormDefinition): string[] =>
  allFields(definition)
    .map((f) => f.internalName)
    .filter((n) => !!n)
    .concat(definition.retiredColumns || []);

/**
 * Record an internal column name as retired so generateInternalName never reuses
 * it. Call when a question that was already published (provisioned) is deleted.
 * Mutates and returns the definition; safe to call twice.
 */
export const retireField = (definition: IFormDefinition, name: string | undefined): IFormDefinition => {
  const clean = (name || '').trim();
  if (clean.length === 0) {
    return definition;
  }
  const list = definition.retiredColumns || [];
  if (list.filter((n) => n.toLowerCase() === clean.toLowerCase()).length === 0) {
    definition.retiredColumns = list.concat([clean]);
  }
  return definition;
};

/** Retire a deleted field's column when it had been provisioned (no-op for never-published fields). */
export const retireDeletedField = (definition: IFormDefinition, field: IFormField): IFormDefinition =>
  field.provisioned && field.internalName ? retireField(definition, field.internalName) : definition;

// ---------------------------------------------------------------------------
// choices
// ---------------------------------------------------------------------------

/** Deterministic PRNG so a shuffled order is stable across re-renders. */
const seededRandom = (seed: number): (() => number) => {
  let state = seed >>> 0;
  return () => {
    state += 0x6d2b79f5;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

export const shuffleWithSeed = <T,>(items: T[], seed: number): T[] => {
  const random = seededRandom(seed);
  const result = items.slice();
  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(random() * (i + 1));
    const swap = result[i];
    result[i] = result[j];
    result[j] = swap;
  }
  return result;
};

export const OTHER_PREFIX = 'Other: ';

/** Non-empty, trimmed, de-duplicated options for a choice-like field. */
export const effectiveChoices = (field: IFormField): string[] => {
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  (field.choices || []).forEach((raw) => {
    const value = (raw || '').trim();
    if (value.length === 0 || seen[value.toLowerCase()]) {
      return;
    }
    seen[value.toLowerCase()] = true;
    result.push(value);
  });
  return result;
};

/** True when a stored answer is a write-in "Other" value. */
export const isOtherValue = (field: IFormField, value: string): boolean =>
  field.allowOther === true && effectiveChoices(field).indexOf(value) === -1 && value.length > 0;

// ---------------------------------------------------------------------------
// emptiness & validation
// ---------------------------------------------------------------------------

export const isEmptyValue = (field: IFormField, value: unknown): boolean => {
  if (value === undefined || value === null) {
    return true;
  }
  switch (field.type) {
    case FieldType.YesNo:
      // a toggle always holds a value — emptiness only matters for Consent
      return false;
    case FieldType.Consent:
      return value !== true;
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
    case FieldType.Ranking:
      return Array.isArray(value) ? value.length === 0 : String(value).trim().length === 0;
    case FieldType.Likert: {
      const answers = value as ILikertValue;
      if (typeof answers !== 'object') {
        return true;
      }
      return Object.keys(answers).filter((row) => !!answers[row]).length === 0;
    }
    case FieldType.Person:
      return !Array.isArray(value) || value.length === 0;
    case FieldType.FileUpload:
      return !Array.isArray(value) || (value as IFormFile[]).length === 0;
    case FieldType.Signature:
      return String(value).trim().length === 0;
    case FieldType.Hyperlink:
      return !(value as IHyperlinkValue).url;
    case FieldType.Address: {
      const address = value as IAddressValue;
      if (typeof address === 'string') {
        return String(address).trim().length === 0;
      }
      if (!address) {
        return true;
      }
      return (
        [address.street, address.street2, address.city, address.state, address.postalCode, address.country]
          .filter((part) => (part || '').trim().length > 0).length === 0
      );
    }
    case FieldType.Rating:
      return !(typeof value === 'number' && value > 0);
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider:
    case FieldType.Scale:
      return typeof value !== 'number' || isNaN(value as number);
    case FieldType.Date:
    case FieldType.Time:
      return !(value instanceof Date) || isNaN((value as Date).getTime());
    case FieldType.RichText:
      return stripHtml(String(value)).length === 0;
    case FieldType.Content:
      return false;
    default:
      return String(value).trim().length === 0;
  }
};

/** Selection count for the multi-select rules. */
const selectionCount = (value: unknown): number =>
  Array.isArray(value) ? value.length : isEmptyValue({ type: FieldType.Text } as IFormField, value) ? 0 : 1;

/** Result of a validation check: a message code plus tokens, or author-supplied text. */
export interface IValidationResult {
  /** Logic_ message key (empty when `message` carries author-supplied text) */
  code: string;
  params?: { [token: string]: string | number };
  /** set when the text was written by the form owner (requiredMessage / patternMessage) */
  message?: string;
}

/**
 * Compile a field's custom validation pattern. The expression is anchored as
 * ^(?:...)$ so "\d{3}" means the whole answer is three digits, not "contains
 * three digits". Returns undefined for an invalid pattern — it must never throw.
 */
export const compilePattern = (pattern: string | undefined): RegExp | undefined => {
  if (!pattern) {
    return undefined;
  }
  try {
    return new RegExp('^(?:' + pattern + ')$');
  } catch {
    return undefined;
  }
};

const failure = (code: string, params?: { [token: string]: string | number }): IValidationResult => ({
  code: code,
  params: params
});

/**
 * Validate one answer and return a message code plus tokens (or undefined when
 * valid). Use this when the caller supplies its own strings; validateField wraps
 * it with the English defaults.
 */
export const validateFieldResult = (field: IFormField, value: unknown): IValidationResult | undefined => {
  if (!isInputType(field.type)) {
    return undefined;
  }
  const label = field.title || '';
  const empty = isEmptyValue(field, value);

  if (field.required && empty) {
    if (field.requiredMessage) {
      return { code: '', message: field.requiredMessage };
    }
    switch (field.type) {
      case FieldType.Consent:
        return failure('Logic_RequiredConsent');
      case FieldType.FileUpload:
        return failure('Logic_RequiredFile');
      case FieldType.Signature:
        return failure('Logic_RequiredSignature');
      case FieldType.Likert:
        return failure('Logic_RequiredLikert');
      default:
        return failure('Logic_Required', { label: label });
    }
  }
  if (empty) {
    return undefined;
  }

  switch (field.type) {
    case FieldType.Email:
      if (!EMAIL_REGEX.test(String(value).trim())) {
        return failure('Logic_InvalidEmail');
      }
      break;

    case FieldType.Phone:
      if (!PHONE_REGEX.test(String(value).trim())) {
        return failure('Logic_InvalidPhone');
      }
      break;

    case FieldType.Hyperlink:
      if (!URL_REGEX.test((value as IHyperlinkValue).url.trim())) {
        return failure('Logic_InvalidUrl');
      }
      break;

    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider:
    case FieldType.Scale:
    case FieldType.Rating: {
      const num = value as number;
      const min = field.type === FieldType.Rating ? undefined : field.min;
      const max = field.type === FieldType.Rating ? field.maxRating : field.max;
      if (typeof min === 'number' && num < min) {
        return failure('Logic_MinValue', { value: formatNumberForMessage(field, min) });
      }
      if (typeof max === 'number' && num > max) {
        return failure('Logic_MaxValue', { value: formatNumberForMessage(field, max) });
      }
      break;
    }

    case FieldType.Text:
    case FieldType.MultilineText:
      if (field.maxLength && String(value).length > field.maxLength) {
        return failure('Logic_MaxLength', { max: field.maxLength });
      }
      break;

    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup: {
      const count = selectionCount(value);
      if (field.allowMultiple && typeof field.minSelections === 'number' && count < field.minSelections) {
        return failure(field.minSelections === 1 ? 'Logic_ChooseAtLeastOne' : 'Logic_ChooseAtLeastMany', {
          count: field.minSelections
        });
      }
      if (field.allowMultiple && typeof field.maxSelections === 'number' && count > field.maxSelections) {
        return failure(field.maxSelections === 1 ? 'Logic_ChooseAtMostOne' : 'Logic_ChooseAtMostMany', {
          count: field.maxSelections
        });
      }
      break;
    }

    case FieldType.Likert: {
      const answers = (value || {}) as ILikertValue;
      const rows = field.likertRows || [];
      const unanswered = rows.filter((row) => !answers[row]);
      if (field.required && unanswered.length > 0) {
        return failure('Logic_LikertRemaining', { count: unanswered.length });
      }
      break;
    }

    case FieldType.FileUpload: {
      const files = (value || []) as IFormFile[];
      const maxFiles = field.maxFiles || 3;
      if (files.length > maxFiles) {
        return failure(maxFiles === 1 ? 'Logic_MaxFilesOne' : 'Logic_MaxFilesMany', { count: maxFiles });
      }
      const maxBytes = (field.maxFileSizeMb || 10) * 1024 * 1024;
      const tooBig = files.filter((f) => f.size > maxBytes)[0];
      if (tooBig) {
        return failure('Logic_FileTooLarge', { name: tooBig.name, size: field.maxFileSizeMb || 10 });
      }
      const allowed = (field.allowedExtensions || []).map((e) => e.toLowerCase().replace(/^\./, ''));
      if (allowed.length > 0) {
        const wrong = files.filter((f) => allowed.indexOf(fileExtension(f.name)) === -1)[0];
        if (wrong) {
          return failure('Logic_FileTypeNotAccepted', { name: wrong.name, types: allowed.join(', ') });
        }
      }
      break;
    }

    case FieldType.Address: {
      const address = value as IAddressValue;
      if (address && typeof address !== 'string' && field.required) {
        if (!(address.street || '').trim() || !(address.city || '').trim()) {
          return failure('Logic_AddressRequired');
        }
      }
      break;
    }
  }

  // custom pattern applies to any text-bearing answer; anchored, never throws
  if (field.pattern) {
    const text = String(typeof value === 'object' ? formatValue(field, value) : value);
    const expression = compilePattern(field.pattern);
    if (expression && !expression.test(text)) {
      return field.patternMessage
        ? { code: '', message: field.patternMessage }
        : failure('Logic_PatternDefault');
    }
  }

  return undefined;
};

/** Returns an error message, or undefined when the value is valid. `messages` overrides the English defaults. */
export const validateField = (
  field: IFormField,
  value: unknown,
  messages?: IMessageBag
): string | undefined => {
  const result = validateFieldResult(field, value);
  if (!result) {
    return undefined;
  }
  if (result.message !== undefined) {
    return result.message;
  }
  const params = { ...(result.params || {}) };
  if (result.code === 'Logic_Required' && !params.label) {
    params.label = msg('Logic_ThisQuestion', undefined, messages);
  }
  return msg(result.code, params, messages);
};

const formatNumberForMessage = (field: IFormField, num: number): string => {
  if (field.numberFormat === 'currency') {
    return (field.currencySymbol || '$') + String(num);
  }
  if (field.numberFormat === 'percent') {
    return String(num) + '%';
  }
  return String(num);
};

export const fileExtension = (name: string): string => {
  const parts = String(name || '').split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
};

// ---------------------------------------------------------------------------
// conditional logic
// ---------------------------------------------------------------------------

/** Numeric projection of a value for range comparisons. */
const asNumber = (value: unknown): number | undefined => {
  if (typeof value === 'number') {
    return isNaN(value) ? undefined : value;
  }
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? undefined : value.getTime();
  }
  if (Array.isArray(value)) {
    return value.length;
  }
  const num = Number(String(value === undefined || value === null ? '' : value).replace(/,/g, '').trim());
  return isNaN(num) ? undefined : num;
};

const MS_PER_DAY = 86400000;

/** Calendar-day index (local date, time of day ignored) so comparisons never depend on the hour or time zone. */
const dayIndex = (date: Date): number =>
  Math.floor(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / MS_PER_DAY);

const minutesOfDay = (date: Date): number => date.getHours() * 60 + date.getMinutes();

/**
 * Parse a date string as a LOCAL date. A bare "2026-03-05" is midnight local
 * time (new Date("2026-03-05") would be UTC midnight, i.e. the previous evening
 * west of Greenwich). "today" gives today's date; strings with a time part are
 * parsed as usual. Returns undefined when unparseable.
 */
export const parseLocalDate = (text: string | undefined): Date | undefined => {
  const trimmed = (text || '').trim();
  if (trimmed.length === 0) {
    return undefined;
  }
  if (/^today$/i.test(trimmed)) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return today;
  }
  const dateOnly = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(trimmed);
  if (dateOnly) {
    const year = Number(dateOnly[1]);
    const month = Number(dateOnly[2]) - 1;
    const day = Number(dateOnly[3]);
    const local = new Date(year, month, day);
    // reject overflow such as 2026-02-31
    return local.getFullYear() === year && local.getMonth() === month && local.getDate() === day
      ? local
      : undefined;
  }
  const parsed = new Date(trimmed);
  return isNaN(parsed.getTime()) ? undefined : parsed;
};

/**
 * Parse a time-of-day operand ("HH:mm", "H:mm", "HH:mm:ss", "h:mm AM/PM", or a
 * full date-time string) into minutes after midnight. Undefined when invalid.
 */
export const parseTimeToMinutes = (text: string | undefined): number | undefined => {
  const trimmed = (text || '').trim();
  if (trimmed.length === 0) {
    return undefined;
  }
  const match = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]m)?$/i.exec(trimmed);
  if (match) {
    let hours = Number(match[1]);
    const minutes = Number(match[2]);
    const meridiem = match[3] ? match[3].toLowerCase() : '';
    if (minutes > 59) {
      return undefined;
    }
    if (meridiem) {
      if (hours < 1 || hours > 12) {
        return undefined;
      }
      hours = (hours % 12) + (meridiem === 'pm' ? 12 : 0);
    } else if (hours > 23) {
      return undefined;
    }
    return hours * 60 + minutes;
  }
  if (/\d{4}-\d{2}-\d{2}T/.test(trimmed)) {
    const parsed = new Date(trimmed);
    return isNaN(parsed.getTime()) ? undefined : minutesOfDay(parsed);
  }
  return undefined;
};

const toDateValue = (raw: unknown): Date | undefined => {
  if (raw instanceof Date) {
    return isNaN(raw.getTime()) ? undefined : raw;
  }
  if (typeof raw === 'string') {
    return parseLocalDate(raw);
  }
  return undefined;
};

const isDateLike = (field: IFormField | undefined): boolean =>
  !!field && (field.type === FieldType.Date || field.type === FieldType.Time);

const NUMERIC_TYPES: FieldType[] = [
  FieldType.Number,
  FieldType.Calculated,
  FieldType.Slider,
  FieldType.Scale,
  FieldType.Rating
];

/** What an answer compares as: a day index for Date, minutes for Time, otherwise a number. */
const projectAnswer = (field: IFormField, raw: unknown): number | undefined => {
  if (field.type === FieldType.Date) {
    const date = toDateValue(raw);
    return date ? dayIndex(date) : undefined;
  }
  if (field.type === FieldType.Time) {
    if (typeof raw === 'string' && parseTimeToMinutes(raw) !== undefined) {
      return parseTimeToMinutes(raw);
    }
    const date = toDateValue(raw);
    return date ? minutesOfDay(date) : undefined;
  }
  return asNumber(raw);
};

/** Parse a rule's comparison operand in the same space as projectAnswer. */
const projectOperand = (field: IFormField, raw: string | undefined): number | undefined => {
  const text = (raw || '').trim();
  if (text.length === 0) {
    return undefined;
  }
  if (field.type === FieldType.Date) {
    const date = parseLocalDate(text);
    return date ? dayIndex(date) : undefined;
  }
  if (field.type === FieldType.Time) {
    return parseTimeToMinutes(text);
  }
  const num = Number(text.replace(/,/g, ''));
  return isNaN(num) ? undefined : num;
};

/** Selections of a multi-value answer, split on "; " so labels containing a bare ";" survive. */
export const splitMultiValue = (text: string): string[] =>
  String(text || '')
    .split(/;\s+|;$/)
    .map((v) => v.trim())
    .filter((v) => v.length > 0);

// ----- visibility (dependency-ordered, hidden drivers count as empty) -----

interface IVisibilityContext {
  definition: IFormDefinition;
  values: IFormValues;
  fieldsById: { [id: string]: IFormField };
  sectionOfField: { [id: string]: IFormSection };
  cache: { [id: string]: boolean };
  inProgress: { [id: string]: boolean };
}

const createVisibilityContext = (definition: IFormDefinition, values: IFormValues): IVisibilityContext => {
  const fieldsById: { [id: string]: IFormField } = {};
  const sectionOfField: { [id: string]: IFormSection } = {};
  (definition.sections || []).forEach((section) => {
    (section.fields || []).forEach((f) => {
      if (!fieldsById[f.id]) {
        fieldsById[f.id] = f;
        sectionOfField[f.id] = section;
      }
    });
  });
  return { definition, values, fieldsById, sectionOfField, cache: {}, inProgress: {} };
};

const splitParts = (driver: IFormField, raw: unknown): string[] => {
  if (Array.isArray(raw)) {
    return raw
      .map((v) => {
        if (typeof v === 'object' && v !== null) {
          const person = v as IPersonInfo & IFormFile;
          return person.displayName || person.name || '';
        }
        return String(v);
      })
      .map((p) => p.trim().toLowerCase())
      .filter((p) => p.length > 0);
  }
  if (driver.allowMultiple) {
    return splitMultiValue(valueAsComparableString(driver, raw)).map((p) => p.toLowerCase());
  }
  return [];
};

const evaluateConditionIn = (condition: ICondition, ctx: IVisibilityContext): boolean => {
  const driver = ctx.fieldsById[condition.fieldId];
  if (!driver) {
    // a rule pointing at a deleted question can never be satisfied — treating
    // it as "true" used to make branched questions unconditionally visible
    return false;
  }
  // a hidden question's stale answer must not drive other questions
  const raw = fieldVisibleIn(driver, ctx) ? ctx.values[driver.id] : undefined;
  const operator = condition.operator;

  if (UNARY_OPERATORS.indexOf(operator) !== -1) {
    const empty = isEmptyValue(driver, raw);
    return operator === 'empty' ? empty : !empty;
  }

  const dateLike = isDateLike(driver);
  const numeric = NUMERIC_TYPES.indexOf(driver.type) !== -1;

  if (RANGE_OPERATORS.indexOf(operator) !== -1) {
    const actual = projectAnswer(driver, raw);
    const bound = projectOperand(driver, condition.value);
    if (actual === undefined || bound === undefined) {
      return false;
    }
    switch (operator) {
      case 'greaterThan':
      case 'after':
        return actual > bound;
      case 'greaterOrEqual':
        return actual >= bound;
      case 'lessThan':
      case 'before':
        return actual < bound;
      case 'lessOrEqual':
        return actual <= bound;
      case 'between': {
        const upper = projectOperand(driver, condition.value2);
        if (upper === undefined) {
          return false;
        }
        return actual >= Math.min(bound, upper) && actual <= Math.max(bound, upper);
      }
      default:
        return false;
    }
  }

  // equality on dates compares calendar days (times compare minutes), numbers compare numerically
  if ((dateLike || numeric) && (operator === 'equals' || operator === 'notEquals')) {
    const actual = projectAnswer(driver, raw);
    const bound = projectOperand(driver, condition.value);
    if (actual !== undefined && bound !== undefined) {
      return operator === 'equals' ? actual === bound : actual !== bound;
    }
    if (dateLike) {
      return operator === 'notEquals';
    }
  }

  const text = valueAsComparableString(driver, raw);
  const expected = (condition.value || '').toLowerCase();
  const haystack = text.toLowerCase();
  // multi-value answers match if *any* selection matches
  const parts = splitParts(driver, raw);
  const multi = Array.isArray(raw) || driver.allowMultiple === true;

  switch (operator) {
    case 'equals':
      return multi ? parts.indexOf(expected) !== -1 : haystack === expected;
    case 'notEquals':
      return multi ? parts.indexOf(expected) === -1 : haystack !== expected;
    case 'contains':
      return haystack.indexOf(expected) !== -1;
    case 'notContains':
      return haystack.indexOf(expected) === -1;
    default:
      return true;
  }
};

const evaluateGroupIn = (
  group: IConditionGroup | undefined,
  ctx: IVisibilityContext,
  selfFieldId?: string
): boolean => {
  if (!group || !group.conditions || group.conditions.length === 0) {
    return true;
  }
  const usable = group.conditions.filter((c) => c.fieldId && c.fieldId !== selfFieldId);
  if (usable.length === 0) {
    return true;
  }
  const results = usable.map((c) => evaluateConditionIn(c, ctx));
  return group.match === 'any'
    ? results.filter((r) => r).length > 0
    : results.filter((r) => !r).length === 0;
};

/** Memoized, dependency-ordered visibility of one field (its section's rule included). */
const fieldVisibleIn = (field: IFormField, ctx: IVisibilityContext): boolean => {
  if (ctx.cache[field.id] !== undefined) {
    return ctx.cache[field.id];
  }
  if (ctx.inProgress[field.id]) {
    // circular rules: break the cycle by treating the field as visible
    return true;
  }
  ctx.inProgress[field.id] = true;
  const section = ctx.sectionOfField[field.id];
  const visible =
    (!section || evaluateGroupIn(section.visibleWhen, ctx)) &&
    evaluateGroupIn(field.visibleWhen, ctx, field.id);
  ctx.inProgress[field.id] = false;
  ctx.cache[field.id] = visible;
  return visible;
};

/**
 * Visibility of every field in one pass, evaluated in dependency order. Prefer
 * this over calling isFieldVisible in a loop.
 */
export const computeVisibility = (
  definition: IFormDefinition,
  values: IFormValues
): { [fieldId: string]: boolean } => {
  const ctx = createVisibilityContext(definition, values);
  const result: { [fieldId: string]: boolean } = {};
  allFields(definition).forEach((f) => {
    result[f.id] = fieldVisibleIn(f, ctx);
  });
  return result;
};

/** Evaluate a condition group. An empty or absent group is always satisfied. */
export const evaluateConditionGroup = (
  group: IConditionGroup | undefined,
  definition: IFormDefinition,
  values: IFormValues,
  selfFieldId?: string
): boolean => evaluateGroupIn(group, createVisibilityContext(definition, values), selfFieldId);

/** Evaluate a field's visibility, including its section's own rule and hidden drivers. */
export const isFieldVisible = (
  field: IFormField,
  definition: IFormDefinition,
  values: IFormValues
): boolean => fieldVisibleIn(field, createVisibilityContext(definition, values));

export const isSectionVisible = (
  section: IFormSection,
  definition: IFormDefinition,
  values: IFormValues
): boolean => evaluateGroupIn(section.visibleWhen, createVisibilityContext(definition, values));

// ----- rule health -----

export type ConditionIssueReason =
  | 'missingDriver'
  | 'selfReference'
  | 'operator'
  | 'value'
  | 'operand'
  | 'incomplete';

export interface IConditionIssue {
  /** owner field of the rule (undefined for a section rule) */
  fieldId?: string;
  /** owner section of a section rule */
  sectionId?: string;
  conditionIndex: number;
  /** the field the rule reads */
  driverId: string;
  reason: ConditionIssueReason;
  /**
   * 'broken' rules can never behave as the author intended (deleted driver,
   * operator no longer offered, value not among the options, unparseable
   * operand); 'incomplete' rules are merely unfinished (blank comparison value).
   */
  severity: 'broken' | 'incomplete';
}

const optionsForValueCheck = (driver: IFormField): string[] | undefined => {
  switch (driver.type) {
    case FieldType.Choice:
      return driver.allowOther ? undefined : effectiveChoices(driver);
    case FieldType.ImageChoice:
      return (driver.imageChoices || []).map((o) => (o.label || '').trim()).filter((l) => l.length > 0);
    case FieldType.YesNo:
    case FieldType.Consent:
      return ['Yes', 'No'];
    default:
      return undefined;
  }
};

const checkCondition = (
  condition: ICondition,
  driver: IFormField | undefined,
  ownerFieldId: string | undefined
): { reason: ConditionIssueReason; severity: 'broken' | 'incomplete' } | undefined => {
  if (!condition.fieldId || !driver) {
    return { reason: 'missingDriver', severity: 'broken' };
  }
  if (ownerFieldId && condition.fieldId === ownerFieldId) {
    return { reason: 'selfReference', severity: 'broken' };
  }
  if (operatorsForField(driver).indexOf(condition.operator) === -1) {
    return { reason: 'operator', severity: 'broken' };
  }
  if (UNARY_OPERATORS.indexOf(condition.operator) !== -1) {
    return undefined;
  }
  const value = (condition.value || '').trim();

  if (RANGE_OPERATORS.indexOf(condition.operator) !== -1 || ((isDateLike(driver) || NUMERIC_TYPES.indexOf(driver.type) !== -1))) {
    if (value.length === 0) {
      return { reason: 'incomplete', severity: 'incomplete' };
    }
    if (projectOperand(driver, value) === undefined) {
      return { reason: 'operand', severity: 'broken' };
    }
    if (condition.operator === 'between') {
      const upper = (condition.value2 || '').trim();
      if (upper.length === 0) {
        return { reason: 'incomplete', severity: 'incomplete' };
      }
      if (projectOperand(driver, upper) === undefined) {
        return { reason: 'operand', severity: 'broken' };
      }
    }
    return undefined;
  }

  const options = optionsForValueCheck(driver);
  if (options) {
    if (value.length === 0) {
      return { reason: 'incomplete', severity: 'incomplete' };
    }
    const lower = value.toLowerCase();
    const known = options.map((o) => o.toLowerCase());
    const fits =
      condition.operator === 'contains' || condition.operator === 'notContains'
        ? known.filter((o) => o.indexOf(lower) !== -1).length > 0
        : known.indexOf(lower) !== -1;
    if (!fits) {
      return { reason: 'value', severity: 'broken' };
    }
  }
  return undefined;
};

/** Every branching rule that no longer makes sense, for the designer to flag or prune. */
export const conditionIssues = (definition: IFormDefinition): IConditionIssue[] => {
  const byId: { [id: string]: IFormField } = {};
  allFields(definition).forEach((f) => {
    byId[f.id] = f;
  });
  const issues: IConditionIssue[] = [];
  const scan = (group: IConditionGroup | undefined, fieldId?: string, sectionId?: string): void => {
    ((group && group.conditions) || []).forEach((condition, index) => {
      const problem = checkCondition(condition, byId[condition.fieldId], fieldId);
      if (problem) {
        issues.push({
          fieldId: fieldId,
          sectionId: sectionId,
          conditionIndex: index,
          driverId: condition.fieldId,
          reason: problem.reason,
          severity: problem.severity
        });
      }
    });
  };
  (definition.sections || []).forEach((section) => {
    scan(section.visibleWhen, undefined, section.id);
    (section.fields || []).forEach((f) => scan(f.visibleWhen, f.id, section.id));
  });
  return issues;
};

/**
 * Remove branching rules that can no longer work: deleted drivers, operators
 * the driver's (possibly changed) type no longer offers, values that are no
 * longer among the options, unparseable operands. Blank-value rules are kept
 * (they are just unfinished) unless `includeIncomplete` is set. Mutates and
 * returns the definition.
 */
export const pruneInvalidConditions = (
  definition: IFormDefinition,
  options?: { includeIncomplete?: boolean }
): IFormDefinition => {
  const dropIncomplete = !!(options && options.includeIncomplete);
  const byId: { [id: string]: IFormField } = {};
  allFields(definition).forEach((f) => {
    byId[f.id] = f;
  });
  const prune = (group: IConditionGroup | undefined, ownerFieldId?: string): IConditionGroup | undefined => {
    if (!group) {
      return undefined;
    }
    const conditions = (group.conditions || []).filter((c) => {
      const problem = checkCondition(c, byId[c.fieldId], ownerFieldId);
      return !problem || (problem.severity === 'incomplete' && !dropIncomplete);
    });
    return conditions.length > 0 ? { match: group.match, conditions } : undefined;
  };
  definition.sections.forEach((section) => {
    const sectionRule = prune(section.visibleWhen);
    if (sectionRule) {
      section.visibleWhen = sectionRule;
    } else {
      delete section.visibleWhen;
    }
    section.fields.forEach((field) => {
      const fieldRule = prune(field.visibleWhen, field.id);
      if (fieldRule) {
        field.visibleWhen = fieldRule;
      } else {
        delete field.visibleWhen;
      }
    });
  });
  return definition;
};

/** Remove rules that reference fields no longer in the definition (and other unworkable rules). */
export const pruneDanglingConditions = (definition: IFormDefinition): IFormDefinition =>
  pruneInvalidConditions(definition);

// ---------------------------------------------------------------------------
// calculated fields
// ---------------------------------------------------------------------------

/**
 * Recompute every Calculated field from the current answers. Iterates until the
 * values stop changing (bounded by the number of calculated fields + 1, so a
 * chain of calculations resolves and a circular one cannot loop forever).
 * Answers to hidden questions are ignored, exactly as they are at submit time.
 */
export const applyCalculatedFields = (definition: IFormDefinition, values: IFormValues): IFormValues => {
  const calculated = allFields(definition).filter((f) => f.type === FieldType.Calculated && f.formula);
  if (calculated.length === 0) {
    return values;
  }
  const byTitle: { [title: string]: IFormField } = {};
  allFields(definition).forEach((f) => {
    const key = (f.title || '').trim().toLowerCase();
    if (key && !byTitle[key]) {
      byTitle[key] = f;
    }
  });
  const next: IFormValues = { ...values };
  const passes = calculated.length + 1;
  for (let pass = 0; pass < passes; pass++) {
    // visibility can itself depend on a calculated answer, so re-derive each pass
    const visibility = computeVisibility(definition, next);
    let changed = false;
    calculated.forEach((field) => {
      const result = evaluateFormula(field.formula || '', (name) => {
        const referenced = byTitle[name.trim().toLowerCase()];
        if (!referenced || visibility[referenced.id] === false) {
          return undefined;
        }
        return asNumber(next[referenced.id]);
      });
      const rounded =
        result === undefined
          ? undefined
          : typeof field.decimalPlaces === 'number'
            ? Number(result.toFixed(field.decimalPlaces))
            : Math.round(result * 1e6) / 1e6;
      if (next[field.id] !== rounded) {
        next[field.id] = rounded;
        changed = true;
      }
    });
    if (!changed) {
      break;
    }
  }
  return next;
};

// ---------------------------------------------------------------------------
// value normalization: SharePoint shape -> form-state shape
// ---------------------------------------------------------------------------

/**
 * Convert a raw value read back from SharePoint into the same shape the form
 * uses while being filled in. Everything downstream (formatting, summaries,
 * CSV, the read-only response renderer) then works against one shape instead
 * of branching on provenance.
 */
export const normalizeFromSharePoint = (field: IFormField, raw: unknown): unknown => {
  if (raw === undefined || raw === null || raw === '') {
    return undefined;
  }
  switch (field.type) {
    case FieldType.Date:
    case FieldType.Time: {
      const date = new Date(String(raw));
      return isNaN(date.getTime()) ? undefined : date;
    }
    case FieldType.YesNo:
    case FieldType.Consent:
      return raw === true || raw === 'true' || raw === 1;
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Rating:
    case FieldType.Slider:
    case FieldType.Scale: {
      const num = Number(raw);
      return isNaN(num) ? undefined : num;
    }
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
      if (Array.isArray(raw)) {
        return (raw as unknown[]).map((v) => String(v));
      }
      return field.allowMultiple ? splitMulti(String(raw)) : String(raw);
    case FieldType.Ranking:
      return Array.isArray(raw) ? (raw as unknown[]).map((v) => String(v)) : splitMulti(String(raw));
    case FieldType.Likert:
      return parseJsonObject(String(raw));
    case FieldType.Person: {
      const toPerson = (entry: { Title?: string; EMail?: string; Email?: string; Name?: string }): IPersonInfo => ({
        loginName: entry.Name || entry.Title || '',
        displayName: entry.Title || '',
        email: entry.EMail || entry.Email
      });
      if (Array.isArray(raw)) {
        return (raw as { Title?: string }[]).map(toPerson);
      }
      return [toPerson(raw as { Title?: string })];
    }
    case FieldType.Hyperlink: {
      const link = raw as { Url?: string; Description?: string };
      if (link && link.Url) {
        return { url: link.Url, description: link.Description };
      }
      return { url: String(raw), description: '' };
    }
    case FieldType.FileUpload:
      return splitMulti(String(raw)).map((name) => ({ name, size: 0, content: '' }));
    default:
      return String(raw);
  }
};

const splitMulti = (text: string): string[] => splitMultiValue(text);

const parseJsonObject = (text: string): ILikertValue => {
  try {
    const parsed = JSON.parse(text);
    return parsed && typeof parsed === 'object' ? (parsed as ILikertValue) : {};
  } catch {
    return {};
  }
};

/** Format a postal address as a single readable line. */
export const formatAddress = (value: IAddressValue | string | undefined): string => {
  if (!value) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  return [value.street, value.street2, value.city, value.state, value.postalCode, value.country]
    .map((part) => (part || '').trim())
    .filter((part) => part.length > 0)
    .join(', ');
};

// ---------------------------------------------------------------------------
// formatting
// ---------------------------------------------------------------------------

const formatNumber = (field: IFormField, num: number): string => {
  const decimals =
    typeof field.decimalPlaces === 'number'
      ? field.decimalPlaces
      : field.numberFormat === 'currency'
        ? 2
        : undefined;
  const text = decimals === undefined ? String(num) : num.toFixed(decimals);
  const grouped = groupThousands(text);
  if (field.numberFormat === 'currency') {
    return (field.currencySymbol || '$') + grouped;
  }
  if (field.numberFormat === 'percent') {
    return grouped + '%';
  }
  return field.unitSuffix ? grouped + ' ' + field.unitSuffix : grouped;
};

const groupThousands = (text: string): string => {
  const negative = text.charAt(0) === '-';
  const body = negative ? text.slice(1) : text;
  const parts = body.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (negative ? '-' : '') + parts.join('.');
};

/**
 * Human-readable rendering of a form value. Accepts the form-state shape, so
 * SharePoint values must go through normalizeFromSharePoint first.
 */
export const formatValue = (
  field: IFormField,
  value: unknown,
  messages?: IMessageBag,
  locale?: string
): string => {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  switch (field.type) {
    case FieldType.YesNo:
      return value === true ? msg('Logic_Yes', undefined, messages) : msg('Logic_No', undefined, messages);
    case FieldType.Consent:
      return value === true
        ? msg('Logic_Agreed', undefined, messages)
        : msg('Logic_NotAgreed', undefined, messages);
    case FieldType.Date:
      if (value instanceof Date) {
        return field.includeTime ? value.toLocaleString(locale) : value.toLocaleDateString(locale);
      }
      return String(value);
    case FieldType.Time:
      return value instanceof Date
        ? value.toLocaleTimeString(locale, { hour: 'numeric', minute: '2-digit' })
        : String(value);
    case FieldType.Number:
    case FieldType.Calculated:
      return typeof value === 'number' ? formatNumber(field, value) : String(value);
    case FieldType.Rating:
      return String(value) + ' / ' + String(field.maxRating || 5);
    case FieldType.Scale:
      return String(value) + ' / ' + String(typeof field.max === 'number' ? field.max : 10);
    case FieldType.Slider:
      return field.unitSuffix ? String(value) + ' ' + field.unitSuffix : String(value);
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
      return Array.isArray(value) ? (value as string[]).join(', ') : String(value);
    case FieldType.Ranking:
      return Array.isArray(value)
        ? (value as string[]).map((v, i) => String(i + 1) + '. ' + v).join('   ')
        : String(value);
    case FieldType.Likert: {
      const answers = value as ILikertValue;
      return Object.keys(answers)
        .filter((row) => !!answers[row])
        .map((row) => row + ': ' + answers[row])
        .join(MULTI_SEPARATOR);
    }
    case FieldType.Person:
      return Array.isArray(value)
        ? (value as IPersonInfo[]).map((p) => p.displayName).filter((n) => n).join(', ')
        : String(value);
    case FieldType.Hyperlink: {
      const link = value as IHyperlinkValue;
      if (!link.url) {
        return '';
      }
      return link.description && link.description !== link.url
        ? link.description + ' (' + link.url + ')'
        : link.url;
    }
    case FieldType.Address:
      return formatAddress(value as IAddressValue | string);
    case FieldType.FileUpload:
      return Array.isArray(value)
        ? (value as IFormFile[]).map((f) => f.name).join(MULTI_SEPARATOR)
        : String(value);
    case FieldType.Signature:
      return String(value).indexOf('data:') === 0 ? msg('Logic_Signed', undefined, messages) : String(value);
    case FieldType.RichText:
      return stripHtml(String(value));
    case FieldType.Content:
      return '';
    default:
      return String(value);
  }
};

/** Convenience wrapper: normalize a SharePoint value and format it. */
export const formatSharePointValue = (
  field: IFormField,
  raw: unknown,
  messages?: IMessageBag,
  locale?: string
): string => formatValue(field, normalizeFromSharePoint(field, raw), messages, locale);

/** String rendering used by conditional rules and search. */
export const valueAsComparableString = (field: IFormField | undefined, value: unknown): string => {
  if (value === undefined || value === null) {
    return '';
  }
  if (Array.isArray(value)) {
    return value
      .map((v) => {
        if (typeof v === 'object' && v !== null) {
          const person = v as IPersonInfo & IFormFile;
          return person.displayName || person.name || '';
        }
        return String(v);
      })
      .join(MULTI_SEPARATOR);
  }
  if (field && (field.type === FieldType.YesNo || field.type === FieldType.Consent)) {
    return value === true ? 'Yes' : 'No';
  }
  if (field && field.type === FieldType.Hyperlink) {
    return (value as IHyperlinkValue).url || '';
  }
  if (field && field.type === FieldType.Address) {
    return formatAddress(value as IAddressValue);
  }
  if (field && field.type === FieldType.Likert) {
    const answers = value as ILikertValue;
    return Object.keys(answers).map((row) => answers[row]).filter((a) => a).join(MULTI_SEPARATOR);
  }
  if (field && field.type === FieldType.RichText) {
    return stripHtml(String(value));
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
};

// ---------------------------------------------------------------------------
// notification email
// ---------------------------------------------------------------------------

const escapeHtml = (value: string): string =>
  String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

export interface IResponseEmailOptions {
  /** heading shown at the top of the email */
  heading: string;
  /** short sentence under the heading */
  intro: string;
  accentColor: string;
  /** when set, a "View in Microsoft Lists" button links to the list item */
  itemUrl?: string;
  /** translated Logic_ strings (optional) */
  messages?: IMessageBag;
  /** locale for dates in the answers (optional) */
  locale?: string;
}

/** Inline-styled HTML email summarizing one response — safe for Outlook. */
export const buildResponseEmailHtml = (
  definition: IFormDefinition,
  values: IFormValues,
  options: IResponseEmailOptions
): string => {
  const accent = options.accentColor || '#0078d4';
  const visibility = computeVisibility(definition, values);
  const rows = inputFields(definition)
    .filter((field) => visibility[field.id] !== false)
    .map((field) => {
      const answer = formatValue(field, values[field.id], options.messages, options.locale);
      return (
        '<tr>' +
        '<td style="padding:8px 16px 8px 0;font-size:13px;color:#605e5c;vertical-align:top;white-space:nowrap;">' +
        escapeHtml(field.title || msg('Logic_EmailQuestion', undefined, options.messages)) +
        '</td>' +
        '<td style="padding:8px 0;font-size:14px;color:#323130;vertical-align:top;">' +
        (answer ? escapeHtml(answer) : '<span style="color:#a19f9d;">&mdash;</span>') +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  const button = options.itemUrl
    ? '<p style="margin:20px 0 0;"><a href="' +
      escapeHtml(options.itemUrl) +
      '" style="background:' +
      accent +
      ';color:#ffffff;text-decoration:none;padding:9px 20px;border-radius:4px;font-size:14px;display:inline-block;">' +
      escapeHtml(msg('Logic_EmailViewItem', undefined, options.messages)) +
      '</a></p>'
    : '';

  return (
    '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;margin:0 auto;">' +
    '<div style="border-top:4px solid ' +
    accent +
    ';background:#faf9f8;border-radius:0 0 8px 8px;padding:20px 24px;">' +
    '<h2 style="margin:0 0 4px;font-size:20px;color:#323130;">' +
    escapeHtml(options.heading) +
    '</h2>' +
    '<p style="margin:0;font-size:13px;color:#605e5c;">' +
    escapeHtml(options.intro) +
    '</p>' +
    '</div>' +
    '<div style="padding:8px 24px 24px;">' +
    '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;">' +
    rows +
    '</table>' +
    button +
    '<p style="margin:24px 0 0;font-size:11px;color:#a19f9d;">' +
    escapeHtml(msg('Logic_EmailFooter', undefined, options.messages)) +
    '</p>' +
    '</div>' +
    '</div>'
  );
};

// ---------------------------------------------------------------------------
// export
// ---------------------------------------------------------------------------

/** A cell that is already a real number (optionally signed / grouped / currency / percent). */
const NUMERIC_CELL = /^[+-]?[$\u20AC\u00A3\u00A5]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][+-]?\d+)?%?$/;

const isNumericCell = (text: string): boolean => /\d/.test(text) && NUMERIC_CELL.test(text);

const csvEscape = (value: string): string => {
  const text = value === undefined || value === null ? '' : String(value);
  // a leading =, +, -, @, tab or CR makes Excel treat the cell as a formula —
  // but a genuine number such as -5 must stay numeric, so only guard text
  const guarded = /^[=+\-@\t\r]/.test(text) && !isNumericCell(text) ? "'" + text : text;
  if (/[",\r\n]/.test(guarded)) {
    return '"' + guarded.replace(/"/g, '""') + '"';
  }
  return guarded;
};

export interface ICsvRow {
  id?: number;
  created: Date;
  createdBy: string;
  values: Record<string, unknown>;
}

export interface ICsvOptions {
  /** locale for the Submitted column and date answers (browser default when omitted) */
  locale?: string;
  /** translated Logic_ strings (headers, Yes/No labels) */
  messages?: IMessageBag;
}

/** Build CSV text for the responses grid. */
export const buildCsv = (fields: IFormField[], rows: ICsvRow[], options?: ICsvOptions): string => {
  const locale = options ? options.locale : undefined;
  const messages = options ? options.messages : undefined;
  const columns = fields.filter((f) => isInputType(f.type));
  const header = [
    msg('Logic_CsvResponseId', undefined, messages),
    msg('Logic_CsvSubmitted', undefined, messages),
    msg('Logic_CsvSubmittedBy', undefined, messages)
  ].concat(columns.map((f) => f.title));
  const lines = [header.map(csvEscape).join(',')];
  rows.forEach((row) => {
    const cells = [
      row.id === undefined ? '' : String(row.id),
      row.created.toLocaleString(locale),
      row.createdBy
    ].concat(columns.map((f) => formatSharePointValue(f, row.values[f.internalName], messages, locale)));
    lines.push(cells.map(csvEscape).join(','));
  });
  // UTF-8 BOM so Excel opens the file with correct encoding
  return '\uFEFF' + lines.join('\r\n');
};

export const downloadTextFile = (fileName: string, mimeType: string, content: string): void => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// ---------------------------------------------------------------------------
// persistence
// ---------------------------------------------------------------------------

/** Safe JSON parse of a persisted form definition, migrated to the current schema. */
export const parseFormDefinition = (json: string | undefined): IFormDefinition | undefined => {
  if (!json) {
    return undefined;
  }
  try {
    const parsed = JSON.parse(json) as IFormDefinition;
    if (parsed && Array.isArray(parsed.sections) && parsed.settings) {
      return migrateDefinition(parsed);
    }
  } catch {
    // fall through — treat unparseable definitions as absent
  }
  return undefined;
};

export const isSchemaCurrent = (definition: IFormDefinition): boolean =>
  definition.schemaVersion === CURRENT_SCHEMA_VERSION;

// ---------------------------------------------------------------------------
// designer pre-flight
// ---------------------------------------------------------------------------

export type IssueSeverity = 'error' | 'warning';

export interface IFormIssue {
  severity: IssueSeverity;
  /** English (or translated, when a message bag was passed) text */
  message: string;
  /** Logic_ message key, so callers can re-render in another language */
  code?: string;
  /** tokens used to build `message` */
  params?: { [token: string]: string | number };
  /** field the owner should jump to, when the issue is field-specific */
  fieldId?: string;
  sectionId?: string;
}

/**
 * Problems worth surfacing before the owner provisions columns and hands out a
 * link. Errors block publishing; warnings are advisory. `messages` optionally
 * supplies translated Logic_Issue_* strings.
 */
export const validateDefinition = (definition: IFormDefinition, messages?: IMessageBag): IFormIssue[] => {
  const issues: IFormIssue[] = [];
  const add = (
    severity: IssueSeverity,
    code: string,
    params?: { [token: string]: string | number },
    where?: { fieldId?: string; sectionId?: string }
  ): void => {
    issues.push({
      severity: severity,
      code: code,
      params: params,
      message: msg(code, params, messages),
      fieldId: where ? where.fieldId : undefined,
      sectionId: where ? where.sectionId : undefined
    });
  };

  const fields = allFields(definition);
  const inputs = fields.filter((f) => isInputType(f.type));

  if (inputs.length === 0) {
    add('error', 'Logic_Issue_NoQuestions');
  }

  const titleCounts: { [key: string]: number } = {};
  inputs.forEach((field) => {
    const key = (field.title || '').trim().toLowerCase();
    if (key) {
      titleCounts[key] = (titleCounts[key] || 0) + 1;
    }
  });

  fields.forEach((field) => {
    const label = field.title
      ? '"' + field.title + '"'
      : msg('Logic_UntitledQuestion', undefined, messages);
    const where = { fieldId: field.id };

    if (isInputType(field.type) && !(field.title || '').trim()) {
      add('warning', 'Logic_Issue_NoTitle', undefined, where);
    }

    if (isInputType(field.type) && titleCounts[(field.title || '').trim().toLowerCase()] > 1) {
      add('warning', 'Logic_Issue_DuplicateTitle', { label: label }, where);
    }

    const needsChoices = field.type === FieldType.Choice || field.type === FieldType.Ranking;
    if (needsChoices && effectiveChoices(field).length < 2) {
      add('error', 'Logic_Issue_NeedTwoOptions', { label: label }, where);
    }

    if (field.type === FieldType.ImageChoice) {
      const options = field.imageChoices || [];
      if (options.length < 2) {
        add('error', 'Logic_Issue_NeedTwoOptions', { label: label }, where);
      }
      if (options.filter((o) => !(o.imageUrl || '').trim()).length > 0) {
        add('warning', 'Logic_Issue_ImageNoImage', { label: label }, where);
      }
    }

    if (field.type === FieldType.Likert) {
      if ((field.likertRows || []).filter((r) => (r || '').trim()).length === 0) {
        add('error', 'Logic_Issue_LikertNeedRow', { label: label }, where);
      }
      if ((field.likertColumns || []).filter((c) => (c || '').trim()).length < 2) {
        add('error', 'Logic_Issue_LikertNeedColumns', { label: label }, where);
      }
    }

    if (field.type === FieldType.Lookup && !field.lookupListId) {
      add('error', 'Logic_Issue_NoLookupList', { label: label }, where);
    }

    if (field.type === FieldType.Calculated) {
      if (!(field.formula || '').trim()) {
        add('error', 'Logic_Issue_NoFormula', { label: label }, where);
      } else if (evaluateFormula(field.formula || '', () => 1) === undefined) {
        add('error', 'Logic_Issue_BadFormula', { label: label }, where);
      } else {
        const missing = (field.formula || '')
          .split(/[{}]/)
          .filter((_part, index) => index % 2 === 1)
          .map((name) => name.trim())
          .filter((name) => name.length > 0 && !findFieldByTitle(definition, name));
        if (missing.length > 0) {
          add('warning', 'Logic_Issue_MissingRefs', { label: label, names: missing.join(', ') }, where);
        }
      }
    }

    if (field.pattern && !compilePattern(field.pattern)) {
      add('error', 'Logic_Issue_BadPattern', { label: label }, where);
    }

    if (field.required && field.readOnly && field.type !== FieldType.Calculated) {
      add('warning', 'Logic_Issue_RequiredReadOnly', { label: label }, where);
    }
  });

  // branching rules (field and section level)
  const sectionTitleOf: { [id: string]: string } = {};
  definition.sections.forEach((section) => {
    sectionTitleOf[section.id] = section.title;
  });
  const reported: { [key: string]: boolean } = {};
  conditionIssues(definition).forEach((issue) => {
    if (issue.severity !== 'broken') {
      return;
    }
    const owner = issue.fieldId ? findField(definition, issue.fieldId) : undefined;
    const label = owner
      ? owner.title
        ? '"' + owner.title + '"'
        : msg('Logic_UntitledQuestion', undefined, messages)
      : '"' + (sectionTitleOf[issue.sectionId || ''] || '') + '"';
    const code =
      issue.reason === 'missingDriver'
        ? 'Logic_Issue_BrokenRule'
        : issue.reason === 'selfReference'
          ? 'Logic_Issue_SelfRule'
          : 'Logic_Issue_RuleMismatch';
    const key = (issue.fieldId || issue.sectionId || '') + '|' + code;
    if (reported[key]) {
      return;
    }
    reported[key] = true;
    add('error', code, { label: label }, { fieldId: issue.fieldId, sectionId: issue.fieldId ? undefined : issue.sectionId });
  });

  definition.sections.forEach((section) => {
    if (definition.sections.length > 1 && (section.fields || []).length === 0) {
      add('warning', 'Logic_Issue_EmptySection', undefined, { sectionId: section.id });
    }
  });

  if (definition.settings.openDate && definition.settings.closeDate) {
    const open = new Date(definition.settings.openDate);
    const close = new Date(definition.settings.closeDate);
    if (!isNaN(open.getTime()) && !isNaN(close.getTime()) && close <= open) {
      add('error', 'Logic_Issue_CloseBeforeOpen');
    }
  }

  if (definition.settings.approvalNotify && !definition.settings.enableApproval) {
    add('warning', 'Logic_Issue_ApprovalNotifyWithoutApproval');
  }

  return issues;
};

// ---------------------------------------------------------------------------
// open / closed state
// ---------------------------------------------------------------------------

export type FormAvailability = 'open' | 'notYetOpen' | 'closed' | 'full' | 'alreadyAnswered';

export interface IAvailability {
  state: FormAvailability;
  message: string;
}

/** Whether the form is currently accepting responses. `context.messages` / `context.locale` localize the text. */
export const formAvailability = (
  definition: IFormDefinition,
  context: {
    responseCount?: number;
    alreadyAnswered?: boolean;
    now?: Date;
    messages?: IMessageBag;
    locale?: string;
  }
): IAvailability => {
  const settings = definition.settings;
  const now = context.now || new Date();
  const fallback = settings.closedMessage || msg('Logic_ClosedDefault', undefined, context.messages);

  if (settings.openDate) {
    const open = new Date(settings.openDate);
    if (!isNaN(open.getTime()) && now < open) {
      return {
        state: 'notYetOpen',
        message: msg('Logic_NotYetOpen', { date: open.toLocaleString(context.locale) }, context.messages)
      };
    }
  }
  if (settings.closeDate) {
    const close = new Date(settings.closeDate);
    if (!isNaN(close.getTime()) && now > close) {
      return { state: 'closed', message: fallback };
    }
  }
  if (
    typeof settings.maxResponses === 'number' &&
    settings.maxResponses > 0 &&
    typeof context.responseCount === 'number' &&
    context.responseCount >= settings.maxResponses
  ) {
    return { state: 'full', message: fallback };
  }
  if (settings.oneResponsePerPerson && context.alreadyAnswered) {
    return {
      state: 'alreadyAnswered',
      message: msg('Logic_AlreadyResponded', undefined, context.messages)
    };
  }
  return { state: 'open', message: '' };
};

// ---------------------------------------------------------------------------
// dashboard helpers
// ---------------------------------------------------------------------------

/** Chart kinds that make sense for a given field type. */
export const availableCharts = (field: IFormField): ChartKind[] => {
  switch (field.type) {
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
    case FieldType.YesNo:
    case FieldType.Consent:
      return ['bar', 'column', 'donut', 'table'];
    case FieldType.Rating:
    case FieldType.Scale:
      return ['bar', 'column', 'gauge', 'stat', 'histogram'];
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider:
      return ['stat', 'histogram', 'column'];
    case FieldType.Likert:
      return ['bar', 'table'];
    case FieldType.Ranking:
      return ['bar', 'table'];
    case FieldType.Date:
    case FieldType.Time:
      return ['column', 'table'];
    case FieldType.Person:
      return ['bar', 'donut', 'table'];
    case FieldType.Text:
    case FieldType.MultilineText:
    case FieldType.RichText:
    case FieldType.Address:
      return ['words', 'table'];
    default:
      return ['table'];
  }
};

/** The chart used when a tile is set to "auto". */
export const defaultChartFor = (field: IFormField): ChartKind => {
  const options = availableCharts(field);
  switch (field.type) {
    case FieldType.YesNo:
    case FieldType.Consent:
      return 'donut';
    case FieldType.Scale:
      return field.scaleAnalytics === 'nps' ? 'gauge' : 'column';
    case FieldType.Rating:
      return 'bar';
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider:
      return 'stat';
    case FieldType.Text:
    case FieldType.MultilineText:
    case FieldType.RichText:
      return 'words';
    default:
      return options[0] || 'table';
  }
};

/** Fields eligible as a cross-tab segment (low-cardinality categoricals). */
export const segmentableFields = (definition: IFormDefinition): IFormField[] =>
  inputFields(definition).filter(
    (f) =>
      f.type === FieldType.Choice ||
      f.type === FieldType.YesNo ||
      f.type === FieldType.Consent ||
      f.type === FieldType.Lookup ||
      f.type === FieldType.Person ||
      f.type === FieldType.ImageChoice ||
      (f.type === FieldType.Rating && (f.maxRating || 5) <= 10)
  );

/** Operator labels shared by the branching editor. */
export const OPERATOR_LABELS: { key: ConditionOperator; text: string }[] = [
  { key: 'equals', text: 'is' },
  { key: 'notEquals', text: 'is not' },
  { key: 'contains', text: 'contains' },
  { key: 'notContains', text: 'does not contain' },
  { key: 'notEmpty', text: 'is answered' },
  { key: 'empty', text: 'is not answered' },
  { key: 'greaterThan', text: 'is greater than' },
  { key: 'greaterOrEqual', text: 'is at least' },
  { key: 'lessThan', text: 'is less than' },
  { key: 'lessOrEqual', text: 'is at most' },
  { key: 'between', text: 'is between' },
  { key: 'before', text: 'is before' },
  { key: 'after', text: 'is after' }
];

/** Operator labels in the caller's language (falls back to the English defaults). */
export const operatorLabels = (messages?: IMessageBag): { key: ConditionOperator; text: string }[] =>
  OPERATOR_LABELS.map((entry) => ({ key: entry.key, text: msg('Logic_Op_' + entry.key, undefined, messages) }));

/** Operators offered for a given driver field type. */
export const operatorsForField = (field: IFormField | undefined): ConditionOperator[] => {
  if (!field) {
    return ['equals', 'notEquals', 'empty', 'notEmpty'];
  }
  switch (field.type) {
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider:
    case FieldType.Scale:
    case FieldType.Rating:
      return ['equals', 'notEquals', 'greaterThan', 'greaterOrEqual', 'lessThan', 'lessOrEqual', 'between', 'empty', 'notEmpty'];
    case FieldType.Date:
    case FieldType.Time:
      return ['equals', 'notEquals', 'before', 'after', 'between', 'empty', 'notEmpty'];
    case FieldType.YesNo:
    case FieldType.Consent:
      return ['equals', 'notEquals'];
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
      return ['equals', 'notEquals', 'contains', 'notContains', 'empty', 'notEmpty'];
    default:
      return ['equals', 'notEquals', 'contains', 'notContains', 'empty', 'notEmpty'];
  }
};
