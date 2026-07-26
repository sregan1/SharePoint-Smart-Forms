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
  isInputType,
  migrateDefinition,
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
 */
export const generateInternalName = (title: string, existing: string[]): string => {
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
  const lower = existing.map((e) => e.toLowerCase());
  while (lower.indexOf(candidate.toLowerCase()) !== -1) {
    candidate = base + suffix;
    suffix++;
  }
  return candidate;
};

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

/** Returns an error message, or undefined when the value is valid. */
export const validateField = (field: IFormField, value: unknown): string | undefined => {
  if (!isInputType(field.type)) {
    return undefined;
  }
  const label = field.title || 'This question';
  const empty = isEmptyValue(field, value);

  if (field.required && empty) {
    if (field.requiredMessage) {
      return field.requiredMessage;
    }
    switch (field.type) {
      case FieldType.Consent:
        return 'Please tick the box to continue';
      case FieldType.FileUpload:
        return 'Please attach at least one file';
      case FieldType.Signature:
        return 'Please add your signature';
      case FieldType.Likert:
        return 'Please answer every row';
      default:
        return label + ' is required';
    }
  }
  if (empty) {
    return undefined;
  }

  switch (field.type) {
    case FieldType.Email:
      if (!EMAIL_REGEX.test(String(value).trim())) {
        return 'Enter a valid email address';
      }
      break;

    case FieldType.Phone:
      if (!PHONE_REGEX.test(String(value).trim())) {
        return 'Enter a valid phone number';
      }
      break;

    case FieldType.Hyperlink:
      if (!URL_REGEX.test((value as IHyperlinkValue).url.trim())) {
        return 'Enter a valid web address starting with http:// or https://';
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
        return 'Value must be at least ' + formatNumberForMessage(field, min);
      }
      if (typeof max === 'number' && num > max) {
        return 'Value must be at most ' + formatNumberForMessage(field, max);
      }
      break;
    }

    case FieldType.Text:
    case FieldType.MultilineText:
      if (field.maxLength && String(value).length > field.maxLength) {
        return 'Maximum length is ' + field.maxLength + ' characters';
      }
      break;

    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup: {
      const count = selectionCount(value);
      if (field.allowMultiple && typeof field.minSelections === 'number' && count < field.minSelections) {
        return 'Choose at least ' + field.minSelections + (field.minSelections === 1 ? ' option' : ' options');
      }
      if (field.allowMultiple && typeof field.maxSelections === 'number' && count > field.maxSelections) {
        return 'Choose no more than ' + field.maxSelections + (field.maxSelections === 1 ? ' option' : ' options');
      }
      break;
    }

    case FieldType.Likert: {
      const answers = (value || {}) as ILikertValue;
      const rows = field.likertRows || [];
      const unanswered = rows.filter((row) => !answers[row]);
      if (field.required && unanswered.length > 0) {
        return 'Please answer every row (' + unanswered.length + ' remaining)';
      }
      break;
    }

    case FieldType.FileUpload: {
      const files = (value || []) as IFormFile[];
      const maxFiles = field.maxFiles || 3;
      if (files.length > maxFiles) {
        return 'Attach no more than ' + maxFiles + (maxFiles === 1 ? ' file' : ' files');
      }
      const maxBytes = (field.maxFileSizeMb || 10) * 1024 * 1024;
      const tooBig = files.filter((f) => f.size > maxBytes)[0];
      if (tooBig) {
        return '"' + tooBig.name + '" is larger than ' + (field.maxFileSizeMb || 10) + ' MB';
      }
      const allowed = (field.allowedExtensions || []).map((e) => e.toLowerCase().replace(/^\./, ''));
      if (allowed.length > 0) {
        const wrong = files.filter((f) => allowed.indexOf(fileExtension(f.name)) === -1)[0];
        if (wrong) {
          return '"' + wrong.name + '" is not an accepted file type (' + allowed.join(', ') + ')';
        }
      }
      break;
    }

    case FieldType.Address: {
      const address = value as IAddressValue;
      if (address && typeof address !== 'string' && field.required) {
        if (!(address.street || '').trim() || !(address.city || '').trim()) {
          return 'Enter at least a street and a city';
        }
      }
      break;
    }
  }

  // custom pattern applies to any text-bearing answer
  if (field.pattern) {
    const text = String(typeof value === 'object' ? formatValue(field, value) : value);
    let expression: RegExp | undefined;
    try {
      expression = new RegExp(field.pattern);
    } catch {
      expression = undefined; // an invalid pattern must never block a submission
    }
    if (expression && !expression.test(text)) {
      return field.patternMessage || 'This answer is not in the expected format';
    }
  }

  return undefined;
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
    return value.getTime();
  }
  if (Array.isArray(value)) {
    return value.length;
  }
  const num = Number(String(value === undefined || value === null ? '' : value).replace(/,/g, '').trim());
  return isNaN(num) ? undefined : num;
};

/** Parse a rule's comparison operand, honoring dates and "today". */
const operandAsNumber = (field: IFormField | undefined, raw: string | undefined): number | undefined => {
  const text = (raw || '').trim();
  if (text.length === 0) {
    return undefined;
  }
  const isDateField = field && (field.type === FieldType.Date || field.type === FieldType.Time);
  if (isDateField || /^today$/i.test(text)) {
    if (/^today$/i.test(text)) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      return today.getTime();
    }
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return parsed.getTime();
    }
  }
  const num = Number(text.replace(/,/g, ''));
  return isNaN(num) ? undefined : num;
};

const evaluateCondition = (
  condition: ICondition,
  definition: IFormDefinition,
  values: IFormValues
): boolean => {
  const driver = findField(definition, condition.fieldId);
  if (!driver) {
    // a rule pointing at a deleted question can never be satisfied — treating
    // it as "true" used to make branched questions unconditionally visible
    return false;
  }
  const raw = values[driver.id];
  const text = valueAsComparableString(driver, raw);
  const operator = condition.operator;

  if (UNARY_OPERATORS.indexOf(operator) !== -1) {
    const empty = isEmptyValue(driver, raw);
    return operator === 'empty' ? empty : !empty;
  }

  if (RANGE_OPERATORS.indexOf(operator) !== -1) {
    const actual = asNumber(raw);
    const bound = operandAsNumber(driver, condition.value);
    if (actual === undefined || bound === undefined) {
      return false;
    }
    switch (operator) {
      case 'greaterThan':
        return actual > bound;
      case 'greaterOrEqual':
        return actual >= bound;
      case 'lessThan':
        return actual < bound;
      case 'lessOrEqual':
        return actual <= bound;
      case 'after':
        return actual > bound;
      case 'before':
        return actual < bound;
      case 'between': {
        const upper = operandAsNumber(driver, condition.value2);
        if (upper === undefined) {
          return false;
        }
        const low = Math.min(bound, upper);
        const high = Math.max(bound, upper);
        return actual >= low && actual <= high;
      }
      default:
        return false;
    }
  }

  const expected = (condition.value || '').toLowerCase();
  // multi-value answers match if *any* selection matches
  const parts = text.split(';').map((p) => p.trim().toLowerCase()).filter((p) => p.length > 0);
  const haystack = text.toLowerCase();

  switch (operator) {
    case 'equals':
      return parts.length > 1 ? parts.indexOf(expected) !== -1 : haystack === expected;
    case 'notEquals':
      return parts.length > 1 ? parts.indexOf(expected) === -1 : haystack !== expected;
    case 'contains':
      return haystack.indexOf(expected) !== -1;
    case 'notContains':
      return haystack.indexOf(expected) === -1;
    default:
      return true;
  }
};

/** Evaluate a condition group. An empty or absent group is always satisfied. */
export const evaluateConditionGroup = (
  group: IConditionGroup | undefined,
  definition: IFormDefinition,
  values: IFormValues,
  selfFieldId?: string
): boolean => {
  if (!group || !group.conditions || group.conditions.length === 0) {
    return true;
  }
  const usable = group.conditions.filter((c) => c.fieldId && c.fieldId !== selfFieldId);
  if (usable.length === 0) {
    return true;
  }
  const results = usable.map((c) => evaluateCondition(c, definition, values));
  return group.match === 'any'
    ? results.filter((r) => r).length > 0
    : results.filter((r) => !r).length === 0;
};

/** Evaluate a field's visibility, including its section's own rule. */
export const isFieldVisible = (
  field: IFormField,
  definition: IFormDefinition,
  values: IFormValues
): boolean => {
  const section = (definition.sections || []).filter(
    (s) => (s.fields || []).filter((f) => f.id === field.id).length > 0
  )[0];
  if (section && !evaluateConditionGroup(section.visibleWhen, definition, values)) {
    return false;
  }
  return evaluateConditionGroup(field.visibleWhen, definition, values, field.id);
};

export const isSectionVisible = (
  section: IFormSection,
  definition: IFormDefinition,
  values: IFormValues
): boolean => evaluateConditionGroup(section.visibleWhen, definition, values);

/** Remove rules that reference fields no longer in the definition. */
export const pruneDanglingConditions = (definition: IFormDefinition): IFormDefinition => {
  const ids: { [id: string]: boolean } = {};
  allFields(definition).forEach((f) => {
    ids[f.id] = true;
  });
  const prune = (group: IConditionGroup | undefined): IConditionGroup | undefined => {
    if (!group) {
      return undefined;
    }
    const conditions = (group.conditions || []).filter((c) => ids[c.fieldId]);
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
      const fieldRule = prune(field.visibleWhen);
      if (fieldRule) {
        field.visibleWhen = fieldRule;
      } else {
        delete field.visibleWhen;
      }
    });
  });
  return definition;
};

// ---------------------------------------------------------------------------
// calculated fields
// ---------------------------------------------------------------------------

/**
 * Recompute every Calculated field from the current answers. Runs a couple of
 * passes so a calculation that references another calculation still resolves.
 */
export const applyCalculatedFields = (definition: IFormDefinition, values: IFormValues): IFormValues => {
  const calculated = allFields(definition).filter((f) => f.type === FieldType.Calculated && f.formula);
  if (calculated.length === 0) {
    return values;
  }
  const next: IFormValues = { ...values };
  const passes = Math.min(calculated.length, 5);
  for (let pass = 0; pass < passes; pass++) {
    let changed = false;
    calculated.forEach((field) => {
      const result = evaluateFormula(field.formula || '', (name) => {
        const referenced = findFieldByTitle(definition, name);
        if (!referenced) {
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

const splitMulti = (text: string): string[] =>
  String(text || '')
    .split(';')
    .map((v) => v.trim())
    .filter((v) => v.length > 0);

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
export const formatValue = (field: IFormField, value: unknown): string => {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  switch (field.type) {
    case FieldType.YesNo:
      return value === true ? 'Yes' : 'No';
    case FieldType.Consent:
      return value === true ? 'Agreed' : 'Not agreed';
    case FieldType.Date:
      if (value instanceof Date) {
        return field.includeTime ? value.toLocaleString() : value.toLocaleDateString();
      }
      return String(value);
    case FieldType.Time:
      return value instanceof Date
        ? value.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
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
      return String(value).indexOf('data:') === 0 ? 'Signed' : String(value);
    case FieldType.RichText:
      return stripHtml(String(value));
    case FieldType.Content:
      return '';
    default:
      return String(value);
  }
};

/** Convenience wrapper: normalize a SharePoint value and format it. */
export const formatSharePointValue = (field: IFormField, raw: unknown): string =>
  formatValue(field, normalizeFromSharePoint(field, raw));

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
      .join(';');
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
    return Object.keys(answers).map((row) => answers[row]).filter((a) => a).join(';');
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
}

/** Inline-styled HTML email summarizing one response — safe for Outlook. */
export const buildResponseEmailHtml = (
  definition: IFormDefinition,
  values: IFormValues,
  options: IResponseEmailOptions
): string => {
  const accent = options.accentColor || '#0078d4';
  const rows = inputFields(definition)
    .filter((field) => isFieldVisible(field, definition, values))
    .map((field) => {
      const answer = formatValue(field, values[field.id]);
      return (
        '<tr>' +
        '<td style="padding:8px 16px 8px 0;font-size:13px;color:#605e5c;vertical-align:top;white-space:nowrap;">' +
        escapeHtml(field.title || 'Question') +
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
      'View in Microsoft Lists</a></p>'
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
    '<p style="margin:24px 0 0;font-size:11px;color:#a19f9d;">Sent automatically by Smart Forms.</p>' +
    '</div>' +
    '</div>'
  );
};

// ---------------------------------------------------------------------------
// export
// ---------------------------------------------------------------------------

const csvEscape = (value: string): string => {
  const text = value === undefined || value === null ? '' : String(value);
  // a leading =, +, - or @ makes Excel treat the cell as a formula
  const guarded = /^[=+\-@]/.test(text) ? "'" + text : text;
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

/** Build CSV text for the responses grid. */
export const buildCsv = (fields: IFormField[], rows: ICsvRow[]): string => {
  const columns = fields.filter((f) => isInputType(f.type));
  const header = ['Response ID', 'Submitted', 'Submitted by'].concat(columns.map((f) => f.title));
  const lines = [header.map(csvEscape).join(',')];
  rows.forEach((row) => {
    const cells = [
      row.id === undefined ? '' : String(row.id),
      row.created.toLocaleString(),
      row.createdBy
    ].concat(columns.map((f) => formatSharePointValue(f, row.values[f.internalName])));
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
  message: string;
  /** field the owner should jump to, when the issue is field-specific */
  fieldId?: string;
  sectionId?: string;
}

/**
 * Problems worth surfacing before the owner provisions columns and hands out a
 * link. Errors block publishing; warnings are advisory.
 */
export const validateDefinition = (definition: IFormDefinition): IFormIssue[] => {
  const issues: IFormIssue[] = [];
  const fields = allFields(definition);
  const inputs = fields.filter((f) => isInputType(f.type));

  if (inputs.length === 0) {
    issues.push({ severity: 'error', message: 'Add at least one question before collecting responses.' });
  }

  const titleCounts: { [key: string]: number } = {};
  inputs.forEach((field) => {
    const key = (field.title || '').trim().toLowerCase();
    if (key) {
      titleCounts[key] = (titleCounts[key] || 0) + 1;
    }
  });

  fields.forEach((field) => {
    const label = field.title ? '"' + field.title + '"' : 'An untitled question';

    if (isInputType(field.type) && !(field.title || '').trim()) {
      issues.push({
        severity: 'warning',
        fieldId: field.id,
        message: 'A question has no text — it will be named automatically when you publish.'
      });
    }

    if (isInputType(field.type) && titleCounts[(field.title || '').trim().toLowerCase()] > 1) {
      issues.push({
        severity: 'warning',
        fieldId: field.id,
        message: label + ' shares its text with another question, which makes results harder to read.'
      });
    }

    const needsChoices =
      field.type === FieldType.Choice || field.type === FieldType.Ranking;
    if (needsChoices && effectiveChoices(field).length < 2) {
      issues.push({
        severity: 'error',
        fieldId: field.id,
        message: label + ' needs at least two options.'
      });
    }

    if (field.type === FieldType.ImageChoice) {
      const options = field.imageChoices || [];
      if (options.length < 2) {
        issues.push({ severity: 'error', fieldId: field.id, message: label + ' needs at least two options.' });
      }
      if (options.filter((o) => !(o.imageUrl || '').trim()).length > 0) {
        issues.push({
          severity: 'warning',
          fieldId: field.id,
          message: label + ' has options with no image — they will show as text tiles.'
        });
      }
    }

    if (field.type === FieldType.Likert) {
      if ((field.likertRows || []).filter((r) => (r || '').trim()).length === 0) {
        issues.push({ severity: 'error', fieldId: field.id, message: label + ' needs at least one row.' });
      }
      if ((field.likertColumns || []).filter((c) => (c || '').trim()).length < 2) {
        issues.push({ severity: 'error', fieldId: field.id, message: label + ' needs at least two scale columns.' });
      }
    }

    if (field.type === FieldType.Lookup && !field.lookupListId) {
      issues.push({
        severity: 'error',
        fieldId: field.id,
        message: label + ' has no source list selected.'
      });
    }

    if (field.type === FieldType.Calculated) {
      if (!(field.formula || '').trim()) {
        issues.push({ severity: 'error', fieldId: field.id, message: label + ' has no formula.' });
      } else if (evaluateFormula(field.formula || '', () => 1) === undefined) {
        issues.push({
          severity: 'error',
          fieldId: field.id,
          message: label + ' has a formula that cannot be worked out. Check the brackets and operators.'
        });
      } else {
        const missing = (field.formula || '')
          .split(/[{}]/)
          .filter((_part, index) => index % 2 === 1)
          .map((name) => name.trim())
          .filter((name) => name.length > 0 && !findFieldByTitle(definition, name));
        if (missing.length > 0) {
          issues.push({
            severity: 'warning',
            fieldId: field.id,
            message: label + ' refers to questions that do not exist: ' + missing.join(', ') + '.'
          });
        }
      }
    }

    if (field.pattern) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-unused-expressions
        new RegExp(field.pattern);
      } catch {
        issues.push({
          severity: 'error',
          fieldId: field.id,
          message: label + ' has an invalid validation pattern.'
        });
      }
    }

    if (field.visibleWhen) {
      const broken = (field.visibleWhen.conditions || []).filter(
        (c) => !c.fieldId || !findField(definition, c.fieldId)
      );
      if (broken.length > 0) {
        issues.push({
          severity: 'error',
          fieldId: field.id,
          message: label + ' has a branching rule pointing at a deleted question.'
        });
      }
      const selfReference = (field.visibleWhen.conditions || []).filter((c) => c.fieldId === field.id);
      if (selfReference.length > 0) {
        issues.push({
          severity: 'error',
          fieldId: field.id,
          message: label + ' has a branching rule that refers to itself.'
        });
      }
    }

    if (field.required && field.readOnly && field.type !== FieldType.Calculated) {
      issues.push({
        severity: 'warning',
        fieldId: field.id,
        message: label + ' is both required and read-only, so it can only be answered by a prefilled link.'
      });
    }
  });

  definition.sections.forEach((section) => {
    if (definition.sections.length > 1 && (section.fields || []).length === 0) {
      issues.push({
        severity: 'warning',
        sectionId: section.id,
        message: 'A section has no questions and will be skipped.'
      });
    }
  });

  if (definition.settings.openDate && definition.settings.closeDate) {
    const open = new Date(definition.settings.openDate);
    const close = new Date(definition.settings.closeDate);
    if (!isNaN(open.getTime()) && !isNaN(close.getTime()) && close <= open) {
      issues.push({
        severity: 'error',
        message: 'The close date is on or before the open date, so the form will never accept responses.'
      });
    }
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

/** Whether the form is currently accepting responses. */
export const formAvailability = (
  definition: IFormDefinition,
  context: { responseCount?: number; alreadyAnswered?: boolean; now?: Date }
): IAvailability => {
  const settings = definition.settings;
  const now = context.now || new Date();
  const fallback = settings.closedMessage || 'This form is no longer accepting responses.';

  if (settings.openDate) {
    const open = new Date(settings.openDate);
    if (!isNaN(open.getTime()) && now < open) {
      return {
        state: 'notYetOpen',
        message: 'This form opens on ' + open.toLocaleString() + '.'
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
    return { state: 'alreadyAnswered', message: 'You have already responded to this form. Thank you!' };
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
      return ['before', 'after', 'between', 'empty', 'notEmpty'];
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
