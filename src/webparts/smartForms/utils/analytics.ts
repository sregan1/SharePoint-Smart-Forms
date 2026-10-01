import { FieldType, IFormField, ILikertValue, IMessageBag, IResponseItem, msg } from '../models';
import {
  effectiveChoices,
  formatSharePointValue,
  normalizeFromSharePoint,
  splitMultiValue
} from './formUtils';
import { stripHtml } from './sanitizeHtml';

/**
 * Response analytics.
 *
 * Pure functions over the loaded response set — no React, no SharePoint — so the
 * views can memoize them and the tests can exercise them directly.
 *
 * Chart rows and segments carry two separate things: a display `label` (which
 * may be translated or decorated, e.g. "4 ★") and the raw `filterValue` used to
 * find the matching responses again (e.g. "4"). Never filter by label; use
 * matchesRowFilter().
 */

export interface IDistributionRow {
  /** display text (localized / decorated) */
  label: string;
  count: number;
  /**
   * The raw answer key this row stands for: "4" for a "4 ★" rating row,
   * "Yes" for a localized Yes/No row, "2026-03-05" for a Date row, "09" for a
   * Time row grouped by hour. Undefined only on a folded "Other" row.
   */
  filterValue?: string;
  /** set on a folded "Other" row: every raw key that was folded into it */
  filterValues?: string[];
}

export const isAnswered = (value: unknown): boolean => {
  if (value === undefined || value === null || value === '') {
    return false;
  }
  return !(Array.isArray(value) && value.length === 0);
};

/** Yes/No answer as its stable key, or undefined when unanswered (null is NOT "No"). */
const yesNoKey = (raw: unknown): 'Yes' | 'No' | undefined => {
  if (raw === true || raw === 'true' || raw === 1 || raw === 'Yes') {
    return 'Yes';
  }
  if (raw === false || raw === 'false' || raw === 0 || raw === 'No') {
    return 'No';
  }
  return undefined;
};

/** Field-aware "was this question answered": an unanswered Yes/No is not a No. */
export const isFieldAnswered = (field: IFormField, value: unknown): boolean => {
  if (field.type === FieldType.YesNo) {
    return yesNoKey(value) !== undefined;
  }
  return isAnswered(value);
};

export const answeredCount = (field: IFormField, items: IResponseItem[]): number =>
  items.filter((item) => isFieldAnswered(field, item.values[field.internalName])).length;

const pad2 = (n: number): string => (n < 10 ? '0' + n : String(n));

const asDate = (field: IFormField, raw: unknown): Date | undefined => {
  if (raw instanceof Date) {
    return isNaN(raw.getTime()) ? undefined : raw;
  }
  const normalized = normalizeFromSharePoint(field, raw);
  return normalized instanceof Date ? normalized : undefined;
};

/** Local calendar date key, "YYYY-MM-DD". */
const localDateKey = (date: Date): string =>
  date.getFullYear() + '-' + pad2(date.getMonth() + 1) + '-' + pad2(date.getDate());

/**
 * The raw keys one answer contributes to a distribution / segment / filter.
 * This is the single source of truth shared by distributionFor, segmentBy and
 * matchesRowFilter so a clicked bar always selects exactly the responses it counted.
 * Empty array = unanswered.
 */
export const answerKeys = (field: IFormField, raw: unknown): string[] => {
  if (!isFieldAnswered(field, raw)) {
    return [];
  }
  switch (field.type) {
    case FieldType.YesNo: {
      const key = yesNoKey(raw);
      return key ? [key] : [];
    }
    case FieldType.Consent:
      return [raw === true || raw === 'true' ? 'Agreed' : 'Not agreed'];
    case FieldType.Person: {
      const normalized = normalizeFromSharePoint(field, raw);
      return Array.isArray(normalized)
        ? (normalized as { displayName?: string }[])
            .map((person) => (person && person.displayName) || '')
            .filter((name) => name.length > 0)
        : [];
    }
    case FieldType.Date: {
      const date = asDate(field, raw);
      return date ? [localDateKey(date)] : [];
    }
    case FieldType.Time: {
      const date = asDate(field, raw);
      return date ? [pad2(date.getHours())] : [];
    }
    case FieldType.Rating: {
      const num = Number(raw);
      return isNaN(num) ? [] : [String(num)];
    }
    case FieldType.Scale: {
      const num = Number(raw);
      return isNaN(num) ? [] : [String(Math.round(num * 100) / 100)];
    }
    default:
      break;
  }
  if (Array.isArray(raw)) {
    return (raw as unknown[]).map((v) => String(v));
  }
  // multi-value text columns hold "A; B" — split so each selection counts (and filters) on its own
  if (
    field.allowMultiple &&
    (field.type === FieldType.Choice || field.type === FieldType.ImageChoice || field.type === FieldType.Lookup)
  ) {
    return splitMultiValue(String(raw));
  }
  return [String(raw)];
};

/** Display label for a raw answer key. */
const labelForKey = (
  field: IFormField,
  key: string,
  labels?: IMessageBag,
  locale?: string
): string => {
  switch (field.type) {
    case FieldType.YesNo:
      return key === 'Yes' ? msg('Logic_Yes', undefined, labels) : msg('Logic_No', undefined, labels);
    case FieldType.Consent:
      return key === 'Agreed' ? msg('Logic_Agreed', undefined, labels) : msg('Logic_NotAgreed', undefined, labels);
    case FieldType.Date: {
      const parts = key.split('-');
      const date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      return isNaN(date.getTime())
        ? key
        : date.toLocaleDateString(locale, { year: 'numeric', month: 'short', day: 'numeric' });
    }
    case FieldType.Time: {
      const date = new Date(2000, 0, 1, Number(key), 0, 0, 0);
      return isNaN(date.getTime()) ? key : date.toLocaleTimeString(locale, { hour: 'numeric' });
    }
    case FieldType.Rating:
      return key + (field.ratingIcon === 'like' ? '' : ' ★');
    default:
      return key;
  }
};

/**
 * Whether a response falls in a clicked chart row / segment. Pass the row or
 * segment object itself (anything with filterValue / filterValues).
 */
export const matchesRowFilter = (
  field: IFormField,
  item: IResponseItem,
  row: { filterValue?: string; filterValues?: string[] }
): boolean => {
  const keys = answerKeys(field, item.values[field.internalName]);
  if (row.filterValues && row.filterValues.length > 0) {
    return keys.filter((k) => row.filterValues!.indexOf(k) !== -1).length > 0;
  }
  if (row.filterValue === undefined) {
    return false;
  }
  if (row.filterValue === '') {
    return keys.length === 0;
  }
  return keys.indexOf(row.filterValue) !== -1;
};

/** Counts per category for a choice-like field, honoring the designed order. */
export const distributionFor = (
  field: IFormField,
  items: IResponseItem[],
  labels?: IMessageBag,
  locale?: string
): IDistributionRow[] => {
  const counts: { [key: string]: number } = {};
  const bump = (key: string): void => {
    counts[key] = (counts[key] || 0) + 1;
  };

  items.forEach((item) => {
    answerKeys(field, item.values[field.internalName]).forEach(bump);
  });

  const row = (key: string): IDistributionRow => ({
    label: labelForKey(field, key, labels, locale),
    count: counts[key] || 0,
    filterValue: key
  });

  let rows: IDistributionRow[];

  if (field.type === FieldType.Choice || field.type === FieldType.ImageChoice) {
    // keep the designer's option order, then any write-in or legacy values
    const designed =
      field.type === FieldType.Choice
        ? effectiveChoices(field)
        : (field.imageChoices || []).map((o) => o.label).filter((l) => l);
    rows = designed.map(row);
    Object.keys(counts).forEach((key) => {
      if (designed.indexOf(key) === -1) {
        rows.push(row(key));
      }
    });
  } else if (field.type === FieldType.YesNo) {
    rows = [row('Yes'), row('No')];
  } else if (field.type === FieldType.Consent) {
    rows = [row('Agreed'), row('Not agreed')];
  } else if (field.type === FieldType.Rating) {
    const max = field.maxRating || 5;
    rows = [];
    for (let star = max; star >= 1; star--) {
      rows.push(row(String(star)));
      if (field.allowHalfRating && counts[String(star - 0.5)] > 0) {
        rows.push(row(String(star - 0.5)));
      }
    }
  } else if (field.type === FieldType.Scale) {
    const min = typeof field.min === 'number' ? field.min : 1;
    const max = typeof field.max === 'number' ? field.max : 5;
    const step = field.step && field.step > 0 ? field.step : 1;
    rows = [];
    for (let value = min; value <= max; value += step) {
      rows.push(row(String(Math.round(value * 100) / 100)));
    }
  } else if (field.type === FieldType.Date || field.type === FieldType.Time) {
    // chronological, grouped by local date (Time: by hour) — never by label text
    rows = Object.keys(counts)
      .sort()
      .map(row);
  } else {
    rows = Object.keys(counts).map(row);
    rows.sort((a, b) => b.count - a.count);
  }

  return rows;
};

/**
 * Fold a long tail into "Other".
 *
 * A generated ninth hue would be indistinguishable from an existing one under
 * colorblind simulation, so charts that color by category cap their series and
 * collapse the rest rather than inventing colors. The "Other" row has no
 * filterValue; its `filterValues` lists every raw key folded into it.
 */
export const foldTail = (
  rows: IDistributionRow[],
  maxSeries: number,
  labels?: IMessageBag
): IDistributionRow[] => {
  if (rows.length <= maxSeries) {
    return rows;
  }
  const sorted = rows.slice().sort((a, b) => b.count - a.count);
  const head = sorted.slice(0, maxSeries - 1);
  const tail = sorted.slice(maxSeries - 1);
  const otherCount = tail.reduce((sum, row) => sum + row.count, 0);
  const folded: string[] = [];
  tail.forEach((row) => {
    if (row.filterValues) {
      row.filterValues.forEach((v) => folded.push(v));
    } else if (row.filterValue !== undefined) {
      folded.push(row.filterValue);
    } else {
      folded.push(row.label);
    }
  });
  return head.concat([
    {
      label: msg('Logic_OtherSlice', { count: tail.length }, labels),
      count: otherCount,
      filterValues: folded
    }
  ]);
};

// ---------------------------------------------------------------------------
// NPS
// ---------------------------------------------------------------------------

export interface INpsStats {
  score: number;
  promoters: number;
  passives: number;
  detractors: number;
  total: number;
}

const numbersFor = (field: IFormField, items: IResponseItem[]): number[] =>
  items
    .map((item) => item.values[field.internalName])
    .filter((v) => v !== undefined && v !== null && v !== '')
    .map((v) => Number(v))
    .filter((n) => !isNaN(n));

/** Classic NPS: % promoters (9-10) minus % detractors (0-6). */
export const npsStats = (field: IFormField, items: IResponseItem[]): INpsStats | undefined => {
  const scores = numbersFor(field, items).filter((n) => n >= 0 && n <= 10);
  if (scores.length === 0) {
    return undefined;
  }
  const promoters = scores.filter((n) => n >= 9).length;
  const passives = scores.filter((n) => n >= 7 && n <= 8).length;
  const detractors = scores.filter((n) => n <= 6).length;
  return {
    score: Math.round(((promoters - detractors) / scores.length) * 100),
    promoters,
    passives,
    detractors,
    total: scores.length
  };
};

// ---------------------------------------------------------------------------
// ranking
// ---------------------------------------------------------------------------

export interface IRankingRow {
  label: string;
  avgPosition: number;
  votes: number;
  /** count of first-place votes */
  firstPlace: number;
}

/**
 * Average position per option across all responses (1 = ranked first).
 *
 * Stored as JSON when the definition is current; older responses hold a
 * semicolon-delimited string, which is parsed as a fallback.
 */
export const rankingStats = (field: IFormField, items: IResponseItem[]): IRankingRow[] => {
  const positions: { [label: string]: number[] } = {};
  items.forEach((item) => {
    const raw = item.values[field.internalName];
    if (!isAnswered(raw)) {
      return;
    }
    const ordered = Array.isArray(raw)
      ? (raw as unknown[]).map((v) => String(v))
      : splitMultiValue(String(raw));
    ordered.forEach((label, index) => {
      (positions[label] = positions[label] || []).push(index + 1);
    });
  });
  const rows: IRankingRow[] = Object.keys(positions).map((label) => {
    const list = positions[label];
    const sum = list.reduce((acc, n) => acc + n, 0);
    return {
      label: label,
      avgPosition: Math.round((sum / list.length) * 10) / 10,
      votes: list.length,
      firstPlace: list.filter((p) => p === 1).length
    };
  });
  rows.sort((a, b) => a.avgPosition - b.avgPosition);
  return rows;
};

// ---------------------------------------------------------------------------
// numeric
// ---------------------------------------------------------------------------

export interface INumericStats {
  count: number;
  min: number;
  max: number;
  avg: number;
  median: number;
  stdDev: number;
  sum: number;
}

export const numericStats = (field: IFormField, items: IResponseItem[]): INumericStats | undefined => {
  const numbers = numbersFor(field, items);
  if (numbers.length === 0) {
    return undefined;
  }
  const sorted = numbers.slice().sort((a, b) => a - b);
  const sum = numbers.reduce((acc, n) => acc + n, 0);
  const avg = sum / numbers.length;
  const middle = Math.floor(sorted.length / 2);
  const median =
    sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
  const variance = numbers.reduce((acc, n) => acc + Math.pow(n - avg, 2), 0) / numbers.length;
  const round = (n: number): number => Math.round(n * 100) / 100;
  return {
    count: numbers.length,
    min: sorted[0],
    max: sorted[sorted.length - 1],
    avg: round(avg),
    median: round(median),
    stdDev: round(Math.sqrt(variance)),
    sum: round(sum)
  };
};

export interface IHistogramBin {
  label: string;
  count: number;
  from: number;
  to: number;
}

/** Equal-width bins over the observed range. */
export const histogram = (
  field: IFormField,
  items: IResponseItem[],
  binCount?: number
): IHistogramBin[] => {
  const numbers = numbersFor(field, items);
  if (numbers.length === 0) {
    return [];
  }
  const min = Math.min(...numbers);
  const max = Math.max(...numbers);
  if (min === max) {
    return [{ label: String(min), count: numbers.length, from: min, to: max }];
  }
  // Sturges' rule, clamped to something readable on a small tile
  const bins = Math.max(4, Math.min(binCount || Math.ceil(Math.log(numbers.length) / Math.LN2) + 1, 12));
  const width = (max - min) / bins;
  const result: IHistogramBin[] = [];
  for (let i = 0; i < bins; i++) {
    const from = min + i * width;
    const to = i === bins - 1 ? max : from + width;
    result.push({
      label: formatBinBound(from) + '–' + formatBinBound(to),
      count: 0,
      from: from,
      to: to
    });
  }
  numbers.forEach((n) => {
    let index = Math.floor((n - min) / width);
    if (index >= bins) {
      index = bins - 1;
    }
    if (index < 0) {
      index = 0;
    }
    result[index].count++;
  });
  return result;
};

const formatBinBound = (value: number): string => {
  const rounded = Math.round(value * 10) / 10;
  return String(rounded % 1 === 0 ? Math.round(rounded) : rounded);
};

// ---------------------------------------------------------------------------
// Likert
// ---------------------------------------------------------------------------

export interface ILikertStats {
  categories: string[];
  rows: { label: string; counts: number[]; total: number; average: number }[];
}

/**
 * Counts per statement per scale column.
 *
 * Answers are stored as JSON — a delimited string couldn't survive a statement
 * label that itself contains a semicolon.
 */
export const likertStats = (field: IFormField, items: IResponseItem[]): ILikertStats => {
  const rowLabels = (field.likertRows || []).map((r) => (r || '').trim()).filter((r) => r.length > 0);
  const categories = (field.likertColumns || []).map((c) => (c || '').trim()).filter((c) => c.length > 0);
  const counts: { [row: string]: number[] } = {};
  rowLabels.forEach((row) => {
    counts[row] = categories.map(() => 0);
  });

  items.forEach((item) => {
    const raw = item.values[field.internalName];
    if (!isAnswered(raw)) {
      return;
    }
    const answers = normalizeFromSharePoint(field, raw) as ILikertValue;
    if (!answers || typeof answers !== 'object') {
      return;
    }
    Object.keys(answers).forEach((row) => {
      const columnIndex = categories.indexOf(answers[row]);
      if (columnIndex >= 0 && counts[row]) {
        counts[row][columnIndex]++;
      }
    });
  });

  return {
    categories: categories,
    rows: rowLabels.map((label) => {
      const rowCounts = counts[label] || categories.map(() => 0);
      const total = rowCounts.reduce((sum, n) => sum + n, 0);
      // treat the columns as an ordered 1..n scale for a mean
      const weighted = rowCounts.reduce((sum, n, index) => sum + n * (index + 1), 0);
      return {
        label: label,
        counts: rowCounts,
        total: total,
        average: total > 0 ? Math.round((weighted / total) * 100) / 100 : 0
      };
    })
  };
};

// ---------------------------------------------------------------------------
// free text
// ---------------------------------------------------------------------------

/** Very common words carry no signal in a word frequency view. */
const STOP_WORDS: string[] = [
  'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'any', 'can', 'had', 'her', 'was', 'one',
  'our', 'out', 'day', 'get', 'has', 'him', 'his', 'how', 'its', 'may', 'new', 'now', 'old', 'see',
  'two', 'way', 'who', 'boy', 'did', 'she', 'use', 'that', 'this', 'with', 'have', 'from', 'they',
  'been', 'were', 'will', 'would', 'there', 'their', 'what', 'about', 'which', 'when', 'make',
  'like', 'time', 'just', 'know', 'take', 'into', 'your', 'some', 'them', 'than', 'then', 'only',
  'other', 'very', 'also', 'more', 'most', 'much', 'because', 'could', 'should', 'these', 'those',
  'being', 'does', 'each', 'over', 'such', 'both', 'while', 'after', 'before', 'where', 'here'
];

export interface IWordCount {
  word: string;
  count: number;
}

/** Word frequency across a text field's answers, stop words removed. */
export const wordFrequency = (
  field: IFormField,
  items: IResponseItem[],
  limit?: number
): IWordCount[] => {
  const counts: { [word: string]: number } = {};
  items.forEach((item) => {
    const raw = item.values[field.internalName];
    if (!isAnswered(raw)) {
      return;
    }
    const text = field.type === FieldType.RichText ? stripHtml(String(raw)) : String(raw);
    text
      .toLowerCase()
      .replace(/[^a-z0-9\s'-]/g, ' ')
      .split(/\s+/)
      .forEach((word) => {
        const cleaned = word.replace(/^[-']+|[-']+$/g, '');
        if (cleaned.length < 3 || STOP_WORDS.indexOf(cleaned) !== -1) {
          return;
        }
        counts[cleaned] = (counts[cleaned] || 0) + 1;
      });
  });
  const rows = Object.keys(counts).map((word) => ({ word: word, count: counts[word] }));
  rows.sort((a, b) => (b.count - a.count) || a.word.localeCompare(b.word));
  return rows.slice(0, limit || 24);
};

/** Every non-empty answer for a text field, for the drill-in list. */
export const textAnswers = (field: IFormField, items: IResponseItem[]): { id: number; text: string }[] =>
  items
    .map((item) => ({
      id: item.id,
      text: formatSharePointValue(field, item.values[field.internalName])
    }))
    .filter((entry) => entry.text.length > 0);

// ---------------------------------------------------------------------------
// timeline
// ---------------------------------------------------------------------------

export type TimeGrain = 'day' | 'week' | 'month';

export interface ITimelinePoint {
  label: string;
  value: number;
  start: Date;
}

const startOfDay = (date: Date): Date => {
  const d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  return d;
};

const startOfWeek = (date: Date): Date => {
  const d = startOfDay(date);
  // ISO-ish: weeks start Monday
  const day = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - day);
  return d;
};

const startOfMonth = (date: Date): Date => {
  const d = startOfDay(date);
  d.setDate(1);
  return d;
};

const bucketStart = (date: Date, grain: TimeGrain): Date =>
  grain === 'month' ? startOfMonth(date) : grain === 'week' ? startOfWeek(date) : startOfDay(date);

const advance = (date: Date, grain: TimeGrain): Date => {
  const d = new Date(date.getTime());
  if (grain === 'month') {
    d.setMonth(d.getMonth() + 1);
  } else if (grain === 'week') {
    d.setDate(d.getDate() + 7);
  } else {
    d.setDate(d.getDate() + 1);
  }
  return d;
};

const labelFor = (date: Date, grain: TimeGrain, locale?: string): string => {
  if (grain === 'month') {
    return date.toLocaleDateString(locale, { month: 'short', year: '2-digit' });
  }
  return date.toLocaleDateString(locale, { month: 'short', day: 'numeric' });
};

/** Most buckets a timeline will ever return. */
export const MAX_TIMELINE_BUCKETS = 400;

export interface ITimelineResult {
  points: ITimelinePoint[];
  /** the grain actually used — coarser than requested when the range was too long */
  grain: TimeGrain;
  /** true when even the coarsest grain overflowed and the oldest buckets were dropped */
  truncated: boolean;
}

const GRAIN_ORDER: TimeGrain[] = ['day', 'week', 'month'];

/**
 * Responses per time bucket, with empty buckets filled in.
 *
 * Gaps matter: skipping days with no responses would draw a line implying steady
 * activity across a quiet week. When the range would need more than
 * MAX_TIMELINE_BUCKETS points the grain is coarsened (day -> week -> month); if
 * even months overflow, the most recent MAX_TIMELINE_BUCKETS are kept and
 * `truncated` is set so the UI can say so.
 */
export const timelineDetailed = (
  items: IResponseItem[],
  grain: TimeGrain,
  locale?: string
): ITimelineResult => {
  if (items.length === 0) {
    return { points: [], grain: grain, truncated: false };
  }
  let used: TimeGrain = grain;
  for (;;) {
    const counts: { [key: number]: number } = {};
    let earliest: Date | undefined;
    let latest: Date | undefined;
    items.forEach((item) => {
      const bucket = bucketStart(item.created, used);
      const key = bucket.getTime();
      counts[key] = (counts[key] || 0) + 1;
      if (!earliest || bucket < earliest) {
        earliest = bucket;
      }
      if (!latest || bucket > latest) {
        latest = bucket;
      }
    });
    if (!earliest || !latest) {
      return { points: [], grain: used, truncated: false };
    }
    const coarsest = used === 'month';
    const points: ITimelinePoint[] = [];
    let cursor: Date = earliest;
    // hard stop far beyond any sane range so a corrupt date cannot spin forever
    let guard = 0;
    while (cursor <= latest && guard < 100000) {
      points.push({
        label: labelFor(cursor, used, locale),
        value: counts[cursor.getTime()] || 0,
        start: cursor
      });
      cursor = advance(cursor, used);
      guard++;
      if (!coarsest && points.length > MAX_TIMELINE_BUCKETS) {
        break;
      }
    }
    if (points.length <= MAX_TIMELINE_BUCKETS) {
      return { points: points, grain: used, truncated: false };
    }
    if (!coarsest) {
      used = GRAIN_ORDER[GRAIN_ORDER.indexOf(used) + 1];
      continue;
    }
    return {
      points: points.slice(points.length - MAX_TIMELINE_BUCKETS),
      grain: used,
      truncated: true
    };
  }
};

/** Points only — see timelineDetailed for the grain actually used and the truncation flag. */
export const timeline = (items: IResponseItem[], grain: TimeGrain, locale?: string): ITimelinePoint[] =>
  timelineDetailed(items, grain, locale).points;

/** Pick a sensible grain from how much history there is. */
export const suggestGrain = (items: IResponseItem[]): TimeGrain => {
  if (items.length === 0) {
    return 'day';
  }
  let earliest = items[0].created;
  let latest = items[0].created;
  items.forEach((item) => {
    if (item.created < earliest) {
      earliest = item.created;
    }
    if (item.created > latest) {
      latest = item.created;
    }
  });
  const days = (latest.getTime() - earliest.getTime()) / 86400000;
  if (days > 240) {
    return 'month';
  }
  if (days > 60) {
    return 'week';
  }
  return 'day';
};

// ---------------------------------------------------------------------------
// KPIs
// ---------------------------------------------------------------------------

export interface IKpiSet {
  total: number;
  last7: number;
  previous7: number;
  last30: number;
  today: number;
  uniqueRespondents: number;
  /** median seconds to complete, when the duration column has data */
  medianDuration?: number;
  /** share of questions answered, averaged across responses */
  completionRate?: number;
  /** daily counts for the last 14 days, for the sparkline */
  trend: number[];
  firstResponse?: Date;
  lastResponse?: Date;
}

/**
 * `fields` (optional) are the input questions: when given, completion counts
 * only answers to those questions, and an unanswered Yes/No is not an answer.
 * Without it every stored value key is counted, as before.
 */
export const computeKpis = (
  items: IResponseItem[],
  questionCount: number,
  now?: Date,
  fields?: IFormField[]
): IKpiSet => {
  const reference = now || new Date();
  const dayMs = 86400000;
  const todayStart = startOfDay(reference).getTime();
  const since = (days: number): number => todayStart - (days - 1) * dayMs;

  const inRange = (item: IResponseItem, from: number, to?: number): boolean => {
    const t = item.created.getTime();
    return t >= from && (to === undefined || t < to);
  };

  const last7 = items.filter((i) => inRange(i, since(7))).length;
  const previous7 = items.filter((i) => inRange(i, since(14), since(7))).length;
  const last30 = items.filter((i) => inRange(i, since(30))).length;
  const today = items.filter((i) => inRange(i, todayStart)).length;

  const respondents: { [key: string]: boolean } = {};
  items.forEach((item) => {
    const key = item.createdByEmail || item.createdBy;
    if (key) {
      respondents[key] = true;
    }
  });

  const durations = items
    .map((i) => i.durationSeconds)
    .filter((d) => typeof d === 'number' && (d as number) > 0) as number[];
  durations.sort((a, b) => a - b);
  const medianDuration =
    durations.length === 0
      ? undefined
      : durations.length % 2 === 0
        ? Math.round((durations[durations.length / 2 - 1] + durations[durations.length / 2]) / 2)
        : durations[Math.floor(durations.length / 2)];

  const completionRate =
    questionCount === 0 || items.length === 0
      ? undefined
      : Math.round(
          (items.reduce((sum, item) => {
            const answered = fields
              ? fields.filter((f) => isFieldAnswered(f, item.values[f.internalName])).length
              : Object.keys(item.values).filter((key) => isAnswered(item.values[key])).length;
            return sum + Math.min(1, answered / questionCount);
          }, 0) /
            items.length) *
            100
        );

  const trend: number[] = [];
  for (let day = 13; day >= 0; day--) {
    const from = todayStart - day * dayMs;
    trend.push(items.filter((i) => inRange(i, from, from + dayMs)).length);
  }

  let firstResponse: Date | undefined;
  let lastResponse: Date | undefined;
  items.forEach((item) => {
    if (!firstResponse || item.created < firstResponse) {
      firstResponse = item.created;
    }
    if (!lastResponse || item.created > lastResponse) {
      lastResponse = item.created;
    }
  });

  return {
    total: items.length,
    last7: last7,
    previous7: previous7,
    last30: last30,
    today: today,
    uniqueRespondents: Object.keys(respondents).length,
    medianDuration: medianDuration,
    completionRate: completionRate,
    trend: trend,
    firstResponse: firstResponse,
    lastResponse: lastResponse
  };
};

export const formatDuration = (seconds: number): string => {
  if (seconds < 60) {
    return seconds + 's';
  }
  const minutes = Math.floor(seconds / 60);
  const rest = seconds % 60;
  if (minutes < 60) {
    return rest === 0 ? minutes + 'm' : minutes + 'm ' + rest + 's';
  }
  const hours = Math.floor(minutes / 60);
  return hours + 'h ' + (minutes % 60) + 'm';
};

// ---------------------------------------------------------------------------
// segmentation (cross-tab)
// ---------------------------------------------------------------------------

export interface ISegment {
  /** display text */
  label: string;
  /** raw answer key ('' for the "no answer" segment); use matchesRowFilter(field, item, segment) */
  filterValue: string;
  items: IResponseItem[];
}

/**
 * Split a response set by another question's answer.
 *
 * This is what makes "NPS by department" possible. Multi-value answers (arrays
 * or "A; B" text) put a response in every selected segment, so segment sizes can
 * sum to more than the total — reported per segment rather than as a share of
 * the whole. Date fields segment by local date and Time by hour, chronologically.
 */
export const segmentBy = (
  field: IFormField,
  items: IResponseItem[],
  labels?: IMessageBag,
  locale?: string
): ISegment[] => {
  const buckets: { [key: string]: IResponseItem[] } = {};
  const order: string[] = [];
  const push = (key: string, item: IResponseItem): void => {
    if (!buckets[key]) {
      buckets[key] = [];
      order.push(key);
    }
    if (buckets[key].indexOf(item) === -1) {
      buckets[key].push(item);
    }
  };

  items.forEach((item) => {
    const keys = answerKeys(field, item.values[field.internalName]);
    if (keys.length === 0) {
      push('', item);
      return;
    }
    keys.forEach((key) => push(key, item));
  });

  // put designed choice order first where we know it
  const designed =
    field.type === FieldType.Choice
      ? effectiveChoices(field)
      : field.type === FieldType.ImageChoice
        ? (field.imageChoices || []).map((o) => o.label)
        : [];
  let ordered = designed
    .filter((key) => !!buckets[key])
    .concat(order.filter((key) => designed.indexOf(key) === -1));
  if (field.type === FieldType.Date || field.type === FieldType.Time) {
    // keys are zero-padded, so a plain sort is chronological; "no answer" goes last
    ordered = order
      .filter((key) => key !== '')
      .sort()
      .concat(order.filter((key) => key === ''));
  }

  return ordered.map((key) => ({
    label: key === '' ? msg('Logic_NoAnswer', undefined, labels) : labelForKey(field, key, labels, locale),
    filterValue: key,
    items: buckets[key]
  }));
};
