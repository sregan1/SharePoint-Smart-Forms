import { FieldType, IFormField, ILikertValue, IResponseItem } from '../models';
import { effectiveChoices, formatSharePointValue, normalizeFromSharePoint } from './formUtils';
import { stripHtml } from './sanitizeHtml';

/**
 * Response analytics.
 *
 * Pure functions over the loaded response set — no React, no SharePoint — so the
 * views can memoize them and the tests can exercise them directly.
 */

export interface IDistributionRow {
  label: string;
  count: number;
}

export const isAnswered = (value: unknown): boolean => {
  if (value === undefined || value === null || value === '') {
    return false;
  }
  return !(Array.isArray(value) && value.length === 0);
};

export const answeredCount = (field: IFormField, items: IResponseItem[]): number =>
  items.filter((item) => isAnswered(item.values[field.internalName])).length;

/** Counts per category for a choice-like field, honoring the designed order. */
export const distributionFor = (field: IFormField, items: IResponseItem[]): IDistributionRow[] => {
  const counts: { [label: string]: number } = {};
  const bump = (label: string): void => {
    counts[label] = (counts[label] || 0) + 1;
  };

  items.forEach((item) => {
    const raw = item.values[field.internalName];
    if (!isAnswered(raw)) {
      return;
    }
    if (field.type === FieldType.YesNo) {
      bump(raw === true || raw === 'true' ? 'Yes' : 'No');
      return;
    }
    if (field.type === FieldType.Consent) {
      bump(raw === true || raw === 'true' ? 'Agreed' : 'Not agreed');
      return;
    }
    if (field.type === FieldType.Person) {
      const normalized = normalizeFromSharePoint(field, raw);
      if (Array.isArray(normalized)) {
        (normalized as { displayName?: string }[]).forEach((person) => {
          if (person && person.displayName) {
            bump(person.displayName);
          }
        });
      }
      return;
    }
    if (Array.isArray(raw)) {
      (raw as unknown[]).forEach((v) => bump(String(v)));
      return;
    }
    // multi-value text columns (image choice, multi lookup) hold "A; B"
    if (
      field.allowMultiple &&
      (field.type === FieldType.ImageChoice || field.type === FieldType.Lookup)
    ) {
      String(raw)
        .split(';')
        .map((v) => v.trim())
        .filter((v) => v.length > 0)
        .forEach(bump);
      return;
    }
    bump(String(raw));
  });

  let rows: IDistributionRow[];

  if (field.type === FieldType.Choice) {
    // keep the designer's option order, then any write-in or legacy values
    const designed = effectiveChoices(field);
    rows = designed.map((c) => ({ label: c, count: counts[c] || 0 }));
    Object.keys(counts).forEach((label) => {
      if (designed.indexOf(label) === -1) {
        rows.push({ label: label, count: counts[label] });
      }
    });
  } else if (field.type === FieldType.ImageChoice) {
    const designed = (field.imageChoices || []).map((o) => o.label).filter((l) => l);
    rows = designed.map((c) => ({ label: c, count: counts[c] || 0 }));
    Object.keys(counts).forEach((label) => {
      if (designed.indexOf(label) === -1) {
        rows.push({ label: label, count: counts[label] });
      }
    });
  } else if (field.type === FieldType.YesNo) {
    rows = [
      { label: 'Yes', count: counts.Yes || 0 },
      { label: 'No', count: counts.No || 0 }
    ];
  } else if (field.type === FieldType.Consent) {
    rows = [
      { label: 'Agreed', count: counts.Agreed || 0 },
      { label: 'Not agreed', count: counts['Not agreed'] || 0 }
    ];
  } else if (field.type === FieldType.Rating) {
    const max = field.maxRating || 5;
    rows = [];
    for (let star = max; star >= 1; star--) {
      const whole = counts[String(star)] || 0;
      const half = counts[String(star - 0.5)] || 0;
      rows.push({ label: star + (field.ratingIcon === 'like' ? '' : ' ★'), count: whole });
      if (field.allowHalfRating && half > 0) {
        rows.push({ label: star - 0.5 + ' ★', count: half });
      }
    }
  } else if (field.type === FieldType.Scale) {
    const min = typeof field.min === 'number' ? field.min : 1;
    const max = typeof field.max === 'number' ? field.max : 5;
    const step = field.step && field.step > 0 ? field.step : 1;
    rows = [];
    for (let value = min; value <= max; value += step) {
      const key = String(Math.round(value * 100) / 100);
      rows.push({ label: key, count: counts[key] || 0 });
    }
  } else {
    rows = Object.keys(counts).map((label) => ({ label: label, count: counts[label] }));
    rows.sort((a, b) => b.count - a.count);
  }

  return rows;
};

/**
 * Fold a long tail into "Other".
 *
 * A generated ninth hue would be indistinguishable from an existing one under
 * colorblind simulation, so charts that color by category cap their series and
 * collapse the rest rather than inventing colors.
 */
export const foldTail = (rows: IDistributionRow[], maxSeries: number): IDistributionRow[] => {
  if (rows.length <= maxSeries) {
    return rows;
  }
  const sorted = rows.slice().sort((a, b) => b.count - a.count);
  const head = sorted.slice(0, maxSeries - 1);
  const tail = sorted.slice(maxSeries - 1);
  const otherCount = tail.reduce((sum, row) => sum + row.count, 0);
  return head.concat([{ label: 'Other (' + tail.length + ')', count: otherCount }]);
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
      : String(raw)
          .split(';')
          .map((v) => v.trim())
          .filter((v) => v.length > 0);
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

const labelFor = (date: Date, grain: TimeGrain): string => {
  if (grain === 'month') {
    return date.toLocaleDateString(undefined, { month: 'short', year: '2-digit' });
  }
  return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
};

/**
 * Responses per time bucket, with empty buckets filled in.
 *
 * Gaps matter: skipping days with no responses would draw a line implying steady
 * activity across a quiet week.
 */
export const timeline = (items: IResponseItem[], grain: TimeGrain): ITimelinePoint[] => {
  if (items.length === 0) {
    return [];
  }
  const counts: { [key: number]: number } = {};
  let earliest: Date | undefined;
  let latest: Date | undefined;

  items.forEach((item) => {
    const bucket = bucketStart(item.created, grain);
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
    return [];
  }

  const points: ITimelinePoint[] = [];
  let cursor = earliest;
  // guard against a pathological range producing thousands of points
  let guard = 0;
  while (cursor <= latest && guard < 400) {
    points.push({
      label: labelFor(cursor, grain),
      value: counts[cursor.getTime()] || 0,
      start: cursor
    });
    cursor = advance(cursor, grain);
    guard++;
  }
  return points;
};

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

export const computeKpis = (
  items: IResponseItem[],
  questionCount: number,
  now?: Date
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
            const answered = Object.keys(item.values).filter((key) => isAnswered(item.values[key])).length;
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
  label: string;
  items: IResponseItem[];
}

/**
 * Split a response set by another question's answer.
 *
 * This is what makes "NPS by department" possible. Multi-value answers put a
 * response in every matching segment, so segment sizes can sum to more than the
 * total — reported per segment rather than as a share of the whole.
 */
export const segmentBy = (field: IFormField, items: IResponseItem[]): ISegment[] => {
  const buckets: { [label: string]: IResponseItem[] } = {};
  const order: string[] = [];
  const push = (label: string, item: IResponseItem): void => {
    if (!buckets[label]) {
      buckets[label] = [];
      order.push(label);
    }
    buckets[label].push(item);
  };

  items.forEach((item) => {
    const raw = item.values[field.internalName];
    if (!isAnswered(raw)) {
      push('(no answer)', item);
      return;
    }
    if (field.type === FieldType.YesNo) {
      push(raw === true || raw === 'true' ? 'Yes' : 'No', item);
      return;
    }
    if (field.type === FieldType.Consent) {
      push(raw === true || raw === 'true' ? 'Agreed' : 'Not agreed', item);
      return;
    }
    if (field.type === FieldType.Person) {
      const normalized = normalizeFromSharePoint(field, raw);
      if (Array.isArray(normalized) && normalized.length > 0) {
        (normalized as { displayName?: string }[]).forEach((person) => {
          push(person.displayName || '(unknown)', item);
        });
      } else {
        push('(no answer)', item);
      }
      return;
    }
    if (Array.isArray(raw)) {
      (raw as unknown[]).forEach((v) => push(String(v), item));
      return;
    }
    push(String(raw), item);
  });

  // put designed choice order first where we know it
  const designed =
    field.type === FieldType.Choice
      ? effectiveChoices(field)
      : field.type === FieldType.ImageChoice
        ? (field.imageChoices || []).map((o) => o.label)
        : [];
  const ordered = designed
    .filter((label) => !!buckets[label])
    .concat(order.filter((label) => designed.indexOf(label) === -1));

  return ordered.map((label) => ({ label: label, items: buckets[label] }));
};
