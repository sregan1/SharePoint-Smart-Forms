import type { IReadonlyTheme } from '@microsoft/sp-component-base';

/**
 * Theme plumbing.
 *
 * Every color in the stylesheets is a `var(--sf-*)` custom property resolved
 * here, once, at the root of the web part. Two things feed the token set:
 *
 *   1. the SharePoint site theme (via ThemeProvider) — surfaces, text, borders,
 *      and whether we're inverted, so the form is legible on a dark site;
 *   2. the form's own accent color, chosen by the owner, plus the tints and
 *      readable foregrounds derived from it.
 *
 * Deriving the accent variants here rather than in CSS keeps the stylesheets
 * free of `color-mix()` (patchy support, hard to reason about) and means the
 * contrast decision — light text or dark text on the accent — is made once with
 * real luminance math instead of being guessed per rule.
 */

export interface IRgb {
  r: number;
  g: number;
  b: number;
}

const clamp255 = (n: number): number => Math.max(0, Math.min(255, Math.round(n)));

export const hexToRgb = (hex: string): IRgb => {
  const value = String(hex || '').replace('#', '').trim();
  const full =
    value.length === 3
      ? value.charAt(0) + value.charAt(0) + value.charAt(1) + value.charAt(1) + value.charAt(2) + value.charAt(2)
      : value;
  const parsed = parseInt(full.slice(0, 6), 16);
  if (isNaN(parsed) || full.length < 6) {
    return { r: 0, g: 120, b: 212 };
  }
  return {
    r: (parsed >> 16) & 255,
    g: (parsed >> 8) & 255,
    b: parsed & 255
  };
};

export const rgbToHex = (rgb: IRgb): string => {
  const part = (n: number): string => {
    const text = clamp255(n).toString(16);
    return text.length === 1 ? '0' + text : text;
  };
  return '#' + part(rgb.r) + part(rgb.g) + part(rgb.b);
};

/** Blend two colors. `weight` is how much of `a` survives (0..1). */
export const mix = (a: string, b: string, weight: number): string => {
  const first = hexToRgb(a);
  const second = hexToRgb(b);
  const w = Math.max(0, Math.min(1, weight));
  return rgbToHex({
    r: first.r * w + second.r * (1 - w),
    g: first.g * w + second.g * (1 - w),
    b: first.b * w + second.b * (1 - w)
  });
};

/** Relative luminance per WCAG 2.1. */
export const luminance = (hex: string): number => {
  const { r, g, b } = hexToRgb(hex);
  const channel = (raw: number): number => {
    const c = raw / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
};

export const contrastRatio = (a: string, b: string): number => {
  const first = luminance(a);
  const second = luminance(b);
  const lighter = Math.max(first, second);
  const darker = Math.min(first, second);
  return (lighter + 0.05) / (darker + 0.05);
};

/** Pick whichever of white / near-black reads better on `background`. */
export const readableOn = (background: string): string =>
  contrastRatio(background, '#ffffff') >= contrastRatio(background, '#201f1e') ? '#ffffff' : '#201f1e';

/** Darken toward black by `amount` (0..1). */
export const darken = (hex: string, amount: number): string => mix(hex, '#000000', 1 - amount);

/** Lighten toward white by `amount` (0..1). */
export const lighten = (hex: string, amount: number): string => mix(hex, '#ffffff', 1 - amount);

/**
 * Nudge an accent color until it has at least `minimum` contrast against the
 * surface it sits on. Owners pick accents from a swatch list, but they can also
 * inherit one from an older form — a mid-tone accent that was fine on white is
 * unreadable on a dark card, so shift it rather than render low-contrast text.
 */
export const ensureContrast = (color: string, background: string, minimum: number): string => {
  if (contrastRatio(color, background) >= minimum) {
    return color;
  }
  const towardLight = luminance(background) < 0.5;
  let candidate = color;
  for (let step = 1; step <= 10; step++) {
    candidate = towardLight ? lighten(color, step / 12) : darken(color, step / 12);
    if (contrastRatio(candidate, background) >= minimum) {
      return candidate;
    }
  }
  return towardLight ? '#ffffff' : '#201f1e';
};

interface IPaletteLike {
  [key: string]: string | undefined;
}

const pick = (source: IPaletteLike | undefined, key: string, fallback: string): string => {
  const value = source ? source[key] : undefined;
  return typeof value === 'string' && value.length > 0 ? value : fallback;
};

/** Light-theme defaults, used when no site theme is available (workbench). */
const LIGHT_FALLBACK = {
  bodyBackground: '#ffffff',
  bodyText: '#323130',
  bodySubtext: '#605e5c',
  disabledBodyText: '#a19f9d',
  bodyDivider: '#edebe9',
  inputBorder: '#8a8886',
  neutralLighterAlt: '#faf9f8',
  neutralLighter: '#f3f2f1',
  neutralQuaternaryAlt: '#e1dfdd',
  neutralTertiaryAlt: '#c8c6c4',
  errorText: '#a4262c',
  successText: '#0b6a0b',
  warningText: '#8a5700'
};

const DARK_FALLBACK = {
  bodyBackground: '#1b1a19',
  bodyText: '#f3f2f1',
  bodySubtext: '#c8c6c4',
  disabledBodyText: '#797775',
  bodyDivider: '#3b3a39',
  inputBorder: '#8a8886',
  neutralLighterAlt: '#252423',
  neutralLighter: '#292827',
  neutralQuaternaryAlt: '#3b3a39',
  neutralTertiaryAlt: '#484644',
  errorText: '#f1707b',
  successText: '#6ccb6c',
  warningText: '#ffd267'
};

/** Flat map of CSS custom properties; cast to React.CSSProperties at the call site. */
export interface ICssVariables {
  [property: string]: string;
}

export interface IThemeInfo {
  isDark: boolean;
  /** CSS custom properties to apply at the root of the web part */
  tokens: ICssVariables;
  /** the accent actually used after any contrast correction */
  accent: string;
  /** readable text color on top of the accent */
  accentForeground: string;
  /** card surface color, needed by charts that blend against it */
  surface: string;
}

/**
 * Build the `--sf-*` token set for a site theme and form accent color.
 */
export const buildTheme = (theme: IReadonlyTheme | undefined, accentColor: string): IThemeInfo => {
  const semantic = (theme ? theme.semanticColors : undefined) as IPaletteLike | undefined;
  const palette = (theme ? theme.palette : undefined) as IPaletteLike | undefined;
  const isDark = !!(theme && theme.isInverted);
  const fallback = isDark ? DARK_FALLBACK : LIGHT_FALLBACK;

  const bg = pick(semantic, 'bodyBackground', fallback.bodyBackground);
  const fg = pick(semantic, 'bodyText', fallback.bodyText);
  const fgMuted = pick(semantic, 'bodySubtext', fallback.bodySubtext);
  const fgSubtle = pick(semantic, 'disabledBodyText', fallback.disabledBodyText);
  const border = pick(semantic, 'bodyDivider', fallback.bodyDivider);
  const borderStrong = pick(semantic, 'inputBorder', fallback.inputBorder);
  const surface = pick(palette, 'neutralLighterAlt', fallback.neutralLighterAlt);
  const surface2 = pick(palette, 'neutralLighter', fallback.neutralLighter);
  const surface3 = pick(palette, 'neutralQuaternaryAlt', fallback.neutralQuaternaryAlt);
  const borderMuted = pick(palette, 'neutralTertiaryAlt', fallback.neutralTertiaryAlt);
  const error = pick(semantic, 'errorText', fallback.errorText);

  // Cards sit slightly above the page background. On a dark theme "above"
  // means lighter; on light it means plain white.
  const card = isDark ? lighten(bg, 0.06) : '#ffffff';
  const cardRaised = isDark ? lighten(bg, 0.1) : '#ffffff';

  const rawAccent = accentColor || '#0078d4';
  // 3:1 is the WCAG minimum for large text and UI component boundaries, which
  // is what the accent is used for (bars, chips, borders, filled buttons).
  const accent = ensureContrast(rawAccent, card, 3);
  const accentForeground = readableOn(accent);
  const accentHover = isDark ? lighten(accent, 0.12) : darken(accent, 0.12);
  const accentTint = mix(accent, card, isDark ? 0.22 : 0.12);
  const accentTintStrong = mix(accent, card, isDark ? 0.34 : 0.2);
  const accentGradientEnd = isDark ? darken(accent, 0.18) : darken(accent, 0.28);

  const success = pick(semantic, 'successText', fallback.successText);
  const warning = pick(semantic, 'warningText', fallback.warningText);

  const tokens: ICssVariables = {
    '--sf-bg': bg,
    '--sf-card': card,
    '--sf-card-raised': cardRaised,
    '--sf-surface': surface,
    '--sf-surface-2': surface2,
    '--sf-surface-3': surface3,
    '--sf-fg': fg,
    '--sf-fg-muted': fgMuted,
    '--sf-fg-subtle': fgSubtle,
    '--sf-border': border,
    '--sf-border-muted': borderMuted,
    '--sf-border-strong': borderStrong,

    '--sf-accent': accent,
    '--sf-accent-fg': accentForeground,
    '--sf-accent-hover': accentHover,
    '--sf-accent-tint': accentTint,
    '--sf-accent-tint-strong': accentTintStrong,
    '--sf-accent-gradient-end': accentGradientEnd,

    '--sf-error': error,
    '--sf-error-tint': mix(error, card, 0.12),
    '--sf-success': success,
    '--sf-success-tint': mix(success, card, 0.14),
    '--sf-warning': warning,
    '--sf-warning-tint': mix(warning, card, 0.14),

    // shadows have to invert too — a black glow is invisible on a dark card
    '--sf-shadow-sm': isDark ? '0 1px 4px rgba(0, 0, 0, 0.5)' : '0 1px 4px rgba(0, 0, 0, 0.06)',
    '--sf-shadow-md': isDark ? '0 3px 12px rgba(0, 0, 0, 0.55)' : '0 3px 12px rgba(0, 0, 0, 0.1)',
    '--sf-shadow-lg': isDark ? '0 6px 20px rgba(0, 0, 0, 0.6)' : '0 6px 20px rgba(0, 0, 0, 0.12)',

    '--sf-focus': accent
  };

  return { isDark, tokens, accent, accentForeground, surface: card };
};

// ---------------------------------------------------------------------------
// chart palettes
// ---------------------------------------------------------------------------
//
// These are not eyeballed. Each set below was run through the data-viz
// validator (lightness band, chroma floor, CVD separation under simulated
// protanopia/deuteranopia, normal-vision separation floor, and contrast vs the
// surface it actually renders on) against this app's own surfaces: #ffffff on a
// light theme and #292827 on a dark one.
//
// Two consequences are load-bearing and must not be "tidied" away:
//
//  1. The slot ORDER is the colorblind-safety mechanism, not a style choice. A
//     two-series chart uses slots 1–2, so re-ordering silently changes which
//     pairs sit next to each other. Never reorder; never generate a 9th hue.
//  2. Three light-mode slots and one dark-mode slot sit just below 3:1 against
//     the surface. That is permitted only because every chart here also ships
//     visible value labels and a table view. If you remove the labels, the
//     palette stops being compliant.

/** Categorical slots — identity only, assigned in order, never cycled. */
export const CATEGORICAL_LIGHT: string[] = [
  '#2a78d6',
  '#eb6834',
  '#1baf7a',
  '#eda100',
  '#e87ba4',
  '#008300',
  '#4a3aa7',
  '#e34948'
];

/** The same eight hues re-stepped for a dark surface, not a different palette. */
export const CATEGORICAL_DARK: string[] = [
  '#3987e5',
  '#d95926',
  '#199e70',
  '#c98500',
  '#d55181',
  '#008300',
  '#9085e9',
  '#e66767'
];

export const categoricalPalette = (isDark: boolean): string[] =>
  isDark ? CATEGORICAL_DARK : CATEGORICAL_LIGHT;

/**
 * Hard ceiling on categorical series. Past this the tail folds into "Other" —
 * a generated ninth hue is indistinguishable from an existing one under CVD.
 */
export const MAX_CATEGORICAL_SERIES = 8;

/**
 * Status colors, with reserved meaning. Never reused as "series 4", and always
 * shipped with a label rather than relying on hue alone — warning sits below
 * 3:1 on a light surface by design, and the label is the mitigation.
 */
export const STATUS_COLORS = {
  good: '#0ca30c',
  warning: '#fab219',
  serious: '#ec835a',
  critical: '#d03b3b'
};

/**
 * Diverging arms for ordered-scale data (Likert agreement, sentiment).
 *
 * Two hues plus a neutral gray midpoint — never a hue in the middle, or the
 * centre reads as a category rather than as "neither side". Each arm is ordered
 * from the step nearest the surface outward, and each was validated as a
 * one-hue ordinal ramp (monotone lightness, adjacent lightness gap, and a
 * light-end floor against the real surface).
 */
export const DIVERGING_LIGHT = {
  negative: ['#eda1a1', '#d03b3b', '#8f2020'],
  positive: ['#86b6ef', '#3987e5', '#1c5cab'],
  neutral: '#9a9894'
};

export const DIVERGING_DARK = {
  negative: ['#b03030', '#d95c5c', '#f0a8a8'],
  positive: ['#1c5cab', '#3987e5', '#9ec5f4'],
  neutral: '#7a7874'
};

export const divergingPalette = (
  isDark: boolean
): { negative: string[]; positive: string[]; neutral: string } =>
  isDark ? DIVERGING_DARK : DIVERGING_LIGHT;

/**
 * Colors for a diverging stacked bar with `count` ordered categories.
 *
 * Splits the categories into a negative arm, an optional neutral middle (only
 * when the count is odd) and a positive arm, drawing each arm from the far end
 * inward so the extremes are the strongest steps.
 */
export const divergingScale = (count: number, isDark: boolean): string[] => {
  const palette = divergingPalette(isDark);
  const n = Math.max(2, count);
  const hasNeutral = n % 2 === 1;
  const perArm = Math.floor(n / 2);
  const armColors = (arm: string[]): string[] => {
    // arms are documented as 3 steps; take the strongest `perArm` of them
    const usable = arm.slice(0, Math.min(perArm, arm.length));
    while (usable.length < perArm) {
      usable.push(arm[arm.length - 1]);
    }
    return usable;
  };
  const negative = armColors(palette.negative).slice().reverse();
  const positive = armColors(palette.positive);
  return hasNeutral ? negative.concat([palette.neutral], positive) : negative.concat(positive);
};

/**
 * Largest number of bands a single-hue ordinal ramp can carry.
 *
 * Six or more steps cannot keep an adjacent lightness gap of 0.06 within the
 * readable band, which the validator flags — so anything ordered with more
 * categories than this uses the diverging form or a table instead of a ramp.
 */
export const MAX_ORDINAL_STEPS = 5;

/** WCAG relative-luminance contrast, used by the ramp generator. */
const contrastWith = (color: string, surface: string): number => contrastRatio(color, surface);

/**
 * Ordinal ramp in the accent's own hue: one hue, monotone lightness, with the
 * step nearest the surface still readable against it.
 *
 * Each step targets a contrast ratio against the surface, spaced geometrically
 * from a light-end floor up to the accent's own (or a deepened variant when the
 * accent alone can't span enough range). The mix weight hitting a target is
 * found by bisection — contrast against a fixed surface is monotone in the mix
 * weight, so lightness ordering is guaranteed by construction rather than by
 * hand-picked hexes, and the light-end floor lives in the formula.
 *
 * Deliberately generated rather than tabulated: the accent is chosen by the form
 * owner (including a custom color), so no fixed table could cover it.
 */
const LIGHT_END_CONTRAST = 2.15;

export const ordinalRamp = (
  accent: string,
  surface: string,
  steps: number,
  isDark: boolean
): string[] => {
  const n = Math.max(1, Math.min(steps, MAX_ORDINAL_STEPS));
  if (n === 1) {
    return [accent];
  }
  const accentContrast = contrastWith(accent, surface);
  // widen the far end when the accent alone is too close to the surface
  const deepEnd = isDark ? mix(accent, '#ffffff', 0.45) : mix(accent, '#000000', 0.45);
  const deepContrast = contrastWith(deepEnd, surface);
  const useDeep = n >= 4 && deepContrast > accentContrast;
  const anchor = useDeep ? deepEnd : accent;
  const top = Math.max(accentContrast, useDeep ? deepContrast : accentContrast);

  const result: string[] = [];
  for (let i = 0; i < n; i++) {
    const t = i / (n - 1);
    const target = LIGHT_END_CONTRAST * Math.pow(top / LIGHT_END_CONTRAST, t);
    let lo = 0;
    let hi = 1;
    let best = anchor;
    for (let iteration = 0; iteration < 24; iteration++) {
      const mid = (lo + hi) / 2;
      const candidate = mix(anchor, surface, mid);
      best = candidate;
      if (contrastWith(candidate, surface) < target) {
        lo = mid;
      } else {
        hi = mid;
      }
    }
    result.push(best);
  }
  return result;
};

/** Recessive chart chrome, derived from the theme rather than hardcoded. */
export interface IChartInk {
  grid: string;
  axis: string;
  label: string;
  muted: string;
  surface: string;
}

export const chartInk = (theme: IThemeInfo): IChartInk => ({
  grid: theme.isDark ? lighten(theme.surface, 0.08) : darken(theme.surface, 0.07),
  axis: theme.isDark ? lighten(theme.surface, 0.16) : darken(theme.surface, 0.16),
  label: theme.tokens['--sf-fg-muted'] || '#605e5c',
  muted: theme.tokens['--sf-fg-subtle'] || '#a19f9d',
  surface: theme.surface
});
