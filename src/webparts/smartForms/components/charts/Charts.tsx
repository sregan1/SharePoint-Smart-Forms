import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';
import { Icon } from '@fluentui/react';
import styles from './Charts.module.scss';
import { IChartInk } from '../../utils/theme';

/**
 * Inline SVG/HTML chart primitives.
 *
 * Deliberately no chart library: the shapes needed here are simple, and an SPFx
 * bundle pays for every dependency on every page load. Colors are always passed
 * in from utils/theme.ts (which validates them) — nothing here invents a color.
 *
 * Every chart ships visible value labels and a table fallback. That isn't
 * decoration: several palette slots sit just below 3:1 against the surface, and
 * the labels are what make that legal.
 */

export interface IChartDatum {
  label: string;
  value: number;
  /** optional secondary text shown in the tooltip */
  detail?: string;
  /** overrides the series color for this datum */
  color?: string;
  /** the analytics row behind this datum, handed back to onSelect handlers for click-to-filter */
  row?: { label: string; filterValue?: string; filterValues?: string[] };
}

export const formatCompact = (value: number): string => {
  const abs = Math.abs(value);
  if (abs >= 1000000) {
    return (value / 1000000).toFixed(abs >= 10000000 ? 0 : 1) + 'M';
  }
  if (abs >= 1000) {
    return (value / 1000).toFixed(abs >= 10000 ? 0 : 1) + 'K';
  }
  return String(Math.round(value * 100) / 100);
};

export const formatPercent = (part: number, total: number): string =>
  total > 0 ? Math.round((part / total) * 100) + '%' : '0%';

// ---------------------------------------------------------------------------
// tooltip
// ---------------------------------------------------------------------------

interface ITooltipState {
  x: number;
  y: number;
  title: string;
  rows: { label: string; value: string; color?: string }[];
}

export interface IUseTooltip {
  tooltip: ITooltipState | undefined;
  show: (event: React.MouseEvent, state: Omit<ITooltipState, 'x' | 'y'>) => void;
  hide: () => void;
  render: () => React.ReactNode;
}

/** Hover tooltip positioned within the chart's own bounding box. */
export const useTooltip = (containerRef: React.RefObject<HTMLElement>): IUseTooltip => {
  const [tooltip, setTooltip] = React.useState<ITooltipState | undefined>(undefined);

  const show = (event: React.MouseEvent, state: Omit<ITooltipState, 'x' | 'y'>): void => {
    const container = containerRef.current;
    if (!container) {
      return;
    }
    const bounds = container.getBoundingClientRect();
    setTooltip({
      ...state,
      x: event.clientX - bounds.left,
      // sit just above the pointer, clamped inside the container
      y: Math.max(28, event.clientY - bounds.top - 8)
    });
  };

  const hide = (): void => setTooltip(undefined);

  const render = (): React.ReactNode => {
    if (!tooltip) {
      return null;
    }
    return (
      <div className={styles.tooltip} style={{ left: tooltip.x, top: tooltip.y }} role="presentation">
        <div className={styles.tooltipTitle}>{tooltip.title}</div>
        {tooltip.rows.map((row, index) => (
          <div key={index} className={styles.tooltipRow}>
            {row.color && <span className={styles.tooltipSwatch} style={{ background: row.color }} />}
            <span>{row.label}</span>
            <span className={styles.tooltipValue}>{row.value}</span>
          </div>
        ))}
      </div>
    );
  };

  return { tooltip, show, hide, render };
};

// ---------------------------------------------------------------------------
// horizontal bars — the workhorse for choice / rating distributions
// ---------------------------------------------------------------------------

export interface IBarChartProps {
  data: IChartDatum[];
  /** single hue for nominal bars; length already encodes the value */
  color: string;
  total: number;
  /** wider label column for long option text */
  wideLabels?: boolean;
  /** invoked when a bar is clicked, for drill-through */
  onSelect?: (datum: IChartDatum, index: number) => void;
  valueFormatter?: (datum: IChartDatum) => string;
}

/**
 * Horizontal bars in one hue.
 *
 * Nominal categories are *not* colored individually: the bar length already
 * encodes the value, so spending the identity channel on it would be
 * re-encoding. HTML rather than SVG so long labels truncate with CSS and the
 * values stay selectable and screen-reader-visible.
 */
export const BarChart: React.FunctionComponent<IBarChartProps> = (props) => {
  const max = Math.max(1, ...props.data.map((d) => d.value));

  if (props.data.length === 0) {
    return <div className={styles.chartEmpty}>{strings.Charts_NoAnswers}</div>;
  }

  return (
    <div className={styles.barList}>
      {props.data.map((datum, index) => {
        const width = max > 0 ? Math.max((datum.value / max) * 100, datum.value > 0 ? 1.5 : 0) : 0;
        const valueText = props.valueFormatter
          ? props.valueFormatter(datum)
          : datum.value + ' · ' + formatPercent(datum.value, props.total);
        const clickable = !!props.onSelect;
        return (
          <button
            key={datum.label + '::' + index}
            type="button"
            className={
              clickable ? styles.barRow + ' ' + styles.barRowClickable : styles.barRow
            }
            disabled={!clickable}
            title={datum.label}
            aria-label={datum.label + ': ' + valueText}
            onClick={() => props.onSelect && props.onSelect(datum, index)}
          >
            <span className={props.wideLabels ? styles.barLabelWide : styles.barLabel}>
              {datum.label}
            </span>
            <span className={styles.barTrack}>
              <span
                className={styles.barFill}
                style={{ width: width + '%', background: datum.color || props.color }}
              />
            </span>
            <span className={styles.barValue}>{valueText}</span>
          </button>
        );
      })}
    </div>
  );
};

// ---------------------------------------------------------------------------
// columns / histogram
// ---------------------------------------------------------------------------

export interface IColumnChartProps {
  data: IChartDatum[];
  color: string;
  total: number;
  ink: IChartInk;
  height?: number;
  /** show every nth label when the axis gets crowded */
  labelEvery?: number;
  onSelect?: (datum: IChartDatum, index: number) => void;
}

const MAX_BAR_THICKNESS = 40;

/** Vertical columns with a value on the cap. Used for scales and histograms. */
export const ColumnChart: React.FunctionComponent<IColumnChartProps> = (props) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const tooltip = useTooltip(containerRef);

  // Measure the real width so the viewBox is in pixels: the svg then renders at
  // exactly `height` px instead of growing with its container's width, and text
  // is no longer stretched horizontally.
  const [width, setWidth] = React.useState<number>(300);
  const isEmpty = props.data.length === 0;
  React.useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) {
      return undefined;
    }
    const measure = (): void => {
      const w = Math.round(el.getBoundingClientRect().width);
      if (w > 0) {
        setWidth(w);
      }
    };
    measure();
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(measure);
      observer.observe(el);
      return () => observer.disconnect();
    }
    window.addEventListener('resize', measure);
    return () => window.removeEventListener('resize', measure);
  }, [isEmpty]);

  const height = props.height || 150;
  const paddingTop = 18;
  const paddingBottom = 22;
  const plotHeight = height - paddingTop - paddingBottom;
  const count = props.data.length;

  if (count === 0) {
    return <div className={styles.chartEmpty}>{strings.Charts_NoAnswers}</div>;
  }

  const max = Math.max(1, ...props.data.map((d) => d.value));
  const bandWidth = width / count;
  const barWidth = Math.min(bandWidth * 0.62, MAX_BAR_THICKNESS);
  const labelEvery = props.labelEvery || 1;

  return (
    <div className={styles.chartRoot} ref={containerRef}>
      <svg
        className={styles.svg}
        viewBox={'0 0 ' + width + ' ' + height}
        height={height}
        role="img"
        aria-label={formatString(strings.Charts_ColumnChartAria, { count: count })}
      >
        {/* baseline only — gridlines would be noise at this size */}
        <line
          x1={0}
          y1={paddingTop + plotHeight}
          x2={width}
          y2={paddingTop + plotHeight}
          stroke={props.ink.axis}
          strokeWidth={0.4}
          vectorEffect="non-scaling-stroke"
        />
        {props.data.map((datum, index) => {
          const barHeight = max > 0 ? (datum.value / max) * plotHeight : 0;
          const x = index * bandWidth + (bandWidth - barWidth) / 2;
          const y = paddingTop + plotHeight - barHeight;
          return (
            <g key={datum.label + '::' + index}>
              <rect
                x={x}
                y={y}
                width={barWidth}
                height={Math.max(barHeight, datum.value > 0 ? 1 : 0)}
                fill={datum.color || props.color}
                // 4px rounded data-end; the baseline end stays square because
                // the radius is small relative to the bar
                rx={1.5}
                style={{ cursor: props.onSelect ? 'pointer' : 'default' }}
                onMouseMove={(event) =>
                  tooltip.show(event, {
                    title: datum.label,
                    rows: [
                      {
                        label: datum.detail || strings.Charts_Responses,
                        value: datum.value + ' · ' + formatPercent(datum.value, props.total),
                        color: datum.color || props.color
                      }
                    ]
                  })
                }
                onMouseLeave={tooltip.hide}
                onClick={() => props.onSelect && props.onSelect(datum, index)}
              />
              {datum.value > 0 && (
                <text
                  className={styles.valueLabel}
                  x={index * bandWidth + bandWidth / 2}
                  y={y - 5}
                  textAnchor="middle"
                  style={{ fontSize: '10px' }}
                >
                  {formatCompact(datum.value)}
                </text>
              )}
              {index % labelEvery === 0 && (
                <text
                  className={styles.axisLabel}
                  x={index * bandWidth + bandWidth / 2}
                  y={height - 7}
                  textAnchor="middle"
                  style={{ fontSize: '10px' }}
                >
                  {datum.label}
                </text>
              )}
            </g>
          );
        })}
      </svg>
      {tooltip.render()}
    </div>
  );
};

// ---------------------------------------------------------------------------
// donut — part to whole
// ---------------------------------------------------------------------------

export interface IDonutChartProps {
  data: IChartDatum[];
  /** categorical slots, assigned in order */
  palette: string[];
  total: number;
  ink: IChartInk;
  centerLabel?: string;
  centerSubLabel?: string;
  onSelect?: (datum: IChartDatum, index: number) => void;
}

const polar = (cx: number, cy: number, r: number, angle: number): { x: number; y: number } => ({
  x: cx + r * Math.cos(angle - Math.PI / 2),
  y: cy + r * Math.sin(angle - Math.PI / 2)
});

/** Donut with a legend. Never relies on color alone — every slice is labelled. */
export const DonutChart: React.FunctionComponent<IDonutChartProps> = (props) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const tooltip = useTooltip(containerRef);

  const size = 132;
  const cx = size / 2;
  const cy = size / 2;
  const outer = 60;
  const inner = 38;
  const total = props.data.reduce((sum, d) => sum + d.value, 0);

  if (total <= 0) {
    return <div className={styles.chartEmpty}>{strings.Charts_NoAnswers}</div>;
  }

  let angle = 0;
  const arcs = props.data.map((datum, index) => {
    const share = datum.value / total;
    const sweep = share * Math.PI * 2;
    const start = angle;
    const end = angle + sweep;
    angle = end;
    const color = datum.color || props.palette[index % props.palette.length];

    // a full-circle slice can't be drawn as an arc — render it as a ring
    if (share >= 0.9999) {
      return { datum, color, path: '', full: true };
    }

    const p1 = polar(cx, cy, outer, start);
    const p2 = polar(cx, cy, outer, end);
    const p3 = polar(cx, cy, inner, end);
    const p4 = polar(cx, cy, inner, start);
    const largeArc = sweep > Math.PI ? 1 : 0;
    const path = [
      'M',
      p1.x,
      p1.y,
      'A',
      outer,
      outer,
      0,
      largeArc,
      1,
      p2.x,
      p2.y,
      'L',
      p3.x,
      p3.y,
      'A',
      inner,
      inner,
      0,
      largeArc,
      0,
      p4.x,
      p4.y,
      'Z'
    ].join(' ');
    return { datum, color, path, full: false };
  });

  return (
    <div className={styles.chartRoot} ref={containerRef}>
      <div className={styles.donutWrap}>
        <svg
          width={size}
          height={size}
          viewBox={'0 0 ' + size + ' ' + size}
          role="img"
          aria-label={formatString(strings.Charts_DonutChartAria, { count: props.data.length })}
          style={{ flexShrink: 0 }}
        >
          {arcs.map((arc, index) =>
            arc.full ? (
              <circle
                key={index}
                cx={cx}
                cy={cy}
                r={(outer + inner) / 2}
                fill="none"
                stroke={arc.color}
                strokeWidth={outer - inner}
              />
            ) : (
              <path
                key={index}
                d={arc.path}
                fill={arc.color}
                // the 2px surface gap separates touching segments
                stroke={props.ink.surface}
                strokeWidth={2}
                style={{ cursor: props.onSelect ? 'pointer' : 'default' }}
                onMouseMove={(event) =>
                  tooltip.show(event, {
                    title: arc.datum.label,
                    rows: [
                      {
                        label: strings.Charts_Responses,
                        value: arc.datum.value + ' · ' + formatPercent(arc.datum.value, total),
                        color: arc.color
                      }
                    ]
                  })
                }
                onMouseLeave={tooltip.hide}
                onClick={() => props.onSelect && props.onSelect(arc.datum, index)}
              />
            )
          )}
          {props.centerLabel && (
            <text className={styles.donutCenter} x={cx} y={cy + 2} textAnchor="middle">
              {props.centerLabel}
            </text>
          )}
          {props.centerSubLabel && (
            <text className={styles.donutCenterSub} x={cx} y={cy + 16} textAnchor="middle">
              {props.centerSubLabel}
            </text>
          )}
        </svg>

        {/* the legend is the dependable identity channel, always present */}
        <div className={styles.donutLegend}>
          {arcs.map((arc, index) => (
            <div key={index} className={styles.donutLegendRow}>
              <span className={styles.legendSwatch} style={{ background: arc.color }} />
              <span className={styles.donutLegendLabel} title={arc.datum.label}>
                {arc.datum.label}
              </span>
              <span className={styles.donutLegendValue}>
                {arc.datum.value} · {formatPercent(arc.datum.value, total)}
              </span>
            </div>
          ))}
        </div>
      </div>
      {tooltip.render()}
    </div>
  );
};

// ---------------------------------------------------------------------------
// area chart — responses over time
// ---------------------------------------------------------------------------

export interface IAreaPoint {
  label: string;
  value: number;
}

export interface IAreaChartProps {
  points: IAreaPoint[];
  color: string;
  ink: IChartInk;
  height?: number;
  valueLabel?: string;
}

/**
 * Single-series area with a crosshair. One series means no legend box — the
 * tile's own title says what is plotted.
 */
export const AreaChart: React.FunctionComponent<IAreaChartProps> = (props) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const tooltip = useTooltip(containerRef);
  const [hoverIndex, setHoverIndex] = React.useState<number>(-1);

  const width = 600;
  const height = props.height || 160;
  const padLeft = 34;
  const padRight = 8;
  const padTop = 12;
  const padBottom = 24;
  const plotWidth = width - padLeft - padRight;
  const plotHeight = height - padTop - padBottom;

  const points = props.points;
  if (points.length === 0) {
    return <div className={styles.chartEmpty}>{strings.Charts_NoResponses}</div>;
  }

  const max = Math.max(1, ...points.map((p) => p.value));
  // round the axis top to something clean
  const magnitude = Math.pow(10, Math.floor(Math.log(max) / Math.LN10));
  const axisMax = Math.ceil(max / magnitude) * magnitude || 1;

  const xAt = (index: number): number =>
    points.length === 1 ? padLeft + plotWidth / 2 : padLeft + (index / (points.length - 1)) * plotWidth;
  const yAt = (value: number): number => padTop + plotHeight - (value / axisMax) * plotHeight;

  const linePath = points
    .map((point, index) => (index === 0 ? 'M' : 'L') + xAt(index) + ' ' + yAt(point.value))
    .join(' ');
  const areaPath =
    linePath +
    ' L' +
    xAt(points.length - 1) +
    ' ' +
    (padTop + plotHeight) +
    ' L' +
    xAt(0) +
    ' ' +
    (padTop + plotHeight) +
    ' Z';

  const ticks = [0, axisMax / 2, axisMax];
  const labelEvery = Math.max(1, Math.ceil(points.length / 8));

  const handleMove = (event: React.MouseEvent<SVGRectElement>): void => {
    const bounds = (event.currentTarget as SVGRectElement).getBoundingClientRect();
    const ratio = (event.clientX - bounds.left) / bounds.width;
    const index = Math.max(0, Math.min(points.length - 1, Math.round(ratio * (points.length - 1))));
    setHoverIndex(index);
    tooltip.show(event, {
      title: points[index].label,
      rows: [
        {
          label: props.valueLabel || strings.Charts_Responses,
          value: String(points[index].value),
          color: props.color
        }
      ]
    });
  };

  return (
    <div className={styles.chartRoot} ref={containerRef}>
      <svg
        className={styles.svg}
        viewBox={'0 0 ' + width + ' ' + height}
        role="img"
        aria-label={strings.Charts_ResponsesOverTime}
      >
        {ticks.map((tick) => (
          <g key={tick}>
            <line
              x1={padLeft}
              y1={yAt(tick)}
              x2={width - padRight}
              y2={yAt(tick)}
              stroke={props.ink.grid}
              strokeWidth={1}
              vectorEffect="non-scaling-stroke"
            />
            <text className={styles.axisLabel} x={padLeft - 6} y={yAt(tick) + 3} textAnchor="end">
              {formatCompact(tick)}
            </text>
          </g>
        ))}

        {/* the area is a wash, never a saturated block */}
        <path d={areaPath} fill={props.color} fillOpacity={0.12} />
        <path
          d={linePath}
          fill="none"
          stroke={props.color}
          strokeWidth={2}
          strokeLinejoin="round"
          strokeLinecap="round"
        />

        {hoverIndex >= 0 && (
          <>
            <line
              x1={xAt(hoverIndex)}
              y1={padTop}
              x2={xAt(hoverIndex)}
              y2={padTop + plotHeight}
              stroke={props.ink.axis}
              strokeWidth={1}
              vectorEffect="non-scaling-stroke"
            />
            <circle
              cx={xAt(hoverIndex)}
              cy={yAt(points[hoverIndex].value)}
              r={4.5}
              fill={props.color}
              // 2px surface ring keeps the marker legible over the line
              stroke={props.ink.surface}
              strokeWidth={2}
            />
          </>
        )}

        {points.map((point, index) =>
          index % labelEvery === 0 ? (
            <text
              key={index}
              className={styles.axisLabel}
              x={xAt(index)}
              y={height - 7}
              textAnchor="middle"
            >
              {point.label}
            </text>
          ) : null
        )}

        {/* one transparent hit area rather than per-point targets */}
        <rect
          x={padLeft}
          y={padTop}
          width={plotWidth}
          height={plotHeight}
          fill="transparent"
          onMouseMove={handleMove}
          onMouseLeave={() => {
            setHoverIndex(-1);
            tooltip.hide();
          }}
        />
      </svg>
      {tooltip.render()}
    </div>
  );
};

// ---------------------------------------------------------------------------
// sparkline
// ---------------------------------------------------------------------------

export interface ISparklineProps {
  values: number[];
  color: string;
  width?: number;
  height?: number;
}

export const Sparkline: React.FunctionComponent<ISparklineProps> = (props) => {
  const width = props.width || 120;
  const height = props.height || 28;
  const values = props.values;
  if (values.length < 2) {
    return null;
  }
  const max = Math.max(...values);
  const min = Math.min(...values);
  const span = max - min || 1;
  const path = values
    .map((value, index) => {
      const x = (index / (values.length - 1)) * width;
      const y = height - ((value - min) / span) * (height - 4) - 2;
      return (index === 0 ? 'M' : 'L') + x.toFixed(1) + ' ' + y.toFixed(1);
    })
    .join(' ');

  return (
    <svg
      className={styles.statSpark}
      width={width}
      height={height}
      viewBox={'0 0 ' + width + ' ' + height}
      aria-hidden={true}
      focusable="false"
    >
      <path d={path} fill="none" stroke={props.color} strokeWidth={2} strokeLinecap="round" />
    </svg>
  );
};

// ---------------------------------------------------------------------------
// stat tile
// ---------------------------------------------------------------------------

export interface IStatTileProps {
  label: string;
  value: string;
  caption?: string;
  /** signed change vs a previous period */
  delta?: number;
  deltaSuffix?: string;
  /** false when a rise is bad (e.g. time to complete) */
  upIsGood?: boolean;
  trend?: number[];
  trendColor?: string;
  hero?: boolean;
}

export const StatTile: React.FunctionComponent<IStatTileProps> = (props) => {
  const hasDelta = typeof props.delta === 'number' && isFinite(props.delta);
  const upIsGood = props.upIsGood !== false;
  const deltaClass = !hasDelta
    ? styles.statDeltaFlat
    : props.delta === 0
      ? styles.statDeltaFlat
      : (props.delta as number) > 0 === upIsGood
        ? styles.statDeltaGood
        : styles.statDeltaBad;

  return (
    <div className={styles.statTile}>
      <span className={styles.statLabel}>{props.label}</span>
      <span className={styles.statValueRow}>
        <span className={props.hero ? styles.statValueHero : styles.statValue}>{props.value}</span>
        {hasDelta && (props.delta as number) !== 0 && (
          <span className={styles.statDelta + ' ' + deltaClass}>
            <Icon iconName={(props.delta as number) > 0 ? 'CaretSolidUp' : 'CaretSolidDown'} />
            {Math.abs(props.delta as number)}
            {props.deltaSuffix || ''}
          </span>
        )}
      </span>
      {props.caption && <span className={styles.statCaption}>{props.caption}</span>}
      {props.trend && props.trend.length > 1 && (
        <Sparkline values={props.trend} color={props.trendColor || 'currentColor'} />
      )}
    </div>
  );
};

// ---------------------------------------------------------------------------
// gauge — NPS and other bounded scores
// ---------------------------------------------------------------------------

export interface IGaugeProps {
  value: number;
  min: number;
  max: number;
  /** color bands from min to max, drawn behind the needle */
  bands: { upTo: number; color: string }[];
  ink: IChartInk;
  caption: string;
  displayValue: string;
}

/** Semi-circular gauge with a value arc. */
export const Gauge: React.FunctionComponent<IGaugeProps> = (props) => {
  const width = 180;
  const height = 104;
  const cx = width / 2;
  const cy = 92;
  const radius = 74;
  const thickness = 14;

  const span = props.max - props.min || 1;
  const clamped = Math.max(props.min, Math.min(props.max, props.value));
  const ratio = (clamped - props.min) / span;

  /** Arc path across a fraction of the top semicircle. */
  const arc = (fromRatio: number, toRatio: number): string => {
    const a1 = Math.PI + fromRatio * Math.PI;
    const a2 = Math.PI + toRatio * Math.PI;
    const outerFrom = { x: cx + radius * Math.cos(a1), y: cy + radius * Math.sin(a1) };
    const outerTo = { x: cx + radius * Math.cos(a2), y: cy + radius * Math.sin(a2) };
    const innerTo = {
      x: cx + (radius - thickness) * Math.cos(a2),
      y: cy + (radius - thickness) * Math.sin(a2)
    };
    const innerFrom = {
      x: cx + (radius - thickness) * Math.cos(a1),
      y: cy + (radius - thickness) * Math.sin(a1)
    };
    return [
      'M',
      outerFrom.x,
      outerFrom.y,
      'A',
      radius,
      radius,
      0,
      0,
      1,
      outerTo.x,
      outerTo.y,
      'L',
      innerTo.x,
      innerTo.y,
      'A',
      radius - thickness,
      radius - thickness,
      0,
      0,
      0,
      innerFrom.x,
      innerFrom.y,
      'Z'
    ].join(' ');
  };

  let cursor = 0;
  const bandPaths = props.bands.map((band) => {
    const to = Math.max(0, Math.min(1, (band.upTo - props.min) / span));
    const path = arc(cursor, to);
    cursor = to;
    return { path, color: band.color };
  });

  const needleAngle = Math.PI + ratio * Math.PI;
  const needleOuter = {
    x: cx + (radius + 3) * Math.cos(needleAngle),
    y: cy + (radius + 3) * Math.sin(needleAngle)
  };
  const needleInner = {
    x: cx + (radius - thickness - 5) * Math.cos(needleAngle),
    y: cy + (radius - thickness - 5) * Math.sin(needleAngle)
  };

  return (
    <div className={styles.gaugeWrap}>
      <svg
        width={width}
        height={height}
        viewBox={'0 0 ' + width + ' ' + height}
        role="img"
        aria-label={props.caption + ': ' + props.displayValue}
        style={{ flexShrink: 0 }}
      >
        {bandPaths.map((band, index) => (
          <path
            key={index}
            d={band.path}
            fill={band.color}
            stroke={props.ink.surface}
            strokeWidth={2}
          />
        ))}
        <line
          x1={needleInner.x}
          y1={needleInner.y}
          x2={needleOuter.x}
          y2={needleOuter.y}
          stroke={props.ink.label}
          strokeWidth={2.5}
          strokeLinecap="round"
        />
      </svg>
      <div className={styles.gaugeFigure}>
        <span className={styles.gaugeValue}>{props.displayValue}</span>
        <span className={styles.gaugeCaption}>{props.caption}</span>
      </div>
    </div>
  );
};

// ---------------------------------------------------------------------------
// diverging stacked bar — the right form for Likert / agreement scales
// ---------------------------------------------------------------------------

export interface IDivergingRow {
  label: string;
  /** counts per ordered category, low to high */
  counts: number[];
}

export interface IDivergingBarProps {
  rows: IDivergingRow[];
  /** ordered category labels, low to high */
  categories: string[];
  /** one color per category, from utils/theme divergingScale */
  colors: string[];
  ink: IChartInk;
}

/**
 * Ordered-scale share, centred on the neutral category.
 *
 * A plain stacked bar would imply the categories are unrelated; centring on
 * neutral is what makes "mostly disagree" visible at a glance. Each row is
 * normalized to its own total so rows with different response counts compare.
 */
export const DivergingBar: React.FunctionComponent<IDivergingBarProps> = (props) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const tooltip = useTooltip(containerRef);

  if (props.rows.length === 0 || props.categories.length === 0) {
    return <div className={styles.chartEmpty}>{strings.Charts_NoAnswers}</div>;
  }

  return (
    <div className={styles.chartRoot} ref={containerRef}>
      {props.rows.map((row) => {
        const total = row.counts.reduce((sum, n) => sum + n, 0);
        return (
          <div key={row.label} className={styles.divergingRow}>
            <span className={styles.divergingLabel} title={row.label}>
              {row.label}
            </span>
            <span className={styles.divergingTrack}>
              {row.counts.map((count, index) => {
                if (count <= 0) {
                  return null;
                }
                const share = total > 0 ? (count / total) * 100 : 0;
                const color = props.colors[index] || props.colors[props.colors.length - 1];
                return (
                  <span
                    key={index}
                    className={styles.divergingSegment}
                    style={{ width: share + '%', background: color }}
                    onMouseMove={(event) =>
                      tooltip.show(event, {
                        title: row.label,
                        rows: [
                          {
                            label: props.categories[index] || '',
                            value: count + ' · ' + formatPercent(count, total),
                            color
                          }
                        ]
                      })
                    }
                    onMouseLeave={tooltip.hide}
                  />
                );
              })}
            </span>
            <span className={styles.barValue}>{total}</span>
          </div>
        );
      })}
      <div className={styles.legend}>
        {props.categories.map((category, index) => (
          <span key={category + '::' + index} className={styles.legendItem}>
            <span
              className={styles.legendSwatch}
              style={{ background: props.colors[index] || props.colors[props.colors.length - 1] }}
            />
            {category}
          </span>
        ))}
      </div>
      {tooltip.render()}
    </div>
  );
};

// ---------------------------------------------------------------------------
// table fallback
// ---------------------------------------------------------------------------

export interface IDataTableProps {
  columns: [string, string];
  rows: { label: string; value: string }[];
}

/**
 * The relief channel. Several palette slots sit just below 3:1 against the
 * surface, which is only permissible because the values are also readable
 * another way — this is that other way.
 */
export const DataTable: React.FunctionComponent<IDataTableProps> = (props) => (
  <table className={styles.dataTable}>
    <thead>
      <tr>
        <th scope="col">{props.columns[0]}</th>
        <th scope="col">{props.columns[1]}</th>
      </tr>
    </thead>
    <tbody>
      {props.rows.map((row, index) => (
        <tr key={row.label + '::' + index}>
          <td>{row.label}</td>
          <td>{row.value}</td>
        </tr>
      ))}
    </tbody>
  </table>
);
