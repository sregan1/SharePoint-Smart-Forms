import * as React from 'react';
import styles from './ResponsesView.module.scss';
import { ChartKind, FieldType, IFormField, IResponseItem } from '../../models';
import { formatValue } from '../../utils/formUtils';
import {
  answeredCount,
  distributionFor,
  foldTail,
  histogram,
  IDistributionRow,
  likertStats,
  npsStats,
  numericStats,
  rankingStats,
  textAnswers,
  wordFrequency
} from '../../utils/analytics';
import {
  categoricalPalette,
  chartInk,
  divergingScale,
  IThemeInfo,
  MAX_CATEGORICAL_SERIES,
  ordinalRamp,
  STATUS_COLORS
} from '../../utils/theme';
import {
  BarChart,
  ColumnChart,
  DataTable,
  DivergingBar,
  DonutChart,
  formatPercent,
  Gauge,
  IChartDatum
} from '../charts/Charts';

export interface IQuestionChartProps {
  field: IFormField;
  items: IResponseItem[];
  chart: ChartKind;
  theme: IThemeInfo;
  /** clicking a category filters the table view to those respondents */
  onDrillDown?: (field: IFormField, categoryLabel: string) => void;
  /** compact rendering for a dashboard tile */
  compact?: boolean;
}

const toDatum = (row: IDistributionRow): IChartDatum => ({ label: row.label, value: row.count });

/**
 * Renders one question's results in whichever form the tile asks for.
 *
 * Shared by the dashboard and the summary so a chart looks and behaves the same
 * in both, and so the color-job decisions live in one place:
 *
 *  - nominal bars/columns are a SINGLE hue — the bar length already encodes the
 *    value, so coloring each bar separately would re-encode it and waste the
 *    identity channel;
 *  - part-to-whole (donut) is the only form here that spends categorical hues,
 *    capped at eight with the tail folded into "Other";
 *  - ordered scales (Likert) use the diverging form centred on neutral, not a
 *    plain stack, which is what makes "mostly disagree" legible at a glance;
 *  - NPS bands wear reserved status colors with labels, never hue alone.
 */
export const QuestionChart: React.FunctionComponent<IQuestionChartProps> = (props) => {
  const { field, items, theme } = props;
  const ink = React.useMemo(() => chartInk(theme), [theme]);
  const accent = theme.accent;
  const total = React.useMemo(() => answeredCount(field, items), [field, items]);

  const drill = props.onDrillDown
    ? (datum: IChartDatum) => {
        if (props.onDrillDown) {
          props.onDrillDown(field, datum.label);
        }
      }
    : undefined;

  // ----- NPS gauge -----
  if (field.type === FieldType.Scale && field.scaleAnalytics === 'nps' && props.chart !== 'table') {
    const stats = npsStats(field, items);
    if (!stats) {
      return <div className={styles.emptyCardText}>No scores yet.</div>;
    }
    const bandRows = [
      { label: 'Detractors (0–6)', count: stats.detractors, color: STATUS_COLORS.critical },
      { label: 'Passives (7–8)', count: stats.passives, color: STATUS_COLORS.warning },
      { label: 'Promoters (9–10)', count: stats.promoters, color: STATUS_COLORS.good }
    ];
    return (
      <div>
        <Gauge
          value={stats.score}
          min={-100}
          max={100}
          bands={[
            { upTo: 0, color: STATUS_COLORS.critical },
            { upTo: 50, color: STATUS_COLORS.warning },
            { upTo: 100, color: STATUS_COLORS.good }
          ]}
          ink={ink}
          caption="NPS score"
          displayValue={stats.score > 0 ? '+' + stats.score : String(stats.score)}
        />
        <BarChart
          data={bandRows.map((row) => ({ label: row.label, value: row.count, color: row.color }))}
          color={accent}
          total={stats.total}
          wideLabels={true}
          onSelect={undefined}
        />
      </div>
    );
  }

  // ----- Likert diverging stack -----
  if (field.type === FieldType.Likert) {
    const stats = likertStats(field, items);
    if (stats.rows.length === 0 || stats.categories.length === 0) {
      return <div className={styles.emptyCardText}>No answers yet.</div>;
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={['Statement', 'Average (1–' + stats.categories.length + ')']}
          rows={stats.rows.map((row) => ({
            label: row.label,
            value: row.total > 0 ? String(row.average) : '—'
          }))}
        />
      );
    }
    return (
      <DivergingBar
        rows={stats.rows.map((row) => ({ label: row.label, counts: row.counts }))}
        categories={stats.categories}
        colors={divergingScale(stats.categories.length, theme.isDark)}
        ink={ink}
      />
    );
  }

  // ----- ranking -----
  if (field.type === FieldType.Ranking) {
    const rows = rankingStats(field, items);
    if (rows.length === 0) {
      return <div className={styles.emptyCardText}>No rankings yet.</div>;
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={['Option', 'Average position']}
          rows={rows.map((row) => ({ label: row.label, value: String(row.avgPosition) }))}
        />
      );
    }
    const optionCount = Math.max(rows.length, 1);
    // invert: a better (lower) average position must read as a longer bar
    return (
      <BarChart
        data={rows.map((row, index) => ({
          label: index + 1 + '. ' + row.label,
          value: Math.max(0.02, (optionCount - row.avgPosition + 1) / optionCount),
          detail: 'Average position'
        }))}
        color={accent}
        total={1}
        wideLabels={true}
        valueFormatter={(datum) => {
          const row = rows[Number(datum.label.split('.')[0]) - 1];
          return row ? 'avg ' + row.avgPosition : '';
        }}
      />
    );
  }

  // ----- free text -----
  if (
    field.type === FieldType.Text ||
    field.type === FieldType.MultilineText ||
    field.type === FieldType.RichText ||
    field.type === FieldType.Address ||
    field.type === FieldType.Email ||
    field.type === FieldType.Phone ||
    field.type === FieldType.Hyperlink ||
    field.type === FieldType.Signature ||
    field.type === FieldType.FileUpload
  ) {
    const answers = textAnswers(field, items);
    if (answers.length === 0) {
      return <div className={styles.emptyCardText}>No answers yet.</div>;
    }
    const useWords =
      props.chart === 'words' &&
      (field.type === FieldType.Text ||
        field.type === FieldType.MultilineText ||
        field.type === FieldType.RichText);

    if (useWords) {
      const words = wordFrequency(field, items, props.compact ? 14 : 30);
      if (words.length === 0) {
        return <div className={styles.emptyCardText}>Not enough text to summarize yet.</div>;
      }
      const maxCount = words[0].count;
      return (
        <div className={styles.wordCloud}>
          {words.map((word) => {
            // scale by count, but keep everything readable
            const size = 12 + Math.round((word.count / maxCount) * 12);
            return (
              <button
                key={word.word}
                type="button"
                className={styles.wordItem}
                style={{ fontSize: size + 'px', fontWeight: word.count > maxCount / 2 ? 600 : 400 }}
                title={word.count + (word.count === 1 ? ' mention' : ' mentions')}
                onClick={() => drill && drill({ label: word.word, value: word.count })}
              >
                {word.word}
              </button>
            );
          })}
        </div>
      );
    }

    const shown = props.compact ? answers.slice(0, 3) : answers.slice(0, 50);
    return (
      <div>
        <ul className={styles.textSampleList}>
          {shown.map((answer) => (
            <li key={answer.id} className={styles.textSample}>
              {answer.text.length > 300 ? answer.text.slice(0, 300) + '…' : answer.text}
            </li>
          ))}
        </ul>
        {answers.length > shown.length && (
          <div className={styles.segmentNote}>
            Showing {shown.length} of {answers.length} answers.
          </div>
        )}
      </div>
    );
  }

  // ----- numeric -----
  if (
    field.type === FieldType.Number ||
    field.type === FieldType.Calculated ||
    field.type === FieldType.Slider
  ) {
    const stats = numericStats(field, items);
    if (!stats) {
      return <div className={styles.emptyCardText}>No numeric answers yet.</div>;
    }
    if (props.chart === 'histogram' || props.chart === 'column') {
      const bins = histogram(field, items, props.compact ? 6 : 10);
      return (
        <ColumnChart
          data={bins.map((bin) => ({ label: bin.label, value: bin.count, detail: 'Responses' }))}
          color={accent}
          total={stats.count}
          ink={ink}
          height={props.compact ? 130 : 170}
        />
      );
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={['Statistic', 'Value']}
          rows={[
            { label: 'Responses', value: String(stats.count) },
            { label: 'Minimum', value: formatValue(field, stats.min) },
            { label: 'Median', value: formatValue(field, stats.median) },
            { label: 'Average', value: formatValue(field, stats.avg) },
            { label: 'Maximum', value: formatValue(field, stats.max) },
            { label: 'Std deviation', value: String(stats.stdDev) },
            { label: 'Total', value: formatValue(field, stats.sum) }
          ]}
        />
      );
    }
    return (
      <div className={styles.statsRow}>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.median)}</span>
          <span className={styles.statLabel}>Median</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.avg)}</span>
          <span className={styles.statLabel}>Average</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.min)}</span>
          <span className={styles.statLabel}>Min</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.max)}</span>
          <span className={styles.statLabel}>Max</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{stats.stdDev}</span>
          <span className={styles.statLabel}>Std dev</span>
        </div>
      </div>
    );
  }

  // ----- date / time: bucket by value -----
  if (field.type === FieldType.Date || field.type === FieldType.Time) {
    const rows = distributionFor(field, items).map((row) => ({
      label: row.label.length > 12 ? row.label.slice(0, 12) : row.label,
      count: row.count
    }));
    if (rows.length === 0) {
      return <div className={styles.emptyCardText}>No answers yet.</div>;
    }
    return (
      <BarChart
        data={rows.slice(0, 12).map(toDatum)}
        color={accent}
        total={total}
        onSelect={drill}
      />
    );
  }

  // ----- everything else is a distribution -----
  const rows = distributionFor(field, items);
  if (rows.length === 0 || total === 0) {
    return <div className={styles.emptyCardText}>No answers yet.</div>;
  }

  if (props.chart === 'table') {
    return (
      <DataTable
        columns={['Answer', 'Responses']}
        rows={rows.map((row) => ({
          label: row.label,
          value: row.count + ' · ' + formatPercent(row.count, total)
        }))}
      />
    );
  }

  if (props.chart === 'donut') {
    const folded = foldTail(rows, MAX_CATEGORICAL_SERIES);
    return (
      <DonutChart
        data={folded.map(toDatum)}
        palette={categoricalPalette(theme.isDark)}
        total={total}
        ink={ink}
        centerLabel={String(total)}
        centerSubLabel={total === 1 ? 'response' : 'responses'}
        onSelect={drill}
      />
    );
  }

  if (props.chart === 'column') {
    return (
      <ColumnChart
        data={rows.map(toDatum)}
        color={accent}
        total={total}
        ink={ink}
        height={props.compact ? 130 : 170}
        labelEvery={rows.length > 12 ? 2 : 1}
        onSelect={drill}
      />
    );
  }

  if (props.chart === 'gauge' && field.type === FieldType.Scale) {
    const stats = numericStats(field, items);
    const min = typeof field.min === 'number' ? field.min : 1;
    const max = typeof field.max === 'number' ? field.max : 5;
    if (!stats) {
      return <div className={styles.emptyCardText}>No answers yet.</div>;
    }
    // an ordered scale reads as a one-hue ramp, low to high
    const ramp = ordinalRamp(accent, theme.surface, 3, theme.isDark);
    return (
      <Gauge
        value={stats.avg}
        min={min}
        max={max}
        bands={[
          { upTo: min + (max - min) / 3, color: ramp[0] },
          { upTo: min + (2 * (max - min)) / 3, color: ramp[1] },
          { upTo: max, color: ramp[2] }
        ]}
        ink={ink}
        caption={'Average of ' + stats.count}
        displayValue={String(stats.avg)}
      />
    );
  }

  // default: single-hue horizontal bars
  return (
    <BarChart
      data={rows.map(toDatum)}
      color={accent}
      total={total}
      wideLabels={rows.some((row) => row.label.length > 14)}
      onSelect={drill}
    />
  );
};
