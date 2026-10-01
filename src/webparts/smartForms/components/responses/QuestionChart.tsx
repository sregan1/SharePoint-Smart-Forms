import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';
import styles from './ResponsesView.module.scss';
import { ChartKind, FieldType, IFormField, IMessageBag, IResponseItem } from '../../models';
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

/** What a click on a chart category stands for; pass it to matchesRowFilter. */
export interface IDrillTarget {
  label: string;
  filterValue?: string;
  filterValues?: string[];
}

export interface IQuestionChartProps {
  field: IFormField;
  items: IResponseItem[];
  chart: ChartKind;
  theme: IThemeInfo;
  /** clicking a category filters the table view to those respondents */
  onDrillDown?: (field: IFormField, target?: IDrillTarget) => void;
  locale?: string;
  /** compact rendering for a dashboard tile */
  compact?: boolean;
}

const messages = strings as unknown as IMessageBag;

const toDatum = (row: IDistributionRow): IChartDatum => ({
  label: row.label,
  value: row.count,
  row: { label: row.label, filterValue: row.filterValue, filterValues: row.filterValues }
});

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
          props.onDrillDown(field, datum.row || { label: datum.label });
        }
      }
    : undefined;

  // ----- NPS gauge -----
  if (field.type === FieldType.Scale && field.scaleAnalytics === 'nps' && props.chart !== 'table') {
    const stats = npsStats(field, items);
    if (!stats) {
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoScores}</div>;
    }
    const bandRows = [
      { label: strings.Responses_Chart_Detractors, count: stats.detractors, color: STATUS_COLORS.critical },
      { label: strings.Responses_Chart_Passives, count: stats.passives, color: STATUS_COLORS.warning },
      { label: strings.Responses_Chart_Promoters, count: stats.promoters, color: STATUS_COLORS.good }
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
          caption={strings.Responses_Chart_NpsScore}
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
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoAnswers}</div>;
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={[
            strings.Responses_Chart_Statement,
            formatString(strings.Responses_Chart_AverageRange, { max: stats.categories.length })
          ]}
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
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoRankings}</div>;
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={[strings.Responses_Chart_Option, strings.Responses_Chart_AveragePosition]}
          rows={rows.map((row) => ({ label: row.label, value: String(row.avgPosition) }))}
        />
      );
    }
    const optionCount = Math.max(rows.length, 1);
    const rankByLabel: { [label: string]: (typeof rows)[number] } = {};
    rows.forEach((row, index) => {
      rankByLabel[index + 1 + '. ' + row.label] = row;
    });
    // invert: a better (lower) average position must read as a longer bar
    return (
      <BarChart
        data={rows.map((row, index) => ({
          label: index + 1 + '. ' + row.label,
          value: Math.max(0.02, (optionCount - row.avgPosition + 1) / optionCount),
          detail: strings.Responses_Chart_AveragePosition
        }))}
        color={accent}
        total={1}
        wideLabels={true}
        valueFormatter={(datum) => {
          const row = rankByLabel[datum.label];
          return row ? formatString(strings.Responses_Chart_AvgValue, { value: row.avgPosition }) : '';
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
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoAnswers}</div>;
    }
    const useWords =
      props.chart === 'words' &&
      (field.type === FieldType.Text ||
        field.type === FieldType.MultilineText ||
        field.type === FieldType.RichText);

    if (useWords) {
      const words = wordFrequency(field, items, props.compact ? 14 : 30);
      if (words.length === 0) {
        return <div className={styles.emptyCardText}>{strings.Responses_Chart_NotEnoughText}</div>;
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
                title={formatString(
                  word.count === 1
                    ? strings.Responses_Chart_MentionOne
                    : strings.Responses_Chart_MentionOther,
                  { count: word.count }
                )}
                onClick={() => props.onDrillDown && props.onDrillDown(field, { label: word.word })}
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
            {formatString(strings.Responses_Chart_ShowingAnswers, {
              shown: shown.length,
              total: answers.length
            })}
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
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoNumeric}</div>;
    }
    if (props.chart === 'histogram' || props.chart === 'column') {
      const bins = histogram(field, items, props.compact ? 6 : 10);
      return (
        <ColumnChart
          data={bins.map((bin) => ({ label: bin.label, value: bin.count, detail: strings.Responses_Chart_Responses }))}
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
          columns={[strings.Responses_Chart_Statistic, strings.Responses_Chart_Value]}
          rows={[
            { label: strings.Responses_Chart_Responses, value: String(stats.count) },
            { label: strings.Responses_Chart_Minimum, value: formatValue(field, stats.min) },
            { label: strings.Responses_Chart_Median, value: formatValue(field, stats.median) },
            { label: strings.Responses_Chart_Average, value: formatValue(field, stats.avg) },
            { label: strings.Responses_Chart_Maximum, value: formatValue(field, stats.max) },
            { label: strings.Responses_Chart_StdDeviation, value: String(stats.stdDev) },
            { label: strings.Responses_Chart_Total, value: formatValue(field, stats.sum) }
          ]}
        />
      );
    }
    return (
      <div className={styles.statsRow}>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.median)}</span>
          <span className={styles.statLabel}>{strings.Responses_Chart_Median}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.avg)}</span>
          <span className={styles.statLabel}>{strings.Responses_Chart_Average}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.min)}</span>
          <span className={styles.statLabel}>{strings.Responses_Chart_Min}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{formatValue(field, stats.max)}</span>
          <span className={styles.statLabel}>{strings.Responses_Chart_Max}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statValue}>{stats.stdDev}</span>
          <span className={styles.statLabel}>{strings.Responses_Chart_StdDev}</span>
        </div>
      </div>
    );
  }

  // ----- date / time: chronological, grouped by local date (Time: by hour) -----
  if (field.type === FieldType.Date || field.type === FieldType.Time) {
    const rows = distributionFor(field, items, messages, props.locale);
    if (rows.length === 0) {
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoAnswers}</div>;
    }
    if (props.chart === 'table') {
      return (
        <DataTable
          columns={[strings.Responses_Chart_Answer, strings.Responses_Chart_Responses]}
          rows={rows.map((row) => ({ label: row.label, value: row.count + ' · ' + formatPercent(row.count, total) }))}
        />
      );
    }
    return (
      <ColumnChart
        data={rows.map(toDatum)}
        color={accent}
        total={total}
        ink={ink}
        height={props.compact ? 130 : 170}
        labelEvery={Math.max(1, Math.ceil(rows.length / 8))}
        onSelect={drill}
      />
    );
  }

  // ----- everything else is a distribution -----
  const rows = distributionFor(field, items, messages, props.locale);
  if (rows.length === 0 || total === 0) {
    return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoAnswers}</div>;
  }

  if (props.chart === 'table') {
    return (
      <DataTable
        columns={[strings.Responses_Chart_Answer, strings.Responses_Chart_Responses]}
        rows={rows.map((row) => ({
          label: row.label,
          value: row.count + ' · ' + formatPercent(row.count, total)
        }))}
      />
    );
  }

  if (props.chart === 'donut') {
    const folded = foldTail(rows, MAX_CATEGORICAL_SERIES, messages);
    return (
      <DonutChart
        data={folded.map(toDatum)}
        palette={categoricalPalette(theme.isDark)}
        total={total}
        ink={ink}
        centerLabel={String(total)}
        centerSubLabel={
          total === 1 ? strings.Responses_Chart_ResponseOne : strings.Responses_Chart_ResponseOther
        }
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
      return <div className={styles.emptyCardText}>{strings.Responses_Chart_NoAnswers}</div>;
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
        caption={formatString(strings.Responses_Chart_AverageOf, { count: stats.count })}
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
