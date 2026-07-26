import * as React from 'react';
import {
  ActionButton,
  ContextualMenuItemType,
  Dropdown,
  Icon,
  IconButton,
  IContextualMenuItem,
  IContextualMenuProps
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import {
  ChartKind,
  FieldType,
  IDashboardSettings,
  IDashboardTile,
  IFormDefinition,
  IFormField,
  IResponseItem
} from '../../models';
import { availableCharts, defaultChartFor, inputFields } from '../../utils/formUtils';
import {
  answeredCount,
  computeKpis,
  formatDuration,
  npsStats,
  numericStats,
  suggestGrain,
  timeline,
  TimeGrain
} from '../../utils/analytics';
import { chartInk, IThemeInfo } from '../../utils/theme';
import { AreaChart, formatCompact, StatTile } from '../charts/Charts';
import { QuestionChart } from './QuestionChart';

export interface IDashboardViewProps {
  definition: IFormDefinition;
  items: IResponseItem[];
  theme: IThemeInfo;
  /** owners can rearrange tiles; viewers get the saved layout read-only */
  canEdit: boolean;
  onSettingsChange?: (dashboard: IDashboardSettings) => void;
  onDrillDown?: (field: IFormField, categoryLabel: string) => void;
}

const CHART_LABELS: { [kind: string]: string } = {
  auto: 'Automatic',
  bar: 'Bars',
  column: 'Columns',
  donut: 'Donut',
  stat: 'Statistics',
  gauge: 'Gauge',
  histogram: 'Histogram',
  words: 'Word frequency',
  table: 'Table'
};

const CHART_ICONS: { [kind: string]: string } = {
  auto: 'Lightbulb',
  bar: 'BarChartHorizontal',
  column: 'BarChartVertical',
  donut: 'DonutChart',
  stat: 'NumberField',
  gauge: 'SpeedHigh',
  histogram: 'Histogram',
  words: 'TextField',
  table: 'Table'
};

/**
 * Score for auto-selecting which questions lead the dashboard.
 *
 * Without this a dashboard is either empty until configured, or shows every
 * question in definition order — which buries the interesting one. "Interesting"
 * here means: a lopsided choice split, a low rating, a wide numeric spread, or
 * any NPS at all.
 */
const interestScore = (field: IFormField, items: IResponseItem[]): number => {
  const answered = answeredCount(field, items);
  if (answered === 0) {
    return -1;
  }
  const coverage = answered / Math.max(1, items.length);

  switch (field.type) {
    case FieldType.Scale:
      // an NPS is always the headline when there is one
      return field.scaleAnalytics === 'nps' ? 100 + coverage * 10 : 60 + coverage * 10;
    case FieldType.Rating: {
      const stats = numericStats(field, items);
      if (!stats) {
        return 0;
      }
      const max = field.maxRating || 5;
      // a low average is more worth surfacing than a high one
      return 70 - (stats.avg / max) * 30 + coverage * 10;
    }
    case FieldType.YesNo:
    case FieldType.Consent:
      return 50 + coverage * 10;
    case FieldType.Choice:
    case FieldType.ImageChoice:
    case FieldType.Lookup:
      return 65 + coverage * 10;
    case FieldType.Likert:
      return 75 + coverage * 5;
    case FieldType.Number:
    case FieldType.Calculated:
    case FieldType.Slider: {
      const stats = numericStats(field, items);
      if (!stats) {
        return 0;
      }
      const spread = stats.max - stats.min;
      const relative = stats.avg !== 0 ? Math.min(Math.abs(spread / stats.avg), 3) / 3 : 0;
      return 45 + relative * 20 + coverage * 5;
    }
    case FieldType.Ranking:
      return 55 + coverage * 5;
    case FieldType.Person:
      return 30 + coverage * 5;
    default:
      // free text is useful but rarely the headline
      return 20 + coverage * 5;
  }
};

const AUTO_TILE_COUNT = 6;

export const DashboardView: React.FunctionComponent<IDashboardViewProps> = (props) => {
  const { definition, items, theme } = props;
  const ink = React.useMemo(() => chartInk(theme), [theme]);
  const settings: IDashboardSettings = definition.settings.dashboard || {
    showKpis: true,
    showTimeline: true,
    timelineGrain: 'day',
    tiles: []
  };

  const fields = React.useMemo(() => inputFields(definition), [definition]);
  const questionCount = fields.length;

  const kpis = React.useMemo(() => computeKpis(items, questionCount), [items, questionCount]);

  const [grain, setGrain] = React.useState<TimeGrain>(() => settings.timelineGrain || suggestGrain(items));
  const points = React.useMemo(() => timeline(items, grain), [items, grain]);

  /** Tiles to render: pinned first, then the highest-scoring questions. */
  const tiles = React.useMemo(() => {
    const configured: { [fieldId: string]: IDashboardTile } = {};
    (settings.tiles || []).forEach((tile) => {
      configured[tile.fieldId] = tile;
    });

    const pinned = (settings.tiles || [])
      .filter((tile) => tile.pinned && !tile.hidden)
      .map((tile) => fields.filter((f) => f.id === tile.fieldId)[0])
      .filter((field) => !!field);

    const hidden: { [fieldId: string]: boolean } = {};
    (settings.tiles || []).forEach((tile) => {
      if (tile.hidden) {
        hidden[tile.fieldId] = true;
      }
    });

    const auto = fields
      .filter((field) => !hidden[field.id] && pinned.filter((p) => p.id === field.id).length === 0)
      .map((field) => ({ field: field, score: interestScore(field, items) }))
      .filter((entry) => entry.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, Math.max(0, AUTO_TILE_COUNT - pinned.length))
      .map((entry) => entry.field);

    return pinned.concat(auto).map((field) => ({
      field: field,
      chart: (configured[field.id] && configured[field.id].chart) || ('auto' as ChartKind),
      pinned: !!(configured[field.id] && configured[field.id].pinned)
    }));
  }, [fields, items, settings.tiles]);

  const patchTile = (fieldId: string, patch: Partial<IDashboardTile>): void => {
    if (!props.onSettingsChange) {
      return;
    }
    const existing = (settings.tiles || []).slice();
    const index = existing.map((t) => t.fieldId).indexOf(fieldId);
    if (index >= 0) {
      existing[index] = { ...existing[index], ...patch };
    } else {
      existing.push({ fieldId: fieldId, chart: 'auto', ...patch });
    }
    props.onSettingsChange({ ...settings, tiles: existing, timelineGrain: grain });
  };

  const chartMenu = (field: IFormField, current: ChartKind): IContextualMenuProps => {
    const options: ChartKind[] = (['auto'] as ChartKind[]).concat(availableCharts(field));
    const isPinned = tiles.filter((t) => t.field.id === field.id && t.pinned).length > 0;
    const items: IContextualMenuItem[] = options.map((kind) => ({
      key: kind,
      text: CHART_LABELS[kind] || kind,
      iconProps: { iconName: CHART_ICONS[kind] || 'BarChartVertical' },
      canCheck: true,
      checked: current === kind,
      onClick: () => patchTile(field.id, { chart: kind })
    }));
    items.push({ key: 'div', itemType: ContextualMenuItemType.Divider });
    items.push({
      key: 'pin',
      text: isPinned ? 'Unpin from dashboard' : 'Pin to dashboard',
      iconProps: { iconName: 'Pin' },
      onClick: () => patchTile(field.id, { pinned: !isPinned })
    });
    items.push({
      key: 'hide',
      text: 'Hide from dashboard',
      iconProps: { iconName: 'Hide3' },
      onClick: () => patchTile(field.id, { hidden: true, pinned: false })
    });
    return { items: items };
  };

  if (items.length === 0) {
    return (
      <div className={styles.empty}>
        <Icon iconName="BarChartVertical" className={styles.emptyIcon} />
        <div className={styles.emptyTitle}>No responses yet</div>
        <p className={styles.emptyBody}>
          Once people start answering, this dashboard fills in automatically — headline numbers, a
          response timeline, and a chart for the questions worth watching.
        </p>
      </div>
    );
  }

  const deltaLast7 = kpis.previous7 > 0 ? kpis.last7 - kpis.previous7 : undefined;

  return (
    <div className={styles.dashboard}>
      {settings.showKpis !== false && (
        <div className={styles.dashboardSection}>
          <div className={styles.statRow}>
              <StatTile
                label="Total responses"
                value={formatCompact(kpis.total)}
                caption={
                  kpis.lastResponse ? 'Latest ' + kpis.lastResponse.toLocaleDateString() : undefined
                }
                trend={kpis.trend}
                trendColor={theme.accent}
                hero={true}
              />
              <StatTile
                label="Last 7 days"
                value={String(kpis.last7)}
                delta={deltaLast7}
                caption={deltaLast7 === undefined ? 'No prior week to compare' : 'vs previous 7 days'}
              />
              <StatTile label="Today" value={String(kpis.today)} caption="Since midnight" />
              <StatTile
                label="Unique respondents"
                value={String(kpis.uniqueRespondents)}
                caption={
                  kpis.uniqueRespondents < kpis.total
                    ? kpis.total - kpis.uniqueRespondents + ' repeat responses'
                    : 'One each'
                }
              />
              {kpis.completionRate !== undefined && (
                <StatTile
                  label="Questions answered"
                  value={kpis.completionRate + '%'}
                  caption="Averaged across responses"
                />
              )}
              {kpis.medianDuration !== undefined && (
                <StatTile
                  label="Median time to complete"
                  value={formatDuration(kpis.medianDuration)}
                  upIsGood={false}
                  caption="Recorded on submit"
                />
              )}
          </div>
        </div>
      )}

      {settings.showTimeline !== false && points.length > 1 && (
        <div className={styles.timelineCard}>
          <div className={styles.timelineHeader}>
            <span className={styles.timelineTitle}>Responses over time</span>
            <Dropdown
              options={[
                { key: 'day', text: 'By day' },
                { key: 'week', text: 'By week' },
                { key: 'month', text: 'By month' }
              ]}
              selectedKey={grain}
              styles={{ root: { minWidth: 110 } }}
              onChange={(_e, option) => {
                if (!option) {
                  return;
                }
                const next = String(option.key) as TimeGrain;
                setGrain(next);
                if (props.onSettingsChange) {
                  props.onSettingsChange({ ...settings, timelineGrain: next });
                }
              }}
            />
          </div>
          <AreaChart points={points} color={theme.accent} ink={ink} height={170} />
        </div>
      )}

      <div className={styles.dashboardSection}>
        <div className={styles.dashboardSectionHeader}>
          <h4 className={styles.dashboardSectionTitle}>Highlights</h4>
          {props.canEdit && (settings.tiles || []).filter((t) => t.hidden).length > 0 && (
            <ActionButton
              iconProps={{ iconName: 'RedEye' }}
              text="Show hidden questions"
              onClick={() =>
                props.onSettingsChange &&
                props.onSettingsChange({
                  ...settings,
                  tiles: (settings.tiles || []).map((t) => ({ ...t, hidden: false }))
                })
              }
            />
          )}
        </div>

        <div className={styles.summaryGrid}>
          {tiles.map((tile) => {
            const chart = tile.chart === 'auto' ? defaultChartFor(tile.field) : tile.chart;
            const wide =
              tile.field.type === FieldType.Likert || tile.field.type === FieldType.Ranking;
            return (
              <div
                key={tile.field.id}
                className={
                  wide ? styles.summaryCard + ' ' + styles.summaryCardSpan2 : styles.summaryCard
                }
              >
                <div className={styles.summaryCardHeader}>
                  <span className={styles.summaryCardIcon}>
                    <Icon iconName={CHART_ICONS[chart] || 'BarChartVertical'} />
                  </span>
                  <div className={styles.summaryCardTitleGroup}>
                    <div className={styles.summaryCardTitle}>
                      {tile.field.title || 'Untitled question'}
                    </div>
                    <div className={styles.summaryCardMeta}>
                      {answeredCount(tile.field, items)} of {items.length} answered
                      {tile.pinned ? ' · pinned' : ''}
                    </div>
                  </div>
                  {props.canEdit && (
                    <div className={styles.summaryCardActions}>
                      <IconButton
                        iconProps={{ iconName: 'MoreVertical' }}
                        title="Chart options"
                        ariaLabel={'Chart options for ' + (tile.field.title || 'question')}
                        menuProps={chartMenu(tile.field, tile.chart)}
                        onRenderMenuIcon={() => null}
                      />
                    </div>
                  )}
                </div>
                <div className={styles.summaryCardBody}>
                  <QuestionChart
                    field={tile.field}
                    items={items}
                    chart={chart}
                    theme={theme}
                    compact={true}
                    onDrillDown={props.onDrillDown}
                  />
                </div>
              </div>
            );
          })}
        </div>

        {tiles.length === 0 && (
          <div className={styles.emptyCardText}>
            Every question is hidden from the dashboard. Use &ldquo;Show hidden questions&rdquo; to bring
            them back.
          </div>
        )}
      </div>
    </div>
  );
};

/** NPS shown in the KPI strip when the form has one, used by the shell. */
export const headlineNps = (
  definition: IFormDefinition,
  items: IResponseItem[]
): { score: number; total: number } | undefined => {
  const field = inputFields(definition).filter(
    (f) => f.type === FieldType.Scale && f.scaleAnalytics === 'nps'
  )[0];
  if (!field) {
    return undefined;
  }
  const stats = npsStats(field, items);
  return stats ? { score: stats.score, total: stats.total } : undefined;
};
