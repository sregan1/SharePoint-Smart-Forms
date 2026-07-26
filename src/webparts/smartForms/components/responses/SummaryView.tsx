import * as React from 'react';
import {
  ContextualMenuItemType,
  Dropdown,
  Icon,
  IconButton,
  IContextualMenuItem,
  IContextualMenuProps,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import {
  ChartKind,
  FieldType,
  getFieldTypeMeta,
  IFormDefinition,
  IFormField,
  IResponseItem
} from '../../models';
import {
  availableCharts,
  defaultChartFor,
  inputFields,
  segmentableFields
} from '../../utils/formUtils';
import { answeredCount, segmentBy } from '../../utils/analytics';
import { IThemeInfo, MAX_CATEGORICAL_SERIES } from '../../utils/theme';
import { QuestionChart } from './QuestionChart';

export interface ISummaryViewProps {
  definition: IFormDefinition;
  items: IResponseItem[];
  theme: IThemeInfo;
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
 * Every question, summarized.
 *
 * The differentiating feature here is **segment by**: pick any low-cardinality
 * question and every other card re-renders split by its answers, giving
 * "NPS by department" or "satisfaction by region" without leaving the page.
 * Segment choices are per-view state, not saved settings — it's an exploration
 * tool, and persisting it would silently change what the next person sees.
 */
export const SummaryView: React.FunctionComponent<ISummaryViewProps> = (props) => {
  const { definition, items, theme } = props;
  const fields = React.useMemo(
    () => inputFields(definition).filter((f) => f.provisioned !== false),
    [definition]
  );
  const segmentOptions = React.useMemo(() => segmentableFields(definition), [definition]);

  const [chartOverrides, setChartOverrides] = React.useState<{ [fieldId: string]: ChartKind }>({});
  const [segmentFieldId, setSegmentFieldId] = React.useState<string>('');

  const segmentField = segmentOptions.filter((f) => f.id === segmentFieldId)[0];

  const segments = React.useMemo(() => {
    if (!segmentField) {
      return undefined;
    }
    const all = segmentBy(segmentField, items);
    // too many segments makes every card unreadable; keep the largest few
    const sorted = all.slice().sort((a, b) => b.items.length - a.items.length);
    return sorted.slice(0, 6);
  }, [segmentField, items]);

  const chartFor = (field: IFormField): ChartKind => {
    const override = chartOverrides[field.id];
    return override && override !== 'auto' ? override : defaultChartFor(field);
  };

  const chartMenu = (field: IFormField): IContextualMenuProps => {
    const current = chartOverrides[field.id] || 'auto';
    const options: ChartKind[] = (['auto'] as ChartKind[]).concat(availableCharts(field));
    const items2: IContextualMenuItem[] = options.map((kind) => ({
      key: kind,
      text: CHART_LABELS[kind] || kind,
      iconProps: { iconName: CHART_ICONS[kind] || 'BarChartVertical' },
      canCheck: true,
      checked: current === kind,
      onClick: () => setChartOverrides((prev) => ({ ...prev, [field.id]: kind }))
    }));
    if (props.onDrillDown) {
      items2.push({ key: 'div', itemType: ContextualMenuItemType.Divider });
      items2.push({
        key: 'drill',
        text: 'See all answers in the table',
        iconProps: { iconName: 'Table' },
        onClick: () => props.onDrillDown && props.onDrillDown(field, '')
      });
    }
    return { items: items2 };
  };

  if (items.length === 0) {
    return (
      <div className={styles.empty}>
        <Icon iconName="BarChartVerticalFill" className={styles.emptyIcon} />
        <div className={styles.emptyTitle}>No responses yet</div>
        <p className={styles.emptyBody}>Every question will get its own summary here.</p>
      </div>
    );
  }

  const renderCard = (field: IFormField): React.ReactNode => {
    const meta = getFieldTypeMeta(field.type);
    const chart = chartFor(field);
    const wide = field.type === FieldType.Likert || field.type === FieldType.Ranking || !!segments;

    return (
      <div
        key={field.id}
        className={wide ? styles.summaryCard + ' ' + styles.summaryCardSpan2 : styles.summaryCard}
      >
        <div className={styles.summaryCardHeader}>
          <span className={styles.summaryCardIcon}>
            <Icon iconName={meta.icon} />
          </span>
          <div className={styles.summaryCardTitleGroup}>
            <div className={styles.summaryCardTitle}>{field.title || 'Untitled question'}</div>
            <div className={styles.summaryCardMeta}>
              {answeredCount(field, items)} of {items.length} answered · {meta.label}
            </div>
          </div>
          <div className={styles.summaryCardActions}>
            <IconButton
              iconProps={{ iconName: 'MoreVertical' }}
              title="Chart options"
              ariaLabel={'Chart options for ' + (field.title || 'question')}
              menuProps={chartMenu(field)}
              onRenderMenuIcon={() => null}
            />
          </div>
        </div>

        <div className={styles.summaryCardBody}>
          {segments ? (
            <div className={styles.segmentGroup}>
              {segments.map((segment) => (
                <div key={segment.label} className={styles.segmentBlock}>
                  <div className={styles.segmentHeader}>
                    <span>{segment.label}</span>
                    <span className={styles.segmentCount}>
                      {segment.items.length}{' '}
                      {segment.items.length === 1 ? 'response' : 'responses'}
                    </span>
                  </div>
                  <QuestionChart
                    field={field}
                    items={segment.items}
                    chart={chart}
                    theme={theme}
                    compact={true}
                    onDrillDown={props.onDrillDown}
                  />
                </div>
              ))}
            </div>
          ) : (
            <QuestionChart
              field={field}
              items={items}
              chart={chart}
              theme={theme}
              onDrillDown={props.onDrillDown}
            />
          )}
        </div>
      </div>
    );
  };

  return (
    <div>
      {segmentOptions.length > 0 && (
        <div className={styles.filterBar}>
          <Dropdown
            className={styles.filterDropdown}
            label={undefined}
            placeholder="Compare by…"
            options={[{ key: '', text: 'No comparison' }].concat(
              segmentOptions.map((f) => ({ key: f.id, text: 'By ' + (f.title || 'untitled') }))
            )}
            selectedKey={segmentFieldId}
            onChange={(_e, option) => setSegmentFieldId(option ? String(option.key) : '')}
          />
          {segmentField && (
            <span className={styles.filterChip}>
              <Icon iconName="GroupedList" /> Split by {segmentField.title || 'question'}
            </span>
          )}
        </div>
      )}

      {segments && segments.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          No responses answered &ldquo;{segmentField ? segmentField.title : ''}&rdquo;, so there is
          nothing to compare.
        </MessageBar>
      )}

      {segmentField && segments && segments.length > 0 && (
        <div className={styles.segmentNote}>
          Showing the {segments.length} largest groups. A response that selected several answers
          appears in every matching group, so group totals can add up to more than{' '}
          {items.length}.
        </div>
      )}

      <div className={segments ? styles.summaryGridWide : styles.summaryGrid}>
        {fields.map(renderCard)}
      </div>

      {fields.length > MAX_CATEGORICAL_SERIES && !segments && (
        <div className={styles.segmentNote}>
          Charts with more than {MAX_CATEGORICAL_SERIES} categories group the smallest into
          &ldquo;Other&rdquo; — switch a card to Table to see every value.
        </div>
      )}
    </div>
  );
};
