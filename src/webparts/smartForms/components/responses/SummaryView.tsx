import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';
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
  IMessageBag,
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
import { IDrillTarget, QuestionChart } from './QuestionChart';

export interface ISummaryViewProps {
  definition: IFormDefinition;
  items: IResponseItem[];
  theme: IThemeInfo;
  onDrillDown?: (field: IFormField, target?: IDrillTarget) => void;
  locale?: string;
}

const messages = strings as unknown as IMessageBag;

const chartLabel = (kind: string): string => {
  const labels: { [kind: string]: string } = {
    auto: strings.Responses_Dashboard_ChartAuto,
    bar: strings.Responses_Dashboard_ChartBars,
    column: strings.Responses_Dashboard_ChartColumns,
    donut: strings.Responses_Dashboard_ChartDonut,
    stat: strings.Responses_Dashboard_ChartStatistics,
    gauge: strings.Responses_Dashboard_ChartGauge,
    histogram: strings.Responses_Dashboard_ChartHistogram,
    words: strings.Responses_Dashboard_ChartWordFrequency,
    table: strings.Responses_Dashboard_ChartTable
  };
  return labels[kind] || kind;
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
    const all = segmentBy(segmentField, items, messages, props.locale);
    // too many segments makes every card unreadable; keep the largest few
    const sorted = all.slice().sort((a, b) => b.items.length - a.items.length);
    return sorted.slice(0, 6);
  }, [segmentField, items, props.locale]);

  const chartFor = (field: IFormField): ChartKind => {
    const override = chartOverrides[field.id];
    return override && override !== 'auto' ? override : defaultChartFor(field);
  };

  const chartMenu = (field: IFormField): IContextualMenuProps => {
    const current = chartOverrides[field.id] || 'auto';
    const options: ChartKind[] = (['auto'] as ChartKind[]).concat(availableCharts(field));
    const items2: IContextualMenuItem[] = options.map((kind) => ({
      key: kind,
      text: chartLabel(kind),
      iconProps: { iconName: CHART_ICONS[kind] || 'BarChartVertical' },
      canCheck: true,
      checked: current === kind,
      onClick: () => setChartOverrides((prev) => ({ ...prev, [field.id]: kind }))
    }));
    if (props.onDrillDown) {
      items2.push({ key: 'div', itemType: ContextualMenuItemType.Divider });
      items2.push({
        key: 'drill',
        text: strings.Responses_Summary_SeeAllAnswers,
        iconProps: { iconName: 'Table' },
        onClick: () => props.onDrillDown && props.onDrillDown(field)
      });
    }
    return { items: items2 };
  };

  if (items.length === 0) {
    return (
      <div className={styles.empty}>
        <Icon iconName="BarChartVerticalFill" className={styles.emptyIcon} />
        <div className={styles.emptyTitle}>{strings.Responses_Dashboard_EmptyTitle}</div>
        <p className={styles.emptyBody}>{strings.Responses_Summary_EmptyBody}</p>
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
            <div className={styles.summaryCardTitle}>{field.title || strings.Responses_Dashboard_UntitledQuestion}</div>
            <div className={styles.summaryCardMeta}>
              {formatString(strings.Responses_Summary_AnsweredOfType, {
                answered: answeredCount(field, items),
                total: items.length,
                type: meta.label
              })}
            </div>
          </div>
          <div className={styles.summaryCardActions}>
            <IconButton
              iconProps={{ iconName: 'MoreVertical' }}
              title={strings.Responses_Dashboard_ChartOptions}
              ariaLabel={formatString(strings.Responses_Dashboard_ChartOptionsFor, {
                title: field.title || strings.Responses_Dashboard_QuestionFallback
              })}
              menuProps={chartMenu(field)}
              onRenderMenuIcon={() => null}
            />
          </div>
        </div>

        <div className={styles.summaryCardBody}>
          {segments ? (
            <div className={styles.segmentGroup}>
              {segments.map((segment) => (
                <div key={segment.filterValue} className={styles.segmentBlock}>
                  <div className={styles.segmentHeader}>
                    <span>{segment.label}</span>
                    <span className={styles.segmentCount}>
                      {formatString(
                        segment.items.length === 1
                          ? strings.Responses_Summary_ResponseCountOne
                          : strings.Responses_Summary_ResponseCountOther,
                        { count: segment.items.length }
                      )}
                    </span>
                  </div>
                  <QuestionChart
                    field={field}
                    items={segment.items}
                    chart={chart}
                    theme={theme}
                    compact={true}
                    locale={props.locale}
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
              locale={props.locale}
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
            placeholder={strings.Responses_Summary_CompareBy}
            options={[{ key: '', text: strings.Responses_Summary_NoComparison }].concat(
              segmentOptions.map((f) => ({ key: f.id, text: formatString(strings.Responses_Summary_ByField, {
                  title: f.title || strings.Responses_Summary_Untitled
                }) }))
            )}
            selectedKey={segmentFieldId}
            onChange={(_e, option) => setSegmentFieldId(option ? String(option.key) : '')}
          />
          {segmentField && (
            <span className={styles.filterChip}>
              <Icon iconName="GroupedList" />{' '}
              {formatString(strings.Responses_Summary_SplitBy, {
                title: segmentField.title || strings.Responses_Dashboard_QuestionFallback
              })}
            </span>
          )}
        </div>
      )}

      {segments && segments.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          {formatString(strings.Responses_Summary_NothingToCompare, {
            title: segmentField ? segmentField.title : ''
          })}
        </MessageBar>
      )}

      {segmentField && segments && segments.length > 0 && (
        <div className={styles.segmentNote}>
          {formatString(strings.Responses_Summary_SegmentNote, {
            groups: segments.length,
            total: items.length
          })}
        </div>
      )}

      <div className={segments ? styles.summaryGridWide : styles.summaryGrid}>
        {fields.map(renderCard)}
      </div>

      {fields.length > MAX_CATEGORICAL_SERIES && !segments && (
        <div className={styles.segmentNote}>
          {formatString(strings.Responses_Summary_OtherNote, { max: MAX_CATEGORICAL_SERIES })}
        </div>
      )}
    </div>
  );
};
