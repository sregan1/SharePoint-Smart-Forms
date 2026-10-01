import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString, isRtlLocale } from '../../utils/localeUtils';
import {
  CommandBar,
  Dropdown,
  Icon,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  SearchBox,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import {
  FieldType,
  IDashboardSettings,
  IFormDefinition,
  IFormField,
  IMessageBag,
  IResponseItem
} from '../../models';
import {
  buildCsv,
  downloadTextFile,
  formatSharePointValue,
  inputFields,
  segmentableFields
} from '../../utils/formUtils';
import { matchesRowFilter, segmentBy } from '../../utils/analytics';
import { IThemeInfo } from '../../utils/theme';
import { ApprovalStatus, IResponseItemEx, SharePointService } from '../../services/SharePointService';
import { debugLog, logError } from '../../utils/debug';
import { IDrillTarget } from './QuestionChart';
import { ResponseDetailPanel } from './ResponseDetailPanel';
import { SummaryView } from './SummaryView';
import { DashboardView } from './DashboardView';
import { ResponseTable } from './ResponseTable';

export interface IResponsesViewProps {
  definition: IFormDefinition;
  listId: string;
  listTitle: string;
  spService: SharePointService;
  theme: IThemeInfo;
  isOwner: boolean;
  /** persists dashboard tile layout back into the form definition */
  onDashboardChange?: (dashboard: IDashboardSettings) => void;
  onCopyShareLink?: () => void;
  onPreview?: () => void;
  /** culture name (e.g. 'de-DE') for date/number formatting and RTL; supplied by the shell */
  locale?: string;
}

const messages = strings as unknown as IMessageBag;
const SEARCH_DEBOUNCE_MS = 250;
const APPROVAL_BATCH = 5;

type ApprovalFilter = '' | ApprovalStatus;

interface IAnswerFilter {
  fieldId: string;
  row: IDrillTarget;
}

type ViewKey = 'dashboard' | 'summary' | 'table';

type DateRange = 'all' | '7' | '30' | '90' | 'today';

const getDateRangeOptions = (): IDropdownOption[] => [
  { key: 'all', text: strings.Responses_View_AllTime },
  { key: 'today', text: strings.Responses_View_Today },
  { key: '7', text: strings.Responses_View_Last7Days },
  { key: '30', text: strings.Responses_View_Last30Days },
  { key: '90', text: strings.Responses_View_Last90Days }
];

/**
 * The owner's results experience.
 *
 * One filter bar sits above every view, so narrowing to "last 30 days, Marketing
 * only" narrows the dashboard, the summary and the table together — previously
 * search applied to the table alone and the summary always showed everything,
 * which quietly reported two different numbers on one screen.
 */
export const ResponsesView: React.FunctionComponent<IResponsesViewProps> = (props) => {
  const { definition, listId, spService, theme } = props;
  const approvalEnabled = definition.settings.enableApproval === true;
  const dir = isRtlLocale(props.locale) ? 'rtl' : undefined;

  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [items, setItems] = React.useState<IResponseItemEx[]>([]);
  const [truncated, setTruncated] = React.useState<boolean>(false);
  const [view, setView] = React.useState<ViewKey>('dashboard');
  const [notice, setNotice] = React.useState<string>('');

  // ----- shared filters -----
  const [search, setSearch] = React.useState<string>('');
  const [debouncedSearch, setDebouncedSearch] = React.useState<string>('');
  const [dateRange, setDateRange] = React.useState<DateRange>('all');
  const [segmentFieldId, setSegmentFieldId] = React.useState<string>('');
  const [answerFilter, setAnswerFilter] = React.useState<IAnswerFilter | undefined>(undefined);
  const [approvalFilter, setApprovalFilter] = React.useState<ApprovalFilter>('');

  const segmentValue =
    answerFilter && answerFilter.fieldId === segmentFieldId ? answerFilter.row.filterValue : undefined;

  const [selectedId, setSelectedId] = React.useState<number | undefined>(undefined);
  const [visibleColumns, setVisibleColumns] = React.useState<string[]>([]);

  React.useEffect(() => {
    const handle = window.setTimeout(() => setDebouncedSearch(search), SEARCH_DEBOUNCE_MS);
    return () => window.clearTimeout(handle);
  }, [search]);

  const fields = React.useMemo(
    () => inputFields(definition).filter((f) => f.provisioned !== false),
    [definition]
  );
  const filterableFields = React.useMemo(() => segmentableFields(definition), [definition]);

  // Depend on the serialized definition rather than the object: the shell
  // re-parses it on every render, so an object dependency refetched constantly.
  const definitionKey = React.useMemo(
    () => fields.map((f) => f.internalName).join('|'),
    [fields]
  );

  // Searchable text is expensive (formats every answer), so compute it once per
  // item object; a new load or optimistic update yields new objects and so a fresh entry.
  const searchTextCache = React.useRef<{ key: string; map: WeakMap<IResponseItem, string> }>({
    key: '',
    map: new WeakMap()
  });

  // A generation counter rather than a per-call `cancelled` flag: `load` is also
  // invoked from the Refresh button, where there is no effect teardown to flip a
  // local flag, so a slow first request could otherwise land after a fast retry
  // and overwrite it with stale rows.
  const loadGeneration = React.useRef<number>(0);

  const load = React.useCallback((): void => {
    const generation = ++loadGeneration.current;
    setLoading(true);
    setError('');
    spService
      .getResponses(listId, definition, { includeApproval: approvalEnabled })
      .then((page) => {
        if (generation !== loadGeneration.current) {
          return;
        }
        debugLog('getResponses loaded', { count: page.items.length, hasMore: page.hasMore });
        setItems(page.items);
        setTruncated(page.hasMore);
        setLoading(false);
      })
      .catch((error) => {
        if (generation !== loadGeneration.current) {
          return;
        }
        logError('getResponses', error);
        setError(strings.Responses_View_LoadError);
        setLoading(false);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listId, definitionKey, approvalEnabled]);

  React.useEffect(() => {
    load();
  }, [load]);

  // ----- filtering -----

  const filtered = React.useMemo(() => {
    let result: IResponseItemEx[] = items;

    if (dateRange !== 'all') {
      const now = new Date();
      const start = new Date(now.getTime());
      start.setHours(0, 0, 0, 0);
      if (dateRange !== 'today') {
        start.setDate(start.getDate() - (Number(dateRange) - 1));
      }
      result = result.filter((item) => item.created >= start);
    }

    if (answerFilter) {
      const field = fields.filter((f) => f.id === answerFilter.fieldId)[0];
      if (field) {
        result = result.filter((item) => matchesRowFilter(field, item, answerFilter.row));
      }
    }

    if (approvalEnabled && approvalFilter) {
      result = result.filter((item) => (item.approvalStatus || 'Pending') === approvalFilter);
    }

    const query = debouncedSearch.trim().toLowerCase();
    if (query) {
      const cacheKey = definitionKey + '@' + (props.locale || '');
      if (searchTextCache.current.key !== cacheKey) {
        searchTextCache.current = { key: cacheKey, map: new WeakMap() };
      }
      const cache = searchTextCache.current.map;
      result = result.filter((item) => {
        let haystack = cache.get(item);
        if (haystack === undefined) {
          haystack = [item.createdBy]
            .concat(
              fields.map((f) => formatSharePointValue(f, item.values[f.internalName], messages, props.locale))
            )
            .join(' ')
            .toLowerCase();
          cache.set(item, haystack);
        }
        return haystack.indexOf(query) !== -1;
      });
    }

    return result;
  }, [items, debouncedSearch, dateRange, answerFilter, approvalFilter, approvalEnabled, fields, definitionKey, props.locale]);

  const segmentValueOptions = React.useMemo(() => {
    const field = filterableFields.filter((f) => f.id === segmentFieldId)[0];
    if (!field) {
      return [];
    }
    return segmentBy(field, items, messages, props.locale).map((segment) => ({
      key: segment.filterValue,
      text: formatString(strings.Responses_View_SegmentOption, {
        label: segment.label,
        count: segment.items.length
      }),
      data: { label: segment.label, filterValue: segment.filterValue } as IDrillTarget
    }));
  }, [filterableFields, segmentFieldId, items, props.locale]);

  const filtersActive =
    dateRange !== 'all' ||
    !!answerFilter ||
    search.trim().length > 0 ||
    (approvalEnabled && approvalFilter !== '');

  const clearFilters = (): void => {
    setSearch('');
    setDebouncedSearch('');
    setDateRange('all');
    setSegmentFieldId('');
    setAnswerFilter(undefined);
    setApprovalFilter('');
  };

  // ----- actions -----

  const handleExport = (): void => {
    // buildCsv has no extra-column option, so approval data rides along as synthetic text columns
    const approvalFields: IFormField[] = approvalEnabled
      ? ([
          { id: '__approvalStatus', internalName: '__approvalStatus', title: strings.Responses_Approval_Column },
          { id: '__approvalReviewer', internalName: '__approvalReviewer', title: strings.Responses_Approval_ReviewerColumn },
          { id: '__approvalComment', internalName: '__approvalComment', title: strings.Responses_Approval_CommentColumn }
        ].map((f) => ({ ...f, type: FieldType.Text })) as IFormField[])
      : [];
    const statusText = (status?: ApprovalStatus): string =>
      status === 'Approved'
        ? strings.Responses_Approval_Approved
        : status === 'Rejected'
          ? strings.Responses_Approval_Rejected
          : strings.Responses_Approval_Pending;
    const csv = buildCsv(
      fields.concat(approvalFields),
      filtered.map((item) => ({
        id: item.id,
        created: item.created,
        createdBy: item.createdBy,
        values: approvalEnabled
          ? {
              ...item.values,
              __approvalStatus: statusText(item.approvalStatus),
              __approvalReviewer: item.reviewedBy || '',
              __approvalComment: item.approvalComment || ''
            }
          : item.values
      })),
      { locale: props.locale, messages: messages }
    );
    const suffix = filtersActive ? '-filtered' : '';
    downloadTextFile(
      'form-responses' + suffix + '.csv',
      'text/csv;charset=utf-8',
      csv
    );
  };

  const handleDeleteMany = async (toDelete: IResponseItem[]): Promise<void> => {
    const ids = toDelete.map((item) => item.id);
    const idSet = new Set<number>(ids);
    await spService.deleteResponses(listId, ids);
    setItems((prev) => prev.filter((item) => !idSet.has(item.id)));
    setSelectedId(undefined);
  };

  const handleDeleteOne = async (item: IResponseItem): Promise<void> => {
    await spService.deleteResponse(listId, item.id);
    setItems((prev) => prev.filter((existing) => existing.id !== item.id));
    setSelectedId(undefined);
  };

  /** Clicking a chart segment narrows every view to that answer. */
  const handleDrillDown = (field: IFormField, target?: IDrillTarget): void => {
    if (
      target &&
      (target.filterValue !== undefined || (target.filterValues && target.filterValues.length > 0))
    ) {
      setAnswerFilter({ fieldId: field.id, row: target });
      setSegmentFieldId(filterableFields.filter((f) => f.id === field.id).length > 0 ? field.id : '');
    } else if (target && target.label) {
      // no stable answer key (e.g. a word from the word cloud): fall back to a text search
      setSearch(target.label);
      setDebouncedSearch(target.label);
    }
    setView('table');
  };

  /** Optimistic approval update; rolls back and rethrows on failure. */
  const applyApproval = async (
    target: IResponseItem,
    status: ApprovalStatus,
    comment: string
  ): Promise<void> => {
    const previous = items.filter((i) => i.id === target.id)[0];
    const patch = (values: Partial<IResponseItemEx>): void =>
      setItems((prev) => prev.map((i) => (i.id === target.id ? { ...i, ...values } : i)));
    patch({
      approvalStatus: status,
      approvalComment: comment,
      reviewedBy: status === 'Pending' ? '' : previous ? previous.reviewedBy : undefined
    });
    try {
      await spService.setResponseApproval(listId, definition, target.id, status, comment, {
        notifyRespondent: definition.settings.approvalNotify === true,
        locale: props.locale
      });
    } catch (e) {
      logError('setResponseApproval', e);
      if (previous) {
        patch({
          approvalStatus: previous.approvalStatus,
          approvalComment: previous.approvalComment,
          reviewedBy: previous.reviewedBy
        });
      }
      throw e;
    }
  };

  const handleBulkApproval = async (targets: IResponseItem[], status: ApprovalStatus): Promise<void> => {
    let failed = 0;
    for (let i = 0; i < targets.length; i += APPROVAL_BATCH) {
      const results = await Promise.all(
        targets.slice(i, i + APPROVAL_BATCH).map((t) =>
          applyApproval(t, status, '').then(
            () => true,
            () => false
          )
        )
      );
      failed += results.filter((ok) => !ok).length;
    }
    setNotice(failed > 0 ? formatString(strings.Responses_Approval_BulkFailed, { count: failed }) : '');
  };

  const selectedIndex = filtered.map((item) => item.id).indexOf(selectedId as number);
  const selected = selectedIndex >= 0 ? filtered[selectedIndex] : undefined;

  if (loading) {
    return (
      <div className={styles.centered} dir={dir}>
        <Spinner size={SpinnerSize.large} label={strings.Responses_View_Loading} />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.responses} dir={dir}>
        <MessageBar
          messageBarType={MessageBarType.warning}
          actions={
            <PrimaryButton text={strings.Responses_View_TryAgain} onClick={load} />
          }
        >
          {error}
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={styles.responses} dir={dir}>
      <div className={styles.toolbar}>
        <Pivot
          headersOnly={true}
          selectedKey={view}
          onLinkClick={(item) => item && setView(item.props.itemKey as ViewKey)}
        >
          <PivotItem headerText={strings.Responses_View_Dashboard} itemKey="dashboard" itemIcon="ViewDashboard" />
          <PivotItem headerText={strings.Responses_View_Summary} itemKey="summary" itemIcon="BarChartVertical" />
          <PivotItem
            headerText={formatString(strings.Responses_View_ResponsesTab, { count: filtered.length })}
            itemKey="table"
            itemIcon="Table"
          />
        </Pivot>
        <CommandBar
          className={styles.commands}
          items={[
            {
              key: 'refresh',
              text: strings.Responses_View_Refresh,
              iconProps: { iconName: 'Refresh' },
              onClick: () => {
                load();
              }
            },
            {
              key: 'export',
              text: strings.Responses_View_ExportCsv,
              iconProps: { iconName: 'ExcelDocument' },
              disabled: filtered.length === 0,
              onClick: handleExport
            }
          ]}
          overflowButtonProps={{ ariaLabel: strings.Responses_View_MoreActions }}
        />
      </div>

      {notice && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          className={styles.noticeBar}
          onDismiss={() => setNotice('')}
        >
          {notice}
        </MessageBar>
      )}

      {truncated && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.truncationNote}>
          {formatString(strings.Responses_View_Truncated, {
            count: items.length.toLocaleString(props.locale)
          })}
        </MessageBar>
      )}

      {/* one filter bar for every view, so all three always agree on the count */}
      <div className={styles.filterBar}>
        <SearchBox
          className={styles.filterSearch}
          placeholder={strings.Responses_View_SearchPlaceholder}
          value={search}
          onChange={(_e, v) => setSearch(v || '')}
          onClear={() => {
            setSearch('');
            setDebouncedSearch('');
          }}
        />
        <Dropdown
          className={styles.filterDropdown}
          options={getDateRangeOptions()}
          selectedKey={dateRange}
          ariaLabel={strings.Responses_View_DateRange}
          onChange={(_e, option) => option && setDateRange(String(option.key) as DateRange)}
        />
        {filterableFields.length > 0 && (
          <Dropdown
            className={styles.filterDropdown}
            placeholder={strings.Responses_View_FilterByAnswer}
            options={[{ key: '', text: strings.Responses_View_AnyAnswer }].concat(
              filterableFields.map((f) => ({ key: f.id, text: f.title || strings.Responses_View_Untitled }))
            )}
            selectedKey={segmentFieldId}
            ariaLabel={strings.Responses_View_FilterByQuestion}
            onChange={(_e, option) => {
              setSegmentFieldId(option ? String(option.key) : '');
              setAnswerFilter(undefined);
            }}
          />
        )}
        {segmentFieldId && (
          <Dropdown
            className={styles.filterDropdown}
            placeholder={strings.Responses_View_SelectValue}
            options={segmentValueOptions}
            selectedKey={segmentValue === undefined ? null : segmentValue}
            ariaLabel={strings.Responses_View_FilterValue}
            onChange={(_e, option) =>
              option
                ? setAnswerFilter({ fieldId: segmentFieldId, row: option.data as IDrillTarget })
                : setAnswerFilter(undefined)
            }
          />
        )}
        {answerFilter && segmentValue === undefined && (
          <span className={styles.filterChip}>
            <Icon iconName="Filter" /> {answerFilter.row.label}
          </span>
        )}
        {approvalEnabled && (
          <Dropdown
            className={styles.filterDropdown}
            options={[
              { key: '', text: strings.Responses_Approval_AnyStatus },
              { key: 'Pending', text: strings.Responses_Approval_Pending },
              { key: 'Approved', text: strings.Responses_Approval_Approved },
              { key: 'Rejected', text: strings.Responses_Approval_Rejected }
            ]}
            selectedKey={approvalFilter}
            ariaLabel={strings.Responses_Approval_FilterLabel}
            onChange={(_e, option) => option && setApprovalFilter(String(option.key) as ApprovalFilter)}
          />
        )}
        {filtersActive && (
          <PrimaryButton
            iconProps={{ iconName: 'ClearFilter' }}
            text={strings.Responses_View_Clear}
            onClick={clearFilters}
          />
        )}
        <span className={styles.filterCount}>
          {filtered.length === items.length
            ? formatString(
                items.length === 1
                  ? strings.Responses_View_CountOne
                  : strings.Responses_View_CountOther,
                { count: items.length }
              )
            : formatString(strings.Responses_View_CountFiltered, {
                shown: filtered.length,
                total: items.length
              })}
        </span>
      </div>

      {items.length === 0 ? (
        <div className={styles.empty}>
          <Icon iconName="Inbox" className={styles.emptyIcon} />
          <div className={styles.emptyTitle}>{strings.Responses_View_EmptyTitle}</div>
          <p className={styles.emptyBody}>
            {formatString(strings.Responses_View_EmptyBody, { list: props.listTitle })}
          </p>
          <div className={styles.emptyActions}>
            {props.onCopyShareLink && (
              <PrimaryButton
                iconProps={{ iconName: 'Link' }}
                text={strings.Responses_View_CopyShareLink}
                onClick={props.onCopyShareLink}
              />
            )}
            {props.onPreview && (
              <PrimaryButton
                iconProps={{ iconName: 'View' }}
                text={strings.Responses_View_PreviewForm}
                onClick={props.onPreview}
              />
            )}
          </div>
        </div>
      ) : (
        <>
          {view === 'dashboard' && (
            <DashboardView
              definition={definition}
              items={filtered}
              theme={theme}
              canEdit={props.isOwner}
              locale={props.locale}
              onSettingsChange={props.onDashboardChange}
              onDrillDown={handleDrillDown}
            />
          )}

          {view === 'summary' && (
            <SummaryView
              definition={definition}
              items={filtered}
              theme={theme}
              locale={props.locale}
              onDrillDown={handleDrillDown}
            />
          )}

          {view === 'table' && (
            <ResponseTable
              fields={fields}
              items={filtered}
              visibleColumns={visibleColumns}
              onVisibleColumnsChange={setVisibleColumns}
              onOpen={(item) => setSelectedId(item.id)}
              onDeleteMany={handleDeleteMany}
              canDelete={props.isOwner}
              locale={props.locale}
              approvalEnabled={approvalEnabled && props.isOwner}
              onBulkApproval={handleBulkApproval}
            />
          )}
        </>
      )}

      {selected && (
        <ResponseDetailPanel
          item={selected}
          definition={definition}
          listId={listId}
          spService={spService}
          canDelete={props.isOwner}
          index={selectedIndex}
          total={filtered.length}
          locale={props.locale}
          approvalEnabled={approvalEnabled && props.isOwner}
          onSetApproval={applyApproval}
          onNavigate={(delta) => {
            const next = filtered[selectedIndex + delta];
            if (next) {
              setSelectedId(next.id);
            }
          }}
          onDelete={handleDeleteOne}
          onDismiss={() => setSelectedId(undefined)}
        />
      )}
    </div>
  );
};
