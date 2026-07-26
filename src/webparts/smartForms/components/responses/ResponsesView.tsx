import * as React from 'react';
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
  IDashboardSettings,
  IFormDefinition,
  IFormField,
  IResponseItem
} from '../../models';
import {
  buildCsv,
  downloadTextFile,
  formatSharePointValue,
  inputFields,
  segmentableFields
} from '../../utils/formUtils';
import { segmentBy } from '../../utils/analytics';
import { IThemeInfo } from '../../utils/theme';
import { SharePointService } from '../../services/SharePointService';
import { debugLog, logError } from '../../utils/debug';
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
}

type ViewKey = 'dashboard' | 'summary' | 'table';

type DateRange = 'all' | '7' | '30' | '90' | 'today';

const DATE_RANGE_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'All time' },
  { key: 'today', text: 'Today' },
  { key: '7', text: 'Last 7 days' },
  { key: '30', text: 'Last 30 days' },
  { key: '90', text: 'Last 90 days' }
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

  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [items, setItems] = React.useState<IResponseItem[]>([]);
  const [truncated, setTruncated] = React.useState<boolean>(false);
  const [view, setView] = React.useState<ViewKey>('dashboard');

  // ----- shared filters -----
  const [search, setSearch] = React.useState<string>('');
  const [dateRange, setDateRange] = React.useState<DateRange>('all');
  const [segmentFieldId, setSegmentFieldId] = React.useState<string>('');
  const [segmentValue, setSegmentValue] = React.useState<string>('');

  const [selectedId, setSelectedId] = React.useState<number | undefined>(undefined);
  const [visibleColumns, setVisibleColumns] = React.useState<string[]>([]);

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
      .getResponses(listId, definition)
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
        setError(
          'Responses could not be loaded. Publish the form once so the list columns exist, then try again.'
        );
        setLoading(false);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listId, definitionKey]);

  React.useEffect(() => {
    load();
  }, [load]);

  // ----- filtering -----

  const filtered = React.useMemo(() => {
    let result = items;

    if (dateRange !== 'all') {
      const now = new Date();
      const start = new Date(now.getTime());
      start.setHours(0, 0, 0, 0);
      if (dateRange !== 'today') {
        start.setDate(start.getDate() - (Number(dateRange) - 1));
      }
      result = result.filter((item) => item.created >= start);
    }

    if (segmentFieldId && segmentValue) {
      const field = filterableFields.filter((f) => f.id === segmentFieldId)[0];
      if (field) {
        const matching = segmentBy(field, result).filter((s) => s.label === segmentValue)[0];
        result = matching ? matching.items : [];
      }
    }

    const query = search.trim().toLowerCase();
    if (query) {
      result = result.filter((item) => {
        const haystack = [item.createdBy]
          .concat(fields.map((f) => formatSharePointValue(f, item.values[f.internalName])))
          .join(' ')
          .toLowerCase();
        return haystack.indexOf(query) !== -1;
      });
    }

    return result;
  }, [items, search, dateRange, segmentFieldId, segmentValue, fields, filterableFields]);

  const segmentValueOptions = React.useMemo(() => {
    const field = filterableFields.filter((f) => f.id === segmentFieldId)[0];
    if (!field) {
      return [];
    }
    return segmentBy(field, items).map((segment) => ({
      key: segment.label,
      text: segment.label + ' (' + segment.items.length + ')'
    }));
  }, [filterableFields, segmentFieldId, items]);

  const filtersActive = dateRange !== 'all' || !!segmentValue || search.trim().length > 0;

  const clearFilters = (): void => {
    setSearch('');
    setDateRange('all');
    setSegmentFieldId('');
    setSegmentValue('');
  };

  // ----- actions -----

  const handleExport = (): void => {
    const csv = buildCsv(
      fields,
      filtered.map((item) => ({
        id: item.id,
        created: item.created,
        createdBy: item.createdBy,
        values: item.values
      }))
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
    await spService.deleteResponses(listId, ids);
    setItems((prev) => prev.filter((item) => ids.indexOf(item.id) === -1));
    setSelectedId(undefined);
  };

  const handleDeleteOne = async (item: IResponseItem): Promise<void> => {
    await spService.deleteResponse(listId, item.id);
    setItems((prev) => prev.filter((existing) => existing.id !== item.id));
    setSelectedId(undefined);
  };

  /** Clicking a chart segment narrows the table to that answer. */
  const handleDrillDown = (field: IFormField, categoryLabel: string): void => {
    const segmentable = filterableFields.filter((f) => f.id === field.id)[0];
    if (segmentable && categoryLabel) {
      setSegmentFieldId(field.id);
      setSegmentValue(categoryLabel);
    } else if (categoryLabel) {
      // not a segmentable question — fall back to a text search
      setSearch(categoryLabel);
    }
    setView('table');
  };

  const selectedIndex = filtered.map((item) => item.id).indexOf(selectedId as number);
  const selected = selectedIndex >= 0 ? filtered[selectedIndex] : undefined;

  if (loading) {
    return (
      <div className={styles.centered}>
        <Spinner size={SpinnerSize.large} label="Loading responses…" />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.responses}>
        <MessageBar
          messageBarType={MessageBarType.warning}
          actions={
            <PrimaryButton text="Try again" onClick={load} />
          }
        >
          {error}
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={styles.responses}>
      <div className={styles.toolbar}>
        <Pivot
          headersOnly={true}
          selectedKey={view}
          onLinkClick={(item) => item && setView(item.props.itemKey as ViewKey)}
        >
          <PivotItem headerText="Dashboard" itemKey="dashboard" itemIcon="ViewDashboard" />
          <PivotItem headerText="Summary" itemKey="summary" itemIcon="BarChartVertical" />
          <PivotItem
            headerText={'Responses (' + filtered.length + ')'}
            itemKey="table"
            itemIcon="Table"
          />
        </Pivot>
        <CommandBar
          className={styles.commands}
          items={[
            {
              key: 'refresh',
              text: 'Refresh',
              iconProps: { iconName: 'Refresh' },
              onClick: () => {
                load();
              }
            },
            {
              key: 'export',
              text: 'Export CSV',
              iconProps: { iconName: 'ExcelDocument' },
              disabled: filtered.length === 0,
              onClick: handleExport
            }
          ]}
          overflowButtonProps={{ ariaLabel: 'More response actions' }}
        />
      </div>

      {truncated && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.truncationNote}>
          This list has more responses than Smart Forms loads at once. The newest{' '}
          {items.length.toLocaleString()} are shown — open the list in Microsoft Lists for the
          complete set.
        </MessageBar>
      )}

      {/* one filter bar for every view, so all three always agree on the count */}
      <div className={styles.filterBar}>
        <SearchBox
          className={styles.filterSearch}
          placeholder="Search all answers"
          value={search}
          onChange={(_e, v) => setSearch(v || '')}
          onClear={() => setSearch('')}
        />
        <Dropdown
          className={styles.filterDropdown}
          options={DATE_RANGE_OPTIONS}
          selectedKey={dateRange}
          ariaLabel="Date range"
          onChange={(_e, option) => option && setDateRange(String(option.key) as DateRange)}
        />
        {filterableFields.length > 0 && (
          <Dropdown
            className={styles.filterDropdown}
            placeholder="Filter by answer"
            options={[{ key: '', text: 'Any answer' }].concat(
              filterableFields.map((f) => ({ key: f.id, text: f.title || 'Untitled' }))
            )}
            selectedKey={segmentFieldId}
            ariaLabel="Filter by question"
            onChange={(_e, option) => {
              setSegmentFieldId(option ? String(option.key) : '');
              setSegmentValue('');
            }}
          />
        )}
        {segmentFieldId && (
          <Dropdown
            className={styles.filterDropdown}
            placeholder="Select a value"
            options={segmentValueOptions}
            selectedKey={segmentValue || null}
            ariaLabel="Filter value"
            onChange={(_e, option) => setSegmentValue(option ? String(option.key) : '')}
          />
        )}
        {filtersActive && (
          <PrimaryButton
            iconProps={{ iconName: 'ClearFilter' }}
            text="Clear"
            onClick={clearFilters}
          />
        )}
        <span className={styles.filterCount}>
          {filtered.length === items.length
            ? items.length + (items.length === 1 ? ' response' : ' responses')
            : filtered.length + ' of ' + items.length + ' responses'}
        </span>
      </div>

      {items.length === 0 ? (
        <div className={styles.empty}>
          <Icon iconName="Inbox" className={styles.emptyIcon} />
          <div className={styles.emptyTitle}>No responses yet</div>
          <p className={styles.emptyBody}>
            Share the form to start collecting. Answers are saved to &ldquo;{props.listTitle}&rdquo;
            and appear here straight away.
          </p>
          <div className={styles.emptyActions}>
            {props.onCopyShareLink && (
              <PrimaryButton
                iconProps={{ iconName: 'Link' }}
                text="Copy share link"
                onClick={props.onCopyShareLink}
              />
            )}
            {props.onPreview && (
              <PrimaryButton
                iconProps={{ iconName: 'View' }}
                text="Preview the form"
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
              onSettingsChange={props.onDashboardChange}
              onDrillDown={handleDrillDown}
            />
          )}

          {view === 'summary' && (
            <SummaryView
              definition={definition}
              items={filtered}
              theme={theme}
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
