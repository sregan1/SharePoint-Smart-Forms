import * as React from 'react';
import {
  CommandBar,
  ContextualMenuItemType,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  DialogType,
  IColumn,
  IContextualMenuItem,
  Icon,
  Panel,
  PanelType,
  PrimaryButton,
  Selection,
  SelectionMode
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import { FieldType, IFormField, IResponseItem } from '../../models';
import { formatSharePointValue } from '../../utils/formUtils';

export interface IResponseTableProps {
  fields: IFormField[];
  items: IResponseItem[];
  /** internal names of the columns to show, in order */
  visibleColumns: string[];
  onVisibleColumnsChange: (internalNames: string[]) => void;
  onOpen: (item: IResponseItem) => void;
  onDeleteMany: (items: IResponseItem[]) => Promise<void>;
  canDelete: boolean;
}

type SortDirection = 'asc' | 'desc';

/** Columns that always appear, before any question columns. */
const SYSTEM_COLUMN_KEYS = ['created', 'createdBy'];

/**
 * The full response grid.
 *
 * Replaces a fixed four-column view that silently dropped every other question:
 * columns are pickable and reorderable, every column sorts, rows select for bulk
 * delete and export, and `DetailsList` virtualizes so a few thousand rows stay
 * responsive.
 */
export const ResponseTable: React.FunctionComponent<IResponseTableProps> = (props) => {
  const { fields, items } = props;
  const [sortKey, setSortKey] = React.useState<string>('created');
  const [sortDirection, setSortDirection] = React.useState<SortDirection>('desc');
  const [selectedIds, setSelectedIds] = React.useState<number[]>([]);
  const [confirmingDelete, setConfirmingDelete] = React.useState<boolean>(false);
  const [deleting, setDeleting] = React.useState<boolean>(false);
  const [columnPickerOpen, setColumnPickerOpen] = React.useState<boolean>(false);

  // Selection is imperative in Fluent's DetailsList; mirror it into state so the
  // bulk action bar can react to it.
  //
  // The callback reads the instance through a ref rather than closing over the
  // binding it is being passed to: Selection can raise onSelectionChanged during
  // construction and while setItems runs, at which point the `const` is not yet
  // initialised and reading it throws.
  const selectionRef = React.useRef<Selection | undefined>(undefined);
  const selection = React.useMemo(() => {
    const instance = new Selection({
      onSelectionChanged: () => {
        const current = selectionRef.current;
        if (!current) {
          return;
        }
        const chosen = current.getSelection() as IResponseItem[];
        setSelectedIds(chosen.map((item) => item.id));
      }
    });
    selectionRef.current = instance;
    return instance;
  }, []);

  const visible = props.visibleColumns.length > 0
    ? props.visibleColumns
    : fields.slice(0, 6).map((f) => f.internalName);

  const shownFields = visible
    .map((internalName) => fields.filter((f) => f.internalName === internalName)[0])
    .filter((field) => !!field);

  const sortValue = (item: IResponseItem, key: string): string | number => {
    if (key === 'created') {
      return item.created.getTime();
    }
    if (key === 'createdBy') {
      return (item.createdBy || '').toLowerCase();
    }
    const field = fields.filter((f) => f.internalName === key)[0];
    if (!field) {
      return '';
    }
    const raw = item.values[key];
    if (raw === undefined || raw === null || raw === '') {
      // unanswered sorts last in both directions rather than jumping to the top
      return sortDirection === 'asc' ? Infinity : -Infinity;
    }
    const numeric =
      field.type === FieldType.Number ||
      field.type === FieldType.Calculated ||
      field.type === FieldType.Rating ||
      field.type === FieldType.Scale ||
      field.type === FieldType.Slider;
    if (numeric) {
      const num = Number(raw);
      return isNaN(num) ? 0 : num;
    }
    if (field.type === FieldType.Date || field.type === FieldType.Time) {
      const date = new Date(String(raw));
      return isNaN(date.getTime()) ? 0 : date.getTime();
    }
    return formatSharePointValue(field, raw).toLowerCase();
  };

  const sorted = React.useMemo(() => {
    const copy = items.slice();
    copy.sort((a, b) => {
      const left = sortValue(a, sortKey);
      const right = sortValue(b, sortKey);
      if (left === right) {
        return 0;
      }
      const ascending = left < right ? -1 : 1;
      return sortDirection === 'asc' ? ascending : -ascending;
    });
    return copy;
  }, [items, sortKey, sortDirection]);

  const toggleSort = (key: string): void => {
    if (sortKey === key) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
      return;
    }
    setSortKey(key);
    // dates read newest-first by default; text reads A→Z
    setSortDirection(key === 'created' ? 'desc' : 'asc');
  };

  const systemColumns: IColumn[] = [
    {
      key: 'created',
      name: 'Submitted',
      minWidth: 120,
      maxWidth: 165,
      isResizable: true,
      isSorted: sortKey === 'created',
      isSortedDescending: sortKey === 'created' && sortDirection === 'desc',
      onColumnClick: () => toggleSort('created'),
      onRender: (item: IResponseItem) => (
        <button
          type="button"
          className={styles.cellLink}
          onClick={() => props.onOpen(item)}
          title="Open this response"
        >
          {item.created.toLocaleString()}
        </button>
      )
    },
    {
      key: 'createdBy',
      name: 'Submitted by',
      minWidth: 110,
      maxWidth: 170,
      isResizable: true,
      isSorted: sortKey === 'createdBy',
      isSortedDescending: sortKey === 'createdBy' && sortDirection === 'desc',
      onColumnClick: () => toggleSort('createdBy'),
      onRender: (item: IResponseItem) => (
        <span className={item.createdBy ? styles.cellText : styles.cellMuted}>
          {item.createdBy || 'Unknown'}
        </span>
      )
    }
  ];

  const questionColumns: IColumn[] = shownFields.map((field) => ({
    key: field.internalName,
    name: field.title || 'Untitled',
    minWidth: 110,
    isResizable: true,
    isMultiline: field.type === FieldType.MultilineText || field.type === FieldType.RichText,
    isSorted: sortKey === field.internalName,
    isSortedDescending: sortKey === field.internalName && sortDirection === 'desc',
    onColumnClick: () => toggleSort(field.internalName),
    onRender: (item: IResponseItem) => {
      const text = formatSharePointValue(field, item.values[field.internalName]);
      return (
        <span className={text ? styles.cellText : styles.cellMuted} title={text}>
          {text || '—'}
        </span>
      );
    }
  }));

  const columns: IColumn[] = systemColumns.concat(questionColumns);

  const selectedItems = items.filter((item) => selectedIds.indexOf(item.id) !== -1);

  const handleDelete = async (): Promise<void> => {
    setDeleting(true);
    try {
      await props.onDeleteMany(selectedItems);
      selection.setAllSelected(false);
      setSelectedIds([]);
    } finally {
      setDeleting(false);
      setConfirmingDelete(false);
    }
  };

  const columnMenuItems: IContextualMenuItem[] = [
    {
      key: 'header',
      itemType: ContextualMenuItemType.Header,
      text: 'Show these questions'
    }
  ];
  fields.forEach((field) => {
    columnMenuItems.push({
      key: field.internalName,
      text: field.title || 'Untitled',
      canCheck: true,
      checked: visible.indexOf(field.internalName) !== -1,
      onClick: () => {
        const isVisible = visible.indexOf(field.internalName) !== -1;
        props.onVisibleColumnsChange(
          isVisible
            ? visible.filter((name) => name !== field.internalName)
            : visible.concat([field.internalName])
        );
        // keep the menu open so several columns can be toggled in one pass
        return true;
      }
    });
  });

  return (
    <div>
      <div className={styles.tableToolbar}>
        <CommandBar
          className={styles.commands}
          items={[
            {
              key: 'columns',
              text: 'Columns (' + shownFields.length + '/' + fields.length + ')',
              iconProps: { iconName: 'ColumnOptions' },
              subMenuProps: { items: columnMenuItems }
            },
            {
              key: 'all',
              text: 'Show all',
              iconProps: { iconName: 'CheckboxComposite' },
              onClick: () => props.onVisibleColumnsChange(fields.map((f) => f.internalName))
            },
            {
              key: 'reset',
              text: 'Reset columns',
              iconProps: { iconName: 'Undo' },
              onClick: () => props.onVisibleColumnsChange([])
            },
            {
              key: 'reorder',
              text: 'Reorder',
              iconProps: { iconName: 'Sort' },
              onClick: () => setColumnPickerOpen(true)
            }
          ]}
        />
      </div>

      {selectedIds.length > 0 && (
        <div className={styles.selectionBar} role="status">
          <Icon iconName="MultiSelect" />
          {selectedIds.length} {selectedIds.length === 1 ? 'response' : 'responses'} selected
          <span className={styles.selectionSpacer} />
          {props.canDelete && (
            <DefaultButton
              iconProps={{ iconName: 'Delete' }}
              text="Delete selected"
              onClick={() => setConfirmingDelete(true)}
            />
          )}
          <DefaultButton
            text="Clear selection"
            onClick={() => {
              selection.setAllSelected(false);
              setSelectedIds([]);
            }}
          />
        </div>
      )}

      {items.length === 0 ? (
        <div className={styles.empty}>
          <Icon iconName="Table" className={styles.emptyIcon} />
          <div className={styles.emptyTitle}>Nothing to show</div>
          <p className={styles.emptyBody}>No responses match the current filters.</p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <DetailsList
            items={sorted}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selection={selection}
            selectionMode={props.canDelete ? SelectionMode.multiple : SelectionMode.none}
            selectionPreservedOnEmptyClick={true}
            // Enter or double-click opens; the previous grid opened the panel on
            // mere focus change, which fired while arrowing through rows
            onItemInvoked={(item: IResponseItem) => props.onOpen(item)}
            getKey={(item: IResponseItem) => String(item.id)}
            ariaLabelForSelectionColumn="Toggle selection"
            checkButtonAriaLabel="Select this response"
            compact={false}
          />
        </div>
      )}

      <Dialog
        hidden={!confirmingDelete}
        onDismiss={() => setConfirmingDelete(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title:
            'Delete ' + selectedIds.length + (selectedIds.length === 1 ? ' response?' : ' responses?'),
          subText: 'The list items move to the site recycle bin, so they can be restored from there.'
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text={deleting ? 'Deleting…' : 'Delete'}
            disabled={deleting}
            onClick={() => {
              void handleDelete();
            }}
          />
          <DefaultButton text="Cancel" disabled={deleting} onClick={() => setConfirmingDelete(false)} />
        </DialogFooter>
      </Dialog>

      {columnPickerOpen && (
        <ColumnOrderPanel
          fields={fields}
          visible={visible}
          onChange={props.onVisibleColumnsChange}
          onDismiss={() => setColumnPickerOpen(false)}
        />
      )}
    </div>
  );
};

interface IColumnOrderPanelProps {
  fields: IFormField[];
  visible: string[];
  onChange: (internalNames: string[]) => void;
  onDismiss: () => void;
}

/** Reorder the visible columns. Kept separate so the grid stays readable. */
const ColumnOrderPanel: React.FunctionComponent<IColumnOrderPanelProps> = (props) => {
  const [order, setOrder] = React.useState<string[]>(props.visible.slice());

  const move = (index: number, delta: number): void => {
    const target = index + delta;
    if (target < 0 || target >= order.length) {
      return;
    }
    const next = order.slice();
    const moved = next.splice(index, 1)[0];
    next.splice(target, 0, moved);
    setOrder(next);
  };

  const titleFor = (internalName: string): string => {
    const field = props.fields.filter((f) => f.internalName === internalName)[0];
    return field ? field.title || 'Untitled' : internalName;
  };

  return (
    <Panel
      isOpen={true}
      type={PanelType.smallFixedFar}
      headerText="Column order"
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton
            text="Apply"
            onClick={() => {
              props.onChange(order);
              props.onDismiss();
            }}
          />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <p className={styles.segmentNote}>
        Submitted and Submitted by always come first. Drag order applies to the question columns.
      </p>
      {order.map((internalName, index) => (
        <div key={internalName} className={styles.attachmentRow}>
          <span className={styles.detailValue} style={{ flex: '1 1 auto', minWidth: 0 }}>
            {index + 1}. {titleFor(internalName)}
          </span>
          <DefaultButton
            iconProps={{ iconName: 'Up' }}
            title="Move up"
            ariaLabel={'Move ' + titleFor(internalName) + ' up'}
            disabled={index === 0}
            onClick={() => move(index, -1)}
          />
          <DefaultButton
            iconProps={{ iconName: 'Down' }}
            title="Move down"
            ariaLabel={'Move ' + titleFor(internalName) + ' down'}
            disabled={index === order.length - 1}
            onClick={() => move(index, 1)}
          />
        </div>
      ))}
      {order.length === 0 && (
        <p className={styles.emptyCardText}>No question columns are visible yet.</p>
      )}
    </Panel>
  );
};

/** Exported so the shell can keep the same default column set. */
export const defaultVisibleColumns = (fields: IFormField[]): string[] =>
  fields.slice(0, 6).map((f) => f.internalName);

export { SYSTEM_COLUMN_KEYS };
