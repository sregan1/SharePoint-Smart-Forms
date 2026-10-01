import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';
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
import { ApprovalStatus, IResponseItemEx } from '../../services/SharePointService';

/** Chip for an approval status; anything unset reads as pending. */
export const ApprovalChip: React.FunctionComponent<{ status?: ApprovalStatus }> = ({ status }) => {
  const value: ApprovalStatus = status || 'Pending';
  const cls =
    value === 'Approved' ? styles.approvalApproved : value === 'Rejected' ? styles.approvalRejected : styles.approvalPending;
  const text =
    value === 'Approved'
      ? strings.Responses_Approval_Approved
      : value === 'Rejected'
        ? strings.Responses_Approval_Rejected
        : strings.Responses_Approval_Pending;
  const icon = value === 'Approved' ? 'CheckMark' : value === 'Rejected' ? 'Cancel' : 'Clock';
  return (
    <span className={styles.approvalChip + ' ' + cls}>
      <Icon iconName={icon} /> {text}
    </span>
  );
};

export interface IResponseTableProps {
  fields: IFormField[];
  items: IResponseItem[];
  /** internal names of the columns to show, in order */
  visibleColumns: string[];
  onVisibleColumnsChange: (internalNames: string[]) => void;
  onOpen: (item: IResponseItem) => void;
  onDeleteMany: (items: IResponseItem[]) => Promise<void>;
  canDelete: boolean;
  locale?: string;
  /** show the approval column and bulk approve/reject */
  approvalEnabled?: boolean;
  onBulkApproval?: (items: IResponseItem[], status: ApprovalStatus) => Promise<void>;
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
  const [reviewing, setReviewing] = React.useState<boolean>(false);
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
    if (key === 'approval') {
      return ((item as IResponseItemEx).approvalStatus || 'Pending').toLowerCase();
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
      name: strings.Responses_Table_Submitted,
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
          title={strings.Responses_Table_OpenResponse}
        >
          {item.created.toLocaleString(props.locale)}
        </button>
      )
    },
    {
      key: 'createdBy',
      name: strings.Responses_Table_SubmittedBy,
      minWidth: 110,
      maxWidth: 170,
      isResizable: true,
      isSorted: sortKey === 'createdBy',
      isSortedDescending: sortKey === 'createdBy' && sortDirection === 'desc',
      onColumnClick: () => toggleSort('createdBy'),
      onRender: (item: IResponseItem) => (
        <span className={item.createdBy ? styles.cellText : styles.cellMuted}>
          {item.createdBy || strings.Responses_Table_Unknown}
        </span>
      )
    }
  ];

  if (props.approvalEnabled) {
    systemColumns.push({
      key: 'approval',
      name: strings.Responses_Approval_Column,
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      isSorted: sortKey === 'approval',
      isSortedDescending: sortKey === 'approval' && sortDirection === 'desc',
      onColumnClick: () => toggleSort('approval'),
      onRender: (item: IResponseItem) => <ApprovalChip status={(item as IResponseItemEx).approvalStatus} />
    });
  }

  const questionColumns: IColumn[] = shownFields.map((field) => ({
    key: field.internalName,
    name: field.title || strings.Responses_Table_Untitled,
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

  const selectedSet = React.useMemo(() => new Set<number>(selectedIds), [selectedIds]);
  const selectedItems = items.filter((item) => selectedSet.has(item.id));

  const handleBulkApproval = async (status: ApprovalStatus): Promise<void> => {
    if (!props.onBulkApproval) {
      return;
    }
    setReviewing(true);
    try {
      await props.onBulkApproval(selectedItems, status);
    } finally {
      setReviewing(false);
    }
  };

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
      text: strings.Responses_Table_ShowTheseQuestions
    }
  ];
  fields.forEach((field) => {
    columnMenuItems.push({
      key: field.internalName,
      text: field.title || strings.Responses_Table_Untitled,
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
              text: formatString(strings.Responses_Table_ColumnsCount, {
                shown: shownFields.length,
                total: fields.length
              }),
              iconProps: { iconName: 'ColumnOptions' },
              subMenuProps: { items: columnMenuItems }
            },
            {
              key: 'all',
              text: strings.Responses_Table_ShowAll,
              iconProps: { iconName: 'CheckboxComposite' },
              onClick: () => props.onVisibleColumnsChange(fields.map((f) => f.internalName))
            },
            {
              key: 'reset',
              text: strings.Responses_Table_ResetColumns,
              iconProps: { iconName: 'Undo' },
              onClick: () => props.onVisibleColumnsChange([])
            },
            {
              key: 'reorder',
              text: strings.Responses_Table_Reorder,
              iconProps: { iconName: 'Sort' },
              onClick: () => setColumnPickerOpen(true)
            }
          ]}
        />
      </div>

      {selectedIds.length > 0 && (
        <div className={styles.selectionBar} role="status">
          <Icon iconName="MultiSelect" />
          {formatString(
            selectedIds.length === 1
              ? strings.Responses_Table_SelectedOne
              : strings.Responses_Table_SelectedOther,
            { count: selectedIds.length }
          )}
          <span className={styles.selectionSpacer} />
          {props.approvalEnabled && props.onBulkApproval && (
            <>
              <DefaultButton
                iconProps={{ iconName: 'CheckMark' }}
                text={strings.Responses_Approval_ApproveSelected}
                disabled={reviewing}
                onClick={() => {
                  void handleBulkApproval('Approved');
                }}
              />
              <DefaultButton
                iconProps={{ iconName: 'Cancel' }}
                text={strings.Responses_Approval_RejectSelected}
                disabled={reviewing}
                onClick={() => {
                  void handleBulkApproval('Rejected');
                }}
              />
            </>
          )}
          {props.canDelete && (
            <DefaultButton
              iconProps={{ iconName: 'Delete' }}
              text={strings.Responses_Table_DeleteSelected}
              onClick={() => setConfirmingDelete(true)}
            />
          )}
          <DefaultButton
            text={strings.Responses_Table_ClearSelection}
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
          <div className={styles.emptyTitle}>{strings.Responses_Table_NothingToShow}</div>
          <p className={styles.emptyBody}>{strings.Responses_Table_NoMatch}</p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <DetailsList
            items={sorted}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selection={selection}
            selectionMode={props.canDelete || props.approvalEnabled ? SelectionMode.multiple : SelectionMode.none}
            selectionPreservedOnEmptyClick={true}
            // Enter or double-click opens; the previous grid opened the panel on
            // mere focus change, which fired while arrowing through rows
            onItemInvoked={(item: IResponseItem) => props.onOpen(item)}
            getKey={(item: IResponseItem) => String(item.id)}
            ariaLabelForSelectionColumn={strings.Responses_Table_ToggleSelection}
            checkButtonAriaLabel={strings.Responses_Table_SelectResponse}
            compact={false}
          />
        </div>
      )}

      <Dialog
        hidden={!confirmingDelete}
        onDismiss={() => setConfirmingDelete(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: formatString(
            selectedIds.length === 1
              ? strings.Responses_Table_DeleteTitleOne
              : strings.Responses_Table_DeleteTitleOther,
            { count: selectedIds.length }
          ),
          subText: strings.Responses_Table_DeleteBody
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text={deleting ? strings.Responses_Table_Deleting : strings.Responses_Table_Delete}
            disabled={deleting}
            onClick={() => {
              void handleDelete();
            }}
          />
          <DefaultButton
            text={strings.Responses_Table_Cancel}
            disabled={deleting}
            onClick={() => setConfirmingDelete(false)}
          />
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
    return field ? field.title || strings.Responses_Table_Untitled : internalName;
  };

  return (
    <Panel
      isOpen={true}
      type={PanelType.smallFixedFar}
      headerText={strings.Responses_Table_ColumnOrder}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton
            text={strings.Responses_Table_Apply}
            onClick={() => {
              props.onChange(order);
              props.onDismiss();
            }}
          />
          <DefaultButton text={strings.Responses_Table_Cancel} onClick={props.onDismiss} />
        </div>
      )}
    >
      <p className={styles.segmentNote}>
        {strings.Responses_Table_OrderNote}
      </p>
      {order.map((internalName, index) => (
        <div key={internalName} className={styles.attachmentRow}>
          <span className={styles.detailValue} style={{ flex: '1 1 auto', minWidth: 0 }}>
            {index + 1}. {titleFor(internalName)}
          </span>
          <DefaultButton
            iconProps={{ iconName: 'Up' }}
            title={strings.Responses_Table_MoveUp}
            ariaLabel={formatString(strings.Responses_Table_MoveUpFor, {
              title: titleFor(internalName)
            })}
            disabled={index === 0}
            onClick={() => move(index, -1)}
          />
          <DefaultButton
            iconProps={{ iconName: 'Down' }}
            title={strings.Responses_Table_MoveDown}
            ariaLabel={formatString(strings.Responses_Table_MoveDownFor, {
              title: titleFor(internalName)
            })}
            disabled={index === order.length - 1}
            onClick={() => move(index, 1)}
          />
        </div>
      ))}
      {order.length === 0 && (
        <p className={styles.emptyCardText}>{strings.Responses_Table_NoColumns}</p>
      )}
    </Panel>
  );
};

/** Exported so the shell can keep the same default column set. */
export const defaultVisibleColumns = (fields: IFormField[]): string[] =>
  fields.slice(0, 6).map((f) => f.internalName);

export { SYSTEM_COLUMN_KEYS };
