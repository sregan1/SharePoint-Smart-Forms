import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  IconButton,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import { FieldType, IFormDefinition, IFormField, IFormValues, IResponseItem } from '../../models';
import { formatSharePointValue, inputFields, normalizeFromSharePoint } from '../../utils/formUtils';
import { formatDuration } from '../../utils/analytics';
import { SharePointService } from '../../services/SharePointService';
import { ResponseFormView } from '../form/FormRenderer';
import { RichTextView } from '../form/RichTextField';
import { ApprovalStatus, IResponseItemEx } from '../../services/SharePointService';
import { ApprovalChip } from './ResponseTable';

export interface IResponseDetailPanelProps {
  item: IResponseItem;
  definition: IFormDefinition;
  listId: string;
  spService: SharePointService;
  canDelete: boolean;
  /** position within the current filtered set, for prev/next */
  index: number;
  total: number;
  onNavigate: (delta: number) => void;
  onDelete: (item: IResponseItem) => Promise<void>;
  onDismiss: () => void;
  locale?: string;
  /** show the approval controls (form has approval enabled and the viewer is the list owner) */
  approvalEnabled?: boolean;
  /** persist a decision; rejects on failure so the panel can show the error */
  onSetApproval?: (item: IResponseItem, status: ApprovalStatus, comment: string) => Promise<void>;
}

interface IAttachment {
  fileName: string;
  url: string;
  internalName: string;
}

/**
 * One response, in detail.
 *
 * Two readings of the same answers: a labelled list for scanning, and the form
 * itself rendered read-only, which is far more legible for a long form because
 * the layout matches what the respondent actually saw. Prev/next moves through
 * the filtered set without closing, so reviewing 30 responses is 30 keystrokes
 * rather than 60 clicks.
 */
export const ResponseDetailPanel: React.FunctionComponent<IResponseDetailPanelProps> = (props) => {
  const { item, definition } = props;
  const [confirmingDelete, setConfirmingDelete] = React.useState<boolean>(false);
  const [deleting, setDeleting] = React.useState<boolean>(false);
  const [comment, setComment] = React.useState<string>('');
  const [approvalBusy, setApprovalBusy] = React.useState<boolean>(false);
  const [approvalError, setApprovalError] = React.useState<string>('');
  const [attachments, setAttachments] = React.useState<IAttachment[] | undefined>(undefined);

  const ex = item as IResponseItemEx;
  React.useEffect(() => {
    setComment(ex.approvalComment || '');
    setApprovalError('');
  }, [item.id, ex.approvalComment]);

  const decide = async (status: ApprovalStatus): Promise<void> => {
    if (!props.onSetApproval) {
      return;
    }
    setApprovalBusy(true);
    setApprovalError('');
    try {
      await props.onSetApproval(item, status, comment.trim());
    } catch {
      setApprovalError(strings.Responses_Approval_Error);
    } finally {
      setApprovalBusy(false);
    }
  };

  const fields = React.useMemo(
    () => inputFields(definition).filter((f) => f.provisioned !== false),
    [definition]
  );

  const hasAttachmentFields = fields.some(
    (f) => f.type === FieldType.FileUpload || f.type === FieldType.Signature
  );

  React.useEffect(() => {
    let cancelled = false;
    if (!hasAttachmentFields || !item.attachmentCount) {
      setAttachments([]);
      return undefined;
    }
    setAttachments(undefined);
    props.spService
      .getResponseAttachments(props.listId, item.id)
      .then((files) => {
        if (!cancelled) {
          setAttachments(files);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setAttachments([]);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [item.id, hasAttachmentFields, item.attachmentCount]);

  /** Normalized answers, so the read-only form view can reuse the real controls. */
  const formValues: IFormValues = React.useMemo(() => {
    const values: IFormValues = {};
    fields.forEach((field) => {
      const normalized = normalizeFromSharePoint(field, item.values[field.internalName]);
      if (normalized !== undefined) {
        values[field.id] = normalized;
      }
    });
    return values;
  }, [fields, item]);

  const handleDelete = async (): Promise<void> => {
    setDeleting(true);
    try {
      await props.onDelete(item);
    } finally {
      setDeleting(false);
      setConfirmingDelete(false);
    }
  };

  const attachmentsFor = (field: IFormField): IAttachment[] =>
    (attachments || []).filter((file) => file.internalName === field.internalName);

  const renderAnswer = (field: IFormField): React.ReactNode => {
    const raw = item.values[field.internalName];

    if (field.type === FieldType.RichText && raw) {
      return <RichTextView html={String(raw)} />;
    }

    if (field.type === FieldType.Signature) {
      const files = attachmentsFor(field);
      if (files.length > 0) {
        return <img className={styles.signaturePreview} src={files[0].url} alt={strings.Responses_Detail_Signature} />;
      }
      return <span className={raw ? styles.detailValue : styles.detailValueEmpty}>{raw ? strings.Responses_Detail_Signed : strings.Responses_Detail_NoAnswer}</span>;
    }

    if (field.type === FieldType.FileUpload) {
      const files = attachmentsFor(field);
      if (files.length === 0) {
        const names = formatSharePointValue(field, raw);
        return (
          <span className={names ? styles.detailValue : styles.detailValueEmpty}>
            {names || strings.Responses_Detail_NoAnswer}
          </span>
        );
      }
      return (
        <div className={styles.attachmentList}>
          {files.map((file) => (
            <div key={file.url} className={styles.attachmentRow}>
              <Icon iconName="Attach" />
              <a
                className={styles.attachmentLink}
                href={file.url}
                target="_blank"
                rel="noopener noreferrer"
                title={file.fileName}
              >
                {file.fileName}
              </a>
            </div>
          ))}
        </div>
      );
    }

    const text = formatSharePointValue(field, raw);
    return (
      <span className={text ? styles.detailValue : styles.detailValueEmpty}>
        {text || strings.Responses_Detail_NoAnswer}
      </span>
    );
  };

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText={formatString(strings.Responses_Detail_Header, {
        index: props.index + 1,
        total: props.total
      })}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            title={strings.Responses_Detail_Previous}
            ariaLabel={strings.Responses_Detail_Previous}
            disabled={props.index <= 0}
            onClick={() => props.onNavigate(-1)}
          />
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            title={strings.Responses_Detail_Next}
            ariaLabel={strings.Responses_Detail_Next}
            disabled={props.index >= props.total - 1}
            onClick={() => props.onNavigate(1)}
          />
          <DefaultButton
            iconProps={{ iconName: 'Print' }}
            text={strings.Responses_Detail_Print}
            onClick={() => window.print()}
          />
          <span className={styles.panelFooterSpacer} />
          {props.canDelete && (
            <DefaultButton
              iconProps={{ iconName: 'Delete' }}
              text={strings.Responses_Detail_Delete}
              disabled={deleting}
              onClick={() => setConfirmingDelete(true)}
            />
          )}
        </div>
      )}
    >
      <div className={styles.detailMeta}>
        <span className={styles.detailMetaItem}>
          <Icon iconName="Clock" /> {item.created.toLocaleString(props.locale)}
        </span>
        {item.createdBy && (
          <span className={styles.detailMetaItem}>
            <Icon iconName="Contact" /> {item.createdBy}
          </span>
        )}
        {item.durationSeconds !== undefined && (
          <span className={styles.detailMetaItem}>
            <Icon iconName="Timer" />{' '}
            {formatString(strings.Responses_Detail_Took, {
              duration: formatDuration(item.durationSeconds)
            })}
          </span>
        )}
        {item.status === 'Draft' && (
          <span className={styles.filterChip}>
            <Icon iconName="Edit" /> {strings.Responses_Detail_Draft}
          </span>
        )}
        <span className={styles.detailMetaItem}>#{item.id}</span>
      </div>

      {props.approvalEnabled && props.onSetApproval && (
        <div className={styles.approvalBox}>
          <ApprovalChip status={ex.approvalStatus} />
          {ex.reviewedBy && ex.approvalStatus && ex.approvalStatus !== 'Pending' && (
            <div className={styles.approvalMeta}>
              {formatString(strings.Responses_Approval_ReviewedBy, { name: ex.reviewedBy })}
            </div>
          )}
          {approvalError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setApprovalError('')}>
              {approvalError}
            </MessageBar>
          )}
          <TextField
            label={strings.Responses_Approval_CommentLabel}
            multiline={true}
            rows={2}
            value={comment}
            disabled={approvalBusy}
            onChange={(_e, v) => setComment(v || '')}
          />
          <div className={styles.approvalActions}>
            <PrimaryButton
              iconProps={{ iconName: 'CheckMark' }}
              text={strings.Responses_Approval_Approve}
              disabled={approvalBusy}
              onClick={() => {
                void decide('Approved');
              }}
            />
            <DefaultButton
              iconProps={{ iconName: 'Cancel' }}
              text={strings.Responses_Approval_Reject}
              disabled={approvalBusy}
              onClick={() => {
                void decide('Rejected');
              }}
            />
            <DefaultButton
              iconProps={{ iconName: 'Undo' }}
              text={strings.Responses_Approval_Reset}
              disabled={approvalBusy || !ex.approvalStatus || ex.approvalStatus === 'Pending'}
              onClick={() => {
                void decide('Pending');
              }}
            />
          </div>
        </div>
      )}

      <Pivot>
        <PivotItem headerText={strings.Responses_Detail_Answers} itemIcon="BulletedList">
          {attachments === undefined && hasAttachmentFields && (
            <Spinner size={SpinnerSize.small} label={strings.Responses_Detail_LoadingAttachments} labelPosition="right" />
          )}
          <div className={styles.detailList}>
            {fields.map((field) => (
              <div key={field.id} className={styles.detailRow}>
                <div className={styles.detailLabel}>{field.title || strings.Responses_Dashboard_UntitledQuestion}</div>
                {renderAnswer(field)}
              </div>
            ))}
          </div>
        </PivotItem>
        <PivotItem headerText={strings.Responses_Detail_AsTheForm} itemIcon="PageLeft">
          <ResponseFormView
            definition={definition}
            values={formValues}
            spService={props.spService}
          />
        </PivotItem>
      </Pivot>

      <Dialog
        hidden={!confirmingDelete}
        onDismiss={() => setConfirmingDelete(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.Responses_Detail_DeleteTitle,
          subText: strings.Responses_Detail_DeleteBody
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text={deleting ? strings.Responses_Detail_Deleting : strings.Responses_Detail_Delete}
            disabled={deleting}
            onClick={() => {
              void handleDelete();
            }}
          />
          <DefaultButton
            text={strings.Responses_Detail_Cancel}
            disabled={deleting}
            onClick={() => setConfirmingDelete(false)}
          />
        </DialogFooter>
      </Dialog>
    </Panel>
  );
};
