import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  IconButton,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import styles from './ResponsesView.module.scss';
import { FieldType, IFormDefinition, IFormField, IFormValues, IResponseItem } from '../../models';
import { formatSharePointValue, inputFields, normalizeFromSharePoint } from '../../utils/formUtils';
import { formatDuration } from '../../utils/analytics';
import { SharePointService } from '../../services/SharePointService';
import { ResponseFormView } from '../form/FormRenderer';
import { RichTextView } from '../form/RichTextField';

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
  const [attachments, setAttachments] = React.useState<IAttachment[] | undefined>(undefined);

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
        return <img className={styles.signaturePreview} src={files[0].url} alt="Signature" />;
      }
      return <span className={raw ? styles.detailValue : styles.detailValueEmpty}>{raw ? 'Signed' : 'No answer'}</span>;
    }

    if (field.type === FieldType.FileUpload) {
      const files = attachmentsFor(field);
      if (files.length === 0) {
        const names = formatSharePointValue(field, raw);
        return (
          <span className={names ? styles.detailValue : styles.detailValueEmpty}>
            {names || 'No answer'}
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
        {text || 'No answer'}
      </span>
    );
  };

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText={'Response ' + (props.index + 1) + ' of ' + props.total}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            title="Previous response"
            ariaLabel="Previous response"
            disabled={props.index <= 0}
            onClick={() => props.onNavigate(-1)}
          />
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            title="Next response"
            ariaLabel="Next response"
            disabled={props.index >= props.total - 1}
            onClick={() => props.onNavigate(1)}
          />
          <DefaultButton
            iconProps={{ iconName: 'Print' }}
            text="Print"
            onClick={() => window.print()}
          />
          <span className={styles.panelFooterSpacer} />
          {props.canDelete && (
            <DefaultButton
              iconProps={{ iconName: 'Delete' }}
              text="Delete"
              disabled={deleting}
              onClick={() => setConfirmingDelete(true)}
            />
          )}
        </div>
      )}
    >
      <div className={styles.detailMeta}>
        <span className={styles.detailMetaItem}>
          <Icon iconName="Clock" /> {item.created.toLocaleString()}
        </span>
        {item.createdBy && (
          <span className={styles.detailMetaItem}>
            <Icon iconName="Contact" /> {item.createdBy}
          </span>
        )}
        {item.durationSeconds !== undefined && (
          <span className={styles.detailMetaItem}>
            <Icon iconName="Timer" /> took {formatDuration(item.durationSeconds)}
          </span>
        )}
        {item.status === 'Draft' && (
          <span className={styles.filterChip}>
            <Icon iconName="Edit" /> Draft
          </span>
        )}
        <span className={styles.detailMetaItem}>#{item.id}</span>
      </div>

      <Pivot>
        <PivotItem headerText="Answers" itemIcon="BulletedList">
          {attachments === undefined && hasAttachmentFields && (
            <Spinner size={SpinnerSize.small} label="Loading attachments…" labelPosition="right" />
          )}
          <div className={styles.detailList}>
            {fields.map((field) => (
              <div key={field.id} className={styles.detailRow}>
                <div className={styles.detailLabel}>{field.title || 'Untitled question'}</div>
                {renderAnswer(field)}
              </div>
            ))}
          </div>
        </PivotItem>
        <PivotItem headerText="As the form" itemIcon="PageLeft">
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
          title: 'Delete this response?',
          subText: 'The list item moves to the site recycle bin, so it can be restored from there.'
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
    </Panel>
  );
};
