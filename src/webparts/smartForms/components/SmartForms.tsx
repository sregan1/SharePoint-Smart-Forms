import * as React from 'react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  IconButton,
  MessageBar,
  MessageBarButton,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import styles from './SmartForms.module.scss';
import { SharePointService } from '../services/SharePointService';
import {
  createEmptyFormDefinition,
  FieldType,
  FIELD_TYPE_META,
  IDashboardSettings,
  IFormDefinition,
  IListInfo,
  isInputType
} from '../models';
import { FORM_TEMPLATES, buildTemplate, templateQuestionCount } from '../models/templates';
import {
  allFields,
  generateInternalName,
  inputFields,
  parseFormDefinition,
  validateDefinition
} from '../utils/formUtils';
import { buildTheme } from '../utils/theme';
import { debugLog, logError } from '../utils/debug';
import { FormRenderer } from './form/FormRenderer';
import { FormDesigner } from './designer/FormDesigner';
import { FormSettingsPanel } from './designer/FormSettingsPanel';
import { ResponsesView } from './responses/ResponsesView';

export interface ISmartFormsProps {
  listId: string;
  formDefinitionJson: string;
  /** stable per web part instance; used as the key for the durable, page-independent definition store */
  instanceId: string;
  spService: SharePointService;
  isEditMode: boolean;
  /** the site theme, so the form can render legibly on a dark page */
  theme: IReadonlyTheme | undefined;
  onConfigure: () => void;
  onFormDefinitionChange: (definitionJson: string) => void;
}

type TabKey = 'questions' | 'responses';

/** True when the page was opened through a share URL (?sfview=fill). */
const isFillView = (): boolean => {
  try {
    return new URLSearchParams(window.location.search).get('sfview') === 'fill';
  } catch {
    return false;
  }
};

/** The current page URL with the fill-view parameter applied. */
const buildShareUrl = (): string => {
  const url = new URL(window.location.href);
  url.searchParams.set('sfview', 'fill');
  url.hash = '';
  return url.toString();
};

const copyText = async (text: string): Promise<void> => {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }
  // legacy fallback for non-secure contexts
  const input = document.createElement('textarea');
  input.value = text;
  input.style.position = 'fixed';
  input.style.opacity = '0';
  document.body.appendChild(input);
  input.select();
  document.execCommand('copy');
  document.body.removeChild(input);
};

/**
 * Final tidy-up before list columns are created: give untitled questions a
 * name, drop empty options, and derive clean internal column names from the
 * final titles.
 */
const prepareForPublish = (definition: IFormDefinition): IFormDefinition => {
  const next: IFormDefinition = JSON.parse(JSON.stringify(definition));
  const taken = inputFields(next)
    .filter((f) => f.provisioned === true)
    .map((f) => f.internalName);
  let questionNumber = 0;
  next.sections.forEach((section) => {
    section.fields.forEach((field) => {
      if (field.choices) {
        field.choices = field.choices.map((c) => (c || '').trim()).filter((c) => c.length > 0);
      }
      if (field.likertRows) {
        field.likertRows = field.likertRows.map((r) => (r || '').trim()).filter((r) => r.length > 0);
      }
      if (field.likertColumns) {
        field.likertColumns = field.likertColumns
          .map((c) => (c || '').trim())
          .filter((c) => c.length > 0);
      }
      if (field.imageChoices) {
        field.imageChoices = field.imageChoices.filter((o) => (o.label || '').trim().length > 0);
      }
      // content blocks store nothing, so they never get a column
      if (field.type === FieldType.Content) {
        return;
      }
      questionNumber++;
      if (field.provisioned !== true) {
        field.title = (field.title || '').trim() || 'Question ' + questionNumber;
        field.internalName = generateInternalName(field.title, taken);
        taken.push(field.internalName);
      }
    });
  });
  return next;
};

export const SmartForms: React.FunctionComponent<ISmartFormsProps> = (props) => {
  const { listId, spService, formDefinitionJson } = props;

  const [loading, setLoading] = React.useState<boolean>(!!listId);
  const [loadError, setLoadError] = React.useState<string>('');
  const [listInfo, setListInfo] = React.useState<IListInfo | undefined>(undefined);
  const [isOwner, setIsOwner] = React.useState<boolean>(false);
  const [definition, setDefinition] = React.useState<IFormDefinition>(
    () => parseFormDefinition(formDefinitionJson) || createEmptyFormDefinition()
  );
  const [activeTab, setActiveTab] = React.useState<TabKey>('questions');
  const [previewing, setPreviewing] = React.useState<boolean>(false);
  const [settingsOpen, setSettingsOpen] = React.useState<boolean>(false);
  const [shareOpen, setShareOpen] = React.useState<boolean>(false);
  const [templatesOpen, setTemplatesOpen] = React.useState<boolean>(false);
  const [publishing, setPublishing] = React.useState<boolean>(false);
  const [publishError, setPublishError] = React.useState<string>('');
  const [publishNotice, setPublishNotice] = React.useState<string>('');
  const [conflictNotice, setConflictNotice] = React.useState<string>('');
  const [linkCopied, setLinkCopied] = React.useState<boolean>(false);
  const [preflightOpen, setPreflightOpen] = React.useState<boolean>(false);
  const [focusFieldId, setFocusFieldId] = React.useState<string | undefined>(undefined);
  // 'clean' once the definition has been saved to SharePoint (independent of page save)
  const [saveState, setSaveState] = React.useState<'clean' | 'saving' | 'error'>('clean');

  const fillView = React.useMemo(isFillView, []);

  const theme = React.useMemo(
    () => buildTheme(props.theme, definition.settings.accentColor),
    [props.theme, definition.settings.accentColor]
  );

  const issues = React.useMemo(() => validateDefinition(definition), [definition]);
  const blockingIssues = issues.filter((issue) => issue.severity === 'error');

  // ----- auto-save: edits persist straight to SharePoint, debounced, independent of page save -----
  const pendingJson = React.useRef<string | undefined>(undefined);
  const saveTimer = React.useRef<number | undefined>(undefined);
  // only the most recently issued save is allowed to move the chip out of 'saving',
  // so a slow earlier request can't clobber the state after a faster later one lands
  const saveSeq = React.useRef(0);

  const persist = (updated: IFormDefinition, immediate?: boolean): void => {
    setDefinition(updated);
    const json = JSON.stringify(updated);
    pendingJson.current = json;
    setSaveState('saving');
    if (saveTimer.current) {
      clearTimeout(saveTimer.current);
    }
    const flush = (): void => {
      pendingJson.current = undefined;
      // lightweight mirror on the page property; not required for durability
      props.onFormDefinitionChange(json);
      const seq = ++saveSeq.current;
      spService
        .saveFormDefinition(props.instanceId, json)
        .then(() => {
          if (saveSeq.current === seq) {
            setSaveState('clean');
          }
        })
        .catch((error) => {
          logError('saveFormDefinition', error);
          if (saveSeq.current === seq) {
            setSaveState('error');
          }
        });
    };
    if (immediate) {
      flush();
      return;
    }
    saveTimer.current = setTimeout(flush, 600) as unknown as number;
  };

  React.useEffect(
    () => () => {
      // flush a pending auto-save when the web part unmounts
      if (saveTimer.current) {
        clearTimeout(saveTimer.current);
      }
      if (pendingJson.current) {
        props.onFormDefinitionChange(pendingJson.current);
        void spService.saveFormDefinition(props.instanceId, pendingJson.current);
      }
    },
    []
  );

  // hydrate from the durable store once on mount; it's the source of truth once
  // anything has been saved through it. A page property with no matching durable
  // record means the form predates this store (or was never saved) — migrate it.
  React.useEffect(() => {
    let cancelled = false;
    debugLog('hydrating form definition', { instanceId: props.instanceId });
    spService
      .loadFormDefinition(props.instanceId)
      .then((stored) => {
        if (cancelled) {
          return;
        }
        debugLog('hydration result', { instanceId: props.instanceId, found: !!stored });
        if (stored) {
          const parsed = parseFormDefinition(stored);
          if (parsed) {
            setDefinition(parsed);
          }
        } else if (formDefinitionJson) {
          void spService.saveFormDefinition(props.instanceId, formDefinitionJson);
        }
      })
      .catch((error) => {
        // the initial paint from the page property already covers this case, but
        // log it so a silent read failure doesn't look like "the data disappeared"
        logError('loadFormDefinition', error);
      });
    return () => {
      cancelled = true;
    };
    // instanceId is stable for the lifetime of this web part instance
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    let cancelled = false;
    if (!listId) {
      setListInfo(undefined);
      setLoading(false);
      return undefined;
    }
    setLoading(true);
    setLoadError('');
    Promise.all([spService.getListInfo(listId), spService.currentUserIsOwner(listId)])
      .then(([info, owner]) => {
        if (cancelled) {
          return;
        }
        debugLog('list info loaded', { listId, found: !!info, isOwner: owner });
        setListInfo(info);
        setIsOwner(owner);
        if (!info) {
          setLoadError(
            'The list configured for this form no longer exists. Pick another list in the web part settings.'
          );
        }
        setLoading(false);
      })
      .catch((error) => {
        if (!cancelled) {
          logError('loading list info', error);
          setLoadError('Something went wrong loading the form configuration.');
          setLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [listId]);

  // ----- publish -----

  const runPublish = async (): Promise<void> => {
    setPublishing(true);
    setPublishError('');
    setPublishNotice('');
    setConflictNotice('');
    try {
      const prepared = prepareForPublish(definition);
      const result = await spService.ensureFields(listId, prepared);
      debugLog('ensureFields complete', {
        created: result.created,
        conflicts: result.conflicts.map((c) => c.field.internalName)
      });
      persist(result.definition, true);
      spService.clearLookupCache();

      if (result.conflicts.length > 0) {
        setConflictNotice(
          result.conflicts
            .map(
              (conflict) =>
                '"' +
                (conflict.field.title || 'a question') +
                '" already has a list column of a different type (' +
                conflict.existingType +
                '). Rename the question so a new column can be created, or change its type back.'
            )
            .join(' ')
        );
      }
      setPublishNotice(
        result.created.length > 0
          ? result.created.length +
              (result.created.length === 1 ? ' column was' : ' columns were') +
              ' added to "' +
              (listInfo ? listInfo.title : 'the list') +
              '".'
          : 'The response list is already up to date.'
      );
      setShareOpen(true);
    } catch (error) {
      logError('ensureFields (Collect responses)', error);
      setPublishError(
        'The list columns could not be created. This usually means you don\'t have Manage Lists permission on the response list, ' +
          'but it can also happen on sites with custom scripting disabled. Check the browser console for the underlying error, or try again.'
      );
    }
    setPublishing(false);
  };

  const handleCollectResponses = (): void => {
    // surface anything that would make the published form misbehave, rather than
    // provisioning columns for a form with no options or broken branching
    if (issues.length > 0) {
      setPreflightOpen(true);
      return;
    }
    void runPublish();
  };

  const handleCopyLink = (): void => {
    copyText(buildShareUrl())
      .then(() => {
        setLinkCopied(true);
        setTimeout(() => setLinkCopied(false), 2500);
      })
      .catch(() => {
        // clipboard unavailable — the URL is still visible to copy manually
      });
  };

  const applyTemplate = (key: string): void => {
    const built = buildTemplate(key);
    // keep the accent the template was designed around
    persist(built, true);
    setTemplatesOpen(false);
    setActiveTab('questions');
  };

  const rootStyle = theme.tokens as unknown as React.CSSProperties;

  // ----- early returns -----

  if (!listId) {
    return (
      <div className={styles.smartForms} style={rootStyle}>
        <div className={styles.placeholder}>
          <div className={styles.placeholderIcon}>
            <Icon iconName="ClipboardList" />
          </div>
          <h2>Smart Forms</h2>
          <p>
            Build a polished form on top of a SharePoint list — {inputTypeCount()} question types,
            branching logic, file uploads, shareable fill-in links, and a results dashboard.
          </p>
          {props.isEditMode ? (
            <div className={styles.placeholderActions}>
              <PrimaryButton
                iconProps={{ iconName: 'Settings' }}
                text="Choose or create a list"
                onClick={props.onConfigure}
              />
            </div>
          ) : (
            <p className={styles.placeholderHint}>
              Edit the page and configure this web part to get started.
            </p>
          )}
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className={styles.smartForms} style={rootStyle}>
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label="Loading form…" />
        </div>
      </div>
    );
  }

  if (loadError || !listInfo) {
    return (
      <div className={styles.smartForms} style={rootStyle}>
        <MessageBar messageBarType={MessageBarType.error}>
          {loadError || 'Unable to load the form.'}
        </MessageBar>
        {props.isEditMode && (
          <div className={styles.errorActions}>
            <PrimaryButton
              iconProps={{ iconName: 'Settings' }}
              text="Open settings"
              onClick={props.onConfigure}
            />
          </div>
        )}
      </div>
    );
  }

  const liveForm = (
    <FormRenderer definition={definition} listId={listId} spService={spService} mode="live" />
  );

  // A share link shows only the form — no tabs, even for owners
  if (fillView || !isOwner) {
    return (
      <div className={styles.smartForms} style={rootStyle}>
        {liveForm}
      </div>
    );
  }

  const questionCount = inputFields(definition).length;
  const hasUnprovisioned = inputFields(definition).some((f) => f.provisioned !== true);

  return (
    <div className={styles.smartForms} style={rootStyle}>
      <div className={styles.ownerBar}>
        <div className={styles.tabs} role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'questions' && !previewing}
            className={activeTab === 'questions' && !previewing ? styles.tabActive : styles.tab}
            onClick={() => {
              setActiveTab('questions');
              setPreviewing(false);
            }}
          >
            Questions{questionCount > 0 ? ' (' + questionCount + ')' : ''}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'responses' && !previewing}
            className={activeTab === 'responses' && !previewing ? styles.tabActive : styles.tab}
            title={'Responses are saved to "' + listInfo.title + '"'}
            onClick={() => {
              setActiveTab('responses');
              setPreviewing(false);
            }}
          >
            Responses
          </button>
        </div>

        <div className={styles.ownerBarRight}>
          {renderSaveChip(saveState)}
          <IconButton
            iconProps={{ iconName: 'Lightbulb' }}
            title="Start from a template"
            ariaLabel="Start from a template"
            onClick={() => setTemplatesOpen(true)}
          />
          <IconButton
            iconProps={{ iconName: 'Settings' }}
            title="Form settings — appearance, notifications, access"
            ariaLabel="Form settings"
            onClick={() => setSettingsOpen(true)}
          />
          <DefaultButton
            iconProps={{ iconName: previewing ? 'Cancel' : 'View' }}
            text={previewing ? 'Close preview' : 'Preview'}
            disabled={questionCount === 0}
            onClick={() => setPreviewing(!previewing)}
          />
          <PrimaryButton
            iconProps={{ iconName: 'Send' }}
            text={publishing ? 'Getting ready…' : hasUnprovisioned ? 'Collect responses' : 'Share'}
            disabled={publishing || questionCount === 0}
            onClick={handleCollectResponses}
          />
        </div>
      </div>

      {saveState === 'error' && (
        <MessageBar
          messageBarType={MessageBarType.error}
          className={styles.publishBanner}
          actions={
            <div>
              <MessageBarButton onClick={() => persist(definition, true)}>Retry</MessageBarButton>
            </div>
          }
        >
          Your last change could not be saved. Check your connection, then try again.
        </MessageBar>
      )}

      {publishError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setPublishError('')}
          className={styles.publishError}
        >
          {publishError}
        </MessageBar>
      )}

      {conflictNotice && (
        <MessageBar
          messageBarType={MessageBarType.severeWarning}
          onDismiss={() => setConflictNotice('')}
          className={styles.publishError}
        >
          {conflictNotice}
        </MessageBar>
      )}

      {publishNotice && !conflictNotice && (
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={() => setPublishNotice('')}
          className={styles.publishError}
        >
          {publishNotice}
        </MessageBar>
      )}

      {previewing ? (
        <>
          <div className={styles.previewHint}>
            <Icon iconName="View" />
            <span>
              Preview — nothing you submit here is saved. Validation and branching behave exactly as
              they will for respondents.
            </span>
            <span className={styles.previewHintSpacer} />
            <DefaultButton text="Close preview" onClick={() => setPreviewing(false)} />
          </div>
          <FormRenderer
            definition={definition}
            listId={listId}
            spService={spService}
            mode="preview"
          />
        </>
      ) : activeTab === 'questions' ? (
        <FormDesigner
          definition={definition}
          spService={spService}
          focusFieldId={focusFieldId}
          onFocusHandled={() => setFocusFieldId(undefined)}
          onChange={(updated) => persist(updated)}
        />
      ) : (
        <ResponsesView
          definition={definition}
          listId={listId}
          listTitle={listInfo.title}
          spService={spService}
          theme={theme}
          isOwner={isOwner}
          onDashboardChange={(dashboard: IDashboardSettings) =>
            persist({ ...definition, settings: { ...definition.settings, dashboard: dashboard } })
          }
          onCopyShareLink={handleCopyLink}
          onPreview={() => {
            setPreviewing(true);
            setActiveTab('questions');
          }}
        />
      )}

      {settingsOpen && (
        <FormSettingsPanel
          settings={definition.settings}
          onSave={(settings) => {
            persist({ ...definition, settings: settings });
            setSettingsOpen(false);
          }}
          onDismiss={() => setSettingsOpen(false)}
        />
      )}

      {/* ----- pre-flight ----- */}
      <Dialog
        hidden={!preflightOpen}
        onDismiss={() => setPreflightOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: blockingIssues.length > 0 ? 'Fix these before collecting' : 'Worth a look first',
          subText:
            blockingIssues.length > 0
              ? 'These would stop the published form working properly.'
              : 'Nothing is broken, but these are usually worth fixing before you share the link.'
        }}
        minWidth={520}
      >
        <ul className={styles.issueList}>
          {issues.map((issue, index) => (
            <li key={index} className={issue.severity === 'error' ? styles.issueError : styles.issueWarning}>
              <Icon iconName={issue.severity === 'error' ? 'ErrorBadge' : 'Warning'} />
              <span className={styles.issueMessage}>{issue.message}</span>
              {issue.fieldId && (
                <DefaultButton
                  className={styles.issueJump}
                  text="Go to"
                  onClick={() => {
                    setFocusFieldId(issue.fieldId);
                    setActiveTab('questions');
                    setPreviewing(false);
                    setPreflightOpen(false);
                  }}
                />
              )}
            </li>
          ))}
        </ul>
        <DialogFooter>
          {blockingIssues.length === 0 && (
            <PrimaryButton
              text="Collect responses anyway"
              onClick={() => {
                setPreflightOpen(false);
                void runPublish();
              }}
            />
          )}
          <DefaultButton
            text={blockingIssues.length > 0 ? 'Back to the form' : 'Cancel'}
            onClick={() => setPreflightOpen(false)}
          />
        </DialogFooter>
      </Dialog>

      {/* ----- share ----- */}
      {shareOpen && (
        <Dialog
          hidden={false}
          onDismiss={() => setShareOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Your form is ready to share',
            subText: 'Send this link to anyone who should fill in the form.'
          }}
          modalProps={{ isBlocking: false }}
          minWidth={520}
        >
          <div className={styles.shareRow}>
            <TextField
              readOnly={true}
              value={buildShareUrl()}
              className={styles.shareField}
              ariaLabel="Share link"
              onClick={(event) => (event.target as HTMLInputElement).select()}
            />
            <PrimaryButton
              iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Copy' }}
              text={linkCopied ? 'Copied' : 'Copy'}
              onClick={handleCopyLink}
            />
          </div>

          {/*
            The commonest support question about a shared form is "why does
            submitting fail for my colleague" — nearly always list permissions.
          */}
          <div className={styles.shareNote}>
            <Icon iconName="Info" />
            <span>
              Respondents need permission to <strong>add items</strong> to &ldquo;{listInfo.title}
              &rdquo;, and access to this page. If someone gets an error on submit, that is almost
              always why.
            </span>
          </div>
          <div className={styles.shareNote}>
            <Icon iconName="Lightbulb" />
            <span>
              You can pre-answer questions in the link by adding the column name, for example{' '}
              <code>&amp;{firstInternalName(definition)}=Marketing</code>.
            </span>
          </div>

          <DialogFooter>
            <DefaultButton text="Done" onClick={() => setShareOpen(false)} />
          </DialogFooter>
        </Dialog>
      )}

      {/* ----- templates ----- */}
      {templatesOpen && (
        <Dialog
          hidden={false}
          onDismiss={() => setTemplatesOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Start from a template',
            subText:
              questionCount > 0
                ? 'Choosing a template replaces the questions you have now.'
                : 'Pick a starting point — you can change anything afterwards.'
          }}
          modalProps={{ isBlocking: false }}
          minWidth={640}
        >
          <div className={styles.templateGrid}>
            {FORM_TEMPLATES.map((template) => (
              <button
                key={template.key}
                type="button"
                className={styles.templateCard}
                onClick={() => applyTemplate(template.key)}
              >
                <span className={styles.templateIcon}>
                  <Icon iconName={template.icon} />
                </span>
                <span className={styles.templateName}>{template.name}</span>
                <span className={styles.templateMeta}>{template.description}</span>
                <span className={styles.templateMeta}>
                  {template.key === 'blank'
                    ? 'Empty'
                    : templateQuestionCount(template) + ' questions'}
                </span>
              </button>
            ))}
          </div>
          <DialogFooter>
            <DefaultButton text="Cancel" onClick={() => setTemplatesOpen(false)} />
          </DialogFooter>
        </Dialog>
      )}
    </div>
  );
};

/** Save-state chip — reflects whether the definition is durably saved to SharePoint. */
const renderSaveChip = (state: 'clean' | 'saving' | 'error'): React.ReactNode => {
  if (state === 'saving') {
    return (
      <span className={styles.saveChipPending} title="Saving your latest edit">
        <Icon iconName="Sync" /> Saving…
      </span>
    );
  }
  if (state === 'error') {
    return (
      <span
        className={styles.saveChipUnpublished}
        title="Your last edit could not be saved. Check your connection and try again."
      >
        <Icon iconName="Warning" /> Save failed
      </span>
    );
  }
  return (
    <span className={styles.saveChipSaved} title="All changes are saved">
      <Icon iconName="CheckMark" /> Saved
    </span>
  );
};

/**
 * How many question types the palette offers, derived rather than written down
 * so the placeholder copy can't drift out of step with the catalogue. Presets
 * are shortcuts onto these types, not additional types, so they don't count.
 */
const inputTypeCount = (): number =>
  FIELD_TYPE_META.filter((meta) => isInputType(meta.type)).length;

/** An internal name to use in the prefill example, if the form has one. */
const firstInternalName = (definition: IFormDefinition): string => {
  const provisioned = allFields(definition).filter(
    (f) => f.provisioned === true && f.internalName
  )[0];
  return provisioned ? provisioned.internalName : 'SFYourQuestion';
};
