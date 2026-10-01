import * as React from 'react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  IconButton,
  ITextField,
  MessageBar,
  MessageBarButton,
  MessageBarType,
  PrimaryButton,
  setRTL,
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
  IMessageBag,
  isInputType
} from '../models';
import {
  FORM_TEMPLATES,
  buildTemplate,
  templateDescription,
  templateName,
  templateQuestionCount
} from '../models/templates';
import {
  allFields,
  generateInternalName,
  inputFields,
  parseFormDefinition,
  retireField,
  validateDefinition
} from '../utils/formUtils';
import { buildTheme } from '../utils/theme';
import { debugLog, logError } from '../utils/debug';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString, isRtlLocale } from '../utils/localeUtils';
import { FormRenderer } from './form/FormRenderer';
import { FormDesigner } from './designer/FormDesigner';
import { FormSettingsPanel } from './designer/FormSettingsPanel';
import { VersionHistoryPanel } from './designer/VersionHistoryPanel';
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
  /** page UI culture (e.g. 'ar-SA'); drives right-to-left layout and date formatting */
  locale?: string;
  /**
   * Page URL to build share links from. Supplied when the web part runs inside
   * Microsoft Teams, where window.location is the Teams iframe, not the page.
   */
  shareBaseUrl?: string;
}

const messageBag = strings as unknown as IMessageBag;

type TabKey = 'questions' | 'responses';

/** True when the page was opened through a share URL (?sfview=fill). */
const isFillView = (): boolean => {
  try {
    return new URLSearchParams(window.location.search).get('sfview') === 'fill';
  } catch {
    return false;
  }
};

/** The page URL (or the supplied Teams-safe page URL) with the fill-view parameter applied. */
const buildShareUrl = (base?: string): string => {
  const url = new URL(base || window.location.href);
  url.searchParams.set('sfview', 'fill');
  url.hash = '';
  return url.toString();
};

/** Fills {token} placeholders in a template with React nodes (for inline emphasis / code). */
const renderRich = (template: string, nodes: { [token: string]: React.ReactNode }): React.ReactNode[] =>
  template.split(/(\{[A-Za-z]+\})/).map((part, i) => {
    const m = /^\{([A-Za-z]+)\}$/.exec(part);
    return m && nodes[m[1]] !== undefined ? <React.Fragment key={i}>{nodes[m[1]]}</React.Fragment> : part;
  });

/** Legacy copy through a temporary textarea; false when the browser refuses. */
const copyViaTextarea = (text: string): boolean => {
  const input = document.createElement('textarea');
  input.value = text;
  input.style.position = 'fixed';
  input.style.opacity = '0';
  document.body.appendChild(input);
  input.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  document.body.removeChild(input);
  return ok;
};

/**
 * Copies text: async clipboard first, then the textarea fallback (an iframe such
 * as Teams often rejects the async API). Resolves false when both fail so the
 * caller can ask the user to copy by hand.
 */
const copyText = async (text: string): Promise<boolean> => {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // permission policy or focus problem; fall through to the legacy path
    }
  }
  return copyViaTextarea(text);
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
    .map((f) => f.internalName)
    .concat(next.retiredColumns || []);
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
        field.title = (field.title || '').trim() || formatString(strings.App_DefaultQuestionTitle, { number: String(questionNumber) });
        field.internalName = generateInternalName(field.title, taken, next.retiredColumns);
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
  const [copyFailed, setCopyFailed] = React.useState<boolean>(false);
  const [historyOpen, setHistoryOpen] = React.useState<boolean>(false);
  // the definition in place before a version restore, so the owner can undo it
  const [restoreUndo, setRestoreUndo] = React.useState<{ previous: IFormDefinition; date: string } | undefined>(
    undefined
  );
  const shareFieldRef = React.useRef<ITextField | null>(null);
  const [preflightOpen, setPreflightOpen] = React.useState<boolean>(false);
  const [focusFieldId, setFocusFieldId] = React.useState<string | undefined>(undefined);
  // 'clean' once the definition has been saved to SharePoint (independent of page save)
  const [saveState, setSaveState] = React.useState<'clean' | 'saving' | 'error'>('clean');

  const fillView = React.useMemo(isFillView, []);
  const rtl = isRtlLocale(props.locale);
  const shareUrl = buildShareUrl(props.shareBaseUrl);

  // Fluent's portal-hosted surfaces (dialogs, panels, menus) read the global RTL flag
  React.useEffect(() => {
    if (rtl) {
      setRTL(true);
    }
  }, [rtl]);

  const theme = React.useMemo(
    () => buildTheme(props.theme, definition.settings.accentColor),
    [props.theme, definition.settings.accentColor]
  );

  const issues = React.useMemo(() => validateDefinition(definition, messageBag), [definition]);
  const blockingIssues = issues.filter((issue) => issue.severity === 'error');

  // ----- auto-save: edits persist straight to SharePoint, debounced, independent of page save -----
  const pendingJson = React.useRef<string | undefined>(undefined);
  const saveTimer = React.useRef<number | undefined>(undefined);
  // only the most recently issued save is allowed to move the chip out of 'saving',
  // so a slow earlier request can't clobber the state after a faster later one lands
  const saveSeq = React.useRef(0);
  // set once the owner edits anything; a slower initial load must never overwrite that
  const editedRef = React.useRef(false);
  // what the durable store held at load time (undefined = nothing / unreadable)
  const storedRef = React.useRef<string | undefined>(undefined);
  const [hydrated, setHydrated] = React.useState<boolean>(false);
  const ownerSetupDone = React.useRef(false);

  const persist = (updated: IFormDefinition, immediate?: boolean): void => {
    editedRef.current = true;
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
      if (pendingJson.current && editedRef.current) {
        props.onFormDefinitionChange(pendingJson.current);
        void spService.saveFormDefinition(props.instanceId, pendingJson.current);
      }
    },
    []
  );

  // hydrate from the durable store once on mount. This read is passive: it never
  // creates the configuration list or writes anything, so readers and share-link
  // users can run it safely. Anything the owner edited before it returns wins.
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
        storedRef.current = stored;
        if (stored && !editedRef.current) {
          const parsed = parseFormDefinition(stored);
          if (parsed) {
            setDefinition(parsed);
          }
        }
        setHydrated(true);
      })
      .catch((error) => {
        // the initial paint from the page property already covers this case, but
        // log it so a silent read failure doesn't look like "the data disappeared"
        logError('loadFormDefinition', error);
        if (!cancelled) {
          setHydrated(true);
        }
      });
    return () => {
      cancelled = true;
    };
    // instanceId is stable for the lifetime of this web part instance
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Owner-only setup, once both the read above and the permission check are back:
  // create/tighten the configuration list, and migrate a page-property definition
  // that predates the durable store. Never runs for readers or share-link users.
  React.useEffect(() => {
    if (!hydrated || !isOwner || !listInfo || ownerSetupDone.current) {
      return;
    }
    ownerSetupDone.current = true;
    if (storedRef.current === undefined && formDefinitionJson && !editedRef.current) {
      spService.saveFormDefinition(props.instanceId, formDefinitionJson).catch((error) => {
        logError('migrating the page definition', error);
      });
    } else {
      spService.loadFormDefinition(props.instanceId, true).catch((error) => {
        logError('ensuring the configuration list', error);
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hydrated, isOwner, listInfo]);

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
            strings.App_Error_ListMissing
          );
        }
        setLoading(false);
      })
      .catch((error) => {
        if (!cancelled) {
          logError('loading list info', error);
          setLoadError(strings.App_Error_LoadConfig);
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
      const result = await spService.ensureFields(listId, prepared, prepared.settings.enableApproval === true);
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
                formatString(strings.App_Publish_ConflictNotice, {
                  title: conflict.field.title || strings.App_Publish_ConflictUnnamedQuestion,
                  type: conflict.existingType
                })
            )
            .join(' ')
        );
      }
      setPublishNotice(
        result.created.length > 0
          ? formatString(
              result.created.length === 1
                ? strings.App_Publish_ColumnsAddedOne
                : strings.App_Publish_ColumnsAddedOther,
              {
                count: String(result.created.length),
                list: listInfo ? listInfo.title : strings.App_Publish_TheList
              }
            )
          : strings.App_Publish_UpToDate
      );
      setShareOpen(true);
    } catch (error) {
      logError('ensureFields (Collect responses)', error);
      setPublishError(
strings.App_Publish_Error
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
    setCopyFailed(false);
    copyText(shareUrl)
      .then((ok) => {
        if (ok) {
          setLinkCopied(true);
          setTimeout(() => setLinkCopied(false), 2500);
          return;
        }
        // both clipboard paths refused: show the link selected so the user can copy it by hand
        setCopyFailed(true);
        setShareOpen(true);
        setTimeout(() => {
          if (shareFieldRef.current) {
            shareFieldRef.current.select();
          }
        }, 150);
      })
      .catch(() => setCopyFailed(true));
  };

  const applyRestore = (restored: IFormDefinition, date: string): void => {
    const next: IFormDefinition = JSON.parse(JSON.stringify(restored));
    // columns of questions that exist now but not in the restored version stay in
    // the list; reserve their names so new questions never adopt them
    const keep: { [id: string]: boolean } = {};
    allFields(next).forEach((f) => {
      keep[f.internalName.toLowerCase()] = true;
    });
    allFields(definition).forEach((f) => {
      if (f.provisioned === true && f.internalName && !keep[f.internalName.toLowerCase()]) {
        retireField(next, f.internalName);
      }
    });
    (definition.retiredColumns || []).forEach((name) => retireField(next, name));
    setRestoreUndo({ previous: definition, date });
    persist(next, true);
    setHistoryOpen(false);
    setActiveTab('questions');
    setPreviewing(false);
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
      <div className={styles.smartForms} style={rootStyle} dir={rtl ? 'rtl' : undefined}>
        <div className={styles.placeholder}>
          <div className={styles.placeholderIcon}>
            <Icon iconName="ClipboardList" />
          </div>
          <h2>{strings.App_Placeholder_Title}</h2>
          <p>
            {formatString(strings.App_Placeholder_Description, { count: String(inputTypeCount()) })}
          </p>
          {props.isEditMode ? (
            <div className={styles.placeholderActions}>
              <PrimaryButton
                iconProps={{ iconName: 'Settings' }}
                text={strings.App_Placeholder_ChooseList}
                onClick={props.onConfigure}
              />
            </div>
          ) : (
            <p className={styles.placeholderHint}>
              {strings.App_Placeholder_Hint}
            </p>
          )}
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className={styles.smartForms} style={rootStyle} dir={rtl ? 'rtl' : undefined}>
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label={strings.App_Loading} />
        </div>
      </div>
    );
  }

  if (loadError || !listInfo) {
    return (
      <div className={styles.smartForms} style={rootStyle} dir={rtl ? 'rtl' : undefined}>
        <MessageBar messageBarType={MessageBarType.error}>
          {loadError || strings.App_Error_Unable}
        </MessageBar>
        {props.isEditMode && (
          <div className={styles.errorActions}>
            <PrimaryButton
              iconProps={{ iconName: 'Settings' }}
              text={strings.App_OpenSettings}
              onClick={props.onConfigure}
            />
          </div>
        )}
      </div>
    );
  }

  const liveForm = (
    <FormRenderer
      definition={definition}
      listId={listId}
      spService={spService}
      mode="live"
      locale={props.locale}
    />
  );

  // A share link shows only the form — no tabs, even for owners
  if (fillView || !isOwner) {
    return (
      <div className={styles.smartForms} style={rootStyle} dir={rtl ? 'rtl' : undefined}>
        {liveForm}
      </div>
    );
  }

  const questionCount = inputFields(definition).length;
  const hasUnprovisioned = inputFields(definition).some((f) => f.provisioned !== true);

  return (
    <div className={styles.smartForms} style={rootStyle} dir={rtl ? 'rtl' : undefined}>
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
            {questionCount > 0
              ? formatString(strings.App_Tab_QuestionsWithCount, { count: String(questionCount) })
              : strings.App_Tab_Questions}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'responses' && !previewing}
            className={activeTab === 'responses' && !previewing ? styles.tabActive : styles.tab}
            title={formatString(strings.App_Tab_Responses_Title, { list: listInfo.title })}
            onClick={() => {
              setActiveTab('responses');
              setPreviewing(false);
            }}
          >
            {strings.App_Tab_Responses}
          </button>
        </div>

        <div className={styles.ownerBarRight}>
          {renderSaveChip(saveState)}
          <IconButton
            iconProps={{ iconName: 'Lightbulb' }}
            title={strings.App_Templates_Start}
            ariaLabel={strings.App_Templates_Start}
            onClick={() => setTemplatesOpen(true)}
          />
          <IconButton
            iconProps={{ iconName: 'History' }}
            title={strings.App_History_Title}
            ariaLabel={strings.App_History_Title}
            onClick={() => setHistoryOpen(true)}
          />
          <IconButton
            iconProps={{ iconName: 'Settings' }}
            title={strings.App_Settings_Title}
            ariaLabel={strings.App_Settings_AriaLabel}
            onClick={() => setSettingsOpen(true)}
          />
          <DefaultButton
            iconProps={{ iconName: previewing ? 'Cancel' : 'View' }}
            text={previewing ? strings.App_Preview_Close : strings.App_Preview}
            disabled={questionCount === 0}
            onClick={() => setPreviewing(!previewing)}
          />
          <PrimaryButton
            iconProps={{ iconName: 'Send' }}
            text={
              publishing
                ? strings.App_Publish_GettingReady
                : hasUnprovisioned
                ? strings.App_Publish_CollectResponses
                : strings.App_Publish_Share
            }
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
              <MessageBarButton onClick={() => persist(definition, true)}>{strings.App_Retry}</MessageBarButton>
            </div>
          }
        >
          {strings.App_SaveError_Banner}
        </MessageBar>
      )}

      {restoreUndo && (
        <MessageBar
          messageBarType={MessageBarType.info}
          onDismiss={() => setRestoreUndo(undefined)}
          className={styles.publishError}
          actions={
            <div>
              <MessageBarButton
                onClick={() => {
                  persist(restoreUndo.previous, true);
                  setRestoreUndo(undefined);
                }}
              >
                {strings.App_History_Undo}
              </MessageBarButton>
            </div>
          }
        >
          {formatString(strings.App_History_RestoredNotice, { date: restoreUndo.date })}
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
              {strings.App_Preview_Hint}
            </span>
            <span className={styles.previewHintSpacer} />
            <DefaultButton text={strings.App_Preview_Close} onClick={() => setPreviewing(false)} />
          </div>
          <FormRenderer
            definition={definition}
            listId={listId}
            spService={spService}
            mode="preview"
            locale={props.locale}
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
          locale={props.locale}
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

      {historyOpen && (
        <VersionHistoryPanel
          spService={spService}
          instanceId={props.instanceId}
          locale={props.locale}
          onRestore={applyRestore}
          onDismiss={() => setHistoryOpen(false)}
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
          title:
            blockingIssues.length > 0
              ? strings.App_Preflight_BlockingTitle
              : strings.App_Preflight_WarningTitle,
          subText:
            blockingIssues.length > 0
              ? strings.App_Preflight_BlockingText
              : strings.App_Preflight_WarningText
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
                  text={strings.App_Preflight_GoTo}
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
              text={strings.App_Preflight_CollectAnyway}
              onClick={() => {
                setPreflightOpen(false);
                void runPublish();
              }}
            />
          )}
          <DefaultButton
            text={blockingIssues.length > 0 ? strings.App_Preflight_Back : strings.App_Cancel}
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
            title: strings.App_Share_Title,
            subText: strings.App_Share_SubText
          }}
          modalProps={{ isBlocking: false }}
          minWidth={520}
        >
          <div className={styles.shareRow}>
            <TextField
              readOnly={true}
              value={shareUrl}
              componentRef={(ref) => {
                shareFieldRef.current = ref;
              }}
              className={styles.shareField}
              ariaLabel={strings.App_Share_LinkAriaLabel}
              onClick={(event) => (event.target as HTMLInputElement).select()}
            />
            <PrimaryButton
              iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Copy' }}
              text={linkCopied ? strings.App_Share_Copied : strings.App_Share_Copy}
              onClick={handleCopyLink}
            />
          </div>

          {copyFailed && (
            <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setCopyFailed(false)}>
              {strings.App_Share_CopyManual}
            </MessageBar>
          )}
          {props.shareBaseUrl && (
            <div className={styles.shareNote}>
              <Icon iconName="TeamsLogo" />
              <span>{strings.App_Share_TeamsNote}</span>
            </div>
          )}

          {/*
            The commonest support question about a shared form is "why does
            submitting fail for my colleague" — nearly always list permissions.
          */}
          <div className={styles.shareNote}>
            <Icon iconName="Info" />
            <span>
              {renderRich(strings.App_Share_PermissionNote, {
                addItems: <strong>{strings.App_Share_AddItemsEmphasis}</strong>,
                title: listInfo.title
              })}
            </span>
          </div>
          <div className={styles.shareNote}>
            <Icon iconName="Lightbulb" />
            <span>
              {renderRich(strings.App_Share_PrefillNote, {
                example: <code>&amp;{firstInternalName(definition)}=Marketing</code>
              })}{' '}
              {strings.App_Share_PrefillNotSecure}
            </span>
          </div>

          <DialogFooter>
            <DefaultButton text={strings.App_Done} onClick={() => setShareOpen(false)} />
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
            title: strings.App_Templates_Start,
            subText:
              questionCount > 0
                ? strings.App_Templates_ReplaceWarning
                : strings.App_Templates_PickStart
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
                <span className={styles.templateName}>{templateName(template, messageBag)}</span>
                <span className={styles.templateMeta}>{templateDescription(template, messageBag)}</span>
                <span className={styles.templateMeta}>
                  {template.key === 'blank'
                    ? strings.App_Templates_Empty
                    : formatString(strings.App_Templates_QuestionCount, {
                        count: String(templateQuestionCount(template))
                      })}
                </span>
              </button>
            ))}
          </div>
          <DialogFooter>
            <DefaultButton text={strings.App_Cancel} onClick={() => setTemplatesOpen(false)} />
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
      <span className={styles.saveChipPending} title={strings.App_SaveChip_SavingTitle}>
        <Icon iconName="Sync" /> {strings.App_SaveChip_Saving}
      </span>
    );
  }
  if (state === 'error') {
    return (
      <span
        className={styles.saveChipUnpublished}
        title={strings.App_SaveChip_FailedTitle}
      >
        <Icon iconName="Warning" /> {strings.App_SaveChip_Failed}
      </span>
    );
  }
  return (
    <span className={styles.saveChipSaved} title={strings.App_SaveChip_SavedTitle}>
      <Icon iconName="CheckMark" /> {strings.App_SaveChip_Saved}
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
