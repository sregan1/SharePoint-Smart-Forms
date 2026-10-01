import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import * as strings from 'SmartFormsWebPartStrings';
import styles from './FormDesigner.module.scss';
import { IFormDefinition, isInputType } from '../../models';
import { allFields, parseFormDefinition } from '../../utils/formUtils';
import { IFormVersionInfo, SharePointService } from '../../services/SharePointService';
import { formatString } from '../../utils/localeUtils';
import { logError } from '../../utils/debug';

export interface IVersionHistoryPanelProps {
  spService: SharePointService;
  /** the web part instance that keys the stored definition */
  instanceId: string;
  /** page UI culture, for date formatting */
  locale?: string;
  /** hands the migrated definition to the designer as an unsaved-to-the-list draft */
  onRestore: (definition: IFormDefinition, versionDate: string) => void;
  onDismiss: () => void;
}

type PreviewState =
  | { status: 'loading' }
  | { status: 'error' }
  | { status: 'ready'; definition: IFormDefinition; questions: number };

const formatDate = (iso: string, locale: string | undefined): string => {
  const parsed = new Date(iso);
  return isNaN(parsed.getTime()) ? iso : parsed.toLocaleString(locale);
};

/**
 * Owner-only list of the saved versions of this form's definition, with a
 * per-version preview (question count) and a confirmed restore. Restoring
 * loads the version, migrates it through the normal parser and hands it to the
 * designer; nothing is published until the owner uses Collect responses.
 */
export const VersionHistoryPanel: React.FunctionComponent<IVersionHistoryPanelProps> = (props) => {
  const { spService, instanceId, locale } = props;
  const [versions, setVersions] = React.useState<IFormVersionInfo[] | undefined>(undefined);
  const [previews, setPreviews] = React.useState<{ [versionId: number]: PreviewState }>({});
  const [confirming, setConfirming] = React.useState<IFormVersionInfo | undefined>(undefined);
  const [restoreError, setRestoreError] = React.useState<string>('');
  const [restoring, setRestoring] = React.useState<boolean>(false);
  const mounted = React.useRef(true);

  React.useEffect(() => {
    mounted.current = true;
    spService
      .listFormVersions(instanceId)
      .then((list) => {
        if (mounted.current) {
          setVersions(list);
        }
      })
      .catch((error) => {
        logError('listFormVersions', error);
        if (mounted.current) {
          setVersions([]);
        }
      });
    return () => {
      mounted.current = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** Loads (and caches) one version, migrated to the current schema. */
  const loadVersion = async (version: IFormVersionInfo): Promise<PreviewState> => {
    setPreviews((prev) => ({ ...prev, [version.versionId]: { status: 'loading' } }));
    let result: PreviewState;
    try {
      const json = await spService.getFormVersion(instanceId, version.versionId);
      const definition = parseFormDefinition(json);
      if (!definition) {
        result = { status: 'error' };
      } else {
        result = {
          status: 'ready',
          definition,
          questions: allFields(definition).filter((f) => isInputType(f.type)).length
        };
      }
    } catch (error) {
      logError('getFormVersion', error);
      result = { status: 'error' };
    }
    if (mounted.current) {
      setPreviews((prev) => ({ ...prev, [version.versionId]: result }));
    }
    return result;
  };

  const startRestore = async (version: IFormVersionInfo): Promise<void> => {
    setRestoreError('');
    const cached = previews[version.versionId];
    const state = cached && cached.status === 'ready' ? cached : await loadVersion(version);
    if (!mounted.current) {
      return;
    }
    if (state.status === 'ready') {
      setConfirming(version);
    } else {
      setRestoreError(strings.Designer_History_LoadFailed);
    }
  };

  const confirmRestore = (): void => {
    if (!confirming) {
      return;
    }
    const state = previews[confirming.versionId];
    if (!state || state.status !== 'ready') {
      return;
    }
    setRestoring(true);
    props.onRestore(state.definition, formatDate(confirming.created, locale));
  };

  const renderRow = (version: IFormVersionInfo, index: number): React.ReactNode => {
    const preview = previews[version.versionId];
    return (
      <div key={version.versionId} className={styles.historyRow}>
        <div className={styles.historyRowHeader}>
          <span>{formatDate(version.created, locale)}</span>
          {index === 0 && <span className={styles.historyMeta}>{strings.Designer_History_Latest}</span>}
        </div>
        <div className={styles.historyMeta}>
          {version.createdBy
            ? formatString(strings.Designer_History_By, { name: version.createdBy })
            : strings.Designer_History_UnknownAuthor}
          {version.label ? ' · ' + formatString(strings.Designer_History_VersionLabel, { label: version.label }) : ''}
        </div>
        {preview && preview.status === 'loading' && <Spinner size={SpinnerSize.small} />}
        {preview && preview.status === 'error' && (
          <div className={styles.historyMeta}>{strings.Designer_History_LoadFailed}</div>
        )}
        {preview && preview.status === 'ready' && (
          <div className={styles.historyMeta}>
            {formatString(
              preview.questions === 1
                ? strings.Designer_History_QuestionsOne
                : strings.Designer_History_QuestionsOther,
              { count: preview.questions }
            )}
          </div>
        )}
        <div className={styles.historyActions}>
          <DefaultButton
            text={strings.Designer_History_Preview}
            disabled={!!preview && preview.status === 'loading'}
            onClick={() => void loadVersion(version)}
          />
          <PrimaryButton
            text={strings.Designer_History_Restore}
            disabled={(!!preview && preview.status === 'loading') || index === 0}
            title={index === 0 ? strings.Designer_History_LatestTitle : undefined}
            onClick={() => void startRestore(version)}
          />
        </div>
      </div>
    );
  };

  const confirmState = confirming ? previews[confirming.versionId] : undefined;

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText={strings.Designer_History_Title}
      onDismiss={props.onDismiss}
      closeButtonAriaLabel={strings.Designer_Common_Cancel}
    >
      <div className={styles.historyList}>
        <p className={styles.panelHint}>{strings.Designer_History_Intro}</p>
        {restoreError && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setRestoreError('')}>
            {restoreError}
          </MessageBar>
        )}
        {versions === undefined && <Spinner size={SpinnerSize.medium} label={strings.Designer_History_Loading} />}
        {versions !== undefined && versions.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>{strings.Designer_History_Empty}</MessageBar>
        )}
        {versions !== undefined && versions.map(renderRow)}
      </div>

      <Dialog
        hidden={!confirming}
        onDismiss={() => setConfirming(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.Designer_History_ConfirmTitle,
          subText:
            confirming && confirmState && confirmState.status === 'ready'
              ? formatString(strings.Designer_History_ConfirmText, {
                  date: formatDate(confirming.created, locale),
                  count: confirmState.questions
                })
              : ''
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text={strings.Designer_History_Restore}
            disabled={restoring}
            onClick={confirmRestore}
          />
          <DefaultButton text={strings.Designer_Common_Cancel} onClick={() => setConfirming(undefined)} />
        </DialogFooter>
      </Dialog>
    </Panel>
  );
};
