import * as React from 'react';
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  Icon,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  TextField,
  Toggle
} from '@fluentui/react';
import * as strings from 'SmartFormsWebPartStrings';
import styles from './FormDesigner.module.scss';
import { FormLayout, IFormSettings } from '../../models';

export interface IFormSettingsPanelProps {
  settings: IFormSettings;
  onSave: (settings: IFormSettings) => void;
  onDismiss: () => void;
}

const ACCENT_COLORS: { name: string; color: string }[] = [
  { name: strings.Designer_Color_Blue, color: '#0078d4' },
  { name: strings.Designer_Color_Teal, color: '#03787c' },
  { name: strings.Designer_Color_Green, color: '#498205' },
  { name: strings.Designer_Color_Purple, color: '#8764b8' },
  { name: strings.Designer_Color_Magenta, color: '#881798' },
  { name: strings.Designer_Color_Red, color: '#d13438' },
  { name: strings.Designer_Color_Orange, color: '#ca5010' },
  { name: strings.Designer_Color_Gray, color: '#69797e' }
];

/** A short, curated icon set — the full Fluent catalogue is unusable as a picker. */
const HEADER_ICONS: string[] = [
  'ClipboardList',
  'Feedback',
  'Survey',
  'PeopleAdd',
  'Calendar',
  'Ticket',
  'Money',
  'ShoppingCart',
  'Health',
  'Shield',
  'Lightbulb',
  'Trophy',
  'Airplane',
  'Home',
  'Suitcase',
  'Education',
  'Repair',
  'Megaphone',
  'Emoji2',
  'CheckList'
];

const toIsoDate = (date: Date | undefined): string | undefined =>
  date ? new Date(date.getTime()).toISOString() : undefined;

const fromIsoDate = (iso: string | undefined): Date | undefined => {
  if (!iso) {
    return undefined;
  }
  const parsed = new Date(iso);
  return isNaN(parsed.getTime()) ? undefined : parsed;
};

const numberOrUndefined = (text: string): number | undefined => {
  const trimmed = (text || '').trim();
  if (!trimmed) {
    return undefined;
  }
  const num = Number(trimmed);
  return isNaN(num) || num <= 0 ? undefined : num;
};

/**
 * Form-wide settings.
 *
 * Grouped into tabs rather than one long scroll with an "Advanced" accordion at
 * the bottom. The old single panel was titled "Style" but also held email
 * notifications, which are not style — the grouping here follows what an owner
 * is actually trying to change.
 */
export const FormSettingsPanel: React.FunctionComponent<IFormSettingsPanelProps> = (props) => {
  const [settings, setSettings] = React.useState<IFormSettings>({ ...props.settings });

  const set = (patch: Partial<IFormSettings>): void => setSettings((prev) => ({ ...prev, ...patch }));

  const layoutOptions: IDropdownOption[] = [
    { key: 'singlePage', text: strings.Designer_Settings_LayoutSinglePage },
    { key: 'wizard', text: strings.Designer_Settings_LayoutWizard }
  ];

  const openDate = fromIsoDate(settings.openDate);
  const closeDate = fromIsoDate(settings.closeDate);
  const datesInverted = !!(openDate && closeDate && closeDate <= openDate);

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText={strings.Designer_Settings_Title}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton text={strings.Designer_Common_Apply} disabled={datesInverted} onClick={() => props.onSave(settings)} />
          <DefaultButton text={strings.Designer_Common_Cancel} onClick={props.onDismiss} />
        </div>
      )}
    >
      <Pivot>
        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText={strings.Designer_Settings_TabBasics} itemIcon="Info">
          <div className={styles.panelBody}>
            <TextField
              label={strings.Designer_Settings_FormTitle}
              value={settings.formTitle}
              onChange={(_e, v) => set({ formTitle: v || '' })}
            />
            <TextField
              label={strings.Designer_Settings_Description}
              multiline={true}
              rows={3}
              value={settings.formDescription || ''}
              onChange={(_e, v) => set({ formDescription: v })}
            />
            <Dropdown
              label={strings.Designer_Settings_Layout}
              options={layoutOptions}
              selectedKey={settings.layout}
              onChange={(_e, option) => option && set({ layout: String(option.key) as FormLayout })}
            />
            <Toggle
              label={strings.Designer_Settings_NumberQuestions}
              checked={settings.showQuestionNumbers !== false}
              onChange={(_e, checked) => set({ showQuestionNumbers: checked !== false })}
            />
            <Toggle
              label={strings.Designer_Settings_ShuffleQuestions}
              checked={settings.shuffleQuestions === true}
              onChange={(_e, checked) => set({ shuffleQuestions: checked === true })}
            />
            {settings.layout === 'wizard' && (
              <Toggle
                label={strings.Designer_Settings_ShowProgressBar}
                checked={settings.showProgressBar}
                onChange={(_e, checked) => set({ showProgressBar: checked === true })}
              />
            )}
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText={strings.Designer_Settings_TabAppearance} itemIcon="Color">
          <div className={styles.panelBody}>
            <span className={styles.swatchLabel}>{strings.Designer_Settings_AccentColor}</span>
            <div className={styles.swatchRow}>
              {ACCENT_COLORS.map((swatch) => (
                <button
                  key={swatch.color}
                  type="button"
                  title={swatch.name}
                  aria-label={swatch.name}
                  aria-pressed={settings.accentColor === swatch.color}
                  className={settings.accentColor === swatch.color ? styles.swatchSelected : styles.swatch}
                  style={{ backgroundColor: swatch.color }}
                  onClick={() => set({ accentColor: swatch.color })}
                />
              ))}
              <input
                type="color"
                className={styles.swatchCustom}
                title={strings.Designer_Settings_CustomColorTitle}
                aria-label={strings.Designer_Settings_CustomColorAria}
                value={settings.accentColor || '#0078d4'}
                onChange={(e) => set({ accentColor: e.target.value })}
              />
            </div>
            <p className={styles.panelHint}>
              {strings.Designer_Settings_AccentHint}
            </p>

            <Toggle
              label={strings.Designer_Settings_ShowHeader}
              checked={settings.showFormHeader}
              onChange={(_e, checked) => set({ showFormHeader: checked === true })}
            />

            {settings.showFormHeader && (
              <>
                <span className={styles.swatchLabel}>{strings.Designer_Settings_HeaderIcon}</span>
                <div className={styles.iconGrid} role="group" aria-label={strings.Designer_Settings_HeaderIcon}>
                  {HEADER_ICONS.map((icon) => (
                    <button
                      key={icon}
                      type="button"
                      title={icon}
                      aria-label={icon}
                      aria-pressed={settings.headerIcon === icon}
                      className={settings.headerIcon === icon ? styles.iconOptionSelected : styles.iconOption}
                      onClick={() => set({ headerIcon: icon })}
                    >
                      <Icon iconName={icon} />
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText={strings.Designer_Settings_TabAfterSubmit} itemIcon="CheckMark">
          <div className={styles.panelBody}>
            <TextField
              label={strings.Designer_Settings_SubmitButtonText}
              value={settings.submitButtonText}
              placeholder={strings.Designer_Settings_SubmitPlaceholder}
              onChange={(_e, v) => set({ submitButtonText: v || 'Submit' })}
            />
            <TextField
              label={strings.Designer_Settings_ThankYouTitle}
              value={settings.confirmationTitle}
              placeholder={strings.Designer_Settings_ThankYouPlaceholder}
              onChange={(_e, v) => set({ confirmationTitle: v || 'Thank you!' })}
            />
            <TextField
              label={strings.Designer_Settings_ThankYouMessage}
              multiline={true}
              rows={3}
              value={settings.confirmationMessage}
              onChange={(_e, v) => set({ confirmationMessage: v || '' })}
            />
            <Toggle
              label={strings.Designer_Settings_ShowSubmitAnother}
              checked={settings.allowAnotherResponse}
              onChange={(_e, checked) => set({ allowAnotherResponse: checked === true })}
            />
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText={strings.Designer_Settings_TabNotifications} itemIcon="Mail">
          <div className={styles.panelBody}>
            <TextField
              label={strings.Designer_Settings_EmailEachResponseTo}
              placeholder={strings.Designer_Settings_EmailPlaceholder}
              description={strings.Designer_Settings_EmailDescription}
              value={settings.notifyEmails || ''}
              onChange={(_e, v) => set({ notifyEmails: v })}
            />
            <Toggle
              label={strings.Designer_Settings_RespondentReceipt}
              checked={settings.respondentReceipt === true}
              onChange={(_e, checked) => set({ respondentReceipt: checked === true })}
            />
            <Toggle
              label={strings.Designer_Settings_EnableApproval}
              checked={settings.enableApproval === true}
              onChange={(_e, checked) =>
                set(
                  checked === true
                    ? { enableApproval: true }
                    : { enableApproval: false, approvalNotify: false }
                )
              }
            />
            <Toggle
              label={strings.Designer_Settings_ApprovalNotify}
              checked={settings.enableApproval === true && settings.approvalNotify === true}
              disabled={settings.enableApproval !== true}
              onChange={(_e, checked) => set({ approvalNotify: checked === true })}
            />
            <p className={styles.panelHint}>{strings.Designer_Settings_ApprovalHint}</p>
            <MessageBar messageBarType={MessageBarType.info}>
              {strings.Designer_Settings_NotificationsInfo}
            </MessageBar>
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText={strings.Designer_Settings_TabAccess} itemIcon="Lock">
          <div className={styles.panelBody}>
            <DatePicker
              label={strings.Designer_Settings_OpenFrom}
              value={openDate}
              placeholder={strings.Designer_Settings_OpenImmediately}
              allowTextInput={true}
              onSelectDate={(date) => set({ openDate: toIsoDate(date || undefined) })}
            />
            <DatePicker
              label={strings.Designer_Settings_CloseAfter}
              value={closeDate}
              placeholder={strings.Designer_Settings_NeverClose}
              allowTextInput={true}
              onSelectDate={(date) => set({ closeDate: toIsoDate(date || undefined) })}
            />
            {(openDate || closeDate) && (
              <DefaultButton
                iconProps={{ iconName: 'Clear' }}
                text={strings.Designer_Settings_ClearDates}
                onClick={() => set({ openDate: undefined, closeDate: undefined })}
              />
            )}
            {datesInverted && (
              <MessageBar messageBarType={MessageBarType.error}>
                {strings.Designer_Settings_DatesInverted}
              </MessageBar>
            )}

            <TextField
              label={strings.Designer_Settings_MaxResponses}
              value={settings.maxResponses === undefined ? '' : String(settings.maxResponses)}
              placeholder={strings.Designer_Settings_NoLimit}
              onChange={(_e, v) => set({ maxResponses: numberOrUndefined(v) })}
            />
            <Toggle
              label={strings.Designer_Settings_OneResponsePerPerson}
              checked={settings.oneResponsePerPerson === true}
              onChange={(_e, checked) => set({ oneResponsePerPerson: checked === true })}
            />
            <Toggle
              label={strings.Designer_Settings_AllowSaveDraft}
              checked={settings.allowSaveDraft === true}
              onChange={(_e, checked) => set({ allowSaveDraft: checked === true })}
            />
            <Toggle
              label={strings.Designer_Settings_AllowEdit}
              checked={settings.allowEdit === true}
              onChange={(_e, checked) => set({ allowEdit: checked === true })}
            />
            <p className={styles.panelHint}>{strings.Designer_Settings_AllowEditHint}</p>
            <TextField
              label={strings.Designer_Settings_ClosedMessage}
              multiline={true}
              rows={2}
              value={settings.closedMessage || ''}
              placeholder={strings.Designer_Settings_ClosedMessagePlaceholder}
              onChange={(_e, v) => set({ closedMessage: v })}
            />
            <MessageBar messageBarType={MessageBarType.info}>
              {strings.Designer_Settings_AccessInfo}
            </MessageBar>
          </div>
        </PivotItem>
      </Pivot>
    </Panel>
  );
};
