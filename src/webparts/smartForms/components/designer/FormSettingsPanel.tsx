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
import styles from './FormDesigner.module.scss';
import { FormLayout, IFormSettings } from '../../models';

export interface IFormSettingsPanelProps {
  settings: IFormSettings;
  onSave: (settings: IFormSettings) => void;
  onDismiss: () => void;
}

const ACCENT_COLORS: { name: string; color: string }[] = [
  { name: 'Blue', color: '#0078d4' },
  { name: 'Teal', color: '#03787c' },
  { name: 'Green', color: '#498205' },
  { name: 'Purple', color: '#8764b8' },
  { name: 'Magenta', color: '#881798' },
  { name: 'Red', color: '#d13438' },
  { name: 'Orange', color: '#ca5010' },
  { name: 'Gray', color: '#69797e' }
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
    { key: 'singlePage', text: 'All questions on one page' },
    { key: 'wizard', text: 'One section per step (wizard)' }
  ];

  const openDate = fromIsoDate(settings.openDate);
  const closeDate = fromIsoDate(settings.closeDate);
  const datesInverted = !!(openDate && closeDate && closeDate <= openDate);

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText="Form settings"
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton text="Apply" disabled={datesInverted} onClick={() => props.onSave(settings)} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <Pivot>
        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText="Basics" itemIcon="Info">
          <div className={styles.panelBody}>
            <TextField
              label="Form title"
              value={settings.formTitle}
              onChange={(_e, v) => set({ formTitle: v || '' })}
            />
            <TextField
              label="Description"
              multiline={true}
              rows={3}
              value={settings.formDescription || ''}
              onChange={(_e, v) => set({ formDescription: v })}
            />
            <Dropdown
              label="Layout"
              options={layoutOptions}
              selectedKey={settings.layout}
              onChange={(_e, option) => option && set({ layout: String(option.key) as FormLayout })}
            />
            <Toggle
              label="Number the questions"
              checked={settings.showQuestionNumbers !== false}
              onChange={(_e, checked) => set({ showQuestionNumbers: checked !== false })}
            />
            <Toggle
              label="Shuffle question order for each respondent"
              checked={settings.shuffleQuestions === true}
              onChange={(_e, checked) => set({ shuffleQuestions: checked === true })}
            />
            {settings.layout === 'wizard' && (
              <Toggle
                label="Show a progress bar"
                checked={settings.showProgressBar}
                onChange={(_e, checked) => set({ showProgressBar: checked === true })}
              />
            )}
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText="Appearance" itemIcon="Color">
          <div className={styles.panelBody}>
            <span className={styles.swatchLabel}>Accent color</span>
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
                title="Custom color"
                aria-label="Custom accent color"
                value={settings.accentColor || '#0078d4'}
                onChange={(e) => set({ accentColor: e.target.value })}
              />
            </div>
            <p className={styles.panelHint}>
              Accent colors are automatically adjusted for contrast on dark sites, so the form stays
              readable whichever theme the page uses.
            </p>

            <Toggle
              label="Show the form header"
              checked={settings.showFormHeader}
              onChange={(_e, checked) => set({ showFormHeader: checked === true })}
            />

            {settings.showFormHeader && (
              <>
                <span className={styles.swatchLabel}>Header icon</span>
                <div className={styles.iconGrid} role="group" aria-label="Header icon">
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
        <PivotItem headerText="After submit" itemIcon="CheckMark">
          <div className={styles.panelBody}>
            <TextField
              label="Submit button text"
              value={settings.submitButtonText}
              placeholder="Submit"
              onChange={(_e, v) => set({ submitButtonText: v || 'Submit' })}
            />
            <TextField
              label="Thank-you title"
              value={settings.confirmationTitle}
              placeholder="Thank you!"
              onChange={(_e, v) => set({ confirmationTitle: v || 'Thank you!' })}
            />
            <TextField
              label="Thank-you message"
              multiline={true}
              rows={3}
              value={settings.confirmationMessage}
              onChange={(_e, v) => set({ confirmationMessage: v || '' })}
            />
            <Toggle
              label='Show a "Submit another response" button'
              checked={settings.allowAnotherResponse}
              onChange={(_e, checked) => set({ allowAnotherResponse: checked === true })}
            />
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText="Notifications" itemIcon="Mail">
          <div className={styles.panelBody}>
            <TextField
              label="Email each response to"
              placeholder="name@company.com; team@company.com"
              description="People in your organization get every response by email, with a link to the list item"
              value={settings.notifyEmails || ''}
              onChange={(_e, v) => set({ notifyEmails: v })}
            />
            <Toggle
              label="Email respondents a copy of their answers"
              checked={settings.respondentReceipt === true}
              onChange={(_e, checked) => set({ respondentReceipt: checked === true })}
            />
            <MessageBar messageBarType={MessageBarType.info}>
              Notifications use SharePoint&apos;s own mail service, so recipients must be users in your
              organization. External addresses are silently dropped.
            </MessageBar>
          </div>
        </PivotItem>

        {/* ----------------------------------------------------------------- */}
        <PivotItem headerText="Access" itemIcon="Lock">
          <div className={styles.panelBody}>
            <DatePicker
              label="Open from"
              value={openDate}
              placeholder="Open immediately"
              allowTextInput={true}
              onSelectDate={(date) => set({ openDate: toIsoDate(date || undefined) })}
            />
            <DatePicker
              label="Close after"
              value={closeDate}
              placeholder="Never close"
              allowTextInput={true}
              onSelectDate={(date) => set({ closeDate: toIsoDate(date || undefined) })}
            />
            {(openDate || closeDate) && (
              <DefaultButton
                iconProps={{ iconName: 'Clear' }}
                text="Clear dates"
                onClick={() => set({ openDate: undefined, closeDate: undefined })}
              />
            )}
            {datesInverted && (
              <MessageBar messageBarType={MessageBarType.error}>
                The close date is on or before the open date, so the form would never accept responses.
              </MessageBar>
            )}

            <TextField
              label="Stop after this many responses"
              value={settings.maxResponses === undefined ? '' : String(settings.maxResponses)}
              placeholder="No limit"
              onChange={(_e, v) => set({ maxResponses: numberOrUndefined(v) })}
            />
            <Toggle
              label="Only allow one response per person"
              checked={settings.oneResponsePerPerson === true}
              onChange={(_e, checked) => set({ oneResponsePerPerson: checked === true })}
            />
            <Toggle
              label="Let respondents save a draft and finish later"
              checked={settings.allowSaveDraft === true}
              onChange={(_e, checked) => set({ allowSaveDraft: checked === true })}
            />
            <TextField
              label="Message shown when the form is closed"
              multiline={true}
              rows={2}
              value={settings.closedMessage || ''}
              placeholder="This form is no longer accepting responses."
              onChange={(_e, v) => set({ closedMessage: v })}
            />
            <MessageBar messageBarType={MessageBarType.info}>
              These limits are enforced by the form, not by SharePoint. Anyone with permission to add
              items to the response list could still add one directly through the list.
            </MessageBar>
          </div>
        </PivotItem>
      </Pivot>
    </Panel>
  );
};
