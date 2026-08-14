import * as React from 'react';
import {
  ActionButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Icon,
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
import {
  ConditionMatch,
  ConditionOperator,
  ContentStyle,
  FieldType,
  ICondition,
  IFormDefinition,
  IFormField,
  IListColumnInfo,
  IListInfo,
  isInputType,
  NumberFormat,
  RatingIcon,
  ScaleAnalytics,
  UNARY_OPERATORS
} from '../../models';
import {
  allFields,
  findField,
  formatValue,
  inputFields,
  OPERATOR_LABELS,
  operatorsForField
} from '../../utils/formUtils';
import { evaluateFormula, formulaReferences } from '../../utils/formula';
import { CURRENCY_OPTIONS } from '../../utils/spFieldXml';
import { SharePointService } from '../../services/SharePointService';

export interface IFieldEditorPanelProps {
  field: IFormField;
  definition: IFormDefinition;
  spService: SharePointService;
  onSave: (field: IFormField) => void;
  onDismiss: () => void;
}

const numberOrUndefined = (text: string): number | undefined => {
  const trimmed = (text || '').trim();
  if (!trimmed) {
    return undefined;
  }
  const num = Number(trimmed);
  return isNaN(num) ? undefined : num;
};

const CONTENT_STYLES: IDropdownOption[] = [
  { key: 'info', text: 'Info (accent)' },
  { key: 'success', text: 'Success (green)' },
  { key: 'warning', text: 'Warning (amber)' },
  { key: 'text', text: 'Plain text' },
  { key: 'divider', text: 'Divider line' }
];

const RATING_ICONS: IDropdownOption[] = [
  { key: 'star', text: 'Stars' },
  { key: 'heart', text: 'Hearts' },
  { key: 'like', text: 'Thumbs up' }
];

const NUMBER_FORMATS: IDropdownOption[] = [
  { key: 'plain', text: 'Plain number' },
  { key: 'currency', text: 'Currency' },
  { key: 'percent', text: 'Percentage' }
];

/**
 * "Branching, validation & more" panel. The essentials (label, type, options,
 * required, width) are edited inline on the question card — this panel holds
 * branching rules, validation, and the long tail of per-type options.
 */
export const FieldEditorPanel: React.FunctionComponent<IFieldEditorPanelProps> = (props) => {
  const [field, setField] = React.useState<IFormField>(() => JSON.parse(JSON.stringify(props.field)));
  const [lists, setLists] = React.useState<IListInfo[]>([]);
  const [lookupColumns, setLookupColumns] = React.useState<IListColumnInfo[]>([]);

  const set = (patch: Partial<IFormField>): void => setField((prev) => ({ ...prev, ...patch }));

  const otherFields = allFields(props.definition).filter(
    (f) => f.id !== field.id && isInputType(f.type)
  );

  // ----- lookup source loading -----

  const needsLists = field.type === FieldType.Lookup;

  React.useEffect(() => {
    let cancelled = false;
    if (!needsLists) {
      return undefined;
    }
    props.spService
      .getAvailableLists()
      .then((available) => {
        if (!cancelled) {
          setLists(available);
        }
      })
      .catch((): undefined => undefined);
    return () => {
      cancelled = true;
    };
  }, [needsLists]);

  React.useEffect(() => {
    let cancelled = false;
    if (!needsLists || !field.lookupListId) {
      setLookupColumns([]);
      return undefined;
    }
    props.spService
      .getListColumns(field.lookupListId)
      .then((columns) => {
        if (!cancelled) {
          setLookupColumns(columns);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setLookupColumns([]);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [needsLists, field.lookupListId]);

  // ----- branching -----

  const conditions = field.visibleWhen ? field.visibleWhen.conditions : [];
  const match: ConditionMatch = field.visibleWhen ? field.visibleWhen.match : 'all';

  const setConditions = (nextConditions: ICondition[], nextMatch?: ConditionMatch): void => {
    if (nextConditions.length === 0) {
      set({ visibleWhen: undefined });
      return;
    }
    set({ visibleWhen: { match: nextMatch || match, conditions: nextConditions } });
  };

  const addCondition = (): void => {
    if (otherFields.length === 0) {
      return;
    }
    const driver = otherFields[0];
    setConditions(
      conditions.concat([
        { fieldId: driver.id, operator: operatorsForField(driver)[0] || 'equals', value: '' }
      ])
    );
  };

  const patchCondition = (index: number, patch: Partial<ICondition>): void => {
    setConditions(conditions.map((c, i) => (i === index ? { ...c, ...patch } : c)));
  };

  const removeCondition = (index: number): void => {
    setConditions(conditions.filter((_c, i) => i !== index));
  };

  /** Plain-English restatement of the rule, so the owner can sanity-check it. */
  const rulePreview = (): string => {
    if (conditions.length === 0) {
      return '';
    }
    const parts = conditions.map((condition) => {
      const driver = findField(props.definition, condition.fieldId);
      const label = OPERATOR_LABELS.filter((o) => o.key === condition.operator)[0];
      const operatorText = label ? label.text : condition.operator;
      const name = driver ? driver.title || 'Untitled question' : 'a deleted question';
      if (UNARY_OPERATORS.indexOf(condition.operator) !== -1) {
        return '"' + name + '" ' + operatorText;
      }
      if (condition.operator === 'between') {
        return '"' + name + '" ' + operatorText + ' ' + (condition.value || '?') + ' and ' + (condition.value2 || '?');
      }
      return '"' + name + '" ' + operatorText + ' "' + (condition.value || '') + '"';
    });
    const joiner = match === 'any' ? ' OR ' : ' AND ';
    return 'Shown when ' + parts.join(joiner) + '.';
  };

  /** Options for a rule's value box when the driver is a choice question. */
  const valueOptionsFor = (driver: IFormField | undefined): IDropdownOption[] | undefined => {
    if (!driver) {
      return undefined;
    }
    if (driver.type === FieldType.YesNo || driver.type === FieldType.Consent) {
      return [
        { key: 'Yes', text: 'Yes' },
        { key: 'No', text: 'No' }
      ];
    }
    if (driver.type === FieldType.Choice) {
      const choices = (driver.choices || []).map((c) => (c || '').trim()).filter((c) => c.length > 0);
      return choices.length > 0 ? choices.map((c) => ({ key: c, text: c })) : undefined;
    }
    if (driver.type === FieldType.ImageChoice) {
      const labels = (driver.imageChoices || []).map((o) => o.label).filter((l) => l);
      return labels.length > 0 ? labels.map((c) => ({ key: c, text: c })) : undefined;
    }
    return undefined;
  };

  const renderCondition = (condition: ICondition, index: number): React.ReactNode => {
    const driver = findField(props.definition, condition.fieldId);
    const allowedOperators = operatorsForField(driver);
    const operatorOptions = OPERATOR_LABELS.filter((o) => allowedOperators.indexOf(o.key) !== -1).map(
      (o) => ({ key: o.key, text: o.text })
    );
    const needsValue = UNARY_OPERATORS.indexOf(condition.operator) === -1;
    const valueOptions = valueOptionsFor(driver);

    return (
      <div key={index} className={styles.ruleRow}>
        <div className={styles.ruleRowFields}>
          <Dropdown
            label={index === 0 ? 'When' : undefined}
            options={otherFields.map((f) => ({ key: f.id, text: f.title || 'Untitled question' }))}
            selectedKey={condition.fieldId}
            onChange={(_e, option) => {
              if (!option) {
                return;
              }
              const nextDriver = findField(props.definition, String(option.key));
              const allowed = operatorsForField(nextDriver);
              // an operator that made sense for the old driver may be invalid now
              const operator =
                allowed.indexOf(condition.operator) !== -1 ? condition.operator : allowed[0] || 'equals';
              patchCondition(index, { fieldId: String(option.key), operator, value: '' });
            }}
          />
          <Dropdown
            options={operatorOptions}
            selectedKey={condition.operator}
            onChange={(_e, option) =>
              option && patchCondition(index, { operator: option.key as ConditionOperator })
            }
          />
          {needsValue &&
            (valueOptions ? (
              <Dropdown
                options={valueOptions}
                selectedKey={condition.value || null}
                placeholder="Select a value"
                onChange={(_e, option) => option && patchCondition(index, { value: String(option.key) })}
              />
            ) : (
              <TextField
                value={condition.value || ''}
                placeholder={
                  driver && (driver.type === FieldType.Date || driver.type === FieldType.Time)
                    ? 'YYYY-MM-DD or "today"'
                    : 'Value'
                }
                onChange={(_e, v) => patchCondition(index, { value: v })}
              />
            ))}
          {condition.operator === 'between' && (
            <TextField
              value={condition.value2 || ''}
              placeholder="Upper bound"
              onChange={(_e, v) => patchCondition(index, { value2: v })}
            />
          )}
        </div>
        <DefaultButton
          iconProps={{ iconName: 'Delete' }}
          title="Remove this condition"
          ariaLabel="Remove this condition"
          onClick={() => removeCondition(index)}
        />
      </div>
    );
  };

  // ----- formula validation -----

  const formulaProblem = (): string => {
    if (field.type !== FieldType.Calculated || !(field.formula || '').trim()) {
      return '';
    }
    if (evaluateFormula(field.formula || '', () => 1) === undefined) {
      return 'This formula cannot be worked out. Check the brackets and operators.';
    }
    const missing = formulaReferences(field.formula || '').filter(
      (name) => !allFields(props.definition).filter((f) => (f.title || '').trim().toLowerCase() === name.toLowerCase())[0]
    );
    if (missing.length > 0) {
      return 'These questions do not exist: ' + missing.join(', ') + '.';
    }
    return '';
  };

  const formulaPreview = (): string => {
    if (field.type !== FieldType.Calculated) {
      return '';
    }
    const result = evaluateFormula(field.formula || '', (name) => {
      const referenced = allFields(props.definition).filter(
        (f) => (f.title || '').trim().toLowerCase() === name.toLowerCase()
      )[0];
      // preview against 10 so the shape of the result is visible
      return referenced ? 10 : undefined;
    });
    return result === undefined ? '' : 'With every input at 10, this shows ' + formatValue(field, result) + '.';
  };

  // ----- per-type capability flags -----

  const isNumeric = field.type === FieldType.Number || field.type === FieldType.Calculated;

  const supportsPlaceholder =
    [
      FieldType.YesNo,
      FieldType.Rating,
      FieldType.Slider,
      FieldType.Ranking,
      FieldType.Scale,
      FieldType.Likert,
      FieldType.Signature,
      FieldType.FileUpload,
      FieldType.Calculated,
      FieldType.Content,
      FieldType.Address
    ].indexOf(field.type) === -1;

  const supportsDefault =
    [
      FieldType.Person,
      FieldType.RichText,
      FieldType.Ranking,
      FieldType.Likert,
      FieldType.Signature,
      FieldType.FileUpload,
      FieldType.Calculated,
      FieldType.Content,
      FieldType.Address,
      FieldType.ImageChoice
    ].indexOf(field.type) === -1;

  const supportsPattern =
    [FieldType.Text, FieldType.MultilineText, FieldType.Email, FieldType.Phone].indexOf(field.type) !== -1;

  const supportsChoiceExtras =
    field.type === FieldType.Choice ||
    field.type === FieldType.ImageChoice ||
    field.type === FieldType.Lookup;

  const defaultValueHint = (): string => {
    switch (field.type) {
      case FieldType.YesNo:
      case FieldType.Consent:
        return 'Use "yes" or "no"';
      case FieldType.Choice:
      case FieldType.Lookup:
        return field.allowMultiple ? 'Separate multiple defaults with ;' : 'Must match one of the options';
      case FieldType.Date:
        return 'Use "today", or a date like 2026-03-01';
      case FieldType.Time:
        return 'Use "now", or a time like 09:30';
      default:
        return '';
    }
  };

  const handleSave = (): void => {
    const result: IFormField = { ...field };
    if (result.visibleWhen && (!result.visibleWhen.conditions || result.visibleWhen.conditions.length === 0)) {
      delete result.visibleWhen;
    }
    props.onSave(result);
  };

  const isContent = !isInputType(field.type);

  return (
    <Panel
      isOpen={true}
      type={PanelType.medium}
      headerText={field.title || 'Question settings'}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton text="Apply" onClick={handleSave} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className={styles.panelBody}>
        {/* ----- Content blocks have their own short form ----- */}
        {isContent ? (
          <>
            <h4 className={styles.panelSectionTitle}>Appearance</h4>
            <Dropdown
              label="Style"
              options={CONTENT_STYLES}
              selectedKey={field.contentStyle || 'info'}
              onChange={(_e, option) =>
                option && set({ contentStyle: String(option.key) as ContentStyle })
              }
            />
            <TextField
              label="Image URL (optional)"
              value={field.contentImageUrl || ''}
              placeholder="https://…"
              onChange={(_e, v) => set({ contentImageUrl: v })}
            />
            <TextField
              label="Content"
              multiline={true}
              rows={5}
              value={field.contentHtml || ''}
              description="Basic formatting is allowed: bold, italic, lists and links."
              onChange={(_e, v) => set({ contentHtml: v })}
            />
          </>
        ) : (
          <>
            <Pivot aria-label="Question settings" defaultSelectedKey="options">
              {/* ----- Per-type options ----- */}
              <PivotItem headerText="Options" itemKey="options">
                <div className={styles.panelTabContent}>{renderOptionsTab()}</div>
              </PivotItem>

              {/* ----- Branching ----- */}
              <PivotItem headerText="Branching" itemKey="branching">
                <div className={styles.panelTabContent}>
                  {otherFields.length === 0 ? (
                    <p className={styles.panelHint}>
                      Add another question first — branching needs something to depend on.
                    </p>
                  ) : (
                    <>
                      <Toggle
                        label="Only show this question when other answers match"
                        checked={conditions.length > 0}
                        onChange={(_e, checked) => {
                          if (checked) {
                            addCondition();
                          } else {
                            setConditions([]);
                          }
                        }}
                      />
                      {conditions.length > 0 && (
                        <div className={styles.ruleGroup}>
                          {conditions.length > 1 && (
                            <div className={styles.ruleMatchRow}>
                              <span>Match</span>
                              <Dropdown
                                className={styles.ruleMatchDropdown}
                                options={[
                                  { key: 'all', text: 'all of these' },
                                  { key: 'any', text: 'any of these' }
                                ]}
                                selectedKey={match}
                                onChange={(_e, option) =>
                                  option && setConditions(conditions, option.key as ConditionMatch)
                                }
                              />
                            </div>
                          )}
                          {conditions.map(renderCondition)}
                          <ActionButton
                            iconProps={{ iconName: 'Add' }}
                            text="Add condition"
                            onClick={addCondition}
                          />
                          <div className={styles.rulePreview}>{rulePreview()}</div>
                        </div>
                      )}
                    </>
                  )}
                </div>
              </PivotItem>

              {/* ----- Help & prompts, validation ----- */}
              <PivotItem headerText="Validation" itemKey="validation">
                <div className={styles.panelTabContent}>
                  <TextField
                    label="Help text"
                    value={field.description || ''}
                    placeholder="Shown under the question label"
                    onChange={(_e, v) => set({ description: v })}
                  />
                  {supportsPlaceholder && (
                    <TextField
                      label={field.type === FieldType.Consent ? 'Checkbox label' : 'Placeholder'}
                      value={field.placeholder || ''}
                      onChange={(_e, v) => set({ placeholder: v })}
                    />
                  )}
                  {supportsDefault && (
                    <TextField
                      label="Default answer"
                      value={field.defaultValue || ''}
                      description={defaultValueHint()}
                      onChange={(_e, v) => set({ defaultValue: v })}
                    />
                  )}
                  <TextField
                    label="Custom required message"
                    value={field.requiredMessage || ''}
                    placeholder={(field.title || 'This question') + ' is required'}
                    description="Shown instead of the standard message when this question is left blank"
                    disabled={field.required !== true}
                    onChange={(_e, v) => set({ requiredMessage: v })}
                  />
                  <Toggle
                    label="Read-only (answer can only come from a prefilled link)"
                    checked={field.readOnly === true}
                    onChange={(_e, checked) => set({ readOnly: checked === true })}
                  />
                  {(supportsPattern ||
                    isNumeric ||
                    field.type === FieldType.Text ||
                    field.type === FieldType.MultilineText) && (
                    <>
                      <h4 className={styles.panelSectionTitle}>Validation rules</h4>
                      {(field.type === FieldType.Text || field.type === FieldType.MultilineText) && (
                        <TextField
                          label="Maximum length"
                          value={field.maxLength === undefined ? '' : String(field.maxLength)}
                          placeholder="255"
                          onChange={(_e, v) => set({ maxLength: numberOrUndefined(v) })}
                        />
                      )}
                      {supportsPattern && (
                        <>
                          <TextField
                            label="Must match this pattern"
                            value={field.pattern || ''}
                            placeholder="^[A-Z]{2}\d{4}$"
                            description="A regular expression. Leave blank for no pattern check."
                            onChange={(_e, v) => set({ pattern: v })}
                          />
                          <TextField
                            label="Message when the pattern doesn't match"
                            value={field.patternMessage || ''}
                            placeholder="This answer is not in the expected format"
                            disabled={!field.pattern}
                            onChange={(_e, v) => set({ patternMessage: v })}
                          />
                        </>
                      )}
                    </>
                  )}
                </div>
              </PivotItem>
            </Pivot>
          </>
        )}
      </div>
    </Panel>
  );

  /** The per-type "Options" tab: the long tail of field-specific settings. */
  function renderOptionsTab(): React.ReactNode {
    return (
      <>
            {isNumeric && (
              <>
                <Dropdown
                  label="Format"
                  options={NUMBER_FORMATS}
                  selectedKey={field.numberFormat || 'plain'}
                  onChange={(_e, option) =>
                    option && set({ numberFormat: String(option.key) as NumberFormat })
                  }
                />
                {field.numberFormat === 'currency' && (
                  <Dropdown
                    label="Currency"
                    options={CURRENCY_OPTIONS.map((c) => ({
                      key: c.symbol,
                      text: c.symbol + '  ' + c.label
                    }))}
                    selectedKey={field.currencySymbol || '$'}
                    onChange={(_e, option) => option && set({ currencySymbol: String(option.key) })}
                  />
                )}
                {field.numberFormat === 'plain' && (
                  <TextField
                    label="Unit shown after the value"
                    value={field.unitSuffix || ''}
                    placeholder="kg, hrs, items…"
                    onChange={(_e, v) => set({ unitSuffix: v })}
                  />
                )}
                <div className={styles.inlineFields}>
                  <TextField
                    label="Minimum"
                    value={field.min === undefined ? '' : String(field.min)}
                    onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                  />
                  <TextField
                    label="Maximum"
                    value={field.max === undefined ? '' : String(field.max)}
                    onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                  />
                  <TextField
                    label="Decimals"
                    value={field.decimalPlaces === undefined ? '' : String(field.decimalPlaces)}
                    onChange={(_e, v) => set({ decimalPlaces: numberOrUndefined(v) })}
                  />
                </div>
              </>
            )}

            {field.type === FieldType.Calculated && (
              <>
                <TextField
                  label="Formula"
                  value={field.formula || ''}
                  placeholder="{Quantity} * {Unit price}"
                  multiline={true}
                  rows={2}
                  description='Reference other questions by their text in braces. Supports + - * / ( ) and round, min, max, sum, avg.'
                  onChange={(_e, v) => set({ formula: v })}
                />
                {formulaProblem() ? (
                  <MessageBar messageBarType={MessageBarType.error}>{formulaProblem()}</MessageBar>
                ) : (
                  formulaPreview() && (
                    <MessageBar messageBarType={MessageBarType.success}>{formulaPreview()}</MessageBar>
                  )
                )}
                <p className={styles.panelHint}>
                  Available questions:{' '}
                  {inputFields(props.definition)
                    .filter((f) => f.id !== field.id && (f.title || '').trim())
                    .map((f) => '{' + f.title + '}')
                    .join(', ') || 'none yet'}
                </p>
              </>
            )}

            {field.type === FieldType.Slider && (
              <div className={styles.inlineFields}>
                <TextField
                  label="Minimum"
                  value={field.min === undefined ? '0' : String(field.min)}
                  onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                />
                <TextField
                  label="Maximum"
                  value={field.max === undefined ? '10' : String(field.max)}
                  onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                />
                <TextField
                  label="Step"
                  value={field.step === undefined ? '1' : String(field.step)}
                  onChange={(_e, v) => set({ step: numberOrUndefined(v) })}
                />
              </div>
            )}

            {field.type === FieldType.Scale && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label="From"
                    value={field.min === undefined ? '1' : String(field.min)}
                    onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                  />
                  <TextField
                    label="To"
                    value={field.max === undefined ? '5' : String(field.max)}
                    onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                  />
                  <TextField
                    label="Step"
                    value={field.step === undefined ? '1' : String(field.step)}
                    onChange={(_e, v) => set({ step: numberOrUndefined(v) })}
                  />
                </div>
                <div className={styles.inlineFields}>
                  <TextField
                    label="Label at the low end"
                    value={field.lowLabel || ''}
                    placeholder="Not at all likely"
                    onChange={(_e, v) => set({ lowLabel: v })}
                  />
                  <TextField
                    label="Label at the high end"
                    value={field.highLabel || ''}
                    placeholder="Extremely likely"
                    onChange={(_e, v) => set({ highLabel: v })}
                  />
                </div>
                <Dropdown
                  label="Summarize as"
                  options={[
                    { key: 'average', text: 'Average and distribution' },
                    { key: 'nps', text: 'Net Promoter Score (needs a 0–10 scale)' }
                  ]}
                  selectedKey={field.scaleAnalytics || 'average'}
                  onChange={(_e, option) =>
                    option && set({ scaleAnalytics: String(option.key) as ScaleAnalytics })
                  }
                />
                {field.scaleAnalytics === 'nps' && (field.min !== 0 || field.max !== 10) && (
                  <MessageBar messageBarType={MessageBarType.warning}>
                    Net Promoter Score is only meaningful on a 0–10 scale. Set From to 0 and To to 10.
                  </MessageBar>
                )}
              </>
            )}

            {field.type === FieldType.Rating && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label="Number of icons"
                    value={String(field.maxRating || 5)}
                    onChange={(_e, v) => set({ maxRating: numberOrUndefined(v) })}
                  />
                  <Dropdown
                    label="Icon"
                    options={RATING_ICONS}
                    selectedKey={field.ratingIcon || 'star'}
                    onChange={(_e, option) =>
                      option && set({ ratingIcon: String(option.key) as RatingIcon })
                    }
                  />
                </div>
                <Toggle
                  label="Allow half values"
                  checked={field.allowHalfRating === true}
                  onChange={(_e, checked) => set({ allowHalfRating: checked === true })}
                />
              </>
            )}

            {supportsChoiceExtras && (
              <>
                {field.type === FieldType.Choice && (
                  <>
                    <Dropdown
                      label="Show options as"
                      options={[
                        { key: 'buttons', text: 'Inline buttons' },
                        { key: 'dropdown', text: 'Dropdown' }
                      ]}
                      selectedKey={field.choiceDisplay || undefined}
                      placeholder="Automatic (inline up to 6 options)"
                      onChange={(_e, option) =>
                        option && set({ choiceDisplay: option.key === 'buttons' ? 'buttons' : 'dropdown' })
                      }
                    />
                    <Toggle
                      label='Add a write-in "Other" option'
                      checked={field.allowOther === true}
                      onChange={(_e, checked) => set({ allowOther: checked === true })}
                    />
                    {field.allowOther && (
                      <TextField
                        label='Label for the "Other" option'
                        value={field.otherLabel || ''}
                        placeholder="Other"
                        onChange={(_e, v) => set({ otherLabel: v })}
                      />
                    )}
                  </>
                )}
                <Toggle
                  label="Shuffle the option order for each respondent"
                  checked={field.shuffleOptions === true}
                  onChange={(_e, checked) => set({ shuffleOptions: checked === true })}
                />
                {field.allowMultiple && (
                  <div className={styles.inlineFields}>
                    <TextField
                      label="Minimum selections"
                      value={field.minSelections === undefined ? '' : String(field.minSelections)}
                      onChange={(_e, v) => set({ minSelections: numberOrUndefined(v) })}
                    />
                    <TextField
                      label="Maximum selections"
                      value={field.maxSelections === undefined ? '' : String(field.maxSelections)}
                      onChange={(_e, v) => set({ maxSelections: numberOrUndefined(v) })}
                    />
                  </div>
                )}
              </>
            )}

            {field.type === FieldType.Lookup && (
              <>
                <Dropdown
                  label="Source list"
                  options={lists.map((l) => ({ key: l.id, text: l.title }))}
                  selectedKey={field.lookupListId || null}
                  placeholder="Choose a list"
                  onChange={(_e, option) =>
                    option && set({ lookupListId: String(option.key), lookupColumn: undefined })
                  }
                />
                <Dropdown
                  label="Column to read options from"
                  options={lookupColumns.map((c) => ({
                    key: c.internalName,
                    text: c.title + ' (' + c.typeAsString + ')'
                  }))}
                  selectedKey={field.lookupColumn || null}
                  placeholder={field.lookupListId ? 'Choose a column' : 'Choose a list first'}
                  disabled={!field.lookupListId}
                  onChange={(_e, option) => option && set({ lookupColumn: String(option.key) })}
                />
                <TextField
                  label="Filter (optional)"
                  value={field.lookupFilter || ''}
                  placeholder="Status eq 'Active'"
                  description="An OData filter narrowing which source items become options"
                  onChange={(_e, v) => set({ lookupFilter: v })}
                />
                <MessageBar messageBarType={MessageBarType.info}>
                  Answers store the option text, not a link to the source item, so responses stay
                  readable even if the source list changes later.
                </MessageBar>
              </>
            )}

            {field.type === FieldType.Consent && (
              <TextField
                label="Terms shown above the checkbox"
                multiline={true}
                rows={5}
                value={field.consentText || ''}
                description="Basic formatting is allowed: bold, italic, lists and links."
                onChange={(_e, v) => set({ consentText: v })}
              />
            )}

            {field.type === FieldType.Person && (
              <>
                <Toggle
                  label="Allow multiple people"
                  checked={field.allowMultiplePeople === true}
                  disabled={field.provisioned === true}
                  onChange={(_e, checked) => set({ allowMultiplePeople: checked === true })}
                />
                <Toggle
                  label="Also allow groups"
                  checked={field.allowGroups === true}
                  onChange={(_e, checked) => set({ allowGroups: checked === true })}
                />
              </>
            )}

            {field.type === FieldType.Date && (
              <>
                <Toggle
                  label="Include a time of day"
                  checked={field.includeTime === true}
                  disabled={field.provisioned === true}
                  onChange={(_e, checked) => set({ includeTime: checked === true })}
                />
                {field.includeTime && (
                  <TextField
                    label="Time step (minutes)"
                    value={String(field.timeStepMinutes || 15)}
                    onChange={(_e, v) => set({ timeStepMinutes: numberOrUndefined(v) })}
                  />
                )}
              </>
            )}

            {field.type === FieldType.Time && (
              <TextField
                label="Time step (minutes)"
                value={String(field.timeStepMinutes || 15)}
                onChange={(_e, v) => set({ timeStepMinutes: numberOrUndefined(v) })}
              />
            )}

            {field.type === FieldType.FileUpload && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label="Maximum files"
                    value={String(field.maxFiles || 3)}
                    onChange={(_e, v) => set({ maxFiles: numberOrUndefined(v) })}
                  />
                  <TextField
                    label="Max size each (MB)"
                    value={String(field.maxFileSizeMb || 10)}
                    onChange={(_e, v) => set({ maxFileSizeMb: numberOrUndefined(v) })}
                  />
                </div>
                <TextField
                  label="Accepted file types"
                  value={(field.allowedExtensions || []).join(', ')}
                  placeholder="pdf, docx, png"
                  description="Comma-separated extensions. Leave blank to accept anything SharePoint allows."
                  onChange={(_e, v) =>
                    set({
                      allowedExtensions: (v || '')
                        .split(',')
                        .map((e) => e.trim().toLowerCase().replace(/^\./, ''))
                        .filter((e) => e.length > 0)
                    })
                  }
                />
                <MessageBar messageBarType={MessageBarType.info}>
                  Files are saved as attachments on the response item. SharePoint blocks some
                  extensions (such as .exe) regardless of what you allow here.
                </MessageBar>
              </>
            )}

            {(field.type === FieldType.MultilineText || field.type === FieldType.RichText) && (
              <TextField
                label="Rows"
                value={String(field.rows || (field.type === FieldType.RichText ? 6 : 4))}
                onChange={(_e, v) => set({ rows: numberOrUndefined(v) })}
              />
            )}

            {field.provisioned && (
              <p className={styles.panelHint}>
                <Icon iconName="Lock" /> This question already has a list column
                (<strong>{field.internalName}</strong>). Settings that change the column&apos;s type
                are locked — duplicate the question if you need a different type.
              </p>
            )}
      </>
    );
  }
};
