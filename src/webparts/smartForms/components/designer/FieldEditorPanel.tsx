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
import * as strings from 'SmartFormsWebPartStrings';
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
  IMessageBag,
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
  operatorLabels,
  operatorsForField
} from '../../utils/formUtils';
import { evaluateFormula, formulaReferences } from '../../utils/formula';
import { CURRENCY_OPTIONS } from '../../utils/spFieldXml';
import { SharePointService } from '../../services/SharePointService';
import { formatString } from '../../utils/localeUtils';

const messageBag = strings as unknown as IMessageBag;

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
  { key: 'info', text: strings.Designer_Field_StyleInfo },
  { key: 'success', text: strings.Designer_Field_StyleSuccess },
  { key: 'warning', text: strings.Designer_Field_StyleWarning },
  { key: 'text', text: strings.Designer_Field_StylePlain },
  { key: 'divider', text: strings.Designer_Field_StyleDivider }
];

const RATING_ICONS: IDropdownOption[] = [
  { key: 'star', text: strings.Designer_Field_IconStars },
  { key: 'heart', text: strings.Designer_Field_IconHearts },
  { key: 'like', text: strings.Designer_Field_IconThumbs }
];

const NUMBER_FORMATS: IDropdownOption[] = [
  { key: 'plain', text: strings.Designer_Field_FormatPlain },
  { key: 'currency', text: strings.Designer_Field_FormatCurrency },
  { key: 'percent', text: strings.Designer_Field_FormatPercent }
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
      const label = operatorLabels(messageBag).filter((o) => o.key === condition.operator)[0];
      const operatorText = label ? label.text : condition.operator;
      const name = driver
        ? driver.title || strings.Designer_Field_UntitledQuestion
        : strings.Designer_Field_DeletedQuestion;
      if (UNARY_OPERATORS.indexOf(condition.operator) !== -1) {
        return formatString(strings.Designer_Rule_Unary, { name, operator: operatorText });
      }
      if (condition.operator === 'between') {
        return formatString(strings.Designer_Rule_Between, {
          name,
          operator: operatorText,
          value: condition.value || '?',
          value2: condition.value2 || '?'
        });
      }
      return formatString(strings.Designer_Rule_Binary, {
        name,
        operator: operatorText,
        value: condition.value || ''
      });
    });
    const joiner = ' ' + (match === 'any' ? strings.Designer_Rule_JoinAny : strings.Designer_Rule_JoinAll) + ' ';
    return formatString(strings.Designer_Rule_Preview, { conditions: parts.join(joiner) });
  };

  /** Options for a rule's value box when the driver is a choice question. */
  const valueOptionsFor = (driver: IFormField | undefined): IDropdownOption[] | undefined => {
    if (!driver) {
      return undefined;
    }
    if (driver.type === FieldType.YesNo || driver.type === FieldType.Consent) {
      return [
        { key: 'Yes', text: strings.Designer_Field_Yes },
        { key: 'No', text: strings.Designer_Field_No }
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
    const operatorOptions = operatorLabels(messageBag).filter((o) => allowedOperators.indexOf(o.key) !== -1).map(
      (o) => ({ key: o.key, text: o.text })
    );
    const needsValue = UNARY_OPERATORS.indexOf(condition.operator) === -1;
    const valueOptions = valueOptionsFor(driver);

    return (
      <div key={index} className={styles.ruleRow}>
        <div className={styles.ruleRowFields}>
          <Dropdown
            label={index === 0 ? strings.Designer_Field_When : undefined}
            options={otherFields.map((f) => ({ key: f.id, text: f.title || strings.Designer_Field_UntitledQuestion }))}
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
                placeholder={strings.Designer_Field_SelectValue}
                onChange={(_e, option) => option && patchCondition(index, { value: String(option.key) })}
              />
            ) : (
              <TextField
                value={condition.value || ''}
                placeholder={
                  driver && (driver.type === FieldType.Date || driver.type === FieldType.Time)
                    ? strings.Designer_Field_DateValuePlaceholder
                    : strings.Designer_Field_ValuePlaceholder
                }
                onChange={(_e, v) => patchCondition(index, { value: v })}
              />
            ))}
          {condition.operator === 'between' && (
            <TextField
              value={condition.value2 || ''}
              placeholder={strings.Designer_Field_UpperBound}
              onChange={(_e, v) => patchCondition(index, { value2: v })}
            />
          )}
        </div>
        <DefaultButton
          iconProps={{ iconName: 'Delete' }}
          title={strings.Designer_Field_RemoveCondition}
          ariaLabel={strings.Designer_Field_RemoveCondition}
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
      return strings.Designer_Field_FormulaInvalid;
    }
    const missing = formulaReferences(field.formula || '').filter(
      (name) => !allFields(props.definition).filter((f) => (f.title || '').trim().toLowerCase() === name.toLowerCase())[0]
    );
    if (missing.length > 0) {
      return formatString(strings.Designer_Field_FormulaMissing, { names: missing.join(', ') });
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
    return result === undefined
      ? ''
      : formatString(strings.Designer_Field_FormulaPreview, { value: formatValue(field, result) });
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
        return strings.Designer_Field_HintYesNo;
      case FieldType.Choice:
      case FieldType.Lookup:
        return field.allowMultiple
          ? strings.Designer_Field_HintMultiDefault
          : strings.Designer_Field_HintSingleDefault;
      case FieldType.Date:
        return strings.Designer_Field_HintDate;
      case FieldType.Time:
        return strings.Designer_Field_HintTime;
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
      headerText={field.title || strings.Designer_Field_QuestionSettings}
      onDismiss={props.onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.panelFooter}>
          <PrimaryButton text={strings.Designer_Common_Apply} onClick={handleSave} />
          <DefaultButton text={strings.Designer_Common_Cancel} onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className={styles.panelBody}>
        {/* ----- Content blocks have their own short form ----- */}
        {isContent ? (
          <>
            <h4 className={styles.panelSectionTitle}>{strings.Designer_Field_Appearance}</h4>
            <Dropdown
              label={strings.Designer_Field_Style}
              options={CONTENT_STYLES}
              selectedKey={field.contentStyle || 'info'}
              onChange={(_e, option) =>
                option && set({ contentStyle: String(option.key) as ContentStyle })
              }
            />
            <TextField
              label={strings.Designer_Field_ImageUrl}
              value={field.contentImageUrl || ''}
              placeholder="https://…"
              onChange={(_e, v) => set({ contentImageUrl: v })}
            />
            <TextField
              label={strings.Designer_Field_Content}
              multiline={true}
              rows={5}
              value={field.contentHtml || ''}
              description={strings.Designer_Field_BasicFormatting}
              onChange={(_e, v) => set({ contentHtml: v })}
            />
          </>
        ) : (
          <>
            <Pivot aria-label={strings.Designer_Field_QuestionSettings} defaultSelectedKey="options">
              {/* ----- Per-type options ----- */}
              <PivotItem headerText={strings.Designer_Field_TabOptions} itemKey="options">
                <div className={styles.panelTabContent}>{renderOptionsTab()}</div>
              </PivotItem>

              {/* ----- Branching ----- */}
              <PivotItem headerText={strings.Designer_Field_TabBranching} itemKey="branching">
                <div className={styles.panelTabContent}>
                  {otherFields.length === 0 ? (
                    <p className={styles.panelHint}>
                      {strings.Designer_Field_BranchingNeedsQuestion}
                    </p>
                  ) : (
                    <>
                      <Toggle
                        label={strings.Designer_Field_BranchingToggle}
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
                              <span>{strings.Designer_Field_Match}</span>
                              <Dropdown
                                className={styles.ruleMatchDropdown}
                                options={[
                                  { key: 'all', text: strings.Designer_Field_MatchAll },
                                  { key: 'any', text: strings.Designer_Field_MatchAny }
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
                            text={strings.Designer_Field_AddCondition}
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
              <PivotItem headerText={strings.Designer_Field_TabValidation} itemKey="validation">
                <div className={styles.panelTabContent}>
                  <TextField
                    label={strings.Designer_Field_HelpText}
                    value={field.description || ''}
                    placeholder={strings.Designer_Field_HelpTextPlaceholder}
                    onChange={(_e, v) => set({ description: v })}
                  />
                  {supportsPlaceholder && (
                    <TextField
                      label={
                        field.type === FieldType.Consent
                          ? strings.Designer_Field_CheckboxLabel
                          : strings.Designer_Field_Placeholder
                      }
                      value={field.placeholder || ''}
                      onChange={(_e, v) => set({ placeholder: v })}
                    />
                  )}
                  {supportsDefault && (
                    <TextField
                      label={strings.Designer_Field_DefaultAnswer}
                      value={field.defaultValue || ''}
                      description={defaultValueHint()}
                      onChange={(_e, v) => set({ defaultValue: v })}
                    />
                  )}
                  <TextField
                    label={strings.Designer_Field_RequiredMessage}
                    value={field.requiredMessage || ''}
                    placeholder={formatString(strings.Designer_Field_RequiredMessagePlaceholder, {
                      title: field.title || strings.Designer_Field_ThisQuestion
                    })}
                    description={strings.Designer_Field_RequiredMessageDescription}
                    disabled={field.required !== true}
                    onChange={(_e, v) => set({ requiredMessage: v })}
                  />
                  <Toggle
                    label={strings.Designer_Field_ReadOnly}
                    checked={field.readOnly === true}
                    onChange={(_e, checked) => set({ readOnly: checked === true })}
                  />
                  {(supportsPattern ||
                    isNumeric ||
                    field.type === FieldType.Text ||
                    field.type === FieldType.MultilineText) && (
                    <>
                      <h4 className={styles.panelSectionTitle}>{strings.Designer_Field_ValidationRules}</h4>
                      {(field.type === FieldType.Text || field.type === FieldType.MultilineText) && (
                        <TextField
                          label={strings.Designer_Field_MaxLength}
                          value={field.maxLength === undefined ? '' : String(field.maxLength)}
                          placeholder="255"
                          onChange={(_e, v) => set({ maxLength: numberOrUndefined(v) })}
                        />
                      )}
                      {supportsPattern && (
                        <>
                          <TextField
                            label={strings.Designer_Field_Pattern}
                            value={field.pattern || ''}
                            placeholder="^[A-Z]{2}\d{4}$"
                            description={strings.Designer_Field_PatternDescription}
                            onChange={(_e, v) => set({ pattern: v })}
                          />
                          <TextField
                            label={strings.Designer_Field_PatternMessage}
                            value={field.patternMessage || ''}
                            placeholder={strings.Designer_Field_PatternMessagePlaceholder}
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

  /** The per-type {strings.Designer_Field_TabOptions} tab: the long tail of field-specific settings. */
  function renderOptionsTab(): React.ReactNode {
    return (
      <>
            {isNumeric && (
              <>
                <Dropdown
                  label={strings.Designer_Field_Format}
                  options={NUMBER_FORMATS}
                  selectedKey={field.numberFormat || 'plain'}
                  onChange={(_e, option) =>
                    option && set({ numberFormat: String(option.key) as NumberFormat })
                  }
                />
                {field.numberFormat === 'currency' && (
                  <Dropdown
                    label={strings.Designer_Field_FormatCurrency}
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
                    label={strings.Designer_Field_UnitSuffix}
                    value={field.unitSuffix || ''}
                    placeholder={strings.Designer_Field_UnitSuffixPlaceholder}
                    onChange={(_e, v) => set({ unitSuffix: v })}
                  />
                )}
                <div className={styles.inlineFields}>
                  <TextField
                    label={strings.Designer_Field_Minimum}
                    value={field.min === undefined ? '' : String(field.min)}
                    onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                  />
                  <TextField
                    label={strings.Designer_Field_Maximum}
                    value={field.max === undefined ? '' : String(field.max)}
                    onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                  />
                  <TextField
                    label={strings.Designer_Field_Decimals}
                    value={field.decimalPlaces === undefined ? '' : String(field.decimalPlaces)}
                    onChange={(_e, v) => set({ decimalPlaces: numberOrUndefined(v) })}
                  />
                </div>
              </>
            )}

            {field.type === FieldType.Calculated && (
              <>
                <TextField
                  label={strings.Designer_Field_Formula}
                  value={field.formula || ''}
                  placeholder={strings.Designer_Field_FormulaPlaceholder}
                  multiline={true}
                  rows={2}
                  description={strings.Designer_Field_FormulaDescription}
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
                  {formatString(strings.Designer_Field_AvailableQuestions, {
                    questions:
                      inputFields(props.definition)
                        .filter((f) => f.id !== field.id && (f.title || '').trim())
                        .map((f) => '{' + f.title + '}')
                        .join(', ') || strings.Designer_Field_NoneYet
                  })}
                </p>
              </>
            )}

            {field.type === FieldType.Slider && (
              <div className={styles.inlineFields}>
                <TextField
                  label={strings.Designer_Field_Minimum}
                  value={field.min === undefined ? '0' : String(field.min)}
                  onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                />
                <TextField
                  label={strings.Designer_Field_Maximum}
                  value={field.max === undefined ? '10' : String(field.max)}
                  onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                />
                <TextField
                  label={strings.Designer_Field_Step}
                  value={field.step === undefined ? '1' : String(field.step)}
                  onChange={(_e, v) => set({ step: numberOrUndefined(v) })}
                />
              </div>
            )}

            {field.type === FieldType.Scale && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label={strings.Designer_Field_From}
                    value={field.min === undefined ? '1' : String(field.min)}
                    onChange={(_e, v) => set({ min: numberOrUndefined(v) })}
                  />
                  <TextField
                    label={strings.Designer_Field_To}
                    value={field.max === undefined ? '5' : String(field.max)}
                    onChange={(_e, v) => set({ max: numberOrUndefined(v) })}
                  />
                  <TextField
                    label={strings.Designer_Field_Step}
                    value={field.step === undefined ? '1' : String(field.step)}
                    onChange={(_e, v) => set({ step: numberOrUndefined(v) })}
                  />
                </div>
                <div className={styles.inlineFields}>
                  <TextField
                    label={strings.Designer_Field_LowLabel}
                    value={field.lowLabel || ''}
                    placeholder={strings.Designer_Field_LowLabelPlaceholder}
                    onChange={(_e, v) => set({ lowLabel: v })}
                  />
                  <TextField
                    label={strings.Designer_Field_HighLabel}
                    value={field.highLabel || ''}
                    placeholder={strings.Designer_Field_HighLabelPlaceholder}
                    onChange={(_e, v) => set({ highLabel: v })}
                  />
                </div>
                <Dropdown
                  label={strings.Designer_Field_SummarizeAs}
                  options={[
                    { key: 'average', text: strings.Designer_Field_SummarizeAverage },
                    { key: 'nps', text: strings.Designer_Field_SummarizeNps }
                  ]}
                  selectedKey={field.scaleAnalytics || 'average'}
                  onChange={(_e, option) =>
                    option && set({ scaleAnalytics: String(option.key) as ScaleAnalytics })
                  }
                />
                {field.scaleAnalytics === 'nps' && (field.min !== 0 || field.max !== 10) && (
                  <MessageBar messageBarType={MessageBarType.warning}>
                    {strings.Designer_Field_NpsWarning}
                  </MessageBar>
                )}
              </>
            )}

            {field.type === FieldType.Rating && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label={strings.Designer_Field_NumberOfIcons}
                    value={String(field.maxRating || 5)}
                    onChange={(_e, v) => set({ maxRating: numberOrUndefined(v) })}
                  />
                  <Dropdown
                    label={strings.Designer_Field_Icon}
                    options={RATING_ICONS}
                    selectedKey={field.ratingIcon || 'star'}
                    onChange={(_e, option) =>
                      option && set({ ratingIcon: String(option.key) as RatingIcon })
                    }
                  />
                </div>
                <Toggle
                  label={strings.Designer_Field_AllowHalf}
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
                      label={strings.Designer_Field_ShowOptionsAs}
                      options={[
                        { key: 'buttons', text: strings.Designer_Field_DisplayButtons },
                        { key: 'dropdown', text: strings.Designer_Field_DisplayDropdown }
                      ]}
                      selectedKey={field.choiceDisplay || undefined}
                      placeholder={strings.Designer_Field_DisplayAutomatic}
                      onChange={(_e, option) =>
                        option && set({ choiceDisplay: option.key === 'buttons' ? 'buttons' : 'dropdown' })
                      }
                    />
                    <Toggle
                      label={strings.Designer_Field_AllowOther}
                      checked={field.allowOther === true}
                      onChange={(_e, checked) => set({ allowOther: checked === true })}
                    />
                    {field.allowOther && (
                      <TextField
                        label={strings.Designer_Field_OtherLabel}
                        value={field.otherLabel || ''}
                        placeholder={strings.Designer_Field_OtherPlaceholder}
                        onChange={(_e, v) => set({ otherLabel: v })}
                      />
                    )}
                  </>
                )}
                <Toggle
                  label={strings.Designer_Field_ShuffleOptions}
                  checked={field.shuffleOptions === true}
                  onChange={(_e, checked) => set({ shuffleOptions: checked === true })}
                />
                {field.allowMultiple && (
                  <div className={styles.inlineFields}>
                    <TextField
                      label={strings.Designer_Field_MinSelections}
                      value={field.minSelections === undefined ? '' : String(field.minSelections)}
                      onChange={(_e, v) => set({ minSelections: numberOrUndefined(v) })}
                    />
                    <TextField
                      label={strings.Designer_Field_MaxSelections}
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
                  label={strings.Designer_Field_SourceList}
                  options={lists.map((l) => ({ key: l.id, text: l.title }))}
                  selectedKey={field.lookupListId || null}
                  placeholder={strings.Designer_Field_ChooseList}
                  onChange={(_e, option) =>
                    option && set({ lookupListId: String(option.key), lookupColumn: undefined })
                  }
                />
                <Dropdown
                  label={strings.Designer_Field_LookupColumn}
                  options={lookupColumns.map((c) => ({
                    key: c.internalName,
                    text: c.title + ' (' + c.typeAsString + ')'
                  }))}
                  selectedKey={field.lookupColumn || null}
                  placeholder={
                    field.lookupListId
                      ? strings.Designer_Field_ChooseColumn
                      : strings.Designer_Field_ChooseListFirst
                  }
                  disabled={!field.lookupListId}
                  onChange={(_e, option) => option && set({ lookupColumn: String(option.key) })}
                />
                <TextField
                  label={strings.Designer_Field_LookupFilter}
                  value={field.lookupFilter || ''}
                  placeholder="Status eq 'Active'"
                  description={strings.Designer_Field_LookupFilterDescription}
                  onChange={(_e, v) => set({ lookupFilter: v })}
                />
                <MessageBar messageBarType={MessageBarType.info}>
                  {strings.Designer_Field_LookupInfo}
                </MessageBar>
              </>
            )}

            {field.type === FieldType.Consent && (
              <TextField
                label={strings.Designer_Field_ConsentText}
                multiline={true}
                rows={5}
                value={field.consentText || ''}
                description={strings.Designer_Field_BasicFormatting}
                onChange={(_e, v) => set({ consentText: v })}
              />
            )}

            {field.type === FieldType.Person && (
              <>
                <Toggle
                  label={strings.Designer_Field_AllowMultiplePeople}
                  checked={field.allowMultiplePeople === true}
                  disabled={field.provisioned === true}
                  onChange={(_e, checked) => set({ allowMultiplePeople: checked === true })}
                />
                <Toggle
                  label={strings.Designer_Field_AllowGroups}
                  checked={field.allowGroups === true}
                  onChange={(_e, checked) => set({ allowGroups: checked === true })}
                />
              </>
            )}

            {field.type === FieldType.Date && (
              <>
                <Toggle
                  label={strings.Designer_Field_IncludeTime}
                  checked={field.includeTime === true}
                  disabled={field.provisioned === true}
                  onChange={(_e, checked) => set({ includeTime: checked === true })}
                />
                {field.includeTime && (
                  <TextField
                    label={strings.Designer_Field_TimeStep}
                    value={String(field.timeStepMinutes || 15)}
                    onChange={(_e, v) => set({ timeStepMinutes: numberOrUndefined(v) })}
                  />
                )}
              </>
            )}

            {field.type === FieldType.Time && (
              <TextField
                label={strings.Designer_Field_TimeStep}
                value={String(field.timeStepMinutes || 15)}
                onChange={(_e, v) => set({ timeStepMinutes: numberOrUndefined(v) })}
              />
            )}

            {field.type === FieldType.FileUpload && (
              <>
                <div className={styles.inlineFields}>
                  <TextField
                    label={strings.Designer_Field_MaxFiles}
                    value={String(field.maxFiles || 3)}
                    onChange={(_e, v) => set({ maxFiles: numberOrUndefined(v) })}
                  />
                  <TextField
                    label={strings.Designer_Field_MaxFileSize}
                    value={String(field.maxFileSizeMb || 10)}
                    onChange={(_e, v) => set({ maxFileSizeMb: numberOrUndefined(v) })}
                  />
                </div>
                <TextField
                  label={strings.Designer_Field_AcceptedTypes}
                  value={(field.allowedExtensions || []).join(', ')}
                  placeholder="pdf, docx, png"
                  description={strings.Designer_Field_AcceptedTypesDescription}
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
                  {strings.Designer_Field_FilesInfo}
                </MessageBar>
              </>
            )}

            {(field.type === FieldType.MultilineText || field.type === FieldType.RichText) && (
              <TextField
                label={strings.Designer_Field_Rows}
                value={String(field.rows || (field.type === FieldType.RichText ? 6 : 4))}
                onChange={(_e, v) => set({ rows: numberOrUndefined(v) })}
              />
            )}

            {field.provisioned && (
              <p className={styles.panelHint}>
                <Icon iconName="Lock" />{' '}
                {strings.Designer_Field_ProvisionedNote.split('{name}')[0]}
                <strong>{field.internalName}</strong>
                {strings.Designer_Field_ProvisionedNote.split('{name}').slice(1).join('{name}')}
              </p>
            )}
      </>
    );
  }
};
