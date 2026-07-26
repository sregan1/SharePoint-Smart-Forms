import * as React from 'react';
import {
  Checkbox,
  ChoiceGroup,
  DatePicker,
  Dropdown,
  Slider,
  TextField,
  Toggle
} from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import {
  FieldType,
  IAddressValue,
  IFormField,
  IFormFile,
  IHyperlinkValue,
  ILikertValue,
  IPersonInfo
} from '../../models';
import { SharePointService } from '../../services/SharePointService';
import { effectiveChoices, formatValue, shuffleWithSeed } from '../../utils/formUtils';
import { PersonField } from './PersonField';
import { RichTextField, RichTextView } from './RichTextField';
import { RankingField } from './RankingField';
import { ScaleField } from './ScaleField';
import { RatingField } from './RatingField';
import { LikertField } from './LikertField';
import { FileUploadField } from './FileUploadField';
import { SignatureField } from './SignatureField';
import { LookupField } from './LookupField';
import { ImageChoiceField } from './ImageChoiceField';
import { AddressField } from './AddressField';
import { ConsentField } from './ConsentField';
import { TimeField } from './TimeField';

export interface IFieldControlProps {
  field: IFormField;
  value: unknown;
  error?: string;
  /** id of the element describing this control (help text and/or error) */
  describedBy?: string;
  /** id applied to the focusable input, so a <label for> can target it */
  controlId?: string;
  /** render as a non-editable value (preview cards, read-only fields) */
  disabled?: boolean;
  /** seed for per-respondent option shuffling */
  shuffleSeed?: number;
  onChange: (value: unknown) => void;
  spService: SharePointService;
}

const parseNumber = (text: string): number | undefined => {
  const trimmed = (text || '').replace(/,/g, '').trim();
  if (trimmed.length === 0) {
    return undefined;
  }
  const num = Number(trimmed);
  return isNaN(num) ? undefined : num;
};

interface INumberTextFieldProps {
  value: unknown;
  placeholder?: string;
  prefix?: string;
  suffix?: string;
  disabled?: boolean;
  id?: string;
  ariaLabel?: string;
  describedBy?: string;
  invalid?: boolean;
  onChange: (value: number | undefined) => void;
}

/**
 * TextField that keeps the raw text while typing (so "12." survives a
 * keystroke) but reports parsed numbers to the form state.
 */
const NumberTextField: React.FunctionComponent<INumberTextFieldProps> = (props) => {
  const [text, setText] = React.useState<string>(
    props.value === undefined || props.value === null ? '' : String(props.value)
  );

  React.useEffect(() => {
    // resync only when the external value diverges from what's being typed
    if (parseNumber(text) !== props.value) {
      setText(props.value === undefined || props.value === null ? '' : String(props.value));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.value]);

  return (
    <TextField
      id={props.id}
      value={text}
      placeholder={props.placeholder}
      prefix={props.prefix}
      suffix={props.suffix}
      disabled={props.disabled}
      inputMode="decimal"
      ariaLabel={props.ariaLabel}
      aria-describedby={props.describedBy}
      aria-invalid={props.invalid ? true : undefined}
      onChange={(_event, v) => {
        setText(v || '');
        props.onChange(parseNumber(v || ''));
      }}
    />
  );
};

export const FieldControl: React.FunctionComponent<IFieldControlProps> = (props) => {
  const { field, value, onChange } = props;
  const disabled = props.disabled === true || field.readOnly === true;
  const invalid = !!props.error;
  const describedBy = props.describedBy;
  const controlId = props.controlId;
  // Fluent controls that render their own label get an aria-label instead, since
  // the visible label lives outside the control in FormRenderer
  const ariaLabel = field.title || undefined;

  /** Options in the order this respondent should see them. */
  const orderedChoices = (): string[] => {
    const choices = effectiveChoices(field);
    if (!field.shuffleOptions || typeof props.shuffleSeed !== 'number') {
      return choices;
    }
    return shuffleWithSeed(choices, props.shuffleSeed);
  };

  switch (field.type) {
    // ----- text -----

    case FieldType.Text:
    case FieldType.Email:
    case FieldType.Phone:
      return (
        <TextField
          id={controlId}
          value={(value as string) || ''}
          placeholder={field.placeholder}
          maxLength={field.maxLength || 255}
          disabled={disabled}
          type={field.type === FieldType.Email ? 'email' : field.type === FieldType.Phone ? 'tel' : 'text'}
          autoComplete={
            field.type === FieldType.Email ? 'email' : field.type === FieldType.Phone ? 'tel' : undefined
          }
          iconProps={
            field.type === FieldType.Email
              ? { iconName: 'Mail' }
              : field.type === FieldType.Phone
                ? { iconName: 'Phone' }
                : undefined
          }
          ariaLabel={ariaLabel}
          aria-describedby={describedBy}
          aria-invalid={invalid ? true : undefined}
          onChange={(_event, v) => onChange(v)}
        />
      );

    case FieldType.MultilineText:
      return (
        <TextField
          id={controlId}
          multiline={true}
          rows={field.rows || 4}
          resizable={true}
          value={(value as string) || ''}
          placeholder={field.placeholder}
          maxLength={field.maxLength}
          disabled={disabled}
          ariaLabel={ariaLabel}
          aria-describedby={describedBy}
          aria-invalid={invalid ? true : undefined}
          onChange={(_event, v) => onChange(v)}
        />
      );

    case FieldType.RichText:
      if (disabled) {
        return <RichTextView html={(value as string) || ''} />;
      }
      return (
        <RichTextField
          value={(value as string) || ''}
          placeholder={field.placeholder}
          rows={field.rows}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    // ----- numeric -----

    case FieldType.Number:
      return (
        <NumberTextField
          id={controlId}
          value={value}
          placeholder={field.placeholder}
          prefix={field.numberFormat === 'currency' ? field.currencySymbol || '$' : undefined}
          suffix={
            field.numberFormat === 'percent' ? '%' : field.unitSuffix ? field.unitSuffix : undefined
          }
          disabled={disabled}
          ariaLabel={ariaLabel}
          describedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Calculated:
      // never editable — the formula owns the value
      return (
        <div className={styles.readOnlyValue} id={controlId} aria-describedby={describedBy}>
          {value === undefined || value === null ? '—' : formatValue(field, value)}
        </div>
      );

    // ----- date & time -----

    case FieldType.Date: {
      const current = value instanceof Date ? (value as Date) : undefined;
      return (
        <div className={styles.dateRow}>
          <DatePicker
            id={controlId}
            value={current}
            placeholder={field.placeholder || 'Select a date'}
            allowTextInput={true}
            disabled={disabled}
            ariaLabel={ariaLabel}
            textField={{
              'aria-describedby': describedBy,
              'aria-invalid': invalid ? true : undefined
            }}
            onSelectDate={(date) => {
              if (!date) {
                onChange(undefined);
                return;
              }
              const next = new Date(date.getTime());
              if (field.includeTime) {
                // preserve a time already chosen; otherwise start at 9am
                if (current) {
                  next.setHours(current.getHours(), current.getMinutes(), 0, 0);
                } else {
                  next.setHours(9, 0, 0, 0);
                }
              } else {
                next.setHours(0, 0, 0, 0);
              }
              onChange(next);
            }}
          />
          {field.includeTime && (
            <div className={styles.timeDropdown}>
              <TimeField
                value={current}
                stepMinutes={field.timeStepMinutes || 15}
                disabled={disabled || !current}
                ariaLabel={ariaLabel ? ariaLabel + ' — time' : 'Time'}
                invalid={invalid}
                onChange={(next) => {
                  if (!next || !current) {
                    return;
                  }
                  const merged = new Date(current.getTime());
                  merged.setHours(next.getHours(), next.getMinutes(), 0, 0);
                  onChange(merged);
                }}
              />
            </div>
          )}
        </div>
      );
    }

    case FieldType.Time:
      return (
        <TimeField
          value={value instanceof Date ? (value as Date) : undefined}
          stepMinutes={field.timeStepMinutes || 15}
          placeholder={field.placeholder}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    // ----- choice -----

    case FieldType.Choice: {
      const choices = orderedChoices();
      // Microsoft Forms-style default: show options inline when the list is
      // short, fall back to a dropdown for long lists. Forms that explicitly
      // chose a display keep it.
      const display = field.choiceDisplay || (choices.length <= 6 ? 'buttons' : 'dropdown');
      const otherLabel = field.otherLabel || 'Other';

      if (field.allowMultiple) {
        const selected = Array.isArray(value) ? (value as string[]) : [];
        const known = selected.filter((s) => choices.indexOf(s) !== -1);
        const otherValue = selected.filter((s) => choices.indexOf(s) === -1)[0] || '';
        const atLimit =
          typeof field.maxSelections === 'number' && selected.length >= field.maxSelections;

        const setKnown = (next: string[]): void =>
          onChange(otherValue ? next.concat([otherValue]) : next);

        return (
          <div>
            {display === 'buttons' ? (
              <div
                className={styles.checkboxStack}
                role="group"
                aria-label={ariaLabel}
                aria-describedby={describedBy}
              >
                {choices.map((choice, index) => {
                  const checked = known.indexOf(choice) !== -1;
                  return (
                    <Checkbox
                      // labels can repeat; index keeps the keys unique
                      key={choice + '::' + index}
                      label={choice}
                      checked={checked}
                      disabled={disabled || (atLimit && !checked)}
                      onChange={(_event, isChecked) =>
                        setKnown(isChecked ? known.concat([choice]) : known.filter((s) => s !== choice))
                      }
                    />
                  );
                })}
              </div>
            ) : (
              <Dropdown
                id={controlId}
                multiSelect={true}
                options={choices.map((c, i) => ({ key: c, text: c, index: i }))}
                selectedKeys={known}
                placeholder={field.placeholder || 'Select options'}
                disabled={disabled}
                ariaLabel={ariaLabel}
                aria-describedby={describedBy}
                errorMessage={invalid ? ' ' : undefined}
                onChange={(_event, option) => {
                  if (!option) {
                    return;
                  }
                  const key = String(option.key);
                  if (option.selected && atLimit) {
                    return;
                  }
                  setKnown(option.selected ? known.concat([key]) : known.filter((s) => s !== key));
                }}
              />
            )}

            {field.allowOther && (
              <div className={styles.otherRow}>
                <Checkbox
                  label={otherLabel}
                  checked={!!otherValue}
                  disabled={disabled || (atLimit && !otherValue)}
                  onChange={(_event, isChecked) =>
                    onChange(isChecked ? known.concat([otherLabel]) : known)
                  }
                />
                {!!otherValue && (
                  <TextField
                    className={styles.otherInput}
                    value={otherValue === otherLabel ? '' : otherValue}
                    placeholder="Please specify"
                    disabled={disabled}
                    ariaLabel={otherLabel + ' — please specify'}
                    onChange={(_event, v) => onChange(known.concat([v || otherLabel]))}
                  />
                )}
              </div>
            )}

            {(typeof field.minSelections === 'number' || typeof field.maxSelections === 'number') && (
              <div className={styles.selectionHint}>
                {selectionHint(field.minSelections, field.maxSelections, selected.length)}
              </div>
            )}
          </div>
        );
      }

      // single select
      const selectedValue = (value as string) || '';
      const isOther = !!selectedValue && choices.indexOf(selectedValue) === -1;
      const groupOptions = choices.map((c, i) => ({ key: c, text: c, index: i }));
      if (field.allowOther) {
        groupOptions.push({ key: otherLabel, text: otherLabel, index: groupOptions.length });
      }

      return (
        <div>
          {display === 'buttons' ? (
            <ChoiceGroup
              options={groupOptions}
              selectedKey={isOther ? otherLabel : selectedValue || undefined}
              disabled={disabled}
              ariaLabelledBy={undefined}
              styles={{ root: { marginTop: 4 } }}
              onChange={(_event, option) => {
                if (!option) {
                  return;
                }
                onChange(String(option.key) === otherLabel && !isOther ? otherLabel : String(option.key));
              }}
            />
          ) : (
            <Dropdown
              id={controlId}
              options={groupOptions}
              selectedKey={isOther ? otherLabel : selectedValue || null}
              placeholder={field.placeholder || 'Select an option'}
              disabled={disabled}
              ariaLabel={ariaLabel}
              aria-describedby={describedBy}
              errorMessage={invalid ? ' ' : undefined}
              onChange={(_event, option) => option && onChange(String(option.key))}
            />
          )}
          {field.allowOther && isOther && (
            <div className={styles.otherRow}>
              <TextField
                className={styles.otherInput}
                value={selectedValue === otherLabel ? '' : selectedValue}
                placeholder="Please specify"
                disabled={disabled}
                ariaLabel={otherLabel + ' — please specify'}
                onChange={(_event, v) => onChange(v || otherLabel)}
              />
            </div>
          )}
        </div>
      );
    }

    case FieldType.ImageChoice:
      return (
        <ImageChoiceField
          options={field.imageChoices || []}
          value={value as string | string[] | undefined}
          allowMultiple={field.allowMultiple === true}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Lookup:
      return (
        <LookupField
          listId={field.lookupListId || ''}
          column={field.lookupColumn || 'Title'}
          filter={field.lookupFilter}
          value={value as string | string[] | undefined}
          allowMultiple={field.allowMultiple === true}
          placeholder={field.placeholder}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          spService={props.spService}
          onChange={onChange}
        />
      );

    case FieldType.YesNo:
      return (
        <Toggle
          id={controlId}
          checked={value === true}
          onText="Yes"
          offText="No"
          disabled={disabled}
          ariaLabel={ariaLabel}
          onChange={(_event, checked) => onChange(checked === true)}
        />
      );

    case FieldType.Consent:
      return (
        <ConsentField
          consentText={field.consentText || ''}
          label={field.placeholder || 'I agree'}
          checked={value === true}
          disabled={disabled}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    // ----- scales -----

    case FieldType.Rating:
      return (
        <RatingField
          value={typeof value === 'number' ? (value as number) : undefined}
          max={field.maxRating || 5}
          icon={field.ratingIcon || 'star'}
          allowHalf={field.allowHalfRating === true}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Scale:
      return (
        <ScaleField
          value={typeof value === 'number' ? (value as number) : undefined}
          min={typeof field.min === 'number' ? field.min : 1}
          max={typeof field.max === 'number' ? field.max : 5}
          step={field.step || 1}
          lowLabel={field.lowLabel}
          highLabel={field.highLabel}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Slider: {
      const min = typeof field.min === 'number' ? field.min : 0;
      const max = typeof field.max === 'number' ? field.max : 10;
      const answered = typeof value === 'number';
      return (
        <div>
          <Slider
            min={min}
            max={max}
            step={field.step || 1}
            // an unanswered slider must not look answered: it parks at the
            // minimum but reports no value until the respondent moves it
            value={answered ? (value as number) : min}
            showValue={answered}
            disabled={disabled}
            ariaLabel={ariaLabel}
            valueFormat={(v) => (field.unitSuffix ? v + ' ' + field.unitSuffix : String(v))}
            onChange={(v) => onChange(v)}
          />
          {!answered && !disabled && (
            <div className={styles.selectionHint}>Drag the slider to answer</div>
          )}
        </div>
      );
    }

    case FieldType.Likert:
      return (
        <LikertField
          rows={field.likertRows || []}
          columns={field.likertColumns || []}
          value={value as ILikertValue | undefined}
          disabled={disabled}
          ariaLabel={ariaLabel}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Ranking:
      return (
        <RankingField
          choices={orderedChoices()}
          value={Array.isArray(value) ? (value as string[]) : undefined}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          onChange={onChange}
        />
      );

    // ----- contact & identity -----

    case FieldType.Hyperlink: {
      const link = (value as IHyperlinkValue) || { url: '', description: '' };
      return (
        <div className={styles.hyperlinkGroup}>
          <TextField
            id={controlId}
            value={link.url || ''}
            placeholder={field.placeholder || 'https://…'}
            iconProps={{ iconName: 'Link' }}
            disabled={disabled}
            type="url"
            ariaLabel={ariaLabel ? ariaLabel + ' — web address' : 'Web address'}
            aria-describedby={describedBy}
            aria-invalid={invalid ? true : undefined}
            onChange={(_event, v) => onChange({ ...link, url: v || '' })}
          />
          <TextField
            value={link.description || ''}
            placeholder="Display text (optional)"
            disabled={disabled}
            ariaLabel={ariaLabel ? ariaLabel + ' — display text' : 'Display text'}
            onChange={(_event, v) => onChange({ ...link, description: v || '' })}
          />
        </div>
      );
    }

    case FieldType.Address:
      return (
        <AddressField
          value={value as IAddressValue | string | undefined}
          disabled={disabled}
          required={field.required}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Person:
      return (
        <PersonField
          value={(value as IPersonInfo[]) || []}
          allowMultiple={field.allowMultiplePeople === true}
          allowGroups={field.allowGroups === true}
          placeholder={field.placeholder}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
          spService={props.spService}
        />
      );

    // ----- rich input -----

    case FieldType.FileUpload:
      return (
        <FileUploadField
          value={(value as IFormFile[]) || []}
          maxFiles={field.maxFiles || 3}
          maxFileSizeMb={field.maxFileSizeMb || 10}
          allowedExtensions={field.allowedExtensions || []}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    case FieldType.Signature:
      return (
        <SignatureField
          value={(value as string) || ''}
          disabled={disabled}
          ariaLabel={ariaLabel}
          ariaDescribedBy={describedBy}
          invalid={invalid}
          onChange={onChange}
        />
      );

    default:
      return (
        <TextField
          id={controlId}
          value={(value as string) || ''}
          placeholder={field.placeholder}
          disabled={disabled}
          ariaLabel={ariaLabel}
          aria-describedby={describedBy}
          onChange={(_event, v) => onChange(v)}
        />
      );
  }
};

const selectionHint = (
  min: number | undefined,
  max: number | undefined,
  count: number
): string => {
  const parts: string[] = [];
  if (typeof min === 'number' && typeof max === 'number') {
    parts.push('Choose between ' + min + ' and ' + max);
  } else if (typeof min === 'number') {
    parts.push('Choose at least ' + min);
  } else if (typeof max === 'number') {
    parts.push('Choose up to ' + max);
  }
  parts.push(count + ' selected');
  return parts.join(' · ');
};
