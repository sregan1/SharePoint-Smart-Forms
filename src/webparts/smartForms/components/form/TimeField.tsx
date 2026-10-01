import * as React from 'react';
import { ComboBox, IComboBoxOption } from '@fluentui/react';
import * as strings from 'SmartFormsWebPartStrings';

export interface ITimeFieldProps {
  /** a Date whose time-of-day carries the answer */
  value: Date | undefined;
  /** minute granularity of the offered options */
  stepMinutes: number;
  placeholder?: string;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: Date | undefined) => void;
}

const pad = (n: number): string => (n < 10 ? '0' + n : String(n));

/** "9:30 AM" for a minutes-past-midnight value. */
export const formatMinutes = (minutes: number): string => {
  const hours24 = Math.floor(minutes / 60) % 24;
  const mins = minutes % 60;
  const hour12 = hours24 % 12 === 0 ? 12 : hours24 % 12;
  return hour12 + ':' + pad(mins) + ' ' + (hours24 < 12 ? 'AM' : 'PM');
};

/**
 * Parse free text into minutes past midnight. Accepts "9", "9:30", "930",
 * "9.30", "9pm", "21:15" — people type times in all of these and rejecting them
 * for punctuation is needless friction.
 */
export const parseTimeText = (text: string): number | undefined => {
  const cleaned = (text || '').trim().toLowerCase();
  if (cleaned.length === 0) {
    return undefined;
  }
  const meridiem = /(am|pm)/.exec(cleaned);
  const digits = cleaned.replace(/[^\d]/g, '');
  if (digits.length === 0) {
    return undefined;
  }

  let hours: number;
  let minutes = 0;
  if (digits.length <= 2) {
    hours = Number(digits);
  } else if (digits.length === 3) {
    hours = Number(digits.slice(0, 1));
    minutes = Number(digits.slice(1));
  } else {
    hours = Number(digits.slice(0, 2));
    minutes = Number(digits.slice(2, 4));
  }

  if (isNaN(hours) || isNaN(minutes) || minutes > 59) {
    return undefined;
  }
  if (meridiem) {
    if (hours > 12) {
      return undefined;
    }
    if (meridiem[1] === 'pm' && hours < 12) {
      hours += 12;
    }
    if (meridiem[1] === 'am' && hours === 12) {
      hours = 0;
    }
  }
  if (hours > 23) {
    return undefined;
  }
  return hours * 60 + minutes;
};

/** A Date carrying only a time of day, anchored to today's date. */
const dateWithMinutes = (minutes: number, basedOn: Date | undefined): Date => {
  const base = basedOn instanceof Date && !isNaN(basedOn.getTime()) ? new Date(basedOn.getTime()) : new Date();
  base.setHours(Math.floor(minutes / 60), minutes % 60, 0, 0);
  return base;
};

/**
 * Time-of-day picker. An editable ComboBox rather than a Dropdown so an odd time
 * ("07:42") can be typed instead of forcing everyone onto the step grid.
 */
export const TimeField: React.FunctionComponent<ITimeFieldProps> = (props) => {
  const step = props.stepMinutes && props.stepMinutes > 0 ? props.stepMinutes : 15;

  const options: IComboBoxOption[] = React.useMemo(() => {
    const result: IComboBoxOption[] = [];
    for (let minutes = 0; minutes < 24 * 60; minutes += step) {
      result.push({ key: String(minutes), text: formatMinutes(minutes) });
    }
    return result;
  }, [step]);

  const current =
    props.value instanceof Date && !isNaN(props.value.getTime())
      ? props.value.getHours() * 60 + props.value.getMinutes()
      : undefined;

  // an off-grid time still needs to appear as the selected option
  const allOptions =
    current !== undefined && options.filter((o) => o.key === String(current)).length === 0
      ? options.concat([{ key: String(current), text: formatMinutes(current) }]).sort(
          (a, b) => Number(a.key) - Number(b.key)
        )
      : options;

  const commit = (minutes: number | undefined): void => {
    props.onChange(minutes === undefined ? undefined : dateWithMinutes(minutes, props.value));
  };

  return (
    <ComboBox
      allowFreeform={true}
      autoComplete="on"
      options={allOptions}
      selectedKey={current === undefined ? null : String(current)}
      text={current === undefined ? '' : formatMinutes(current)}
      placeholder={props.placeholder || strings.Form_Time_Placeholder}
      disabled={props.disabled}
      ariaLabel={props.ariaLabel}
      errorMessage={props.invalid ? ' ' : undefined}
      useComboBoxAsMenuWidth={true}
      onChange={(_event, option, _index, value) => {
        if (option) {
          commit(Number(option.key));
          return;
        }
        if (value !== undefined) {
          const parsed = parseTimeText(value);
          if (parsed !== undefined) {
            commit(parsed);
          } else if (value.trim().length === 0) {
            commit(undefined);
          }
        }
      }}
    />
  );
};
