import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { IImageChoiceOption } from '../../models';

export interface IImageChoiceFieldProps {
  options: IImageChoiceOption[];
  value: string | string[] | undefined;
  allowMultiple: boolean;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: string | string[] | undefined) => void;
}

/**
 * Picture tiles. Answers are stored as the option *label*, so results stay
 * readable and an image URL can be swapped without rewriting history.
 *
 * Roles follow the selection mode: a single-select group is a radiogroup, a
 * multi-select group is a plain group of checkboxes — using radio semantics for
 * a multi-select would misreport the control to a screen reader.
 */
export const ImageChoiceField: React.FunctionComponent<IImageChoiceFieldProps> = (props) => {
  const options = (props.options || []).filter((o) => (o.label || '').trim().length > 0);
  const selected = Array.isArray(props.value)
    ? (props.value as string[])
    : props.value
      ? [String(props.value)]
      : [];

  const toggle = (label: string): void => {
    if (props.disabled) {
      return;
    }
    if (props.allowMultiple) {
      props.onChange(
        selected.indexOf(label) !== -1 ? selected.filter((s) => s !== label) : selected.concat([label])
      );
      return;
    }
    props.onChange(selected[0] === label ? undefined : label);
  };

  if (options.length === 0) {
    return null;
  }

  return (
    <div
      className={styles.imageChoiceGrid}
      role={props.allowMultiple ? 'group' : 'radiogroup'}
      aria-label={props.ariaLabel}
      aria-describedby={props.ariaDescribedBy}
      aria-invalid={props.invalid ? true : undefined}
    >
      {options.map((option, index) => {
        const isSelected = selected.indexOf(option.label) !== -1;
        return (
          <button
            key={option.label + '-' + index}
            type="button"
            role={props.allowMultiple ? 'checkbox' : 'radio'}
            aria-checked={isSelected}
            disabled={props.disabled}
            className={isSelected ? styles.imageChoiceTileSelected : styles.imageChoiceTile}
            onClick={() => toggle(option.label)}
          >
            {option.imageUrl ? (
              <img className={styles.imageChoiceThumb} src={option.imageUrl} alt="" loading="lazy" />
            ) : (
              <span className={styles.imageChoicePlaceholder} aria-hidden={true}>
                <Icon iconName="FileImage" />
              </span>
            )}
            <span className={styles.imageChoiceLabel}>{option.label}</span>
          </button>
        );
      })}
    </div>
  );
};
