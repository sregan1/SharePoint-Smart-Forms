import * as React from 'react';
import styles from './FormRenderer.module.scss';

export interface IScaleFieldProps {
  value: number | undefined;
  min: number;
  max: number;
  step: number;
  lowLabel?: string;
  highLabel?: string;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: number) => void;
}

/**
 * A row of numbered buttons — the generalized opinion scale. NPS is this control
 * with min 0 / max 10 and "not at all likely / extremely likely" end labels,
 * rather than a separate field type.
 *
 * Implemented as a real radiogroup: arrow keys move between options and only the
 * selected (or first, when nothing is chosen) button is in the tab order, which
 * is what the ARIA radiogroup pattern requires. The previous NPS control set the
 * roles but left every button tabbable with no keyboard selection.
 */
export const ScaleField: React.FunctionComponent<IScaleFieldProps> = (props) => {
  const { min, max, step } = props;
  const refs = React.useRef<(HTMLButtonElement | null)[]>([]);

  const scores = React.useMemo(() => {
    const values: number[] = [];
    const safeStep = step > 0 ? step : 1;
    // guard against a pathological range producing thousands of buttons
    const limit = 61;
    for (let value = min; value <= max && values.length < limit; value += safeStep) {
      values.push(Math.round(value * 100) / 100);
    }
    return values;
  }, [min, max, step]);

  const selectedIndex = props.value === undefined ? -1 : scores.indexOf(props.value);
  // the radiogroup needs exactly one tab stop
  const tabIndexOf = (index: number): number => {
    if (selectedIndex >= 0) {
      return index === selectedIndex ? 0 : -1;
    }
    return index === 0 ? 0 : -1;
  };

  const focusAndSelect = (index: number): void => {
    const clamped = Math.max(0, Math.min(scores.length - 1, index));
    const target = refs.current[clamped];
    if (target) {
      target.focus();
    }
    props.onChange(scores[clamped]);
  };

  const handleKeyDown = (event: React.KeyboardEvent, index: number): void => {
    if (props.disabled) {
      return;
    }
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        focusAndSelect(index + 1);
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        focusAndSelect(index - 1);
        break;
      case 'Home':
        event.preventDefault();
        focusAndSelect(0);
        break;
      case 'End':
        event.preventDefault();
        focusAndSelect(scores.length - 1);
        break;
      default:
        break;
    }
  };

  return (
    <div>
      <div
        className={styles.scaleRow}
        role="radiogroup"
        aria-label={props.ariaLabel}
        aria-describedby={props.ariaDescribedBy}
        aria-invalid={props.invalid ? true : undefined}
      >
        {scores.map((score, index) => (
          <button
            key={score}
            ref={(element) => {
              refs.current[index] = element;
            }}
            type="button"
            role="radio"
            aria-checked={props.value === score}
            tabIndex={tabIndexOf(index)}
            disabled={props.disabled}
            className={props.value === score ? styles.scaleButtonSelected : styles.scaleButton}
            onClick={() => props.onChange(score)}
            onKeyDown={(event) => handleKeyDown(event, index)}
          >
            {score}
          </button>
        ))}
      </div>
      {(props.lowLabel || props.highLabel) && (
        <div className={styles.scaleLabels}>
          <span>{props.lowLabel || ''}</span>
          <span>{props.highLabel || ''}</span>
        </div>
      )}
    </div>
  );
};
