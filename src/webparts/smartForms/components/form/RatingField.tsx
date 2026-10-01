import * as React from 'react';
import { Icon, IconButton } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { RatingIcon } from '../../models';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

export interface IRatingFieldProps {
  value: number | undefined;
  max: number;
  icon: RatingIcon;
  allowHalf: boolean;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: number | undefined) => void;
}

const ICONS: { [key: string]: { on: string; off: string } } = {
  star: { on: 'FavoriteStarFill', off: 'FavoriteStar' },
  heart: { on: 'HeartFill', off: 'Heart' },
  like: { on: 'LikeSolid', off: 'Like' }
};

/**
 * Star / heart / thumb rating with optional half steps.
 *
 * Fluent 8's `Rating` can't do half values or swap the glyph, so this is a
 * purpose-built radiogroup. Half steps are rendered by overlaying a 50%-width
 * clipped "on" glyph over the "off" one, which keeps a single element per point
 * on the scale and avoids fractional-width layout drift.
 */
export const RatingField: React.FunctionComponent<IRatingFieldProps> = (props) => {
  const max = Math.max(1, Math.min(props.max || 5, 10));
  const glyphs = ICONS[props.icon] || ICONS.star;
  const value = typeof props.value === 'number' ? props.value : 0;
  const refs = React.useRef<(HTMLButtonElement | null)[]>([]);
  const [hover, setHover] = React.useState<number>(0);

  const display = hover > 0 ? hover : value;
  const step = props.allowHalf ? 0.5 : 1;

  const points: number[] = [];
  for (let i = 1; i <= max; i++) {
    points.push(i);
  }

  const setValue = (next: number): void => {
    // clicking the current value clears it, which is the only way to unset a
    // rating that isn't required
    props.onChange(next === value ? undefined : next);
  };

  const nudge = (delta: number): void => {
    const next = Math.max(0, Math.min(max, Math.round((value + delta) / step) * step));
    props.onChange(next === 0 ? undefined : next);
  };

  const handleKeyDown = (event: React.KeyboardEvent): void => {
    if (props.disabled) {
      return;
    }
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowUp':
        event.preventDefault();
        nudge(step);
        break;
      case 'ArrowLeft':
      case 'ArrowDown':
        event.preventDefault();
        nudge(-step);
        break;
      case 'Home':
        event.preventDefault();
        props.onChange(step);
        break;
      case 'End':
        event.preventDefault();
        props.onChange(max);
        break;
      case 'Delete':
      case 'Backspace':
        event.preventDefault();
        props.onChange(undefined);
        break;
      default:
        break;
    }
  };

  /** Which half of the glyph a pointer event landed on. */
  const valueFromPointer = (event: React.MouseEvent, point: number): number => {
    if (!props.allowHalf) {
      return point;
    }
    const bounds = (event.currentTarget as HTMLElement).getBoundingClientRect();
    return event.clientX - bounds.left < bounds.width / 2 ? point - 0.5 : point;
  };

  return (
    <div
      className={styles.ratingRow}
      role="slider"
      aria-label={props.ariaLabel}
      aria-describedby={props.ariaDescribedBy}
      aria-valuemin={0}
      aria-valuemax={max}
      aria-valuenow={value}
      aria-valuetext={value > 0 ? formatString(strings.Form_Rating_ValueText, { value, max }) : strings.Form_Rating_NotRated}
      aria-invalid={props.invalid ? true : undefined}
      aria-disabled={props.disabled ? true : undefined}
      tabIndex={props.disabled ? -1 : 0}
      onKeyDown={handleKeyDown}
      onMouseLeave={() => setHover(0)}
    >
      {points.map((point, index) => {
        const filled = display >= point;
        const half = !filled && display >= point - 0.5;
        return (
          <button
            key={point}
            ref={(element) => {
              refs.current[index] = element;
            }}
            type="button"
            tabIndex={-1}
            disabled={props.disabled}
            aria-hidden={true}
            className={filled ? styles.ratingButton + ' ' + styles.ratingButtonOn : styles.ratingButton}
            onMouseMove={(event) => !props.disabled && setHover(valueFromPointer(event, point))}
            onClick={(event) => !props.disabled && setValue(valueFromPointer(event, point))}
          >
            {half ? (
              <span className={styles.ratingHalfWrap}>
                <Icon iconName={glyphs.off} />
                <span className={styles.ratingHalfOverlay}>
                  <Icon iconName={glyphs.on} />
                </span>
              </span>
            ) : (
              <Icon iconName={filled ? glyphs.on : glyphs.off} />
            )}
          </button>
        );
      })}
      {value > 0 && <span className={styles.ratingValue}>{value}</span>}
      {value > 0 && !props.disabled && (
        <IconButton
          className={styles.ratingClear}
          iconProps={{ iconName: 'Clear' }}
          title={strings.Form_Rating_Clear}
          ariaLabel={strings.Form_Rating_Clear}
          tabIndex={-1}
          onClick={() => props.onChange(undefined)}
        />
      )}
    </div>
  );
};
