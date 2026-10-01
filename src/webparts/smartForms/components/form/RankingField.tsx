import * as React from 'react';
import { Icon, IconButton } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { reorder, useDragList } from '../../hooks/useDragList';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

export interface IRankingFieldProps {
  /** options in their designed order */
  choices: string[];
  /** current order; falls back to the designed order */
  value: string[] | undefined;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  onChange: (ordered: string[]) => void;
}

/**
 * Reconcile a saved order with the current option list: keep the respondent's
 * ordering for options that still exist, append any new options at the end.
 */
const effectiveOrder = (choices: string[], value: string[] | undefined): string[] => {
  const kept = (Array.isArray(value) ? value : []).filter((v) => choices.indexOf(v) !== -1);
  const missing = choices.filter((c) => kept.indexOf(c) === -1);
  return kept.concat(missing);
};

/** Drag-to-reorder (with keyboard-friendly up/down buttons) ranking control. */
export const RankingField: React.FunctionComponent<IRankingFieldProps> = (props) => {
  const order = effectiveOrder(props.choices, props.value);
  const [announcement, setAnnouncement] = React.useState<string>('');

  const applyMove = (from: number, to: number): void => {
    const next = reorder(order, from, to);
    props.onChange(next);
    setAnnouncement(formatString(strings.Form_Ranking_Moved, { item: next[to], position: to + 1, total: next.length }));
  };

  const drag = useDragList({
    count: order.length,
    enabled: !props.disabled,
    dataTransferText: (index) => order[index],
    onReorder: applyMove
  });

  const classFor = (index: number): string => {
    if (drag.isDragging(index)) {
      return styles.rankItemDragging;
    }
    return drag.dropIndicator(index) ? styles.rankItemOver : styles.rankItem;
  };

  return (
    <div>
      <div className={styles.rankList} role="list" aria-label={props.ariaLabel} aria-describedby={props.ariaDescribedBy}>
        {order.map((option, index) => (
          <div
            // options can legitimately repeat their text, so the index has to be
            // part of the key — a bare label collided and swapped the wrong rows
            key={option + '::' + index}
            role="listitem"
            className={classFor(index)}
            {...drag.rowProps(index)}
          >
            <span className={styles.rankBadge}>{index + 1}</span>
            <span className={styles.rankLabel} title={option}>
              {option}
            </span>
            {!props.disabled && <Icon iconName="GripperDotsVertical" className={styles.rankGrip} />}
            {!props.disabled && (
              <span className={styles.rankButtons}>
                <IconButton
                  iconProps={{ iconName: 'Up' }}
                  title={strings.Form_Ranking_MoveUp}
                  ariaLabel={formatString(strings.Form_Ranking_MoveItemUpAria, { item: option })}
                  disabled={index === 0}
                  onClick={() => applyMove(index, index - 1)}
                />
                <IconButton
                  iconProps={{ iconName: 'Down' }}
                  title={strings.Form_Ranking_MoveDown}
                  ariaLabel={formatString(strings.Form_Ranking_MoveItemDownAria, { item: option })}
                  disabled={index === order.length - 1}
                  onClick={() => applyMove(index, index + 1)}
                />
              </span>
            )}
          </div>
        ))}
      </div>
      {/* keyboard reordering is silent without this */}
      <div aria-live="polite" role="status" className={styles.hiddenInput}>
        {announcement}
      </div>
    </div>
  );
};
