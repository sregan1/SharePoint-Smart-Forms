import * as React from 'react';
import styles from './FormRenderer.module.scss';
import { ILikertValue } from '../../models';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

export interface ILikertFieldProps {
  rows: string[];
  columns: string[];
  value: ILikertValue | undefined;
  disabled?: boolean;
  ariaLabel?: string;
  invalid?: boolean;
  onChange: (value: ILikertValue) => void;
}

/**
 * Matrix of statements (rows) rated on one shared scale (columns).
 *
 * Rendered as a real `<table>` with a radiogroup per row: screen readers get
 * proper row/column header association, which a div grid can't give you, and
 * each row's radios share a name so arrow keys move within the row rather than
 * across the whole matrix.
 */
export const LikertField: React.FunctionComponent<ILikertFieldProps> = (props) => {
  const rows = (props.rows || []).map((r) => (r || '').trim()).filter((r) => r.length > 0);
  const columns = (props.columns || []).map((c) => (c || '').trim()).filter((c) => c.length > 0);
  const answers = props.value || {};
  // stable prefix so two Likert questions on one page don't share radio names
  const groupId = React.useMemo(() => 'lk' + Math.random().toString(36).slice(2, 8), []);

  if (rows.length === 0 || columns.length === 0) {
    return null;
  }

  const select = (row: string, column: string): void => {
    props.onChange({ ...answers, [row]: column });
  };

  const answeredCount = rows.filter((row) => !!answers[row]).length;

  return (
    <div className={styles.likertWrap}>
      <table className={styles.likertTable} aria-label={props.ariaLabel}>
        <thead>
          <tr>
            <th scope="col">
              <span className={styles.likertRowLabel}>&nbsp;</span>
            </th>
            {columns.map((column) => (
              <th key={column} scope="col">
                {column}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={row + '-' + rowIndex}>
              <th scope="row">
                <span className={styles.likertRowLabel}>{row}</span>
              </th>
              {columns.map((column) => {
                const checked = answers[row] === column;
                return (
                  <td key={column}>
                    <button
                      type="button"
                      role="radio"
                      name={groupId + '-' + rowIndex}
                      aria-checked={checked}
                      aria-label={formatString(strings.Form_Likert_RadioAria, { row, column })}
                      disabled={props.disabled}
                      tabIndex={checked || (!answers[row] && column === columns[0]) ? 0 : -1}
                      className={checked ? styles.likertRadio + ' ' + styles.likertRadioOn : styles.likertRadio}
                      onClick={() => select(row, column)}
                      onKeyDown={(event) => {
                        if (props.disabled) {
                          return;
                        }
                        const current = columns.indexOf(answers[row]);
                        const from = current < 0 ? 0 : current;
                        if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
                          event.preventDefault();
                          select(row, columns[Math.min(columns.length - 1, from + 1)]);
                        } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
                          event.preventDefault();
                          select(row, columns[Math.max(0, from - 1)]);
                        }
                      }}
                    />
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      {answeredCount < rows.length && (
        <div className={props.invalid ? styles.likertMissing : styles.selectionHint}>
          {formatString(strings.Form_Likert_RowsAnswered, { answered: answeredCount, total: rows.length })}
        </div>
      )}
    </div>
  );
};
