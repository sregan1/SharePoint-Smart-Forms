import * as React from 'react';
import { Checkbox } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { sanitizeHtml } from '../../utils/sanitizeHtml';

export interface IConsentFieldProps {
  /** rich text terms shown above the checkbox */
  consentText: string;
  label: string;
  checked: boolean;
  disabled?: boolean;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (checked: boolean) => void;
}

/**
 * A single acknowledgement checkbox with terms above it.
 *
 * Distinct from Yes/No on purpose. A Yes/No toggle always holds a value, so
 * "required" can't mean anything for it; consent is genuinely unanswered until
 * ticked, and results read as "98% consented" rather than "98% Yes".
 */
export const ConsentField: React.FunctionComponent<IConsentFieldProps> = (props) => {
  const html = React.useMemo(() => sanitizeHtml(props.consentText || ''), [props.consentText]);
  const textId = React.useMemo(() => 'consent' + Math.random().toString(36).slice(2, 8), []);

  return (
    <div className={styles.consentBlock}>
      {html && (
        <div
          id={textId}
          className={styles.consentText}
          // sanitized above; owners author this, respondents only read it
          dangerouslySetInnerHTML={{ __html: html }}
        />
      )}
      <Checkbox
        label={props.label || 'I agree'}
        checked={props.checked === true}
        disabled={props.disabled}
        ariaDescribedBy={[html ? textId : '', props.ariaDescribedBy || ''].filter((v) => v).join(' ') || undefined}
        inputProps={{ 'aria-invalid': props.invalid ? true : undefined }}
        onChange={(_event, checked) => props.onChange(checked === true)}
      />
    </div>
  );
};
