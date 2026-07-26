import * as React from 'react';
import { TextField } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { IAddressValue } from '../../models';

export interface IAddressFieldProps {
  value: IAddressValue | string | undefined;
  disabled?: boolean;
  required?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: IAddressValue) => void;
}

/**
 * Structured postal address.
 *
 * Stored as one readable line in a Note column rather than six columns: a single
 * question that provisioned six SharePoint fields would make the list view and
 * every export unmanageable, and address parts are never analyzed
 * independently. The parts are kept separate in form state so the inputs behave
 * properly, then joined on submit.
 */
export const AddressField: React.FunctionComponent<IAddressFieldProps> = (props) => {
  // a value read back from SharePoint arrives as the formatted line; only the
  // live form holds the structured shape
  const address: IAddressValue =
    typeof props.value === 'string' || !props.value ? {} : (props.value as IAddressValue);

  const set = (patch: Partial<IAddressValue>): void => {
    props.onChange({ ...address, ...patch });
  };

  const label = props.ariaLabel ? props.ariaLabel + ' — ' : '';

  return (
    <div
      className={styles.addressGrid}
      role="group"
      aria-label={props.ariaLabel}
      aria-describedby={props.ariaDescribedBy}
    >
      <TextField
        className={styles.addressWide}
        placeholder="Street address"
        ariaLabel={label + 'street address'}
        value={address.street || ''}
        disabled={props.disabled}
        required={props.required}
        aria-invalid={props.invalid ? true : undefined}
        autoComplete="address-line1"
        onChange={(_event, v) => set({ street: v })}
      />
      <TextField
        className={styles.addressWide}
        placeholder="Apartment, suite, etc. (optional)"
        ariaLabel={label + 'address line 2'}
        value={address.street2 || ''}
        disabled={props.disabled}
        autoComplete="address-line2"
        onChange={(_event, v) => set({ street2: v })}
      />
      <TextField
        placeholder="City"
        ariaLabel={label + 'city'}
        value={address.city || ''}
        disabled={props.disabled}
        required={props.required}
        autoComplete="address-level2"
        onChange={(_event, v) => set({ city: v })}
      />
      <TextField
        placeholder="State / county"
        ariaLabel={label + 'state or county'}
        value={address.state || ''}
        disabled={props.disabled}
        autoComplete="address-level1"
        onChange={(_event, v) => set({ state: v })}
      />
      <TextField
        placeholder="Postcode / ZIP"
        ariaLabel={label + 'postcode'}
        value={address.postalCode || ''}
        disabled={props.disabled}
        autoComplete="postal-code"
        onChange={(_event, v) => set({ postalCode: v })}
      />
      <TextField
        placeholder="Country"
        ariaLabel={label + 'country'}
        value={address.country || ''}
        disabled={props.disabled}
        autoComplete="country-name"
        onChange={(_event, v) => set({ country: v })}
      />
    </div>
  );
};
