import * as React from 'react';
import { TextField } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { IAddressValue } from '../../models';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

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

  const aria = (part: string): string =>
    props.ariaLabel ? formatString(strings.Form_Address_AriaWithLabel, { label: props.ariaLabel, part }) : part;

  return (
    <div
      className={styles.addressGrid}
      role="group"
      aria-label={props.ariaLabel}
      aria-describedby={props.ariaDescribedBy}
    >
      <TextField
        className={styles.addressWide}
        placeholder={strings.Form_Address_StreetPlaceholder}
        ariaLabel={aria(strings.Form_Address_StreetAria)}
        value={address.street || ''}
        disabled={props.disabled}
        required={props.required}
        aria-invalid={props.invalid ? true : undefined}
        autoComplete="address-line1"
        onChange={(_event, v) => set({ street: v })}
      />
      <TextField
        className={styles.addressWide}
        placeholder={strings.Form_Address_Street2Placeholder}
        ariaLabel={aria(strings.Form_Address_Street2Aria)}
        value={address.street2 || ''}
        disabled={props.disabled}
        autoComplete="address-line2"
        onChange={(_event, v) => set({ street2: v })}
      />
      <TextField
        placeholder={strings.Form_Address_CityPlaceholder}
        ariaLabel={aria(strings.Form_Address_CityAria)}
        value={address.city || ''}
        disabled={props.disabled}
        required={props.required}
        autoComplete="address-level2"
        onChange={(_event, v) => set({ city: v })}
      />
      <TextField
        placeholder={strings.Form_Address_StatePlaceholder}
        ariaLabel={aria(strings.Form_Address_StateAria)}
        value={address.state || ''}
        disabled={props.disabled}
        autoComplete="address-level1"
        onChange={(_event, v) => set({ state: v })}
      />
      <TextField
        placeholder={strings.Form_Address_PostalCodePlaceholder}
        ariaLabel={aria(strings.Form_Address_PostalCodeAria)}
        value={address.postalCode || ''}
        disabled={props.disabled}
        autoComplete="postal-code"
        onChange={(_event, v) => set({ postalCode: v })}
      />
      <TextField
        placeholder={strings.Form_Address_CountryPlaceholder}
        ariaLabel={aria(strings.Form_Address_CountryAria)}
        value={address.country || ''}
        disabled={props.disabled}
        autoComplete="country-name"
        onChange={(_event, v) => set({ country: v })}
      />
    </div>
  );
};
