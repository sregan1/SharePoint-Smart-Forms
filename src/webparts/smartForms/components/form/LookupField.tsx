import * as React from 'react';
import { Dropdown, IDropdownOption, Spinner, SpinnerSize } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { SharePointService } from '../../services/SharePointService';
import * as strings from 'SmartFormsWebPartStrings';

export interface ILookupFieldProps {
  listId: string;
  column: string;
  filter?: string;
  value: string | string[] | undefined;
  allowMultiple: boolean;
  placeholder?: string;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  spService: SharePointService;
  onChange: (value: string | string[] | undefined) => void;
}

/**
 * Options read live from another SharePoint list.
 *
 * The *display text* is what gets stored, not a list item id. A real SharePoint
 * Lookup column stores the id, which silently breaks when the source item is
 * renamed or deleted and can't be read at all by a respondent without access to
 * the source list. Storing the resolved text means a response stays meaningful
 * forever, at the cost of not tracking later renames — the right trade for a
 * form archive.
 */
export const LookupField: React.FunctionComponent<ILookupFieldProps> = (props) => {
  const [options, setOptions] = React.useState<string[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let cancelled = false;
    if (!props.listId || !props.column) {
      setOptions([]);
      setLoading(false);
      return undefined;
    }
    setLoading(true);
    setError('');
    props.spService
      .getLookupOptions(props.listId, props.column, props.filter)
      .then((values) => {
        if (cancelled) {
          return;
        }
        setOptions(values);
        setLoading(false);
      })
      .catch(() => {
        if (cancelled) {
          return;
        }
        setError(strings.Form_Lookup_LoadError);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [props.listId, props.column, props.filter]);

  if (loading) {
    return <Spinner size={SpinnerSize.small} label={strings.Form_Lookup_Loading} labelPosition="right" />;
  }

  if (error) {
    return <div className={styles.fieldError}>{error}</div>;
  }

  // a stored answer whose source item has since been removed must still show,
  // or the respondent's own answer would vanish from the form
  const selected = Array.isArray(props.value) ? props.value : props.value ? [String(props.value)] : [];
  const orphans = selected.filter((v) => options.indexOf(v) === -1);
  const dropdownOptions: IDropdownOption[] = options
    .concat(orphans)
    .map((value) => ({ key: value, text: value }));

  if (props.allowMultiple) {
    return (
      <Dropdown
        multiSelect={true}
        options={dropdownOptions}
        selectedKeys={selected}
        placeholder={props.placeholder || strings.Form_Lookup_SelectOptionsPlaceholder}
        disabled={props.disabled}
        ariaLabel={props.ariaLabel}
        aria-describedby={props.ariaDescribedBy}
        errorMessage={props.invalid ? ' ' : undefined}
        onChange={(_event, option) => {
          if (!option) {
            return;
          }
          const key = String(option.key);
          props.onChange(
            option.selected ? selected.concat([key]) : selected.filter((s) => s !== key)
          );
        }}
      />
    );
  }

  return (
    <Dropdown
      options={dropdownOptions}
      selectedKey={selected[0] || null}
      placeholder={props.placeholder || strings.Form_Lookup_SelectOptionPlaceholder}
      disabled={props.disabled}
      ariaLabel={props.ariaLabel}
      aria-describedby={props.ariaDescribedBy}
      errorMessage={props.invalid ? ' ' : undefined}
      onChange={(_event, option) => props.onChange(option ? String(option.key) : undefined)}
    />
  );
};
