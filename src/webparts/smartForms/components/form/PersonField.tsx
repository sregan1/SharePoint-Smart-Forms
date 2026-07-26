import * as React from 'react';
import { IPersonaProps, NormalPeoplePicker } from '@fluentui/react';
import { IPersonInfo } from '../../models';
import { SharePointService } from '../../services/SharePointService';

export interface IPersonFieldProps {
  value: IPersonInfo[];
  allowMultiple: boolean;
  allowGroups?: boolean;
  placeholder?: string;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (value: IPersonInfo[]) => void;
  spService: SharePointService;
}

const toPersona = (person: IPersonInfo): IPersonaProps => ({
  key: person.loginName,
  text: person.displayName,
  secondaryText: person.email,
  // itemProp round-trips the principal kind through the persona so a group
  // selection isn't silently downgraded to a user on the way back out
  itemProp: person.isGroup ? 'group' : 'user'
});

const toPerson = (persona: IPersonaProps): IPersonInfo => ({
  loginName: String(persona.key),
  displayName: persona.text || '',
  email: persona.secondaryText,
  isGroup: persona.itemProp === 'group'
});

export const PersonField: React.FunctionComponent<IPersonFieldProps> = (props) => {
  const resolveSuggestions = async (
    filter: string,
    selected?: IPersonaProps[]
  ): Promise<IPersonaProps[]> => {
    try {
      const people = await props.spService.searchPeople(filter, props.allowGroups === true);
      const selectedKeys = (selected || []).map((s) => String(s.key));
      return people.filter((p) => selectedKeys.indexOf(p.loginName) === -1).map(toPersona);
    } catch {
      return [];
    }
  };

  const hasValue = props.value && props.value.length > 0;

  return (
    <NormalPeoplePicker
      onResolveSuggestions={resolveSuggestions}
      selectedItems={(props.value || []).map(toPersona)}
      onChange={(items) => props.onChange((items || []).map(toPerson))}
      itemLimit={props.allowMultiple ? 20 : 1}
      disabled={props.disabled}
      inputProps={{
        placeholder: hasValue ? '' : props.placeholder || 'Start typing a name…',
        'aria-label': props.ariaLabel,
        'aria-describedby': props.ariaDescribedBy,
        'aria-invalid': props.invalid ? true : undefined
      }}
      pickerSuggestionsProps={{
        suggestionsHeaderText: props.allowGroups ? 'Suggested people and groups' : 'Suggested people',
        noResultsFoundText: props.allowGroups ? 'No people or groups found' : 'No people found',
        loadingText: 'Searching…'
      }}
      resolveDelay={300}
    />
  );
};
