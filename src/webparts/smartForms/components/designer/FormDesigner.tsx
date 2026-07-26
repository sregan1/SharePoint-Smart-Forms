import * as React from 'react';
import {
  ActionButton,
  DefaultButton,
  Icon,
  IconButton,
  IContextualMenuItem,
  IContextualMenuProps,
  ITextField,
  SearchBox,
  TextField,
  Toggle
} from '@fluentui/react';
import styles from './FormDesigner.module.scss';
import {
  defaultsForType,
  FIELD_PRESETS,
  FIELD_TYPE_META,
  FieldCategory,
  FieldType,
  FieldWidth,
  getFieldTypeMeta,
  IFieldTypeMeta,
  IFormDefinition,
  IFormField,
  IFormSection,
  IImageChoiceOption,
  isInputType,
  newId,
  TYPE_SPECIFIC_KEYS
} from '../../models';
import {
  allFields,
  buildNumberMap,
  generateInternalName,
  IFormIssue,
  pruneDanglingConditions,
  validateDefinition
} from '../../utils/formUtils';
import { reorder, useDragList } from '../../hooks/useDragList';
import { SharePointService } from '../../services/SharePointService';
import { FieldControl } from '../form/FieldControl';
import { ContentBlock } from '../form/ContentBlock';
import { FieldEditorPanel } from './FieldEditorPanel';

export interface IFormDesignerProps {
  definition: IFormDefinition;
  spService: SharePointService;
  /** field the shell asked us to scroll to and open (from a pre-flight issue) */
  focusFieldId?: string;
  onChange: (definition: IFormDefinition) => void;
  onFocusHandled?: () => void;
}

const clone = <T,>(value: T): T => JSON.parse(JSON.stringify(value));

/** The types surfaced directly on the "Add new" button, like Microsoft Forms. */
const POPULAR_TYPES: FieldType[] = [
  FieldType.Choice,
  FieldType.Text,
  FieldType.Rating,
  FieldType.Date,
  FieldType.MultilineText
];

const CATEGORY_ORDER: FieldCategory[] = ['Text', 'Choice', 'Scales', 'Date', 'People', 'Advanced', 'Layout'];

/** Types whose options are edited as a plain string list on the card. */
const LIST_OPTION_TYPES: FieldType[] = [FieldType.Choice, FieldType.Ranking];

/**
 * WYSIWYG form editor: the canvas *is* the form. Click a question to edit it
 * in place — no separate designer view, no edit panels for the essentials.
 */
export const FormDesigner: React.FunctionComponent<IFormDesignerProps> = (props) => {
  const { definition } = props;
  const [activeFieldId, setActiveFieldId] = React.useState<string | undefined>(undefined);
  const [panelFieldId, setPanelFieldId] = React.useState<string | undefined>(undefined);
  const [outlineFilter, setOutlineFilter] = React.useState<string>('');

  // keyboard-first option editing: refs let Enter/Tab/Backspace move focus
  // between option rows of the active question
  const optionRefs = React.useRef<{ [key: string]: ITextField | null }>({});
  const [focusOption, setFocusOption] = React.useState<string | undefined>(undefined);
  const cardRefs = React.useRef<{ [fieldId: string]: HTMLDivElement | null }>({});

  React.useEffect(() => {
    if (focusOption === undefined) {
      return;
    }
    const target = optionRefs.current[focusOption];
    if (target) {
      target.focus();
    }
    setFocusOption(undefined);
  }, [focusOption]);

  // jump to a question the pre-flight check flagged
  React.useEffect(() => {
    if (!props.focusFieldId) {
      return;
    }
    setActiveFieldId(props.focusFieldId);
    const card = cardRefs.current[props.focusFieldId];
    if (card) {
      card.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
    if (props.onFocusHandled) {
      props.onFocusHandled();
    }
  }, [props.focusFieldId]);

  const issues = React.useMemo(() => validateDefinition(definition), [definition]);
  const issuesByField = React.useMemo(() => {
    const map: { [fieldId: string]: IFormIssue[] } = {};
    issues.forEach((issue) => {
      if (!issue.fieldId) {
        return;
      }
      (map[issue.fieldId] = map[issue.fieldId] || []).push(issue);
    });
    return map;
  }, [issues]);

  const orderedFields = React.useMemo(() => allFields(definition), [definition]);
  const numberMap = React.useMemo(() => buildNumberMap(orderedFields), [orderedFields]);

  // ----- mutation helpers -----

  const update = (mutate: (next: IFormDefinition) => void): void => {
    const next = clone(definition);
    mutate(next);
    props.onChange(next);
  };

  const findFieldIn = (def: IFormDefinition, fieldId: string): IFormField | undefined => {
    let found: IFormField | undefined;
    def.sections.forEach((s) =>
      s.fields.forEach((f) => {
        if (f.id === fieldId) {
          found = f;
        }
      })
    );
    return found;
  };

  const patchField = (fieldId: string, patch: Partial<IFormField>): void => {
    update((next) => {
      const f = findFieldIn(next, fieldId);
      if (!f) {
        return;
      }
      Object.keys(patch).forEach((key) => {
        const value = (patch as Record<string, unknown>)[key];
        const target = f as unknown as Record<string, unknown>;
        if (value === undefined) {
          delete target[key];
        } else {
          target[key] = value;
        }
      });
    });
  };

  /**
   * Switch a question's type.
   *
   * Every per-type key is cleared before the new type's defaults are seeded.
   * Patching only `type` used to leave the old configuration behind — a Text
   * turned into a Choice arrived with no options at all and rendered an empty
   * option group, and a Choice turned into a Slider had no min/max/step.
   */
  const changeFieldType = (fieldId: string, type: FieldType): void => {
    update((next) => {
      const field = findFieldIn(next, fieldId);
      if (!field) {
        return;
      }
      const target = field as unknown as Record<string, unknown>;
      TYPE_SPECIFIC_KEYS.forEach((key) => {
        delete target[key as string];
      });
      field.type = type;
      const defaults = defaultsForType(type);
      Object.keys(defaults).forEach((key) => {
        target[key] = (defaults as Record<string, unknown>)[key];
      });
    });
  };

  const applyPreset = (sectionId: string, presetKey: string): void => {
    const preset = FIELD_PRESETS.filter((p) => p.key === presetKey)[0];
    if (!preset) {
      return;
    }
    addField(sectionId, preset.type, preset.defaults);
  };

  const addField = (sectionId: string, type: FieldType, extraDefaults?: Partial<IFormField>): void => {
    const id = newId();
    update((next) => {
      const target = next.sections.filter((s) => s.id === sectionId)[0];
      if (!target) {
        return;
      }
      const field: IFormField = {
        id,
        internalName: generateInternalName(
          getFieldTypeMeta(type).label,
          allFields(next).map((f) => f.internalName)
        ),
        title: '',
        type,
        required: false,
        provisioned: false,
        ...defaultsForType(type),
        ...(extraDefaults || {})
      };
      target.fields.push(field);
    });
    setActiveFieldId(id);
  };

  const duplicateField = (section: IFormSection, field: IFormField): void => {
    const id = newId();
    update((next) => {
      const target = next.sections.filter((s) => s.id === section.id)[0];
      if (!target) {
        return;
      }
      const index = target.fields.map((f) => f.id).indexOf(field.id);
      const copy: IFormField = clone(field);
      copy.id = id;
      copy.internalName = generateInternalName(
        copy.title || 'Question',
        allFields(next).map((f) => f.internalName)
      );
      copy.provisioned = false;
      target.fields.splice(index + 1, 0, copy);
    });
    setActiveFieldId(id);
  };

  const removeField = (fieldId: string): void => {
    update((next) => {
      next.sections.forEach((section) => {
        section.fields = section.fields.filter((f) => f.id !== fieldId);
      });
      pruneDanglingConditions(next);
    });
    if (activeFieldId === fieldId) {
      setActiveFieldId(undefined);
    }
    if (panelFieldId === fieldId) {
      setPanelFieldId(undefined);
    }
  };

  const moveFieldWithin = (sectionId: string, from: number, to: number): void => {
    update((next) => {
      const target = next.sections.filter((s) => s.id === sectionId)[0];
      if (!target) {
        return;
      }
      target.fields = reorder(target.fields, from, to);
    });
  };

  /**
   * Move a field to an absolute position in the flattened question order,
   * which may land it in a different section.
   *
   * Each entry carries the section it belongs to. The moved field adopts the
   * section of whatever now sits at its destination, then entries are regrouped
   * by section — so a drag across a section boundary lands where it was dropped
   * and every other field keeps its own section.
   */
  const moveFieldGlobally = (from: number, to: number): void => {
    update((next) => {
      const entries: { sectionIndex: number; field: IFormField }[] = [];
      next.sections.forEach((section, sectionIndex) =>
        section.fields.forEach((field) => entries.push({ sectionIndex, field }))
      );
      if (from < 0 || from >= entries.length || to < 0 || to >= entries.length || from === to) {
        return;
      }

      const [moved] = entries.splice(from, 1);
      // whatever occupies the destination decides which section the drop lands in
      const neighbour = entries[Math.min(to, entries.length - 1)];
      moved.sectionIndex = neighbour ? neighbour.sectionIndex : moved.sectionIndex;
      entries.splice(to, 0, moved);

      next.sections.forEach((section) => {
        section.fields = [];
      });
      entries.forEach((entry) => {
        const index = Math.max(0, Math.min(entry.sectionIndex, next.sections.length - 1));
        next.sections[index].fields.push(entry.field);
      });
    });
  };

  const addSection = (): void => {
    update((next) => {
      next.sections.push({ id: newId(), title: '', description: '', fields: [] });
    });
  };

  const moveSection = (section: IFormSection, delta: number): void => {
    update((next) => {
      const index = next.sections.map((s) => s.id).indexOf(section.id);
      const newIndex = index + delta;
      if (newIndex < 0 || newIndex >= next.sections.length) {
        return;
      }
      next.sections = reorder(next.sections, index, newIndex);
    });
  };

  const removeSection = (section: IFormSection): void => {
    update((next) => {
      if (next.sections.length <= 1) {
        return;
      }
      const index = next.sections.map((s) => s.id).indexOf(section.id);
      const doomed = next.sections[index];
      next.sections = next.sections.filter((s) => s.id !== section.id);
      // questions in a deleted section take their branching rules with them, so
      // any rule pointing *at* them has to go too — otherwise the rule dangles
      // and its owner becomes permanently visible
      if (doomed) {
        pruneDanglingConditions(next);
      }
    });
  };

  // ----- option editing -----

  const optionKey = (fieldId: string, index: number): string => fieldId + ':' + index;

  const setOption = (field: IFormField, index: number, value: string): void => {
    const choices = (field.choices || []).slice();
    choices[index] = value;
    patchField(field.id, { choices });
  };

  const insertOption = (field: IFormField, at: number): void => {
    const choices = (field.choices || []).slice();
    choices.splice(at, 0, '');
    patchField(field.id, { choices });
    setFocusOption(optionKey(field.id, at));
  };

  const removeOption = (field: IFormField, index: number): void => {
    const choices = (field.choices || []).slice();
    choices.splice(index, 1);
    patchField(field.id, { choices });
  };

  const handleOptionKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>,
    field: IFormField,
    index: number
  ): void => {
    const choices = field.choices || [];
    const text = (choices[index] || '').trim();
    if (e.key === 'Enter' && text.length > 0) {
      // Enter → new option right below, ready to type
      e.preventDefault();
      insertOption(field, index + 1);
    } else if (e.key === 'Tab' && !e.shiftKey && index === choices.length - 1 && text.length > 0) {
      // Tab on the last option → keep going, add another
      e.preventDefault();
      insertOption(field, index + 1);
    } else if (e.key === 'Backspace' && text.length === 0 && choices.length > 1) {
      // Backspace on an empty option → remove the row, hop to the previous one
      e.preventDefault();
      removeOption(field, index);
      setFocusOption(optionKey(field.id, Math.max(0, index - 1)));
    }
  };

  // image-choice options are objects, so they get their own setters
  const setImageOption = (field: IFormField, index: number, patch: Partial<IImageChoiceOption>): void => {
    const options = (field.imageChoices || []).slice();
    options[index] = { ...options[index], ...patch };
    patchField(field.id, { imageChoices: options });
  };

  // likert rows/columns are plain string lists under different keys
  const setGridList = (
    field: IFormField,
    key: 'likertRows' | 'likertColumns',
    index: number,
    value: string
  ): void => {
    const list = ((field[key] as string[]) || []).slice();
    list[index] = value;
    patchField(field.id, { [key]: list } as Partial<IFormField>);
  };

  // ----- menus -----

  const addMenu = (section: IFormSection): IContextualMenuProps => {
    const item = (meta: IFieldTypeMeta): IContextualMenuItem => ({
      key: meta.type,
      text: meta.label,
      iconProps: { iconName: meta.icon },
      secondaryText: meta.description,
      onClick: () => addField(section.id, meta.type)
    });

    const popular = POPULAR_TYPES.map((t) => item(getFieldTypeMeta(t)));

    const presets: IContextualMenuItem[] = FIELD_PRESETS.map((preset) => ({
      key: 'preset-' + preset.key,
      text: preset.label,
      iconProps: { iconName: preset.icon },
      secondaryText: preset.description,
      onClick: () => applyPreset(section.id, preset.key)
    }));

    const byCategory: IContextualMenuItem[] = [];
    CATEGORY_ORDER.forEach((category) => {
      const inCategory = FIELD_TYPE_META.filter((m) => m.category === category);
      if (inCategory.length === 0) {
        return;
      }
      byCategory.push({
        key: 'cat-' + category,
        itemType: 1 /* Header */,
        text: category
      });
      inCategory.forEach((meta) => byCategory.push(item(meta)));
    });

    return {
      items: popular.concat([
        { key: 'div1', itemType: 0 /* Divider */ },
        {
          key: 'presets',
          text: 'Common presets',
          iconProps: { iconName: 'Lightbulb' },
          subMenuProps: { items: presets }
        },
        {
          key: 'all',
          text: 'All question types',
          iconProps: { iconName: 'AllApps' },
          subMenuProps: { items: byCategory }
        }
      ])
    };
  };

  const overflowMenu = (field: IFormField): IContextualMenuProps => ({
    items: [
      {
        key: 'type',
        text: 'Change type',
        iconProps: { iconName: getFieldTypeMeta(field.type).icon },
        disabled: field.provisioned === true,
        title: field.provisioned
          ? 'The list column already exists, so the type is locked. Duplicate the question to change it.'
          : undefined,
        subMenuProps: {
          items: CATEGORY_ORDER.reduce((acc: IContextualMenuItem[], category) => {
            const inCategory = FIELD_TYPE_META.filter((m) => m.category === category);
            if (inCategory.length === 0) {
              return acc;
            }
            acc.push({ key: 'h-' + category, itemType: 1, text: category });
            inCategory.forEach((meta) =>
              acc.push({
                key: 'type-' + meta.type,
                text: meta.label,
                iconProps: { iconName: meta.icon },
                canCheck: true,
                checked: field.type === meta.type,
                onClick: () => changeFieldType(field.id, meta.type)
              })
            );
            return acc;
          }, [])
        }
      },
      {
        key: 'settings',
        text: 'Branching, validation & more',
        iconProps: { iconName: 'Flow' },
        onClick: () => setPanelFieldId(field.id)
      },
      { key: 'div', itemType: 0 },
      {
        key: 'duplicate',
        text: 'Duplicate',
        iconProps: { iconName: 'Copy' },
        onClick: () => {
          const section = definition.sections.filter(
            (s) => s.fields.filter((f) => f.id === field.id).length > 0
          )[0];
          if (section) {
            duplicateField(section, field);
          }
        }
      },
      {
        key: 'delete',
        text: 'Delete question',
        iconProps: { iconName: 'Delete' },
        onClick: () => removeField(field.id)
      }
    ]
  });

  // ----- outline rail -----

  const outlineDrag = useDragList({
    count: orderedFields.length,
    onReorder: moveFieldGlobally
  });

  const filteredOutline = React.useMemo(() => {
    const query = outlineFilter.trim().toLowerCase();
    if (!query) {
      return undefined;
    }
    const matches: { [fieldId: string]: boolean } = {};
    orderedFields.forEach((field) => {
      const haystack = (field.title || '') + ' ' + getFieldTypeMeta(field.type).label;
      if (haystack.toLowerCase().indexOf(query) !== -1) {
        matches[field.id] = true;
      }
    });
    return matches;
  }, [outlineFilter, orderedFields]);

  const renderOutline = (): React.ReactNode => {
    let globalIndex = -1;
    return (
      <aside className={styles.outline} aria-label="Question outline">
        <div className={styles.outlineHeader}>
          <span>
            {orderedFields.filter((f) => isInputType(f.type)).length} question
            {orderedFields.filter((f) => isInputType(f.type)).length === 1 ? '' : 's'}
          </span>
          {issues.length > 0 && (
            <span title={issues.length + ' issue(s) to review'}>
              <Icon iconName="Warning" style={{ color: 'var(--sf-warning)' }} />
            </span>
          )}
        </div>
        <div className={styles.outlineSearch}>
          <SearchBox
            placeholder="Find a question"
            value={outlineFilter}
            underlined={true}
            onChange={(_e, v) => setOutlineFilter(v || '')}
          />
        </div>
        <div className={styles.outlineList}>
          {orderedFields.length === 0 && <div className={styles.outlineEmpty}>No questions yet.</div>}
          {definition.sections.map((section, sectionIndex) => (
            <div key={section.id}>
              {definition.sections.length > 1 && (
                <div className={styles.outlineSectionLabel}>
                  {section.title || 'Section ' + (sectionIndex + 1)}
                </div>
              )}
              {section.fields.map((field) => {
                globalIndex++;
                const index = globalIndex;
                if (filteredOutline && !filteredOutline[field.id]) {
                  return null;
                }
                const indicator = outlineDrag.dropIndicator(index);
                const className = outlineDrag.isDragging(index)
                  ? styles.outlineItem
                  : indicator === 'before'
                    ? styles.outlineItemDropBefore
                    : indicator === 'after'
                      ? styles.outlineItemDropAfter
                      : field.id === activeFieldId
                        ? styles.outlineItemActive
                        : styles.outlineItem;
                return (
                  <button
                    key={field.id}
                    type="button"
                    className={className}
                    title={field.title || getFieldTypeMeta(field.type).label}
                    {...outlineDrag.rowProps(index)}
                    onClick={() => {
                      setActiveFieldId(field.id);
                      const card = cardRefs.current[field.id];
                      if (card) {
                        card.scrollIntoView({ block: 'center', behavior: 'smooth' });
                      }
                    }}
                  >
                    <span className={styles.outlineNumber}>
                      {isInputType(field.type) ? numberMap[field.id] : '—'}
                    </span>
                    <span className={styles.outlineTitle}>
                      {field.title || getFieldTypeMeta(field.type).label}
                    </span>
                    {field.visibleWhen && <Icon iconName="Flow" className={styles.outlineIcon} />}
                    {issuesByField[field.id] && (
                      <Icon iconName="Warning" className={styles.outlineIcon} />
                    )}
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      </aside>
    );
  };

  // ----- card rendering -----

  const renderPreviewControl = (field: IFormField): React.ReactNode => (
    <div className={styles.qPreviewControl}>
      <FieldControl
        field={field}
        value={undefined}
        disabled={true}
        onChange={() => undefined}
        spService={props.spService}
      />
    </div>
  );

  const renderInactiveCard = (field: IFormField): React.ReactNode => {
    if (!isInputType(field.type)) {
      return (
        <div>
          <div className={styles.qLabelRow}>
            <Icon iconName={getFieldTypeMeta(field.type).icon} className={styles.optionIcon} />
            <span className={styles.qLabelEmpty}>{getFieldTypeMeta(field.type).label}</span>
          </div>
          <div className={styles.qPreviewControl}>
            <ContentBlock
              html={field.contentHtml}
              imageUrl={field.contentImageUrl}
              style={field.contentStyle}
            />
          </div>
        </div>
      );
    }
    return (
      <>
        <div className={styles.qLabelRow}>
          <span className={styles.qNumber}>{numberMap[field.id]}.</span>
          <span className={field.title ? styles.qLabel : styles.qLabelEmpty}>
            {field.title || 'Question'}
          </span>
          {field.required && <span className={styles.requiredMark}>*</span>}
          {field.visibleWhen && (
            <span
              className={styles.branchChip}
              title="This question only appears when a branching rule matches"
            >
              <Icon iconName="Flow" /> Branching
            </span>
          )}
          {field.provisioned && (
            <span className={styles.lockedChip} title="The list column exists, so the type is locked">
              <Icon iconName="Lock" /> Live
            </span>
          )}
          {issuesByField[field.id] && (
            <span className={styles.issueChip} title={issuesByField[field.id].map((i) => i.message).join('\n')}>
              <Icon iconName="Warning" /> Check
            </span>
          )}
        </div>
        {field.description && <div className={styles.qHelpText}>{field.description}</div>}
        {renderPreviewControl(field)}
      </>
    );
  };

  const renderOptionEditor = (field: IFormField): React.ReactNode => {
    const choices = field.choices || [];
    return (
      <div className={styles.optionList}>
        {choices.map((option, index) => (
          <div key={index} className={styles.optionRow}>
            <Icon
              iconName={
                field.type === FieldType.Ranking
                  ? 'NumberedList'
                  : field.allowMultiple
                    ? 'CheckboxComposite'
                    : 'RadioBtnOff'
              }
              className={styles.optionIcon}
            />
            <TextField
              className={styles.optionInput}
              borderless={true}
              placeholder={'Option ' + (index + 1)}
              value={option}
              ariaLabel={'Option ' + (index + 1)}
              componentRef={(ref) => {
                optionRefs.current[optionKey(field.id, index)] = ref;
              }}
              onFocus={(e) => (e.target as HTMLInputElement).select()}
              onKeyDown={(e) => handleOptionKeyDown(e, field, index)}
              onChange={(_e, v) => setOption(field, index, v || '')}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Remove option"
              ariaLabel={'Remove option ' + (index + 1)}
              tabIndex={-1}
              disabled={choices.length <= 1}
              onClick={() => removeOption(field, index)}
            />
          </div>
        ))}
        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add option"
          onClick={() => insertOption(field, choices.length)}
        />
        {field.allowOther && (
          <div className={styles.optionHint}>
            Respondents can also write in an answer labelled &ldquo;{field.otherLabel || 'Other'}&rdquo;.
          </div>
        )}
      </div>
    );
  };

  const renderImageOptionEditor = (field: IFormField): React.ReactNode => {
    const options = field.imageChoices || [];
    return (
      <div className={styles.optionList}>
        {options.map((option, index) => (
          <div key={index} className={styles.optionRow}>
            {option.imageUrl ? (
              <img className={styles.optionThumb} src={option.imageUrl} alt="" />
            ) : (
              <Icon iconName="FileImage" className={styles.optionIcon} />
            )}
            <TextField
              className={styles.optionInput}
              borderless={true}
              placeholder={'Option ' + (index + 1)}
              value={option.label}
              ariaLabel={'Option ' + (index + 1) + ' label'}
              onChange={(_e, v) => setImageOption(field, index, { label: v || '' })}
            />
            <TextField
              className={styles.optionImageInput}
              borderless={true}
              placeholder="Image URL"
              value={option.imageUrl}
              ariaLabel={'Option ' + (index + 1) + ' image URL'}
              onChange={(_e, v) => setImageOption(field, index, { imageUrl: v || '' })}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Remove option"
              ariaLabel={'Remove option ' + (index + 1)}
              tabIndex={-1}
              disabled={options.length <= 1}
              onClick={() =>
                patchField(field.id, { imageChoices: options.filter((_o, i) => i !== index) })
              }
            />
          </div>
        ))}
        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add option"
          onClick={() =>
            patchField(field.id, { imageChoices: options.concat([{ label: '', imageUrl: '' }]) })
          }
        />
      </div>
    );
  };

  const renderGridEditor = (field: IFormField): React.ReactNode => {
    const renderColumn = (key: 'likertRows' | 'likertColumns', label: string): React.ReactNode => {
      const list = (field[key] as string[]) || [];
      return (
        <div className={styles.gridEditorColumn}>
          <div className={styles.gridEditorLabel}>{label}</div>
          {list.map((entry, index) => (
            <div key={index} className={styles.optionRow}>
              <TextField
                className={styles.optionInput}
                borderless={true}
                placeholder={label + ' ' + (index + 1)}
                value={entry}
                ariaLabel={label + ' ' + (index + 1)}
                onChange={(_e, v) => setGridList(field, key, index, v || '')}
              />
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Remove"
                ariaLabel={'Remove ' + label + ' ' + (index + 1)}
                tabIndex={-1}
                disabled={list.length <= 1}
                onClick={() =>
                  patchField(field.id, {
                    [key]: list.filter((_x, i) => i !== index)
                  } as Partial<IFormField>)
                }
              />
            </div>
          ))}
          <ActionButton
            iconProps={{ iconName: 'Add' }}
            text={'Add ' + label.toLowerCase().replace(/s$/, '')}
            onClick={() => patchField(field.id, { [key]: list.concat(['']) } as Partial<IFormField>)}
          />
        </div>
      );
    };
    return (
      <div className={styles.gridEditor}>
        {renderColumn('likertRows', 'Statements')}
        {renderColumn('likertColumns', 'Scale')}
      </div>
    );
  };

  const renderWidthPicker = (field: IFormField): React.ReactNode => {
    const options: { key: FieldWidth; label: string; title: string }[] = [
      { key: 'full', label: '1/1', title: 'Full width' },
      { key: 'half', label: '1/2', title: 'Half width' },
      { key: 'third', label: '1/3', title: 'One third width' }
    ];
    const current = field.width || 'full';
    return (
      <div className={styles.widthPicker} role="group" aria-label="Question width">
        {options.map((option) => (
          <button
            key={option.key}
            type="button"
            title={option.title}
            aria-pressed={current === option.key}
            className={current === option.key ? styles.widthOptionActive : styles.widthOption}
            onClick={() => patchField(field.id, { width: option.key === 'full' ? undefined : option.key })}
          >
            {option.label}
          </button>
        ))}
      </div>
    );
  };

  const renderActiveCard = (
    section: IFormSection,
    field: IFormField,
    fieldIndex: number
  ): React.ReactNode => {
    const hasListOptions = LIST_OPTION_TYPES.indexOf(field.type) !== -1;
    const isContent = !isInputType(field.type);

    return (
      <>
        <div className={styles.qLabelRow}>
          {!isContent && <span className={styles.qNumber}>{numberMap[field.id]}.</span>}
          <TextField
            className={styles.qTitleInput}
            borderless={true}
            placeholder={isContent ? 'Block label (not shown to respondents)' : 'Question'}
            value={field.title}
            autoFocus={true}
            ariaLabel="Question text"
            onChange={(_e, v) => patchField(field.id, { title: v || '' })}
          />
        </div>

        {isContent ? (
          <div className={styles.optionList}>
            <TextField
              multiline={true}
              rows={3}
              label="Content"
              value={field.contentHtml || ''}
              onChange={(_e, v) => patchField(field.id, { contentHtml: v || '' })}
            />
            <div className={styles.qPreviewControlLive}>
              <ContentBlock
                html={field.contentHtml}
                imageUrl={field.contentImageUrl}
                style={field.contentStyle}
              />
            </div>
          </div>
        ) : hasListOptions ? (
          renderOptionEditor(field)
        ) : field.type === FieldType.ImageChoice ? (
          renderImageOptionEditor(field)
        ) : field.type === FieldType.Likert ? (
          renderGridEditor(field)
        ) : (
          renderPreviewControl(field)
        )}

        <div className={styles.qFooter}>
          <IconButton
            iconProps={{ iconName: 'Up' }}
            title="Move up"
            ariaLabel="Move question up"
            disabled={fieldIndex === 0}
            onClick={() => moveFieldWithin(section.id, fieldIndex, fieldIndex - 1)}
          />
          <IconButton
            iconProps={{ iconName: 'Down' }}
            title="Move down"
            ariaLabel="Move question down"
            disabled={fieldIndex === section.fields.length - 1}
            onClick={() => moveFieldWithin(section.id, fieldIndex, fieldIndex + 1)}
          />
          <IconButton
            iconProps={{ iconName: 'Copy' }}
            title="Duplicate"
            ariaLabel="Duplicate question"
            onClick={() => duplicateField(section, field)}
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Delete question"
            ariaLabel="Delete question"
            onClick={() => removeField(field.id)}
          />

          <span className={styles.footerSpacer} />

          {renderWidthPicker(field)}

          {(field.type === FieldType.Choice ||
            field.type === FieldType.ImageChoice ||
            field.type === FieldType.Lookup) && (
            <Toggle
              className={styles.footerToggle}
              label="Multiple answers"
              inlineLabel={true}
              checked={field.allowMultiple === true}
              disabled={field.provisioned === true}
              onChange={(_e, checked) => patchField(field.id, { allowMultiple: checked === true })}
            />
          )}

          {!isContent && (
            <Toggle
              className={styles.footerToggle}
              label="Required"
              inlineLabel={true}
              checked={field.required === true}
              onChange={(_e, checked) => patchField(field.id, { required: checked === true })}
            />
          )}

          <IconButton
            iconProps={{ iconName: 'MoreVertical' }}
            title="More options"
            ariaLabel="More options for this question"
            menuProps={overflowMenu(field)}
            onRenderMenuIcon={() => null}
          />
        </div>
      </>
    );
  };

  const showSectionChrome = definition.sections.length > 1;
  const panelField = panelFieldId ? findFieldIn(definition, panelFieldId) : undefined;
  let runningIndex = -1;

  return (
    <div className={styles.designerShell}>
      <div
        className={styles.designer}
        onClick={(e) => {
          // clicking the canvas background collapses the active question
          if (e.target === e.currentTarget) {
            setActiveFieldId(undefined);
          }
        }}
      >
        {/* form title lives on the canvas, not three clicks deep in a panel */}
        <div className={styles.titleCard}>
          <TextField
            className={styles.titleInput}
            borderless={true}
            placeholder="Form title"
            value={definition.settings.formTitle}
            ariaLabel="Form title"
            onChange={(_e, v) =>
              update((next) => {
                next.settings.formTitle = v || '';
              })
            }
          />
          <TextField
            className={styles.titleDescriptionInput}
            borderless={true}
            placeholder="Add a description (optional)"
            value={definition.settings.formDescription || ''}
            ariaLabel="Form description"
            onChange={(_e, v) =>
              update((next) => {
                next.settings.formDescription = v || '';
              })
            }
          />
        </div>

        {orderedFields.length === 0 && (
          <div className={styles.emptyCanvas}>
            <Icon iconName="PageAdd" className={styles.emptyCanvasIcon} />
            <h3>Add your first question</h3>
            <p>
              Pick a question type to get started. You can reorder questions by dragging them, and add
              branching so people only see what applies to them.
            </p>
          </div>
        )}

        {definition.sections.map((section, sectionIndex) => (
          <div key={section.id} className={styles.sectionBlock}>
            {showSectionChrome && (
              <div className={styles.sectionCard}>
                <div className={styles.sectionChip}>
                  Section {sectionIndex + 1} of {definition.sections.length}
                </div>
                <div className={styles.sectionHeaderRow}>
                  <div className={styles.sectionTitleFields}>
                    <TextField
                      borderless={true}
                      value={section.title}
                      placeholder="Section title"
                      ariaLabel={'Section ' + (sectionIndex + 1) + ' title'}
                      className={styles.sectionTitleInput}
                      onChange={(_e, v) =>
                        update((next) => {
                          next.sections[sectionIndex].title = v || '';
                        })
                      }
                    />
                    <TextField
                      borderless={true}
                      value={section.description || ''}
                      placeholder="Description (optional)"
                      ariaLabel={'Section ' + (sectionIndex + 1) + ' description'}
                      className={styles.sectionDescriptionInput}
                      onChange={(_e, v) =>
                        update((next) => {
                          next.sections[sectionIndex].description = v || '';
                        })
                      }
                    />
                  </div>
                  <div className={styles.sectionActions}>
                    <IconButton
                      iconProps={{ iconName: 'Up' }}
                      title="Move section up"
                      ariaLabel="Move section up"
                      disabled={sectionIndex === 0}
                      onClick={() => moveSection(section, -1)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title="Move section down"
                      ariaLabel="Move section down"
                      disabled={sectionIndex === definition.sections.length - 1}
                      onClick={() => moveSection(section, 1)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Delete section"
                      ariaLabel="Delete section"
                      onClick={() => removeSection(section)}
                    />
                  </div>
                </div>
              </div>
            )}

            {section.fields.map((field, fieldIndex) => {
              runningIndex++;
              const globalIndex = runningIndex;
              const active = field.id === activeFieldId;
              const indicator = outlineDrag.dropIndicator(globalIndex);
              const className = outlineDrag.isDragging(globalIndex)
                ? styles.qCardDragging
                : indicator === 'before'
                  ? styles.qCardDropBefore
                  : indicator === 'after'
                    ? styles.qCardDropAfter
                    : active
                      ? styles.qCardActive
                      : issuesByField[field.id]
                        ? styles.qCardFlagged
                        : styles.qCard;
              const dragHandlers = outlineDrag.rowProps(globalIndex);
              return (
                <div
                  key={field.id}
                  ref={(element) => {
                    cardRefs.current[field.id] = element;
                  }}
                  className={className}
                  onClick={() => !active && setActiveFieldId(field.id)}
                  onDragOver={dragHandlers.onDragOver}
                  onDragLeave={dragHandlers.onDragLeave}
                  onDrop={dragHandlers.onDrop}
                >
                  {/*
                    Only the grip starts a drag. Making the whole card draggable
                    would hijack text selection inside the inputs of the active
                    card.
                  */}
                  <button
                    type="button"
                    className={styles.qDragHandle}
                    title="Drag to reorder"
                    aria-label={'Reorder ' + (field.title || 'question')}
                    draggable={true}
                    onDragStart={dragHandlers.onDragStart}
                    onDragEnd={dragHandlers.onDragEnd}
                    onKeyDown={(event) => {
                      if (event.key === 'ArrowUp') {
                        event.preventDefault();
                        moveFieldGlobally(globalIndex, globalIndex - 1);
                      } else if (event.key === 'ArrowDown') {
                        event.preventDefault();
                        moveFieldGlobally(globalIndex, globalIndex + 1);
                      }
                    }}
                  >
                    <Icon iconName="GripperDotsVertical" />
                  </button>
                  {active
                    ? renderActiveCard(section, field, fieldIndex)
                    : renderInactiveCard(field)}
                </div>
              );
            })}

            <div className={styles.addNewRow}>
              <DefaultButton
                className={styles.addNewButton}
                iconProps={{ iconName: 'Add' }}
                text="Add new"
                menuProps={addMenu(section)}
              />
            </div>
          </div>
        ))}

        <div className={styles.addSectionRow}>
          <ActionButton
            iconProps={{ iconName: 'DoubleChevronDown8' }}
            text="Add section"
            onClick={addSection}
          />
        </div>

        {panelField && (
          <FieldEditorPanel
            field={panelField}
            definition={definition}
            spService={props.spService}
            onSave={(updated) => {
              update((next) => {
                next.sections.forEach((s) => {
                  s.fields = s.fields.map((f) => (f.id === updated.id ? updated : f));
                });
              });
              setPanelFieldId(undefined);
            }}
            onDismiss={() => setPanelFieldId(undefined)}
          />
        )}
      </div>

      {orderedFields.length > 3 && renderOutline()}
    </div>
  );
};
