import * as React from 'react';
import {
  ActionButton,
  DefaultButton,
  Icon,
  IconButton,
  IContextualMenuItem,
  IContextualMenuProps,
  ITextField,
  MessageBar,
  MessageBarType,
  SearchBox,
  TextField,
  Toggle
} from '@fluentui/react';
import * as strings from 'SmartFormsWebPartStrings';
import styles from './FormDesigner.module.scss';
import {
  defaultsForType,
  FIELD_PRESETS,
  FIELD_TYPE_META,
  FieldCategory,
  FieldType,
  FieldWidth,
  fieldCategoryLabel,
  fieldPresetDescription,
  fieldPresetLabel,
  fieldTypeDescription,
  fieldTypeLabel,
  getFieldTypeMeta,
  IFieldTypeMeta,
  IFormDefinition,
  IFormField,
  IFormSection,
  IImageChoiceOption,
  IMessageBag,
  isInputType,
  newId,
  TYPE_SPECIFIC_KEYS
} from '../../models';
import {
  allFields,
  buildNumberMap,
  generateInternalName,
  IFormIssue,
  pruneInvalidConditions,
  retireDeletedField,
  validateDefinition
} from '../../utils/formUtils';
import { reorder, useDragList } from '../../hooks/useDragList';
import { SharePointService } from '../../services/SharePointService';
import { FieldControl } from '../form/FieldControl';
import { ContentBlock } from '../form/ContentBlock';
import { FieldEditorPanel } from './FieldEditorPanel';
import { formatString } from '../../utils/localeUtils';

export interface IFormDesignerProps {
  definition: IFormDefinition;
  spService: SharePointService;
  /** field the shell asked us to scroll to and open (from a pre-flight issue) */
  focusFieldId?: string;
  onChange: (definition: IFormDefinition) => void;
  onFocusHandled?: () => void;
}

const clone = <T,>(value: T): T => JSON.parse(JSON.stringify(value));

/** Localized heading for a question-type category. */
const messageBag = strings as unknown as IMessageBag;

const categoryLabel = (category: FieldCategory): string =>
  messageBag['Designer_Category_' + category] || fieldCategoryLabel(category, messageBag);

/** Total branching rules in a definition (sections and questions). */
const countRules = (def: IFormDefinition): number => {
  let total = 0;
  def.sections.forEach((section) => {
    total += section.visibleWhen ? section.visibleWhen.conditions.length : 0;
    section.fields.forEach((f) => {
      total += f.visibleWhen ? f.visibleWhen.conditions.length : 0;
    });
  });
  return total;
};

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
  // number of branching rules the last edit removed because they no longer fit
  const [removedRules, setRemovedRules] = React.useState<number>(0);
  const [dropSectionId, setDropSectionId] = React.useState<string | undefined>(undefined);

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

  const issues = React.useMemo(() => validateDefinition(definition, messageBag), [definition]);
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

  const update = (mutate: (next: IFormDefinition) => void, options?: { prune?: boolean }): void => {
    const next = clone(definition);
    mutate(next);
    if (options && options.prune) {
      const before = countRules(next);
      pruneInvalidConditions(next);
      const removed = before - countRules(next);
      if (removed > 0) {
        setRemovedRules(removed);
      }
    }
    props.onChange(next);
  };

  /** Drop branching rules that no longer fit (call after an option is renamed or removed). */
  const pruneRules = (): void => {
    const trial = clone(definition);
    const before = countRules(trial);
    pruneInvalidConditions(trial);
    const removed = before - countRules(trial);
    if (removed > 0) {
      setRemovedRules(removed);
      props.onChange(trial);
    }
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
    }, { prune: true });
  };

  const applyPreset = (sectionId: string, presetKey: string, atIndex?: number): void => {
    const preset = FIELD_PRESETS.filter((p) => p.key === presetKey)[0];
    if (!preset) {
      return;
    }
    addField(sectionId, preset.type, preset.defaults, atIndex);
  };

  /**
   * Add a new question to a section. Omit `atIndex` to append at the end (the
   * "Add new" button); pass `fieldIndex + 1` to insert right after an existing
   * question (the per-question "Insert question below" menu item).
   */
  const addField = (
    sectionId: string,
    type: FieldType,
    extraDefaults?: Partial<IFormField>,
    atIndex?: number
  ): void => {
    const id = newId();
    update((next) => {
      const target = next.sections.filter((s) => s.id === sectionId)[0];
      if (!target) {
        return;
      }
      const field: IFormField = {
        id,
        internalName: generateInternalName(
          fieldTypeLabel(type, messageBag),
          allFields(next).map((f) => f.internalName),
          next.retiredColumns
        ),
        title: '',
        type,
        required: false,
        provisioned: false,
        ...defaultsForType(type),
        ...(extraDefaults || {})
      };
      if (atIndex === undefined || atIndex >= target.fields.length) {
        target.fields.push(field);
      } else {
        target.fields.splice(Math.max(0, atIndex), 0, field);
      }
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
        copy.title || strings.Designer_Card_Question,
        allFields(next).map((f) => f.internalName),
        next.retiredColumns
      );
      copy.provisioned = false;
      target.fields.splice(index + 1, 0, copy);
    });
    setActiveFieldId(id);
  };

  const removeField = (fieldId: string): void => {
    update((next) => {
      next.sections.forEach((section) => {
        section.fields.forEach((f) => {
          if (f.id === fieldId) {
            // a published question's column stays in the list; never reuse its name
            retireDeletedField(next, f);
          }
        });
        section.fields = section.fields.filter((f) => f.id !== fieldId);
      });
    }, { prune: true });
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

  /**
   * Move a question to an exact spot: section + index within that section.
   * `from` is the position in the flattened question order. Works for the end
   * of any section and for empty sections, which a flat index cannot express.
   */
  const moveFieldToSection = (from: number, sectionId: string, indexInSection: number): void => {
    update((next) => {
      const flat: { section: IFormSection; field: IFormField }[] = [];
      next.sections.forEach((section) => section.fields.forEach((field) => flat.push({ section, field })));
      const source = flat[from];
      const target = next.sections.filter((s) => s.id === sectionId)[0];
      if (!source || !target) {
        return;
      }
      const sourceIndex = source.section.fields.indexOf(source.field);
      source.section.fields.splice(sourceIndex, 1);
      let at = indexInSection;
      if (source.section === target && sourceIndex < at) {
        at--;
      }
      at = Math.max(0, Math.min(at, target.fields.length));
      target.fields.splice(at, 0, source.field);
    });
  };

  /** Drop handler for a card/outline row: lands before or after that row, within its own section. */
  const dropOnRow = (event: React.DragEvent, sectionId: string, fieldIndex: number): void => {
    event.preventDefault();
    event.stopPropagation();
    const { dragIndex, dropSide } = outlineDrag.state;
    outlineDrag.cancel();
    setDropSectionId(undefined);
    if (dragIndex < 0) {
      return;
    }
    moveFieldToSection(dragIndex, sectionId, dropSide === 'after' ? fieldIndex + 1 : fieldIndex);
  };

  /** Handlers that make a section's tail (or an empty section) accept a dragged question. */
  const sectionDropProps = (section: IFormSection): React.HTMLAttributes<HTMLDivElement> => ({
    onDragOver: (event: React.DragEvent) => {
      if (outlineDrag.state.dragIndex < 0) {
        return;
      }
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      if (dropSectionId !== section.id) {
        setDropSectionId(section.id);
      }
    },
    onDragLeave: () => setDropSectionId((prev) => (prev === section.id ? undefined : prev)),
    onDrop: (event: React.DragEvent) => {
      event.preventDefault();
      const { dragIndex } = outlineDrag.state;
      outlineDrag.cancel();
      setDropSectionId(undefined);
      if (dragIndex >= 0) {
        moveFieldToSection(dragIndex, section.id, section.fields.length);
      }
    }
  });

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
        doomed.fields.forEach((f) => retireDeletedField(next, f));
      }
    }, { prune: true });
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
    update((next) => {
      const f = findFieldIn(next, field.id);
      if (f) {
        f.choices = choices;
      }
    }, { prune: true });
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

  /** Builds the question-type picker. Pass `atIndex` to insert mid-section instead of appending. */
  const addMenu = (section: IFormSection, atIndex?: number): IContextualMenuProps => {
    const item = (meta: IFieldTypeMeta): IContextualMenuItem => ({
      key: meta.type,
      text: fieldTypeLabel(meta.type, messageBag),
      iconProps: { iconName: meta.icon },
      secondaryText: fieldTypeDescription(meta.type, messageBag),
      onClick: () => addField(section.id, meta.type, undefined, atIndex)
    });

    const popular = POPULAR_TYPES.map((t) => item(getFieldTypeMeta(t)));

    const presets: IContextualMenuItem[] = FIELD_PRESETS.map((preset) => ({
      key: 'preset-' + preset.key,
      text: fieldPresetLabel(preset, messageBag),
      iconProps: { iconName: preset.icon },
      secondaryText: fieldPresetDescription(preset, messageBag),
      onClick: () => applyPreset(section.id, preset.key, atIndex)
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
        text: categoryLabel(category)
      });
      inCategory.forEach((meta) => byCategory.push(item(meta)));
    });

    return {
      items: popular.concat([
        { key: 'div1', itemType: 0 /* Divider */ },
        {
          key: 'presets',
          text: strings.Designer_Menu_CommonPresets,
          iconProps: { iconName: 'Lightbulb' },
          subMenuProps: { items: presets }
        },
        {
          key: 'all',
          text: strings.Designer_Menu_AllTypes,
          iconProps: { iconName: 'AllApps' },
          subMenuProps: { items: byCategory }
        }
      ])
    };
  };

  const overflowMenu = (section: IFormSection, field: IFormField, fieldIndex: number): IContextualMenuProps => ({
    items: [
      {
        key: 'type',
        text: strings.Designer_Menu_ChangeType,
        iconProps: { iconName: getFieldTypeMeta(field.type).icon },
        disabled: field.provisioned === true,
        title: field.provisioned
          ? strings.Designer_Menu_TypeLockedTitle
          : undefined,
        subMenuProps: {
          items: CATEGORY_ORDER.reduce((acc: IContextualMenuItem[], category) => {
            const inCategory = FIELD_TYPE_META.filter((m) => m.category === category);
            if (inCategory.length === 0) {
              return acc;
            }
            acc.push({ key: 'h-' + category, itemType: 1, text: categoryLabel(category) });
            inCategory.forEach((meta) =>
              acc.push({
                key: 'type-' + meta.type,
                text: fieldTypeLabel(meta.type, messageBag),
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
        text: strings.Designer_Menu_MoreSettings,
        iconProps: { iconName: 'Flow' },
        onClick: () => setPanelFieldId(field.id)
      },
      { key: 'div', itemType: 0 },
      {
        key: 'insert-below',
        text: strings.Designer_Menu_InsertBelow,
        iconProps: { iconName: 'Add' },
        subMenuProps: addMenu(section, fieldIndex + 1)
      },
      {
        key: 'duplicate',
        text: strings.Designer_Card_Duplicate,
        iconProps: { iconName: 'Copy' },
        onClick: () => duplicateField(section, field)
      },
      {
        key: 'delete',
        text: strings.Designer_Card_DeleteQuestion,
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
      const haystack = (field.title || '') + ' ' + fieldTypeLabel(field.type, messageBag);
      if (haystack.toLowerCase().indexOf(query) !== -1) {
        matches[field.id] = true;
      }
    });
    return matches;
  }, [outlineFilter, orderedFields]);

  const renderOutline = (): React.ReactNode => {
    let globalIndex = -1;
    return (
      <aside className={styles.outline} aria-label={strings.Designer_Outline_Aria}>
        <div className={styles.outlineHeader}>
          <span>
            {formatString(
              orderedFields.filter((f) => isInputType(f.type)).length === 1
                ? strings.Designer_Outline_QuestionCountOne
                : strings.Designer_Outline_QuestionCountOther,
              { count: orderedFields.filter((f) => isInputType(f.type)).length }
            )}
          </span>
          {issues.length > 0 && (
            <span title={formatString(strings.Designer_Outline_IssuesTitle, { count: issues.length })}>
              <Icon iconName="Warning" style={{ color: 'var(--sf-warning)' }} />
            </span>
          )}
        </div>
        <div className={styles.outlineSearch}>
          <SearchBox
            placeholder={strings.Designer_Outline_Find}
            value={outlineFilter}
            underlined={true}
            onChange={(_e, v) => setOutlineFilter(v || '')}
          />
        </div>
        <div className={styles.outlineList}>
          {orderedFields.length === 0 && <div className={styles.outlineEmpty}>{strings.Designer_Outline_Empty}</div>}
          {definition.sections.map((section, sectionIndex) => (
            <div key={section.id}>
              {definition.sections.length > 1 && (
                <div className={styles.outlineSectionLabel}>
                  {section.title || formatString(strings.Designer_Section_DefaultTitle, { number: sectionIndex + 1 })}
                </div>
              )}
              {section.fields.map((field, fieldIndexInSection) => {
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
                    title={field.title || fieldTypeLabel(field.type, messageBag)}
                    {...outlineDrag.rowProps(index)}
                    onDrop={(event) => dropOnRow(event, section.id, fieldIndexInSection)}
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
                      {field.title || fieldTypeLabel(field.type, messageBag)}
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
            <span className={styles.qLabelEmpty}>{fieldTypeLabel(field.type, messageBag)}</span>
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
            {field.title || strings.Designer_Card_Question}
          </span>
          {field.required && <span className={styles.requiredMark}>*</span>}
          {field.visibleWhen && (
            <span
              className={styles.branchChip}
              title={strings.Designer_Card_BranchingTitle}
            >
              <Icon iconName="Flow" /> {strings.Designer_Card_BranchingChip}
            </span>
          )}
          {field.provisioned && (
            <span className={styles.lockedChip} title={strings.Designer_Card_LiveTitle}>
              <Icon iconName="Lock" /> {strings.Designer_Card_LiveChip}
            </span>
          )}
          {issuesByField[field.id] && (
            <span className={styles.issueChip} title={issuesByField[field.id].map((i) => i.message).join('\n')}>
              <Icon iconName="Warning" /> {strings.Designer_Card_CheckChip}
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
              placeholder={formatString(strings.Designer_Option_Placeholder, { number: index + 1 })}
              value={option}
              ariaLabel={formatString(strings.Designer_Option_Placeholder, { number: index + 1 })}
              componentRef={(ref) => {
                optionRefs.current[optionKey(field.id, index)] = ref;
              }}
              onFocus={(e) => (e.target as HTMLInputElement).select()}
              onKeyDown={(e) => handleOptionKeyDown(e, field, index)}
              onChange={(_e, v) => setOption(field, index, v || '')}
              onBlur={pruneRules}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title={strings.Designer_Option_Remove}
              ariaLabel={formatString(strings.Designer_Option_RemoveAria, { number: index + 1 })}
              tabIndex={-1}
              disabled={choices.length <= 1}
              onClick={() => removeOption(field, index)}
            />
          </div>
        ))}
        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text={strings.Designer_Option_Add}
          onClick={() => insertOption(field, choices.length)}
        />
        {field.allowOther && (
          <div className={styles.optionHint}>
            {formatString(strings.Designer_Option_OtherHint, {
              label: field.otherLabel || strings.Designer_Option_OtherFallback
            })}
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
              placeholder={formatString(strings.Designer_Option_Placeholder, { number: index + 1 })}
              value={option.label}
              ariaLabel={formatString(strings.Designer_Option_LabelAria, { number: index + 1 })}
              onChange={(_e, v) => setImageOption(field, index, { label: v || '' })}
            />
            <TextField
              className={styles.optionImageInput}
              borderless={true}
              placeholder={strings.Designer_Option_ImageUrl}
              value={option.imageUrl}
              ariaLabel={formatString(strings.Designer_Option_ImageUrlAria, { number: index + 1 })}
              onChange={(_e, v) => setImageOption(field, index, { imageUrl: v || '' })}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title={strings.Designer_Option_Remove}
              ariaLabel={formatString(strings.Designer_Option_RemoveAria, { number: index + 1 })}
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
          text={strings.Designer_Option_Add}
          onClick={() =>
            patchField(field.id, { imageChoices: options.concat([{ label: '', imageUrl: '' }]) })
          }
        />
      </div>
    );
  };

  const renderGridEditor = (field: IFormField): React.ReactNode => {
    const renderColumn = (
      key: 'likertRows' | 'likertColumns',
      label: string,
      addText: string
    ): React.ReactNode => {
      const list = (field[key] as string[]) || [];
      return (
        <div className={styles.gridEditorColumn}>
          <div className={styles.gridEditorLabel}>{label}</div>
          {list.map((entry, index) => (
            <div key={index} className={styles.optionRow}>
              <TextField
                className={styles.optionInput}
                borderless={true}
                placeholder={formatString(strings.Designer_Grid_ItemPlaceholder, { label, number: index + 1 })}
                value={entry}
                ariaLabel={formatString(strings.Designer_Grid_ItemPlaceholder, { label, number: index + 1 })}
                onChange={(_e, v) => setGridList(field, key, index, v || '')}
              />
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title={strings.Designer_Grid_Remove}
                ariaLabel={formatString(strings.Designer_Grid_RemoveAria, { label, number: index + 1 })}
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
            text={addText}
            onClick={() => patchField(field.id, { [key]: list.concat(['']) } as Partial<IFormField>)}
          />
        </div>
      );
    };
    return (
      <div className={styles.gridEditor}>
        {renderColumn('likertRows', strings.Designer_Grid_Statements, strings.Designer_Grid_AddStatement)}
        {renderColumn('likertColumns', strings.Designer_Grid_Scale, strings.Designer_Grid_AddScale)}
      </div>
    );
  };

  const renderWidthPicker = (field: IFormField): React.ReactNode => {
    const options: { key: FieldWidth; label: string; title: string }[] = [
      { key: 'full', label: '1/1', title: strings.Designer_Width_Full },
      { key: 'half', label: '1/2', title: strings.Designer_Width_Half },
      { key: 'third', label: '1/3', title: strings.Designer_Width_Third }
    ];
    const current = field.width || 'full';
    return (
      <div className={styles.widthPicker} role="group" aria-label={strings.Designer_Width_Aria}>
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
            placeholder={
              isContent ? strings.Designer_Card_BlockLabelPlaceholder : strings.Designer_Card_Question
            }
            value={field.title}
            autoFocus={true}
            ariaLabel={strings.Designer_Card_QuestionText}
            onChange={(_e, v) => patchField(field.id, { title: v || '' })}
          />
        </div>

        {isContent ? (
          <div className={styles.optionList}>
            <TextField
              multiline={true}
              rows={3}
              label={strings.Designer_Field_Content}
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
            title={strings.Designer_Card_MoveUp}
            ariaLabel={strings.Designer_Card_MoveQuestionUp}
            disabled={fieldIndex === 0}
            onClick={() => moveFieldWithin(section.id, fieldIndex, fieldIndex - 1)}
          />
          <IconButton
            iconProps={{ iconName: 'Down' }}
            title={strings.Designer_Card_MoveDown}
            ariaLabel={strings.Designer_Card_MoveQuestionDown}
            disabled={fieldIndex === section.fields.length - 1}
            onClick={() => moveFieldWithin(section.id, fieldIndex, fieldIndex + 1)}
          />
          <IconButton
            iconProps={{ iconName: 'Copy' }}
            title={strings.Designer_Card_Duplicate}
            ariaLabel={strings.Designer_Card_DuplicateQuestion}
            onClick={() => duplicateField(section, field)}
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title={strings.Designer_Card_DeleteQuestion}
            ariaLabel={strings.Designer_Card_DeleteQuestion}
            onClick={() => removeField(field.id)}
          />

          <span className={styles.footerSpacer} />

          {renderWidthPicker(field)}

          {(field.type === FieldType.Choice ||
            field.type === FieldType.ImageChoice ||
            field.type === FieldType.Lookup) && (
            <Toggle
              className={styles.footerToggle}
              label={strings.Designer_Card_MultipleAnswers}
              inlineLabel={true}
              checked={field.allowMultiple === true}
              disabled={field.provisioned === true}
              onChange={(_e, checked) => patchField(field.id, { allowMultiple: checked === true })}
            />
          )}

          {!isContent && (
            <Toggle
              className={styles.footerToggle}
              label={strings.Designer_Card_Required}
              inlineLabel={true}
              checked={field.required === true}
              onChange={(_e, checked) => patchField(field.id, { required: checked === true })}
            />
          )}

          <IconButton
            iconProps={{ iconName: 'MoreVertical' }}
            title={strings.Designer_Card_MoreOptions}
            ariaLabel={strings.Designer_Card_MoreOptionsAria}
            menuProps={overflowMenu(section, field, fieldIndex)}
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
            placeholder={strings.Designer_Canvas_FormTitle}
            value={definition.settings.formTitle}
            ariaLabel={strings.Designer_Canvas_FormTitle}
            onChange={(_e, v) =>
              update((next) => {
                next.settings.formTitle = v || '';
              })
            }
          />
          <TextField
            className={styles.titleDescriptionInput}
            borderless={true}
            placeholder={strings.Designer_Canvas_DescriptionPlaceholder}
            value={definition.settings.formDescription || ''}
            ariaLabel={strings.Designer_Canvas_DescriptionAria}
            onChange={(_e, v) =>
              update((next) => {
                next.settings.formDescription = v || '';
              })
            }
          />
        </div>

        {removedRules > 0 && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            onDismiss={() => setRemovedRules(0)}
            dismissButtonAriaLabel={strings.Designer_Common_Cancel}
          >
            {formatString(
              removedRules === 1 ? strings.Designer_Rules_RemovedOne : strings.Designer_Rules_RemovedOther,
              { count: removedRules }
            )}
          </MessageBar>
        )}

        {orderedFields.length === 0 && (
          <div className={styles.emptyCanvas}>
            <Icon iconName="PageAdd" className={styles.emptyCanvasIcon} />
            <h3>{strings.Designer_Canvas_EmptyTitle}</h3>
            <p>
              {strings.Designer_Canvas_EmptyBody}
            </p>
          </div>
        )}

        {definition.sections.map((section, sectionIndex) => (
          <div key={section.id} className={styles.sectionBlock}>
            {showSectionChrome && (
              <div className={styles.sectionCard}>
                <div className={styles.sectionChip}>
                  {formatString(strings.Designer_Section_Chip, {
                    number: sectionIndex + 1,
                    total: definition.sections.length
                  })}
                </div>
                <div className={styles.sectionHeaderRow}>
                  <div className={styles.sectionTitleFields}>
                    <TextField
                      borderless={true}
                      value={section.title}
                      placeholder={strings.Designer_Section_TitlePlaceholder}
                      ariaLabel={formatString(strings.Designer_Section_TitleAria, { number: sectionIndex + 1 })}
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
                      placeholder={strings.Designer_Section_DescriptionPlaceholder}
                      ariaLabel={formatString(strings.Designer_Section_DescriptionAria, { number: sectionIndex + 1 })}
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
                      title={strings.Designer_Section_MoveUp}
                      ariaLabel={strings.Designer_Section_MoveUp}
                      disabled={sectionIndex === 0}
                      onClick={() => moveSection(section, -1)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title={strings.Designer_Section_MoveDown}
                      ariaLabel={strings.Designer_Section_MoveDown}
                      disabled={sectionIndex === definition.sections.length - 1}
                      onClick={() => moveSection(section, 1)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title={strings.Designer_Section_Delete}
                      ariaLabel={strings.Designer_Section_Delete}
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
                  onDrop={(event) => dropOnRow(event, section.id, fieldIndex)}
                >
                  {/*
                    Only the grip starts a drag. Making the whole card draggable
                    would hijack text selection inside the inputs of the active
                    card.
                  */}
                  <button
                    type="button"
                    className={styles.qDragHandle}
                    title={strings.Designer_Card_DragToReorder}
                    aria-label={formatString(strings.Designer_Card_ReorderAria, {
                      title: field.title || strings.Designer_Card_ReorderFallback
                    })}
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

            {outlineDrag.state.dragIndex >= 0 && (
              <div
                className={dropSectionId === section.id ? styles.sectionDropZoneActive : styles.sectionDropZone}
                {...sectionDropProps(section)}
              >
                {section.fields.length === 0 ? strings.Designer_Section_DropHere : strings.Designer_Section_DropEnd}
              </div>
            )}

            <div className={styles.addNewRow}>
              <DefaultButton
                className={styles.addNewButton}
                iconProps={{ iconName: 'Add' }}
                text={strings.Designer_Canvas_AddNew}
                menuProps={addMenu(section)}
              />
            </div>
          </div>
        ))}

        <div className={styles.addSectionRow}>
          <ActionButton
            iconProps={{ iconName: 'DoubleChevronDown8' }}
            text={strings.Designer_Canvas_AddSection}
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
              }, { prune: true });
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
