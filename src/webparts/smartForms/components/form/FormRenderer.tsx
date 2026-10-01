import * as React from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import {
  FieldType,
  IFormDefinition,
  IFormField,
  IFormSection,
  IFormValues,
  IMessageBag,
  isInputType
} from '../../models';
import {
  applyCalculatedFields,
  buildNumberMap,
  formAvailability,
  formatValue,
  computeVisibility,
  effectiveChoices,
  IAvailability,
  inputFields,
  parseLocalDate,
  parseTimeToMinutes,
  shuffleWithSeed,
  validateField
} from '../../utils/formUtils';
import { IDraft, ISubmitResult, SharePointService } from '../../services/SharePointService';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString, isRtlLocale } from '../../utils/localeUtils';
import { logError } from '../../utils/debug';
import { FieldControl } from './FieldControl';
import { ContentBlock } from './ContentBlock';

export type RendererMode = 'live' | 'preview';

export interface IFormRendererProps {
  definition: IFormDefinition;
  listId: string;
  spService: SharePointService;
  /** 'preview' validates and shows the confirmation without writing anything */
  mode?: RendererMode;
  /** page UI culture (e.g. 'ar-SA'); drives right-to-left layout and date formatting */
  locale?: string;
  /** notified after a real submission, so the shell can refresh counts */
  onSubmitted?: () => void;
}

const messageBag = strings as unknown as IMessageBag;

/** Today's date at the given minutes-past-midnight, as the Time control expects. */
const dateAtMinutes = (minutes: number): Date => {
  const d = new Date();
  d.setHours(Math.floor(minutes / 60), minutes % 60, 0, 0);
  return d;
};

/** Convert raw text (query string or default) into a Date/Time answer, or undefined. */
const parseDateOrTime = (field: IFormField, raw: string): Date | undefined => {
  if (field.type === FieldType.Time) {
    if (/^now$/i.test(raw.trim())) {
      return new Date();
    }
    const minutes = parseTimeToMinutes(raw);
    return minutes === undefined ? undefined : dateAtMinutes(minutes);
  }
  return parseLocalDate(raw);
};

/** Options a text answer may be chosen from, or undefined when they can't be known statically. */
const knownOptions = (field: IFormField): string[] | undefined => {
  if (field.type === FieldType.Choice) {
    return effectiveChoices(field);
  }
  if (field.type === FieldType.ImageChoice) {
    return (field.imageChoices || []).map((c) => c.label);
  }
  // Lookup options load from another list at runtime, so they can't be checked here
  return undefined;
};

/** Keep only the parts of `raw` (';'-separated when multi) that match a real option, case-insensitively. */
const choiceFromText = (field: IFormField, raw: string): string | string[] | undefined => {
  const parts = field.allowMultiple
    ? raw.split(';').map((s) => s.trim()).filter((s) => s.length > 0)
    : [raw.trim()];
  const options = knownOptions(field);
  const kept = options
    ? parts
        .map((p) => options.filter((o) => o.toLowerCase() === p.toLowerCase())[0])
        .filter((o) => o !== undefined)
    : parts;
  if (kept.length === 0) {
    return undefined;
  }
  return field.allowMultiple ? kept : kept[0];
};

/**
 * Prefill values from the page's query string: any parameter named after a
 * question's internal column name (e.g. ?SFYourName=Alex) becomes that
 * question's initial answer. Lets owners hand out links that pre-answer
 * routing questions — something Microsoft Forms can't do.
 *
 * Security note: this is a convenience, not a security boundary. Anyone can
 * edit the URL or the answer afterwards, so never rely on a prefilled (or
 * read-only-looking) value being trustworthy; validate on the consuming side.
 */
const applyPrefillFromUrl = (definition: IFormDefinition, values: IFormValues): void => {
  let params: URLSearchParams;
  try {
    params = new URLSearchParams(window.location.search);
  } catch {
    return;
  }
  inputFields(definition).forEach((field) => {
    const raw = params.get(field.internalName);
    if (raw === null || raw === '') {
      return;
    }
    switch (field.type) {
      case FieldType.Number:
      case FieldType.Rating:
      case FieldType.Slider:
      case FieldType.Scale: {
        const num = Number(raw);
        if (!isNaN(num)) {
          values[field.id] = num;
        }
        break;
      }
      case FieldType.YesNo:
      case FieldType.Consent:
        values[field.id] = /^(yes|true|1|on)$/i.test(raw);
        break;
      case FieldType.Choice:
      case FieldType.Lookup:
      case FieldType.ImageChoice: {
        const chosen = choiceFromText(field, raw);
        if (chosen !== undefined) {
          values[field.id] = chosen;
        }
        break;
      }
      case FieldType.Date:
      case FieldType.Time: {
        const parsed = parseDateOrTime(field, raw);
        if (parsed) {
          values[field.id] = parsed;
        }
        break;
      }
      case FieldType.Text:
      case FieldType.MultilineText:
      case FieldType.Email:
      case FieldType.Phone:
        values[field.id] = raw;
        break;
      case FieldType.Hyperlink:
        values[field.id] = { url: raw, description: '' };
        break;
      default:
        // people, rankings, signatures, files and grids can't come from a URL
        break;
    }
  });
};

const buildDefaultValues = (definition: IFormDefinition): IFormValues => {
  const values: IFormValues = {};
  inputFields(definition).forEach((field) => {
    if (field.type === FieldType.Ranking) {
      // the designed order is the implicit starting answer
      values[field.id] = (field.choices || []).map((c) => c.trim()).filter((c) => c.length > 0);
      return;
    }
    if (field.type === FieldType.Likert) {
      values[field.id] = {};
      return;
    }
    if (field.defaultValue === undefined || field.defaultValue === '') {
      if (field.type === FieldType.YesNo) {
        values[field.id] = false;
      }
      return;
    }
    switch (field.type) {
      case FieldType.Number:
      case FieldType.Rating:
      case FieldType.Slider:
      case FieldType.Scale: {
        const num = Number(field.defaultValue);
        if (!isNaN(num)) {
          values[field.id] = num;
        }
        break;
      }
      case FieldType.YesNo:
      case FieldType.Consent:
        values[field.id] = /^(yes|true|1|on)$/i.test(field.defaultValue);
        break;
      case FieldType.Choice:
      case FieldType.Lookup:
      case FieldType.ImageChoice:
        {
          const chosen = choiceFromText(field, field.defaultValue);
          if (chosen !== undefined) {
            values[field.id] = chosen;
          }
        }
        break;
      case FieldType.Date:
      case FieldType.Time:
        if (/^now$/i.test(field.defaultValue.trim())) {
          values[field.id] = new Date();
        } else {
          const parsed = parseDateOrTime(field, field.defaultValue);
          if (parsed) {
            values[field.id] = parsed;
          }
        }
        break;
      case FieldType.Person:
      case FieldType.Signature:
      case FieldType.FileUpload:
        break;
      case FieldType.Hyperlink:
        values[field.id] = { url: field.defaultValue, description: '' };
        break;
      default:
        values[field.id] = field.defaultValue;
    }
  });
  applyPrefillFromUrl(definition, values);
  return applyCalculatedFields(definition, values);
};

export const FormRenderer: React.FunctionComponent<IFormRendererProps> = (props) => {
  const { definition, listId, spService } = props;
  const settings = definition.settings;
  const isPreview = props.mode === 'preview';

  const [values, setValues] = React.useState<IFormValues>(() => buildDefaultValues(definition));
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [step, setStep] = React.useState<number>(0);
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [submitted, setSubmitted] = React.useState<boolean>(false);
  const [submitError, setSubmitError] = React.useState<string>('');
  const [showErrorSummary, setShowErrorSummary] = React.useState<boolean>(false);
  const [availability, setAvailability] = React.useState<IAvailability | undefined>(undefined);
  const [checkingAccess, setCheckingAccess] = React.useState<boolean>(false);
  const [draftId, setDraftId] = React.useState<number | undefined>(undefined);
  const [draftState, setDraftState] = React.useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  // the form is held back until any saved draft has been restored, so a late
  // draft can never overwrite what the respondent has already typed
  const [draftLoading, setDraftLoading] = React.useState<boolean>(!isPreview && settings.allowSaveDraft === true);
  // file/signature questions whose stored file couldn't be restored (id + title)
  const [missing, setMissing] = React.useState<{ ids: string[]; titles: string[] }>({ ids: [], titles: [] });
  const [failedAttachments, setFailedAttachments] = React.useState<string[]>([]);
  // edit-my-response
  const [myResponseId, setMyResponseId] = React.useState<number | undefined>(undefined);
  const [checkingEdit, setCheckingEdit] = React.useState<boolean>(!isPreview && settings.allowEdit === true);
  const [editingId, setEditingId] = React.useState<number | undefined>(undefined);
  const [loadingEdit, setLoadingEdit] = React.useState<boolean>(false);
  const [updated, setUpdated] = React.useState<boolean>(false);
  const submitGuard = React.useRef<boolean>(false);
  const saveGuard = React.useRef<boolean>(false);
  const locale = props.locale;

  // how long the response took, for the dashboard's completion-time stat
  const startedAt = React.useRef<number>(Date.now());
  // per-respondent seed so shuffled options and questions stay put across renders
  const shuffleSeed = React.useRef<number>(Math.floor(Math.random() * 0x7fffffff));
  const fieldRefs = React.useRef<{ [fieldId: string]: HTMLDivElement | null }>({});

  // ----- availability (open/close dates, response cap, one-per-person) -----

  const needsAccessCheck =
    !isPreview &&
    (!!settings.openDate ||
      !!settings.closeDate ||
      (typeof settings.maxResponses === 'number' && settings.maxResponses > 0) ||
      settings.oneResponsePerPerson === true);

  React.useEffect(() => {
    let cancelled = false;
    if (!needsAccessCheck) {
      setAvailability(undefined);
      return undefined;
    }
    setCheckingAccess(true);
    const wantsCount = typeof settings.maxResponses === 'number' && settings.maxResponses > 0;
    Promise.all([
      wantsCount ? spService.countResponses(listId) : Promise.resolve(undefined),
      settings.oneResponsePerPerson ? spService.currentUserHasResponded(listId) : Promise.resolve(false)
    ])
      .then(([responseCount, alreadyAnswered]) => {
        if (cancelled) {
          return;
        }
        setAvailability(
          formAvailability(definition, { responseCount, alreadyAnswered, messages: messageBag, locale: props.locale })
        );
        setCheckingAccess(false);
      })
      .catch(() => {
        if (cancelled) {
          return;
        }
        // a failed check must not lock people out of an otherwise open form
        setAvailability(formAvailability(definition, { messages: messageBag, locale: props.locale }));
        setCheckingAccess(false);
      });
    return () => {
      cancelled = true;
    };
  }, [listId, needsAccessCheck, settings.openDate, settings.closeDate, settings.maxResponses, settings.oneResponsePerPerson]);

  // ----- section / field visibility -----

  // one dependency-ordered pass, so hidden drivers behave identically everywhere
  const visibility = React.useMemo(() => computeVisibility(definition, values), [definition, values]);

  const visibleSections = React.useMemo(() => {
    // a field's visibility already folds in its section's rule; a section whose
    // every field is branched away would otherwise render as a blank wizard step
    return (definition.sections || []).filter(
      (section) => section.fields.filter((f) => visibility[f.id] !== false).length > 0
    );
  }, [definition, visibility]);

  const isWizard = settings.layout === 'wizard' && visibleSections.length > 1;

  const orderedFieldsOf = React.useCallback(
    (section: IFormSection): IFormField[] => {
      const visible = section.fields.filter((f) => visibility[f.id] !== false);
      if (!settings.shuffleQuestions) {
        return visible;
      }
      // content blocks are anchors, not questions — shuffling them would scramble
      // the instructions they introduce
      const questions = visible.filter((f) => isInputType(f.type));
      const shuffled = shuffleWithSeed(questions, shuffleSeed.current);
      let next = 0;
      return visible.map((field) => (isInputType(field.type) ? shuffled[next++] : field));
    },
    [visibility, settings.shuffleQuestions]
  );

  /** Every visible field across every visible section, in display order. */
  const visibleOrdered = React.useMemo(() => {
    const result: IFormField[] = [];
    visibleSections.forEach((section) => orderedFieldsOf(section).forEach((f) => result.push(f)));
    return result;
  }, [visibleSections, orderedFieldsOf]);

  const numberMap = React.useMemo(() => buildNumberMap(visibleOrdered), [visibleOrdered]);

  // ----- value updates -----

  const setValue = (field: IFormField, value: unknown): void => {
    setValues((prev) => {
      const next = { ...prev, [field.id]: value };
      // recompute totals immediately so a Calculated field tracks as you type
      return applyCalculatedFields(definition, next);
    });
    setMissing((prev) =>
      prev.ids.indexOf(field.id) === -1
        ? prev
        : {
            ids: prev.ids.filter((id) => id !== field.id),
            titles: prev.titles.filter((_t, i) => prev.ids[i] !== field.id)
          }
    );
    setErrors((prev) => {
      if (!prev[field.id]) {
        return prev;
      }
      const next = { ...prev };
      delete next[field.id];
      return next;
    });
  };

  const validateSections = (sectionsToCheck: IFormSection[]): boolean => {
    const nextErrors: Record<string, string> = {};
    sectionsToCheck.forEach((section) => {
      orderedFieldsOf(section).forEach((field) => {
        if (!isInputType(field.type)) {
          return;
        }
        const error = validateField(field, values[field.id], messageBag);
        if (error) {
          nextErrors[field.id] = error;
        }
      });
    });
    setErrors(nextErrors);
    const ok = Object.keys(nextErrors).length === 0;
    setShowErrorSummary(!ok);
    if (!ok) {
      // move focus to the summary so a keyboard or screen reader user isn't left
      // wondering why nothing happened
      const firstId = Object.keys(nextErrors)[0];
      focusField(firstId);
    }
    return ok;
  };

  const focusField = (fieldId: string): void => {
    const container = fieldRefs.current[fieldId];
    if (!container) {
      return;
    }
    container.scrollIntoView({ block: 'center', behavior: 'smooth' });
    const focusable = container.querySelector<HTMLElement>(
      'input, textarea, select, button, [contenteditable="true"], [tabindex]:not([tabindex="-1"])'
    );
    if (focusable) {
      focusable.focus();
    }
  };

  // ----- submit / draft -----

  /** Apply a restored draft/response; file questions that couldn't be restored must be re-attached. */
  const applyRestored = (restored: IDraft, base: IFormValues): IFormValues => {
    const next: IFormValues = { ...base, ...restored.values };
    const ids = restored.missingFileFieldIds || [];
    ids.forEach((id) => {
      const field = inputFields(definition).filter((f) => f.id === id)[0];
      // a required question with only some files back would pass validation
      // while silently dropping the rest, so make the respondent redo it
      if (field && field.required) {
        delete next[id];
      }
    });
    setMissing({ ids: ids.slice(), titles: (restored.missingFiles || []).slice() });
    return applyCalculatedFields(definition, next);
  };

  const handleSubmit = async (): Promise<void> => {
    if (submitGuard.current || draftLoading || loadingEdit) {
      return;
    }
    if (!validateSections(visibleSections)) {
      return;
    }
    if (isPreview) {
      // a preview must never write to the response list
      setSubmitted(true);
      return;
    }
    submitGuard.current = true;
    setSubmitting(true);
    setSubmitError('');
    try {
      let result: ISubmitResult;
      if (editingId !== undefined) {
        result = await spService.updateResponse(listId, definition, editingId, values, settings.allowEdit === true, {
          locale: locale
        });
        setUpdated(true);
      } else {
        const durationSeconds = Math.max(1, Math.round((Date.now() - startedAt.current) / 1000));
        result = await spService.submitResponseDetailed(listId, definition, values, {
          durationSeconds,
          replaceItemId: draftId,
          locale: locale
        });
        // fire-and-forget: notification emails never block or fail the submission,
        // but a failure is still worth a trace when someone is troubleshooting
        void spService
          .sendResponseNotifications(listId, definition, values, result.id)
          .catch((error) => logError('sendResponseNotifications', error));
      }
      setFailedAttachments(result.failedAttachments || []);
      setSubmitted(true);
      if (props.onSubmitted) {
        props.onSubmitted();
      }
    } catch (error) {
      logError('submitResponse', error);
      setSubmitError(strings.Form_Submit_SaveError);
    }
    submitGuard.current = false;
    setSubmitting(false);
  };

  const handleSaveDraft = async (): Promise<void> => {
    if (isPreview) {
      setDraftState('saved');
      return;
    }
    if (saveGuard.current || submitGuard.current || draftLoading) {
      return;
    }
    saveGuard.current = true;
    setDraftState('saving');
    setFailedAttachments([]);
    try {
      const result = await spService.saveDraftDetailed(listId, definition, values, draftId, locale);
      setDraftId(result.id);
      setFailedAttachments(result.failedAttachments || []);
      setDraftState('saved');
    } catch (error) {
      logError('saveDraft', error);
      setDraftState('error');
    }
    saveGuard.current = false;
  };

  // resume an existing draft for this respondent; the form stays blocked until
  // this settles, so restoring can never clobber anything the user typed
  React.useEffect(() => {
    let cancelled = false;
    if (isPreview || settings.allowSaveDraft !== true) {
      setDraftLoading(false);
      return undefined;
    }
    setDraftLoading(true);
    spService
      .loadDraft(listId, definition)
      .then((draft) => {
        if (cancelled) {
          return;
        }
        if (draft) {
          setDraftId(draft.id);
          setValues((prev) => applyRestored(draft, prev));
        }
        setDraftLoading(false);
      })
      .catch((error) => {
        logError('loadDraft', error);
        if (!cancelled) {
          setDraftLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [listId, isPreview, settings.allowSaveDraft]);

  // does this respondent have a completed response they may edit?
  React.useEffect(() => {
    let cancelled = false;
    if (isPreview || settings.allowEdit !== true) {
      setMyResponseId(undefined);
      setCheckingEdit(false);
      return undefined;
    }
    setCheckingEdit(true);
    spService
      .getMyResponses(listId, definition, 1)
      .then((mine) => {
        if (!cancelled) {
          setMyResponseId(mine.length > 0 ? mine[0].id : undefined);
          setCheckingEdit(false);
        }
      })
      .catch((error) => {
        logError('getMyResponses', error);
        if (!cancelled) {
          setCheckingEdit(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [listId, isPreview, settings.allowEdit]);

  const startEdit = (): void => {
    if (myResponseId === undefined || loadingEdit) {
      return;
    }
    setLoadingEdit(true);
    setSubmitError('');
    spService
      .loadResponseForEdit(listId, definition, myResponseId)
      .then((existing) => {
        if (!existing) {
          setSubmitError(strings.Form_Edit_LoadError);
          setLoadingEdit(false);
          return;
        }
        setValues(applyRestored(existing, buildDefaultValues(definition)));
        setEditingId(existing.id);
        setErrors({});
        setStep(0);
        setShowErrorSummary(false);
        setLoadingEdit(false);
      })
      .catch((error) => {
        logError('loadResponseForEdit', error);
        setSubmitError(strings.Form_Edit_LoadError);
        setLoadingEdit(false);
      });
  };

  const handleReset = (): void => {
    setValues(buildDefaultValues(definition));
    setErrors({});
    setStep(0);
    setSubmitted(false);
    setSubmitError('');
    setShowErrorSummary(false);
    setDraftId(undefined);
    setDraftState('idle');
    setMissing({ ids: [], titles: [] });
    setFailedAttachments([]);
    setEditingId(undefined);
    setUpdated(false);
    startedAt.current = Date.now();
    shuffleSeed.current = Math.floor(Math.random() * 0x7fffffff);
  };

  // ----- early returns -----

  const questionCount = inputFields(definition).length;

  if (questionCount === 0) {
    return (
      <div className={styles.formCard}>
        <div className={styles.emptyForm}>
          <Icon iconName="PageEdit" className={styles.emptyFormIcon} />
          <p>{strings.Form_Empty_Message}</p>
        </div>
      </div>
    );
  }

  if (checkingAccess || checkingEdit || draftLoading || loadingEdit) {
    return (
      <div className={styles.formCard}>
        <div className={styles.emptyForm}>
          <Spinner size={SpinnerSize.large} label={strings.Form_Checking} />
        </div>
      </div>
    );
  }

  // editing your own response is exempt from the response cap and one-per-person rule
  if (availability && availability.state !== 'open' && editingId === undefined) {
    const canEdit =
      myResponseId !== undefined &&
      settings.allowEdit === true &&
      (availability.state === 'alreadyAnswered' || availability.state === 'full');
    const icon = availability.state === 'notYetOpen' ? 'Clock' : availability.state === 'alreadyAnswered' ? 'CheckMark' : 'Lock';
    return (
      <div className={styles.formCard}>
        <div className={styles.closedNotice}>
          <div className={styles.closedIcon}>
            <Icon iconName={icon} />
          </div>
          <h2>{availability.state === 'notYetOpen'
              ? strings.Form_Closed_NotYetOpen
              : availability.state === 'alreadyAnswered'
                ? strings.Form_Closed_AlreadyAnswered
                : strings.Form_Closed_Title}</h2>
          <p>{availability.message}</p>
          {canEdit && (
            <DefaultButton iconProps={{ iconName: 'Edit' }} text={strings.Form_Edit_Button} onClick={startEdit} />
          )}
        </div>
      </div>
    );
  }

  if (submitted) {
    return (
      <div className={styles.formCard}>
        <div className={styles.confirmation}>
          <div className={styles.confirmationIcon}>
            <Icon iconName="CheckMark" />
          </div>
          <h2>{updated ? strings.Form_Edit_UpdatedTitle : settings.confirmationTitle}</h2>
          <p>{updated ? strings.Form_Edit_UpdatedMessage : settings.confirmationMessage}</p>
          {failedAttachments.length > 0 && (
            <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 16 } }}>
              {formatString(strings.Form_Attachments_Failed, { names: failedAttachments.join(', ') })}
            </MessageBar>
          )}
          {isPreview && (
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 16 } }}>
              {strings.Form_Confirmation_PreviewNotice}
            </MessageBar>
          )}
          <div className={styles.confirmationActions}>
            {((settings.allowAnotherResponse && !updated) || isPreview) && (
              <DefaultButton
                iconProps={{ iconName: 'Refresh' }}
                text={isPreview ? strings.Form_Confirmation_PreviewAgain : strings.Form_Confirmation_SubmitAnother}
                onClick={handleReset}
              />
            )}
          </div>
        </div>
      </div>
    );
  }

  // ----- rendering -----

  const renderField = (field: IFormField): React.ReactNode => {
    const widthClass =
      field.width === 'half' ? styles.fieldHalf : field.width === 'third' ? styles.fieldThird : styles.fieldFull;

    if (!isInputType(field.type)) {
      return (
        <div key={field.id} className={widthClass}>
          <ContentBlock html={field.contentHtml} imageUrl={field.contentImageUrl} style={field.contentStyle} />
        </div>
      );
    }

    const error = errors[field.id];
    const labelId = 'lbl-' + field.id;
    const helpId = field.description ? 'hlp-' + field.id : undefined;
    const errorId = error ? 'err-' + field.id : undefined;
    const describedBy = [helpId, errorId].filter((id) => !!id).join(' ') || undefined;
    const number = numberMap[field.id];

    return (
      <div
        key={field.id}
        className={widthClass}
        ref={(element) => {
          fieldRefs.current[field.id] = element;
        }}
      >
        {/*
          A role=group labelled by the question text is what makes composite
          controls (address, likert, ranking, choice sets) announce their question
          at all. A bare <label for> only works for single-input fields, and the
          previous markup wrapped nothing and pointed at nothing.
        */}
        <div
          role="group"
          aria-labelledby={labelId}
          aria-describedby={describedBy}
        >
          <span className={styles.fieldLabel} id={labelId}>
            {settings.showQuestionNumbers && number ? (
              <span className={styles.fieldNumber}>{number}.</span>
            ) : null}
            {field.title}
            {field.required && (
              <span className={styles.requiredMark} aria-hidden={true}>
                {' '}
                *
              </span>
            )}
            {field.required && <span className={styles.hiddenInput}>{strings.Form_Field_RequiredAria}</span>}
          </span>
          {field.description && (
            <div className={styles.fieldDescription} id={helpId}>
              {field.description}
            </div>
          )}
          <FieldControl
            field={field}
            value={values[field.id]}
            error={error}
            describedBy={describedBy}
            controlId={'ctl-' + field.id}
            shuffleSeed={shuffleSeed.current}
            onChange={(v) => setValue(field, v)}
            spService={spService}
          />
          {error && (
            <div className={styles.fieldError} id={errorId} role="alert">
              <Icon iconName="ErrorBadge" /> {error}
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderSection = (section: IFormSection): React.ReactNode => {
    const fields = orderedFieldsOf(section);
    if (fields.length === 0) {
      return null;
    }
    return (
      <div key={section.id} className={styles.section}>
        {(section.title || section.description) && !isWizard && (
          <div className={styles.sectionHeader}>
            {section.title && <h3>{section.title}</h3>}
            {section.description && <p>{section.description}</p>}
          </div>
        )}
        <div className={styles.fieldGrid}>{fields.map(renderField)}</div>
      </div>
    );
  };

  const currentStep = Math.min(step, Math.max(0, visibleSections.length - 1));
  const currentSection = visibleSections[currentStep];

  const goNext = (): void => {
    if (!currentSection || !validateSections([currentSection])) {
      return;
    }
    setShowErrorSummary(false);
    setStep(Math.min(currentStep + 1, visibleSections.length - 1));
  };

  const goBack = (): void => {
    setShowErrorSummary(false);
    setStep(Math.max(0, currentStep - 1));
  };

  const errorEntries = Object.keys(errors).map((fieldId) => ({
    fieldId,
    message: errors[fieldId],
    field: visibleOrdered.filter((f) => f.id === fieldId)[0]
  }));

  const rtl = isRtlLocale(locale);

  return (
    <div className={styles.formCard} dir={rtl ? 'rtl' : undefined}>
      {settings.showFormHeader && (
        <div className={styles.formHeader}>
          <div className={styles.formHeaderIcon}>
            <Icon iconName={settings.headerIcon || 'ClipboardList'} />
          </div>
          <div>
            <h2>{settings.formTitle}</h2>
            {settings.formDescription && <p>{settings.formDescription}</p>}
          </div>
        </div>
      )}

      {isWizard && (
        <div className={styles.wizardHeader}>
          <div className={styles.stepDots} role="list" aria-label={strings.Form_Wizard_StepsAria}>
            {visibleSections.map((section, index) => {
              const state =
                index === currentStep ? styles.stepDotActive : index < currentStep ? styles.stepDotDone : styles.stepDot;
              return (
                <button
                  key={section.id}
                  type="button"
                  role="listitem"
                  className={state}
                  disabled={index > currentStep}
                  aria-current={index === currentStep ? 'step' : undefined}
                  onClick={() => index < currentStep && setStep(index)}
                >
                  <span
                    className={
                      index === currentStep ? styles.stepDotIndexActive : styles.stepDotIndex
                    }
                  >
                    {index < currentStep ? '✓' : index + 1}
                  </span>
                  {section.title || formatString(strings.Form_Wizard_StepFallback, { step: index + 1 })}
                </button>
              );
            })}
          </div>
          {settings.showProgressBar && (
            <ProgressIndicator
              ariaValueText={formatString(strings.Form_Wizard_StepOf, { step: currentStep + 1, total: visibleSections.length })}
              percentComplete={(currentStep + 1) / visibleSections.length}
            />
          )}
          {currentSection && currentSection.description && (
            <div className={styles.sectionHeader}>
              {currentSection.title && <h3>{currentSection.title}</h3>}
              <p>{currentSection.description}</p>
            </div>
          )}
        </div>
      )}

      {showErrorSummary && errorEntries.length > 1 && (
        <MessageBar
          className={styles.errorSummary}
          messageBarType={MessageBarType.error}
          onDismiss={() => setShowErrorSummary(false)}
        >
          {formatString(strings.Form_ErrorSummary_Title, { count: errorEntries.length })}
          <ul className={styles.errorSummaryList}>
            {errorEntries.map((entry) => (
              <li key={entry.fieldId}>
                <button
                  type="button"
                  className={styles.errorSummaryLink}
                  onClick={() => focusField(entry.fieldId)}
                >
                  {entry.field ? entry.field.title || strings.Form_ErrorSummary_UntitledQuestion : strings.Form_ErrorSummary_Question}
                </button>
                {formatString(strings.Form_ErrorSummary_Detail, { message: entry.message })}
              </li>
            ))}
          </ul>
        </MessageBar>
      )}

      {editingId !== undefined ? (
        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 12 } }}>
          {strings.Form_Edit_Banner}
        </MessageBar>
      ) : (
        myResponseId !== undefined &&
        settings.allowEdit === true && (
          <MessageBar
            messageBarType={MessageBarType.info}
            styles={{ root: { marginBottom: 12 } }}
            actions={
              <DefaultButton iconProps={{ iconName: 'Edit' }} text={strings.Form_Edit_Button} onClick={startEdit} />
            }
          >
            {strings.Form_Edit_Available}
          </MessageBar>
        )
      )}

      {missing.titles.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }} role="status">
          {formatString(strings.Form_Draft_MissingFiles, { names: missing.titles.join(', ') })}
        </MessageBar>
      )}

      {failedAttachments.length > 0 && draftState === 'saved' && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }} role="alert">
          {formatString(strings.Form_Attachments_Failed, { names: failedAttachments.join(', ') })}
        </MessageBar>
      )}

      <div className={styles.formBody}>
        {isWizard ? renderSection(currentSection) : visibleSections.map(renderSection)}
      </div>

      {submitError && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.submitError}>
          {submitError}
        </MessageBar>
      )}

      <div className={styles.formFooter}>
        {isWizard && currentStep > 0 && (
          <DefaultButton
            iconProps={{ iconName: 'ChevronLeft' }}
            text={strings.Form_Nav_Back}
            disabled={submitting}
            onClick={goBack}
          />
        )}

        {settings.allowSaveDraft && !isWizard && (
          <DefaultButton
            iconProps={{ iconName: draftState === 'saved' ? 'CheckMark' : 'Save' }}
            text={
              draftState === 'saving'
                ? strings.Form_Draft_Saving
                : draftState === 'saved'
                  ? strings.Form_Draft_Saved
                  : strings.Form_Draft_SaveAndFinishLater
            }
            disabled={submitting || draftState === 'saving' || draftLoading}
            onClick={() => {
              void handleSaveDraft();
            }}
          />
        )}

        <div className={styles.footerSpacer} />

        {draftState === 'error' && (
          <span className={styles.saveIndicator + ' ' + styles.saveIndicatorError}>
            <Icon iconName="Warning" /> {strings.Form_Draft_NotSaved}
          </span>
        )}

        {isWizard && currentStep < visibleSections.length - 1 ? (
          <PrimaryButton
            text={strings.Form_Nav_Next}
            onRenderIcon={() => <Icon iconName="ChevronRight" />}
            onClick={goNext}
          />
        ) : (
          <PrimaryButton
            className={styles.submitButton}
            disabled={submitting || draftLoading}
            onClick={() => {
              void handleSubmit();
            }}
            text={submitting ? undefined : isPreview
                ? strings.Form_Submit_Preview
                : editingId !== undefined
                  ? strings.Form_Edit_Save
                  : settings.submitButtonText || strings.Form_Submit_Button}
          >
            {submitting && <Spinner size={SpinnerSize.small} />}
          </PrimaryButton>
        )}
      </div>
    </div>
  );
};

/** Read-only rendering of one submitted response, laid out as the form itself. */
export interface IResponseFormViewProps {
  definition: IFormDefinition;
  values: IFormValues;
  spService: SharePointService;
}

export const ResponseFormView: React.FunctionComponent<IResponseFormViewProps> = (props) => {
  const { definition, values } = props;
  const fields = inputFields(definition);
  const numberMap = buildNumberMap(fields);

  return (
    <div className={styles.formBody}>
      {fields.map((field) => {
        const answered = values[field.id] !== undefined && values[field.id] !== null && values[field.id] !== '';
        return (
          <div key={field.id} className={styles.fieldFull}>
            <span className={styles.fieldLabel}>
              <span className={styles.fieldNumber}>{numberMap[field.id]}.</span>
              {field.title}
            </span>
            {answered ? (
              <FieldControl
                field={field}
                value={values[field.id]}
                disabled={true}
                onChange={() => undefined}
                spService={props.spService}
              />
            ) : (
              <div className={styles.readOnlyValue}>{strings.Form_Response_NoAnswer}</div>
            )}
          </div>
        );
      })}
    </div>
  );
};

/** Exported for the response detail panel, which shows values without controls. */
export const plainAnswer = (field: IFormField, value: unknown, locale?: string): string =>
  formatValue(field, value, messageBag, locale);
