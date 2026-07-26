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
  isInputType
} from '../../models';
import {
  applyCalculatedFields,
  buildNumberMap,
  formAvailability,
  formatValue,
  IAvailability,
  inputFields,
  isFieldVisible,
  isSectionVisible,
  shuffleWithSeed,
  validateField
} from '../../utils/formUtils';
import { SharePointService } from '../../services/SharePointService';
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
  /** notified after a real submission, so the shell can refresh counts */
  onSubmitted?: () => void;
}

/**
 * Prefill values from the page's query string: any parameter named after a
 * question's internal column name (e.g. ?SFYourName=Alex) becomes that
 * question's initial answer. Lets owners hand out links that pre-answer
 * routing questions — something Microsoft Forms can't do.
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
      case FieldType.ImageChoice:
        values[field.id] = field.allowMultiple
          ? raw.split(';').map((s) => s.trim()).filter((s) => s.length > 0)
          : raw;
        break;
      case FieldType.Date:
      case FieldType.Time: {
        const parsed = /^today$/i.test(raw) ? new Date() : new Date(raw);
        if (!isNaN(parsed.getTime())) {
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
        values[field.id] = field.allowMultiple
          ? field.defaultValue.split(';').map((s) => s.trim()).filter((s) => s.length > 0)
          : field.defaultValue;
        break;
      case FieldType.Date:
      case FieldType.Time:
        if (/^today$|^now$/i.test(field.defaultValue)) {
          values[field.id] = new Date();
        } else {
          const parsed = new Date(field.defaultValue);
          if (!isNaN(parsed.getTime())) {
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
        setAvailability(formAvailability(definition, { responseCount, alreadyAnswered }));
        setCheckingAccess(false);
      })
      .catch(() => {
        if (cancelled) {
          return;
        }
        // a failed check must not lock people out of an otherwise open form
        setAvailability(formAvailability(definition, {}));
        setCheckingAccess(false);
      });
    return () => {
      cancelled = true;
    };
  }, [listId, needsAccessCheck, settings.openDate, settings.closeDate, settings.maxResponses, settings.oneResponsePerPerson]);

  // ----- section / field visibility -----

  const visibleSections = React.useMemo(() => {
    return (definition.sections || []).filter((section) => {
      if (!isSectionVisible(section, definition, values)) {
        return false;
      }
      // a section whose every field is branched away would otherwise render as a
      // blank wizard step with nothing but a Next button
      return section.fields.filter((f) => isFieldVisible(f, definition, values)).length > 0;
    });
  }, [definition, values]);

  const isWizard = settings.layout === 'wizard' && visibleSections.length > 1;

  const orderedFieldsOf = React.useCallback(
    (section: IFormSection): IFormField[] => {
      const visible = section.fields.filter((f) => isFieldVisible(f, definition, values));
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
    [definition, values, settings.shuffleQuestions]
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
        const error = validateField(field, values[field.id]);
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

  const handleSubmit = async (): Promise<void> => {
    if (!validateSections(visibleSections)) {
      return;
    }
    if (isPreview) {
      // a preview must never write to the response list
      setSubmitted(true);
      return;
    }
    setSubmitting(true);
    setSubmitError('');
    try {
      const durationSeconds = Math.max(1, Math.round((Date.now() - startedAt.current) / 1000));
      const itemId = await spService.submitResponse(listId, definition, values, {
        durationSeconds,
        replaceItemId: draftId
      });
      // fire-and-forget: notification emails never block or fail the submission,
      // but a failure is still worth a trace when someone is troubleshooting
      void spService
        .sendResponseNotifications(listId, definition, values, itemId)
        .catch((error) => logError('sendResponseNotifications', error));
      setSubmitted(true);
      if (props.onSubmitted) {
        props.onSubmitted();
      }
    } catch (error) {
      logError('submitResponse', error);
      setSubmitError(
        'Your response could not be saved. Make sure the form has been published from the designer and that you have permission to add items to the response list.'
      );
    }
    setSubmitting(false);
  };

  const handleSaveDraft = async (): Promise<void> => {
    if (isPreview) {
      setDraftState('saved');
      return;
    }
    setDraftState('saving');
    try {
      const id = await spService.saveDraft(listId, definition, values, draftId);
      setDraftId(id);
      setDraftState('saved');
    } catch (error) {
      logError('saveDraft', error);
      setDraftState('error');
    }
  };

  // resume an existing draft for this respondent
  React.useEffect(() => {
    let cancelled = false;
    if (isPreview || settings.allowSaveDraft !== true) {
      return undefined;
    }
    spService
      .loadDraft(listId, definition)
      .then((draft) => {
        if (cancelled || !draft) {
          return;
        }
        setDraftId(draft.id);
        setValues((prev) => applyCalculatedFields(definition, { ...prev, ...draft.values }));
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [listId, isPreview, settings.allowSaveDraft]);

  const handleReset = (): void => {
    setValues(buildDefaultValues(definition));
    setErrors({});
    setStep(0);
    setSubmitted(false);
    setSubmitError('');
    setShowErrorSummary(false);
    setDraftId(undefined);
    setDraftState('idle');
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
          <p>This form doesn&apos;t have any questions yet. The form owner can add them in the Questions tab.</p>
        </div>
      </div>
    );
  }

  if (checkingAccess) {
    return (
      <div className={styles.formCard}>
        <div className={styles.emptyForm}>
          <Spinner size={SpinnerSize.large} label="Checking the form…" />
        </div>
      </div>
    );
  }

  if (availability && availability.state !== 'open') {
    const icon = availability.state === 'notYetOpen' ? 'Clock' : availability.state === 'alreadyAnswered' ? 'CheckMark' : 'Lock';
    return (
      <div className={styles.formCard}>
        <div className={styles.closedNotice}>
          <div className={styles.closedIcon}>
            <Icon iconName={icon} />
          </div>
          <h2>{availability.state === 'notYetOpen' ? 'Not open yet' : availability.state === 'alreadyAnswered' ? 'Already answered' : 'Form closed'}</h2>
          <p>{availability.message}</p>
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
          <h2>{settings.confirmationTitle}</h2>
          <p>{settings.confirmationMessage}</p>
          {isPreview && (
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 16 } }}>
              This was a preview — nothing was saved to the response list.
            </MessageBar>
          )}
          <div className={styles.confirmationActions}>
            {(settings.allowAnotherResponse || isPreview) && (
              <DefaultButton
                iconProps={{ iconName: 'Refresh' }}
                text={isPreview ? 'Preview again' : 'Submit another response'}
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
            {field.required && <span className={styles.hiddenInput}>(required)</span>}
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

  return (
    <div className={styles.formCard}>
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
          <div className={styles.stepDots} role="list" aria-label="Form steps">
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
                  {section.title || 'Step ' + (index + 1)}
                </button>
              );
            })}
          </div>
          {settings.showProgressBar && (
            <ProgressIndicator
              ariaValueText={'Step ' + (currentStep + 1) + ' of ' + visibleSections.length}
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
          Please check {errorEntries.length} answers:
          <ul className={styles.errorSummaryList}>
            {errorEntries.map((entry) => (
              <li key={entry.fieldId}>
                <button
                  type="button"
                  className={styles.errorSummaryLink}
                  onClick={() => focusField(entry.fieldId)}
                >
                  {entry.field ? entry.field.title || 'Untitled question' : 'Question'}
                </button>
                {' — ' + entry.message}
              </li>
            ))}
          </ul>
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
            text="Back"
            disabled={submitting}
            onClick={goBack}
          />
        )}

        {settings.allowSaveDraft && !isWizard && (
          <DefaultButton
            iconProps={{ iconName: draftState === 'saved' ? 'CheckMark' : 'Save' }}
            text={
              draftState === 'saving'
                ? 'Saving…'
                : draftState === 'saved'
                  ? 'Draft saved'
                  : 'Save and finish later'
            }
            disabled={submitting || draftState === 'saving'}
            onClick={() => {
              void handleSaveDraft();
            }}
          />
        )}

        <div className={styles.footerSpacer} />

        {draftState === 'error' && (
          <span className={styles.saveIndicator + ' ' + styles.saveIndicatorError}>
            <Icon iconName="Warning" /> Draft not saved
          </span>
        )}

        {isWizard && currentStep < visibleSections.length - 1 ? (
          <PrimaryButton
            text="Next"
            onRenderIcon={() => <Icon iconName="ChevronRight" />}
            onClick={goNext}
          />
        ) : (
          <PrimaryButton
            className={styles.submitButton}
            disabled={submitting}
            onClick={() => {
              void handleSubmit();
            }}
            text={submitting ? undefined : isPreview ? 'Submit (preview)' : settings.submitButtonText || 'Submit'}
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
              <div className={styles.readOnlyValue}>No answer</div>
            )}
          </div>
        );
      })}
    </div>
  );
};

/** Exported for the response detail panel, which shows values without controls. */
export const plainAnswer = (field: IFormField, value: unknown): string => formatValue(field, value);
