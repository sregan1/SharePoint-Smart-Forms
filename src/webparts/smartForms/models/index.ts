/**
 * Core data models for Smart Forms.
 *
 * A form is described by an IFormDefinition (persisted as JSON in the web part
 * properties). Each input field maps to a real SharePoint column that the
 * SharePointService provisions on demand.
 *
 * Schema history
 *   v1 — original 19 discrete field types, single-rule branching.
 *   v2 — Currency/Percent folded into Number, DateTime folded into Date,
 *        Location dropped, Nps generalized to Scale; branching became a
 *        multi-condition group. See migrateDefinition().
 */

export const CURRENT_SCHEMA_VERSION = 2;

export enum FieldType {
  // ----- text -----
  Text = 'Text',
  MultilineText = 'MultilineText',
  RichText = 'RichText',
  // ----- numeric (currency / percent are formats, not types) -----
  Number = 'Number',
  Calculated = 'Calculated',
  // ----- date & time -----
  Date = 'Date',
  Time = 'Time',
  // ----- choice family -----
  Choice = 'Choice',
  ImageChoice = 'ImageChoice',
  Lookup = 'Lookup',
  YesNo = 'YesNo',
  Consent = 'Consent',
  // ----- scales -----
  Rating = 'Rating',
  Slider = 'Slider',
  Scale = 'Scale',
  Likert = 'Likert',
  Ranking = 'Ranking',
  // ----- contact & identity -----
  Email = 'Email',
  Phone = 'Phone',
  Hyperlink = 'Hyperlink',
  Address = 'Address',
  Person = 'Person',
  // ----- rich input -----
  FileUpload = 'FileUpload',
  Signature = 'Signature',
  // ----- layout (no stored value) -----
  Content = 'Content'
}

/** Field types that collect an answer and therefore need a list column. */
export const isInputType = (type: FieldType): boolean => type !== FieldType.Content;

export type ConditionOperator =
  | 'equals'
  | 'notEquals'
  | 'contains'
  | 'notContains'
  | 'empty'
  | 'notEmpty'
  | 'greaterThan'
  | 'greaterOrEqual'
  | 'lessThan'
  | 'lessOrEqual'
  | 'between'
  | 'before'
  | 'after';

/** Operators that need no comparison value. */
export const UNARY_OPERATORS: ConditionOperator[] = ['empty', 'notEmpty'];

/** Operators that compare numerically or chronologically rather than as text. */
export const RANGE_OPERATORS: ConditionOperator[] = [
  'greaterThan',
  'greaterOrEqual',
  'lessThan',
  'lessOrEqual',
  'between',
  'before',
  'after'
];

export interface ICondition {
  /** id of the field whose value drives the rule */
  fieldId: string;
  operator: ConditionOperator;
  /** comparison value (unused for empty / notEmpty) */
  value?: string;
  /** upper bound, only used by 'between' */
  value2?: string;
}

export type ConditionMatch = 'all' | 'any';

export interface IConditionGroup {
  match: ConditionMatch;
  conditions: ICondition[];
}

export type FieldWidth = 'full' | 'half' | 'third';
export type ChoiceDisplay = 'dropdown' | 'buttons';
export type NumberFormat = 'plain' | 'currency' | 'percent';
export type RatingIcon = 'star' | 'heart' | 'like';
export type ScaleAnalytics = 'nps' | 'average';
export type ContentStyle = 'text' | 'info' | 'success' | 'warning' | 'divider';

/** One uploaded file held in form state before it becomes a list attachment. */
export interface IFormFile {
  name: string;
  size: number;
  /** base64 payload without the data: prefix */
  content: string;
}

/** Structured postal address held in form state. */
export interface IAddressValue {
  street?: string;
  street2?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  country?: string;
}

/** Likert answers keyed by row label. */
export interface ILikertValue {
  [row: string]: string;
}

export interface IImageChoiceOption {
  label: string;
  /** absolute or site-relative image URL */
  imageUrl: string;
}

export interface IFormField {
  /** stable unique id within the form definition */
  id: string;
  /** SharePoint internal column name (generated, letters and digits only) */
  internalName: string;
  title: string;
  type: FieldType;
  /** help text shown under the label */
  description?: string;
  placeholder?: string;
  required?: boolean;
  /** overrides the generic "<title> is required" message */
  requiredMessage?: string;
  defaultValue?: string;
  width?: FieldWidth;
  /** prefilled from the URL but not editable by the respondent */
  readOnly?: boolean;

  // ----- Choice / ImageChoice / Ranking / Likert -----
  choices?: string[];
  imageChoices?: IImageChoiceOption[];
  allowMultiple?: boolean;
  choiceDisplay?: ChoiceDisplay;
  /** append a write-in "Other" option */
  allowOther?: boolean;
  otherLabel?: string;
  minSelections?: number;
  maxSelections?: number;
  /** randomize option order per respondent */
  shuffleOptions?: boolean;

  // ----- Lookup -----
  lookupListId?: string;
  /** internal name of the column supplying the option text */
  lookupColumn?: string;
  /** optional OData filter narrowing the source items */
  lookupFilter?: string;

  // ----- Likert -----
  likertRows?: string[];
  likertColumns?: string[];

  // ----- Number / Calculated / Slider / Scale -----
  min?: number;
  max?: number;
  step?: number;
  decimalPlaces?: number;
  numberFormat?: NumberFormat;
  currencySymbol?: string;
  /** trailing unit shown after the value, e.g. "kg" or "hrs" */
  unitSuffix?: string;
  /** arithmetic over other fields, e.g. "{Quantity} * {Unit price}" */
  formula?: string;

  // ----- Scale (generalizes the old NPS field) -----
  lowLabel?: string;
  highLabel?: string;
  /** which summary treatment the dashboard applies */
  scaleAnalytics?: ScaleAnalytics;

  // ----- Rating -----
  maxRating?: number;
  ratingIcon?: RatingIcon;
  allowHalfRating?: boolean;

  // ----- Date / Time -----
  /** Date fields with a time-of-day component (was the DateTime type) */
  includeTime?: boolean;
  /** minute granularity of the time dropdown */
  timeStepMinutes?: number;

  // ----- Text -----
  maxLength?: number;
  rows?: number;
  /** regular expression the answer must match */
  pattern?: string;
  patternMessage?: string;

  // ----- Person -----
  allowMultiplePeople?: boolean;
  /** include SharePoint / Entra groups in the picker */
  allowGroups?: boolean;

  // ----- FileUpload -----
  maxFiles?: number;
  maxFileSizeMb?: number;
  /** lowercase extensions without the dot */
  allowedExtensions?: string[];

  // ----- Consent -----
  consentText?: string;

  // ----- Content (static block, stores nothing) -----
  contentHtml?: string;
  contentStyle?: ContentStyle;
  contentImageUrl?: string;

  /** show this field only when the rule group is satisfied */
  visibleWhen?: IConditionGroup;

  /** true once the matching SharePoint column has been created */
  provisioned?: boolean;
}

export interface IFormSection {
  id: string;
  title: string;
  description?: string;
  fields: IFormField[];
  /** wizard-only: jump straight to this section id after this one */
  nextSectionId?: string;
  /** show this whole section only when the rule group is satisfied */
  visibleWhen?: IConditionGroup;
}

export type FormLayout = 'singlePage' | 'wizard';

// ----- dashboard configuration -----

export type ChartKind =
  | 'auto'
  | 'bar'
  | 'column'
  | 'donut'
  | 'stat'
  | 'gauge'
  | 'histogram'
  | 'words'
  | 'table';

export type TimelineGrain = 'day' | 'week' | 'month';

export interface IDashboardTile {
  fieldId: string;
  chart: ChartKind;
  /** hidden from the dashboard (still available in Summary) */
  hidden?: boolean;
  /** pinned tiles render before auto-selected ones */
  pinned?: boolean;
}

export interface IDashboardSettings {
  showKpis: boolean;
  showTimeline: boolean;
  timelineGrain: TimelineGrain;
  /** explicit tile order / chart choices; fields not listed fall back to auto */
  tiles: IDashboardTile[];
}

export const DEFAULT_DASHBOARD_SETTINGS: IDashboardSettings = {
  showKpis: true,
  showTimeline: true,
  timelineGrain: 'day',
  tiles: []
};

export interface IFormSettings {
  layout: FormLayout;
  formTitle: string;
  formDescription?: string;
  showFormHeader: boolean;
  headerIcon: string;
  accentColor: string;
  submitButtonText: string;
  confirmationTitle: string;
  confirmationMessage: string;
  allowAnotherResponse: boolean;
  showProgressBar: boolean;
  /** number every question in the rendered form */
  showQuestionNumbers: boolean;
  /** randomize question order within each section */
  shuffleQuestions?: boolean;

  /** semicolon/comma-separated org emails notified with every response */
  notifyEmails?: string;
  /** email the respondent a copy of their answers */
  respondentReceipt?: boolean;

  // ----- access & lifecycle -----
  /** ISO date; before this the form is closed */
  openDate?: string;
  /** ISO date; after this the form is closed */
  closeDate?: string;
  /** stop accepting responses after this many submissions */
  maxResponses?: number;
  /** block a second submission from the same account */
  oneResponsePerPerson?: boolean;
  /** shown instead of the form when closed */
  closedMessage?: string;
  /** let respondents save a partial response and resume later */
  allowSaveDraft?: boolean;
  /** let respondents open and edit their own earlier response */
  allowEdit?: boolean;
  /** route completed responses through an approve / reject step */
  enableApproval?: boolean;
  /** email the respondent when their response is approved or rejected (needs enableApproval) */
  approvalNotify?: boolean;

  dashboard?: IDashboardSettings;
}

export interface IFormDefinition {
  schemaVersion: number;
  sections: IFormSection[];
  settings: IFormSettings;
  /**
   * Internal column names once used by a published question that has since been
   * deleted. Their SharePoint columns still exist, so generateInternalName must
   * never hand the same name to a new question (it would silently adopt the old
   * column's type and data). See retireField() in formUtils.
   */
  retiredColumns?: string[];
}

export const DEFAULT_FORM_SETTINGS: IFormSettings = {
  layout: 'singlePage',
  formTitle: 'Untitled form',
  formDescription: '',
  showFormHeader: true,
  headerIcon: 'ClipboardList',
  accentColor: '#0078d4',
  submitButtonText: 'Submit',
  confirmationTitle: 'Thank you!',
  confirmationMessage: 'Your response has been recorded.',
  allowAnotherResponse: true,
  showProgressBar: true,
  showQuestionNumbers: true,
  shuffleQuestions: false,
  notifyEmails: '',
  respondentReceipt: false,
  oneResponsePerPerson: false,
  closedMessage: 'This form is no longer accepting responses.',
  allowSaveDraft: false,
  allowEdit: false,
  enableApproval: false,
  approvalNotify: false,
  dashboard: { ...DEFAULT_DASHBOARD_SETTINGS }
};


// ---------------------------------------------------------------------------
// user-visible strings produced by browser-free logic
// ---------------------------------------------------------------------------

/**
 * These modules are compiled for Node tests and must not import the SPFx loc
 * module, so every English string they can produce is defined here and callers
 * may pass a translated bag (same keys, prefixed "Logic_") in its place. Missing
 * keys fall back to English. Tokens look like {label}.
 */
export interface IMessageBag {
  [key: string]: string;
}

const STATIC_MESSAGES: IMessageBag = {
  // validation
  Logic_ThisQuestion: 'This question',
  Logic_Required: '{label} is required',
  Logic_RequiredConsent: 'Please tick the box to continue',
  Logic_RequiredFile: 'Please attach at least one file',
  Logic_RequiredSignature: 'Please add your signature',
  Logic_RequiredLikert: 'Please answer every row',
  Logic_InvalidEmail: 'Enter a valid email address',
  Logic_InvalidPhone: 'Enter a valid phone number',
  Logic_InvalidUrl: 'Enter a valid web address starting with http:// or https://',
  Logic_MinValue: 'Value must be at least {value}',
  Logic_MaxValue: 'Value must be at most {value}',
  Logic_MaxLength: 'Maximum length is {max} characters',
  Logic_ChooseAtLeastOne: 'Choose at least {count} option',
  Logic_ChooseAtLeastMany: 'Choose at least {count} options',
  Logic_ChooseAtMostOne: 'Choose no more than {count} option',
  Logic_ChooseAtMostMany: 'Choose no more than {count} options',
  Logic_LikertRemaining: 'Please answer every row ({count} remaining)',
  Logic_MaxFilesOne: 'Attach no more than {count} file',
  Logic_MaxFilesMany: 'Attach no more than {count} files',
  Logic_FileTooLarge: '"{name}" is larger than {size} MB',
  Logic_FileTypeNotAccepted: '"{name}" is not an accepted file type ({types})',
  Logic_AddressRequired: 'Enter at least a street and a city',
  Logic_PatternDefault: 'This answer is not in the expected format',
  // value formatting
  Logic_Yes: 'Yes',
  Logic_No: 'No',
  Logic_Agreed: 'Agreed',
  Logic_NotAgreed: 'Not agreed',
  Logic_Signed: 'Signed',
  // chart / segment labels
  Logic_NoAnswer: '(no answer)',
  Logic_Unknown: '(unknown)',
  Logic_OtherSlice: 'Other ({count})',
  // CSV
  Logic_CsvResponseId: 'Response ID',
  Logic_CsvSubmitted: 'Submitted',
  Logic_CsvSubmittedBy: 'Submitted by',
  // notification email
  Logic_EmailQuestion: 'Question',
  Logic_EmailViewItem: 'View in Microsoft Lists',
  Logic_EmailFooter: 'Sent automatically by Smart Forms.',
  // availability
  Logic_ClosedDefault: 'This form is no longer accepting responses.',
  Logic_NotYetOpen: 'This form opens on {date}.',
  Logic_AlreadyResponded: 'You have already responded to this form. Thank you!',
  // pre-flight checks
  Logic_UntitledQuestion: 'An untitled question',
  Logic_Issue_NoQuestions: 'Add at least one question before collecting responses.',
  Logic_Issue_NoTitle: 'A question has no text — it will be named automatically when you publish.',
  Logic_Issue_DuplicateTitle: '{label} shares its text with another question, which makes results harder to read.',
  Logic_Issue_NeedTwoOptions: '{label} needs at least two options.',
  Logic_Issue_ImageNoImage: '{label} has options with no image — they will show as text tiles.',
  Logic_Issue_LikertNeedRow: '{label} needs at least one row.',
  Logic_Issue_LikertNeedColumns: '{label} needs at least two scale columns.',
  Logic_Issue_NoLookupList: '{label} has no source list selected.',
  Logic_Issue_NoFormula: '{label} has no formula.',
  Logic_Issue_BadFormula: '{label} has a formula that cannot be worked out. Check the brackets and operators.',
  Logic_Issue_MissingRefs: '{label} refers to questions that do not exist: {names}.',
  Logic_Issue_BadPattern: '{label} has an invalid validation pattern.',
  Logic_Issue_BrokenRule: '{label} has a branching rule pointing at a deleted question.',
  Logic_Issue_SelfRule: '{label} has a branching rule that refers to itself.',
  Logic_Issue_RuleMismatch: '{label} has a branching rule that no longer fits the question it depends on.',
  Logic_Issue_RequiredReadOnly:
    '{label} is both required and read-only, so it can only be answered by a prefilled link.',
  Logic_Issue_EmptySection: 'A section has no questions and will be skipped.',
  Logic_Issue_CloseBeforeOpen:
    'The close date is on or before the open date, so the form will never accept responses.',
  Logic_Issue_ApprovalNotifyWithoutApproval: 'Approval emails are switched on but the approval step is off.',
  // branching operators
  Logic_Op_equals: 'is',
  Logic_Op_notEquals: 'is not',
  Logic_Op_contains: 'contains',
  Logic_Op_notContains: 'does not contain',
  Logic_Op_notEmpty: 'is answered',
  Logic_Op_empty: 'is not answered',
  Logic_Op_greaterThan: 'is greater than',
  Logic_Op_greaterOrEqual: 'is at least',
  Logic_Op_lessThan: 'is less than',
  Logic_Op_lessOrEqual: 'is at most',
  Logic_Op_between: 'is between',
  Logic_Op_before: 'is before',
  Logic_Op_after: 'is after'
};

/** Replace {token} placeholders. Unknown tokens are left visible to make a missing param obvious. */
export const formatMessage = (template: string, params?: { [token: string]: string | number }): string =>
  String(template).replace(/\{(\w+)\}/g, (whole: string, name: string) =>
    params && params[name] !== undefined ? String(params[name]) : whole
  );

/** Fill in defaults for settings added after a form was saved, and keep dependent flags consistent. */
export const normalizeSettings = (settings: Partial<IFormSettings> | undefined): IFormSettings => {
  const merged: IFormSettings = { ...DEFAULT_FORM_SETTINGS, ...(settings || {}) } as IFormSettings;
  merged.allowEdit = merged.allowEdit === true;
  merged.enableApproval = merged.enableApproval === true;
  // approval emails only make sense when an approval step exists
  merged.approvalNotify = merged.enableApproval && merged.approvalNotify === true;
  if (!merged.dashboard) {
    merged.dashboard = { ...DEFAULT_DASHBOARD_SETTINGS };
  }
  return merged;
};

/** Lightweight unique id (no external dependency). */
export const newId = (): string =>
  'f' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

export const createEmptyFormDefinition = (): IFormDefinition => ({
  schemaVersion: CURRENT_SCHEMA_VERSION,
  sections: [
    {
      id: newId(),
      // single-section forms show no section chrome — the title stays empty
      // until the owner adds a second section
      title: '',
      description: '',
      fields: []
    }
  ],
  settings: { ...DEFAULT_FORM_SETTINGS, dashboard: { ...DEFAULT_DASHBOARD_SETTINGS } }
});

// ----- field palette metadata -----

export type FieldCategory = 'Popular' | 'Text' | 'Choice' | 'Scales' | 'Date' | 'People' | 'Advanced' | 'Layout';

export interface IFieldTypeMeta {
  type: FieldType;
  label: string;
  icon: string;
  description: string;
  category: FieldCategory;
  /** seeded onto a newly created field of this type */
  defaults?: Partial<IFormField>;
}

export const FIELD_TYPE_META: IFieldTypeMeta[] = [
  // Text
  {
    type: FieldType.Text,
    label: 'Short answer',
    icon: 'TextField',
    description: 'A single line of text',
    category: 'Text'
  },
  {
    type: FieldType.MultilineText,
    label: 'Long answer',
    icon: 'AlignLeft',
    description: 'Several lines of plain text',
    category: 'Text',
    defaults: { rows: 4 }
  },
  {
    type: FieldType.RichText,
    label: 'Formatted text',
    icon: 'FontColorA',
    description: 'Bold, lists and links',
    category: 'Text',
    defaults: { rows: 6 }
  },
  // Choice
  {
    type: FieldType.Choice,
    label: 'Choice',
    icon: 'RadioBtnOn',
    description: 'Pick from a list of options',
    category: 'Choice',
    defaults: { choices: ['Option 1', 'Option 2'] }
  },
  {
    type: FieldType.ImageChoice,
    label: 'Image choice',
    icon: 'PhotoCollection',
    description: 'Pick from picture tiles',
    category: 'Choice',
    defaults: { imageChoices: [{ label: 'Option 1', imageUrl: '' }, { label: 'Option 2', imageUrl: '' }] }
  },
  {
    type: FieldType.Lookup,
    label: 'Lookup from a list',
    icon: 'DocumentSearch',
    description: 'Options read live from another SharePoint list',
    category: 'Choice'
  },
  {
    type: FieldType.YesNo,
    label: 'Yes / No',
    icon: 'ToggleRight',
    description: 'A simple on/off toggle',
    category: 'Choice'
  },
  {
    type: FieldType.Consent,
    label: 'Consent',
    icon: 'CheckboxComposite',
    description: 'A checkbox the respondent must tick to continue',
    category: 'Choice',
    defaults: { required: true, consentText: 'I agree to the terms above.' }
  },
  {
    type: FieldType.Ranking,
    label: 'Ranking',
    icon: 'NumberedList',
    description: 'Drag options into order of preference',
    category: 'Choice',
    defaults: { choices: ['Option 1', 'Option 2', 'Option 3'] }
  },
  // Scales
  {
    type: FieldType.Rating,
    label: 'Rating',
    icon: 'FavoriteStar',
    description: 'Stars, hearts or thumbs',
    category: 'Scales',
    defaults: { maxRating: 5, ratingIcon: 'star' }
  },
  {
    type: FieldType.Scale,
    label: 'Opinion scale',
    icon: 'NumberSymbol',
    description: 'A numbered scale with labelled ends',
    category: 'Scales',
    defaults: { min: 1, max: 5, step: 1, scaleAnalytics: 'average' }
  },
  {
    type: FieldType.Slider,
    label: 'Slider',
    icon: 'Slider',
    description: 'A value on a sliding scale',
    category: 'Scales',
    defaults: { min: 0, max: 10, step: 1 }
  },
  {
    type: FieldType.Likert,
    label: 'Likert grid',
    icon: 'GridViewMedium',
    description: 'Rate several statements on one shared scale',
    category: 'Scales',
    defaults: {
      likertRows: ['Statement 1', 'Statement 2'],
      likertColumns: ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree']
    }
  },
  // Numeric
  {
    type: FieldType.Number,
    label: 'Number',
    icon: 'NumberField',
    description: 'Plain, currency or percentage',
    category: 'Advanced',
    defaults: { numberFormat: 'plain' }
  },
  {
    type: FieldType.Calculated,
    label: 'Calculated',
    icon: 'Calculator',
    description: 'A read-only total worked out from other answers',
    category: 'Advanced',
    defaults: { readOnly: true, numberFormat: 'plain' }
  },
  // Date
  {
    type: FieldType.Date,
    label: 'Date',
    icon: 'Calendar',
    description: 'A calendar date, optionally with a time',
    category: 'Date'
  },
  {
    type: FieldType.Time,
    label: 'Time',
    icon: 'Clock',
    description: 'A time of day on its own',
    category: 'Date',
    defaults: { timeStepMinutes: 15 }
  },
  // People & contact
  {
    type: FieldType.Person,
    label: 'Person',
    icon: 'Contact',
    description: 'Pick someone from the organization',
    category: 'People'
  },
  {
    type: FieldType.Email,
    label: 'Email',
    icon: 'Mail',
    description: 'An email address, validated',
    category: 'People'
  },
  {
    type: FieldType.Phone,
    label: 'Phone',
    icon: 'Phone',
    description: 'A phone number, validated',
    category: 'People'
  },
  {
    type: FieldType.Address,
    label: 'Address',
    icon: 'Home',
    description: 'Street, city, postcode and country',
    category: 'People'
  },
  {
    type: FieldType.Hyperlink,
    label: 'Link',
    icon: 'Link',
    description: 'A web address',
    category: 'Advanced'
  },
  // Rich input
  {
    type: FieldType.FileUpload,
    label: 'File upload',
    icon: 'CloudUpload',
    description: 'Attach one or more files to the response',
    category: 'Advanced',
    defaults: { maxFiles: 3, maxFileSizeMb: 10 }
  },
  {
    type: FieldType.Signature,
    label: 'Signature',
    icon: 'InkingTool',
    description: 'Sign with a mouse, pen or finger',
    category: 'Advanced'
  },
  // Layout
  {
    type: FieldType.Content,
    label: 'Text block',
    icon: 'TextParagraph',
    description: 'Instructions or an image — collects no answer',
    category: 'Layout',
    defaults: { contentStyle: 'info', contentHtml: 'Add your instructions here.' }
  }
];

/**
 * Palette presets — a preset creates a field of an underlying type with
 * particular defaults, so the picker can offer familiar names (NPS, Likert
 * satisfaction) without multiplying the type system.
 */
export interface IFieldPreset {
  key: string;
  label: string;
  icon: string;
  description: string;
  type: FieldType;
  defaults: Partial<IFormField>;
}

export const FIELD_PRESETS: IFieldPreset[] = [
  {
    key: 'nps',
    label: 'NPS (0–10)',
    icon: 'SpeedHigh',
    description: 'How likely are you to recommend…',
    type: FieldType.Scale,
    defaults: {
      min: 0,
      max: 10,
      step: 1,
      lowLabel: 'Not at all likely',
      highLabel: 'Extremely likely',
      scaleAnalytics: 'nps'
    }
  },
  {
    key: 'satisfaction',
    label: 'Satisfaction (1–5)',
    icon: 'Emoji2',
    description: 'Very dissatisfied to very satisfied',
    type: FieldType.Scale,
    defaults: {
      min: 1,
      max: 5,
      step: 1,
      lowLabel: 'Very dissatisfied',
      highLabel: 'Very satisfied',
      scaleAnalytics: 'average'
    }
  },
  {
    key: 'currency',
    label: 'Currency',
    icon: 'Money',
    description: 'A money amount',
    type: FieldType.Number,
    defaults: { numberFormat: 'currency', currencySymbol: '$', decimalPlaces: 2 }
  },
  {
    key: 'percent',
    label: 'Percentage',
    icon: 'CalculatorPercentage',
    description: 'A percentage value',
    type: FieldType.Number,
    defaults: { numberFormat: 'percent', min: 0, max: 100 }
  }
];

const buildDefaultMessages = (): IMessageBag => {
  const bag: IMessageBag = { ...STATIC_MESSAGES };
  FIELD_TYPE_META.forEach((meta) => {
    bag['Logic_Type_' + meta.type] = meta.label;
    bag['Logic_TypeDesc_' + meta.type] = meta.description;
    bag['Logic_Category_' + meta.category] = meta.category;
  });
  FIELD_PRESETS.forEach((preset) => {
    bag['Logic_Preset_' + preset.key] = preset.label;
    bag['Logic_PresetDesc_' + preset.key] = preset.description;
  });
  return bag;
};

/** Every Logic_ key with its English text; pass a translated bag with the same keys to override. */
export const DEFAULT_MESSAGES: IMessageBag = buildDefaultMessages();

/** Look up a message (bag first, English default second) and fill its tokens. */
export const msg = (
  key: string,
  params?: { [token: string]: string | number },
  messages?: IMessageBag
): string => {
  const template = (messages && messages[key]) || DEFAULT_MESSAGES[key] || key;
  return formatMessage(template, params);
};

/** Localized display label for a field type. */
export const fieldTypeLabel = (type: FieldType, messages?: IMessageBag): string =>
  msg('Logic_Type_' + type, undefined, messages);

/** Localized one-line description for a field type. */
export const fieldTypeDescription = (type: FieldType, messages?: IMessageBag): string =>
  msg('Logic_TypeDesc_' + type, undefined, messages);

/** Localized palette category name. */
export const fieldCategoryLabel = (category: FieldCategory, messages?: IMessageBag): string =>
  msg('Logic_Category_' + category, undefined, messages);

/** Localized preset label / description. */
export const fieldPresetLabel = (preset: IFieldPreset, messages?: IMessageBag): string =>
  msg('Logic_Preset_' + preset.key, undefined, messages);
export const fieldPresetDescription = (preset: IFieldPreset, messages?: IMessageBag): string =>
  msg('Logic_PresetDesc_' + preset.key, undefined, messages);

const FALLBACK_META: IFieldTypeMeta = {
  type: FieldType.Text,
  label: 'Question',
  icon: 'FieldEmpty',
  description: '',
  category: 'Text'
};

export const getFieldTypeMeta = (type: FieldType): IFieldTypeMeta => {
  const found = FIELD_TYPE_META.filter((m) => m.type === type)[0];
  return found || { ...FALLBACK_META, type, label: String(type) };
};

/** Defaults seeded when a field is created or switched to a new type. */
export const defaultsForType = (type: FieldType): Partial<IFormField> => {
  const meta = FIELD_TYPE_META.filter((m) => m.type === type)[0];
  return meta && meta.defaults ? JSON.parse(JSON.stringify(meta.defaults)) : {};
};

/**
 * Every per-type configuration key. Switching a field's type clears these
 * before seeding the new type's defaults, so a Choice turned into a Rating
 * doesn't drag its options along.
 */
export const TYPE_SPECIFIC_KEYS: (keyof IFormField)[] = [
  'choices',
  'imageChoices',
  'allowMultiple',
  'choiceDisplay',
  'allowOther',
  'otherLabel',
  'minSelections',
  'maxSelections',
  'shuffleOptions',
  'lookupListId',
  'lookupColumn',
  'lookupFilter',
  'likertRows',
  'likertColumns',
  'min',
  'max',
  'step',
  'decimalPlaces',
  'numberFormat',
  'currencySymbol',
  'unitSuffix',
  'formula',
  'lowLabel',
  'highLabel',
  'scaleAnalytics',
  'maxRating',
  'ratingIcon',
  'allowHalfRating',
  'includeTime',
  'timeStepMinutes',
  'maxLength',
  'rows',
  'pattern',
  'patternMessage',
  'allowMultiplePeople',
  'allowGroups',
  'maxFiles',
  'maxFileSizeMb',
  'allowedExtensions',
  'consentText',
  'contentHtml',
  'contentStyle',
  'contentImageUrl'
];

// ----- SharePoint shapes -----

/** A SharePoint list available as a form target. */
export interface IListInfo {
  id: string;
  title: string;
  itemCount?: number;
  defaultViewUrl?: string;
}

/** A column available as a lookup source. */
export interface IListColumnInfo {
  internalName: string;
  title: string;
  typeAsString: string;
}

/** A person suggestion returned from the people picker search. */
export interface IPersonInfo {
  loginName: string;
  displayName: string;
  email?: string;
  /** true when the principal is a group rather than a user */
  isGroup?: boolean;
}

export type ResponseStatus = 'Draft' | 'Complete';

/** One submitted list item, shaped for the responses views. */
export interface IResponseItem {
  id: number;
  created: Date;
  modified: Date;
  createdBy: string;
  createdByEmail?: string;
  status: ResponseStatus;
  /** how long the respondent took, in seconds, when recorded */
  durationSeconds?: number;
  attachmentCount?: number;
  /** raw values keyed by field internal name */
  values: Record<string, unknown>;
}

/** A page of responses plus the information needed to keep loading. */
export interface IResponsePage {
  items: IResponseItem[];
  /** total matching items reported by SharePoint, when known */
  total?: number;
  hasMore: boolean;
}

export interface IFormValues {
  [fieldId: string]: unknown;
}

/** Value shape used for hyperlink fields inside form state. */
export interface IHyperlinkValue {
  url: string;
  description?: string;
}

// ----- migration -----

/** v1 field types that no longer exist, mapped onto their v2 replacements. */
const LEGACY_TYPE_PATCHES: {
  [legacy: string]: { type: FieldType; patch: Partial<IFormField> };
} = {
  Currency: { type: FieldType.Number, patch: { numberFormat: 'currency' } },
  Percent: { type: FieldType.Number, patch: { numberFormat: 'percent' } },
  DateTime: { type: FieldType.Date, patch: { includeTime: true } },
  Location: { type: FieldType.Text, patch: {} },
  Nps: {
    type: FieldType.Scale,
    patch: { min: 0, max: 10, step: 1, scaleAnalytics: 'nps' }
  }
};

interface ILegacyRule {
  fieldId?: string;
  operator?: ConditionOperator;
  value?: string;
}

const normalizeRetired = (definition: IFormDefinition): void => {
  if (Array.isArray(definition.retiredColumns)) {
    definition.retiredColumns = definition.retiredColumns.filter((n) => typeof n === 'string' && n.length > 0);
  } else {
    delete definition.retiredColumns;
  }
};

/**
 * Bring a persisted definition up to CURRENT_SCHEMA_VERSION.
 *
 * Migration must never change a provisioned field's internalName or its
 * underlying SharePoint type — every mapping here is chosen so the CAML that
 * spFieldXml produces for the migrated field matches the column that already
 * exists (Currency stays Currency, DateTime stays DateTime, Location was
 * always Text).
 */
export const migrateDefinition = (input: IFormDefinition): IFormDefinition => {
  const definition: IFormDefinition = JSON.parse(JSON.stringify(input));
  const from = typeof definition.schemaVersion === 'number' ? definition.schemaVersion : 1;

  if (from >= CURRENT_SCHEMA_VERSION) {
    // still normalize settings so older builds that predate a setting get it
    definition.settings = normalizeSettings(definition.settings);
    normalizeRetired(definition);
    return definition;
  }

  definition.sections = (definition.sections || []).map((section) => {
    const fields = (section.fields || []).map((raw) => {
      const field: IFormField = { ...raw };
      const legacy = LEGACY_TYPE_PATCHES[String(field.type)];
      if (legacy) {
        field.type = legacy.type;
        Object.keys(legacy.patch).forEach((key) => {
          const value = (legacy.patch as Record<string, unknown>)[key];
          const target = field as unknown as Record<string, unknown>;
          if (target[key] === undefined) {
            target[key] = value;
          }
        });
      }

      // legacy NPS end labels lived under their own keys
      const anyField = field as unknown as Record<string, unknown>;
      if (anyField.npsLowLabel !== undefined) {
        field.lowLabel = field.lowLabel || String(anyField.npsLowLabel);
        delete anyField.npsLowLabel;
      }
      if (anyField.npsHighLabel !== undefined) {
        field.highLabel = field.highLabel || String(anyField.npsHighLabel);
        delete anyField.npsHighLabel;
      }

      // plain Number fields gain an explicit format
      if (field.type === FieldType.Number && !field.numberFormat) {
        field.numberFormat = 'plain';
      }

      // single-rule branching becomes a one-condition group
      const rule = field.visibleWhen as unknown as ILegacyRule | IConditionGroup | undefined;
      if (rule && (rule as ILegacyRule).fieldId !== undefined) {
        const legacyRule = rule as ILegacyRule;
        field.visibleWhen = legacyRule.fieldId
          ? {
              match: 'all',
              conditions: [
                {
                  fieldId: legacyRule.fieldId,
                  operator: legacyRule.operator || 'equals',
                  value: legacyRule.value
                }
              ]
            }
          : undefined;
        if (!field.visibleWhen) {
          delete field.visibleWhen;
        }
      }

      return field;
    });
    return { ...section, fields };
  });

  definition.settings = normalizeSettings(definition.settings);
  normalizeRetired(definition);
  definition.schemaVersion = CURRENT_SCHEMA_VERSION;
  return definition;
};
