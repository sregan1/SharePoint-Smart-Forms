import {
  CURRENT_SCHEMA_VERSION,
  DEFAULT_DASHBOARD_SETTINGS,
  DEFAULT_FORM_SETTINGS,
  FieldType,
  IFormDefinition,
  IFormField,
  IFormSection,
  newId
} from './index';

/**
 * Starter templates.
 *
 * A blank canvas is the worst first-run experience for a forms product — the
 * owner has to know the field catalogue before they can begin. Each template is
 * a real, complete form that can be published as-is or edited down.
 *
 * Internal names are deliberately left blank: prepareForPublish derives them
 * from the final titles, so a template that renamed its questions would
 * otherwise provision columns under the template's wording rather than the
 * owner's.
 */

export interface IFormTemplate {
  key: string;
  name: string;
  description: string;
  icon: string;
  accentColor: string;
  headerIcon: string;
  build: () => IFormDefinition;
}

type FieldSpec = Partial<IFormField> & { title: string; type: FieldType };

const field = (spec: FieldSpec): IFormField => ({
  id: newId(),
  internalName: '',
  title: spec.title,
  type: spec.type,
  provisioned: false,
  ...spec
});

const section = (title: string, description: string, fields: IFormField[]): IFormSection => ({
  id: newId(),
  title: title,
  description: description,
  fields: fields
});

const definition = (
  overrides: {
    formTitle: string;
    formDescription: string;
    accentColor: string;
    headerIcon: string;
    submitButtonText?: string;
    confirmationTitle?: string;
    confirmationMessage?: string;
    layout?: 'singlePage' | 'wizard';
    oneResponsePerPerson?: boolean;
    allowSaveDraft?: boolean;
  },
  sections: IFormSection[]
): IFormDefinition => ({
  schemaVersion: CURRENT_SCHEMA_VERSION,
  sections: sections,
  settings: {
    ...DEFAULT_FORM_SETTINGS,
    ...overrides,
    dashboard: { ...DEFAULT_DASHBOARD_SETTINGS }
  }
});

export const FORM_TEMPLATES: IFormTemplate[] = [
  {
    key: 'blank',
    name: 'Blank form',
    description: 'Start from nothing',
    icon: 'PageAdd',
    accentColor: '#0078d4',
    headerIcon: 'ClipboardList',
    build: () =>
      definition(
        {
          formTitle: 'Untitled form',
          formDescription: '',
          accentColor: '#0078d4',
          headerIcon: 'ClipboardList'
        },
        [section('', '', [])]
      )
  },

  {
    key: 'feedback',
    name: 'Customer feedback',
    description: 'NPS, ratings and open comments',
    icon: 'Feedback',
    accentColor: '#0078d4',
    headerIcon: 'Feedback',
    build: () =>
      definition(
        {
          formTitle: 'How did we do?',
          formDescription: 'Your feedback takes about a minute and genuinely shapes what we build next.',
          accentColor: '#0078d4',
          headerIcon: 'Feedback',
          confirmationTitle: 'Thanks for the feedback!',
          confirmationMessage: 'We read every response.'
        },
        [
          section('', '', [
            field({
              title: 'How likely are you to recommend us to a friend or colleague?',
              type: FieldType.Scale,
              required: true,
              min: 0,
              max: 10,
              step: 1,
              lowLabel: 'Not at all likely',
              highLabel: 'Extremely likely',
              scaleAnalytics: 'nps'
            }),
            field({
              title: 'Overall, how satisfied are you?',
              type: FieldType.Rating,
              required: true,
              maxRating: 5,
              ratingIcon: 'star'
            }),
            field({
              title: 'Which parts worked well for you?',
              type: FieldType.Choice,
              allowMultiple: true,
              choices: ['Ease of use', 'Speed', 'Support', 'Value for money', 'Features'],
              allowOther: true
            }),
            field({
              title: 'What is the one thing we should improve?',
              type: FieldType.MultilineText,
              rows: 4,
              placeholder: 'Tell us anything — the more specific the better'
            }),
            field({
              title: 'May we contact you about your feedback?',
              type: FieldType.YesNo
            }),
            field({
              title: 'Email address',
              type: FieldType.Email,
              placeholder: 'you@company.com',
              visibleWhen: {
                match: 'all',
                conditions: []
              }
            })
          ])
        ]
      )
  },

  {
    key: 'event',
    name: 'Event registration',
    description: 'Attendee details, sessions and dietary needs',
    icon: 'Calendar',
    accentColor: '#8764b8',
    headerIcon: 'Calendar',
    build: () =>
      definition(
        {
          formTitle: 'Event registration',
          formDescription: 'Reserve your place. Registration closes when we reach capacity.',
          accentColor: '#8764b8',
          headerIcon: 'Calendar',
          submitButtonText: 'Register',
          confirmationTitle: "You're registered!",
          confirmationMessage: 'A confirmation is on its way to your inbox.',
          layout: 'wizard',
          oneResponsePerPerson: true
        },
        [
          section('Your details', 'How we should address you and reach you.', [
            field({ title: 'Full name', type: FieldType.Text, required: true, width: 'half' }),
            field({ title: 'Email address', type: FieldType.Email, required: true, width: 'half' }),
            field({ title: 'Job title', type: FieldType.Text, width: 'half' }),
            field({ title: 'Organization', type: FieldType.Text, width: 'half' }),
            field({ title: 'Phone number', type: FieldType.Phone, width: 'half' })
          ]),
          section('Your visit', 'Help us plan the day.', [
            field({
              title: 'Which day will you attend?',
              type: FieldType.Choice,
              required: true,
              choices: ['Day one', 'Day two', 'Both days']
            }),
            field({
              title: 'Which sessions interest you?',
              type: FieldType.Choice,
              allowMultiple: true,
              choices: ['Keynote', 'Workshops', 'Panel discussion', 'Networking dinner']
            }),
            field({
              title: 'Do you have any dietary requirements?',
              type: FieldType.MultilineText,
              rows: 2,
              placeholder: 'Allergies, preferences, or none'
            }),
            field({
              title: 'Do you need any accessibility arrangements?',
              type: FieldType.MultilineText,
              rows: 2
            })
          ]),
          section('Confirm', '', [
            field({
              title: 'Terms',
              type: FieldType.Consent,
              required: true,
              consentText:
                'I agree to the event terms and understand my details will be used to manage my registration.',
              placeholder: 'I agree to the event terms'
            })
          ])
        ]
      )
  },

  {
    key: 'itrequest',
    name: 'IT service request',
    description: 'Categorized requests with priority and attachments',
    icon: 'Repair',
    accentColor: '#03787c',
    headerIcon: 'Repair',
    build: () =>
      definition(
        {
          formTitle: 'IT service request',
          formDescription: 'Raise a request with the IT team. We aim to respond within one working day.',
          accentColor: '#03787c',
          headerIcon: 'Repair',
          submitButtonText: 'Submit request',
          confirmationTitle: 'Request received',
          confirmationMessage: 'We have logged your request and will be in touch shortly.'
        },
        [
          section('', '', [
            field({
              title: 'What do you need help with?',
              type: FieldType.Choice,
              required: true,
              choices: [
                'Hardware problem',
                'Software or access',
                'Network or connectivity',
                'New equipment request',
                'Something else'
              ]
            }),
            field({
              title: 'How urgent is this?',
              type: FieldType.Choice,
              required: true,
              choices: [
                'Low — no rush',
                'Normal — within a few days',
                'High — blocking my work',
                'Critical — a team is blocked'
              ],
              choiceDisplay: 'buttons'
            }),
            field({
              title: 'Describe the problem',
              type: FieldType.MultilineText,
              required: true,
              rows: 5,
              placeholder: 'What happened, what you expected, and any error messages'
            }),
            field({
              title: 'When did it start?',
              type: FieldType.Date,
              includeTime: true,
              width: 'half'
            }),
            field({
              title: 'Which device is affected?',
              type: FieldType.Text,
              width: 'half',
              placeholder: 'Asset tag or device name'
            }),
            field({
              title: 'Screenshots or logs',
              type: FieldType.FileUpload,
              maxFiles: 5,
              maxFileSizeMb: 10,
              allowedExtensions: ['png', 'jpg', 'jpeg', 'pdf', 'txt', 'log', 'zip']
            }),
            field({
              title: 'Who should we contact if it is not you?',
              type: FieldType.Person
            })
          ])
        ]
      )
  },

  {
    key: 'expense',
    name: 'Expense claim',
    description: 'Line items, receipts and a calculated total',
    icon: 'Money',
    accentColor: '#498205',
    headerIcon: 'Money',
    build: () =>
      definition(
        {
          formTitle: 'Expense claim',
          formDescription: 'Submit expenses with receipts attached. Claims are reviewed weekly.',
          accentColor: '#498205',
          headerIcon: 'Money',
          submitButtonText: 'Submit claim',
          confirmationTitle: 'Claim submitted',
          confirmationMessage: 'Finance will review your claim and confirm by email.',
          allowSaveDraft: true
        },
        [
          section('', '', [
            field({
              title: 'Claim period',
              type: FieldType.Text,
              required: true,
              width: 'half',
              placeholder: 'e.g. March 2026'
            }),
            field({
              title: 'Cost centre',
              type: FieldType.Text,
              width: 'half'
            }),
            field({
              title: 'Expense category',
              type: FieldType.Choice,
              required: true,
              choices: ['Travel', 'Accommodation', 'Meals', 'Equipment', 'Training', 'Other'],
              allowOther: true
            }),
            field({
              title: 'Net amount',
              type: FieldType.Number,
              required: true,
              numberFormat: 'currency',
              currencySymbol: '$',
              decimalPlaces: 2,
              min: 0,
              width: 'half'
            }),
            field({
              title: 'Tax rate',
              type: FieldType.Number,
              numberFormat: 'percent',
              min: 0,
              max: 100,
              defaultValue: '20',
              width: 'half'
            }),
            field({
              title: 'Total claimed',
              type: FieldType.Calculated,
              readOnly: true,
              numberFormat: 'currency',
              currencySymbol: '$',
              decimalPlaces: 2,
              formula: '{Net amount} * (1 + {Tax rate} / 100)',
              description: 'Worked out automatically from the net amount and tax rate'
            }),
            field({
              title: 'Date of expense',
              type: FieldType.Date,
              required: true,
              width: 'half'
            }),
            field({
              title: 'Receipts',
              type: FieldType.FileUpload,
              required: true,
              maxFiles: 10,
              maxFileSizeMb: 10,
              allowedExtensions: ['pdf', 'png', 'jpg', 'jpeg'],
              description: 'Attach a receipt for every line item'
            }),
            field({
              title: 'Approving manager',
              type: FieldType.Person,
              required: true
            }),
            field({
              title: 'Declaration',
              type: FieldType.Consent,
              required: true,
              consentText:
                'I confirm these expenses were incurred on company business and that the receipts attached are genuine.',
              placeholder: 'I confirm this claim is accurate'
            })
          ])
        ]
      )
  },

  {
    key: 'pulse',
    name: 'Employee pulse survey',
    description: 'Likert grid, engagement scale and anonymity note',
    icon: 'Emoji2',
    accentColor: '#881798',
    headerIcon: 'Emoji2',
    build: () =>
      definition(
        {
          formTitle: 'Team pulse check',
          formDescription: 'Five questions, two minutes. Results are reviewed as a team, not individually.',
          accentColor: '#881798',
          headerIcon: 'Emoji2',
          confirmationTitle: 'Thank you',
          confirmationMessage: 'Your input goes straight into this quarter’s team review.',
          oneResponsePerPerson: true
        },
        [
          section('', '', [
            field({
              title: 'Please read first',
              type: FieldType.Content,
              contentStyle: 'info',
              contentHtml:
                'Answers are stored in a SharePoint list your team administrator can see. If you need to raise something confidentially, speak to your manager or HR directly instead.'
            }),
            field({
              title: 'How are you feeling about work right now?',
              type: FieldType.Scale,
              required: true,
              min: 1,
              max: 5,
              step: 1,
              lowLabel: 'Struggling',
              highLabel: 'Thriving',
              scaleAnalytics: 'average'
            }),
            field({
              title: 'How much do you agree with the following?',
              type: FieldType.Likert,
              required: true,
              likertRows: [
                'I have what I need to do my job well',
                'My work is recognized',
                'I can raise concerns openly',
                'My workload is manageable',
                'I can see how my work matters'
              ],
              likertColumns: [
                'Strongly disagree',
                'Disagree',
                'Neutral',
                'Agree',
                'Strongly agree'
              ]
            }),
            field({
              title: 'What would make the biggest difference to your week?',
              type: FieldType.MultilineText,
              rows: 4
            }),
            field({
              title: 'Anything to celebrate?',
              type: FieldType.MultilineText,
              rows: 3,
              placeholder: 'A win, or someone who deserves credit'
            })
          ])
        ]
      )
  },

  {
    key: 'safety',
    name: 'Safety inspection',
    description: 'Checklist, photos and a signature',
    icon: 'Shield',
    accentColor: '#ca5010',
    headerIcon: 'Shield',
    build: () =>
      definition(
        {
          formTitle: 'Site safety inspection',
          formDescription: 'Complete one record per inspection. Photograph anything you flag.',
          accentColor: '#ca5010',
          headerIcon: 'Shield',
          submitButtonText: 'Submit inspection',
          confirmationTitle: 'Inspection recorded',
          confirmationMessage: 'Thank you — flagged items are routed to the site manager.',
          allowSaveDraft: true
        },
        [
          section('Inspection details', '', [
            field({ title: 'Site or area', type: FieldType.Text, required: true, width: 'half' }),
            field({
              title: 'Inspection date and time',
              type: FieldType.Date,
              required: true,
              includeTime: true,
              defaultValue: 'today',
              width: 'half'
            }),
            field({ title: 'Inspector', type: FieldType.Person, required: true })
          ]),
          section('Checklist', 'Mark each item as you check it.', [
            field({
              title: 'Rate the following areas',
              type: FieldType.Likert,
              required: true,
              likertRows: [
                'Walkways clear',
                'Fire exits unobstructed',
                'PPE available and worn',
                'Equipment guards in place',
                'Spill kits stocked',
                'First aid supplies in date'
              ],
              likertColumns: ['Fail', 'Needs attention', 'Pass']
            }),
            field({
              title: 'Were any issues found?',
              type: FieldType.YesNo,
              required: false
            }),
            field({
              title: 'Describe the issues found',
              type: FieldType.MultilineText,
              rows: 4,
              visibleWhen: { match: 'all', conditions: [] }
            }),
            field({
              title: 'Photographs',
              type: FieldType.FileUpload,
              maxFiles: 8,
              maxFileSizeMb: 10,
              allowedExtensions: ['png', 'jpg', 'jpeg']
            })
          ]),
          section('Sign off', '', [
            field({
              title: 'Inspector signature',
              type: FieldType.Signature,
              required: true
            }),
            field({
              title: 'Declaration',
              type: FieldType.Consent,
              required: true,
              consentText: 'I confirm this inspection was carried out in full and the record above is accurate.',
              placeholder: 'I confirm this record is accurate'
            })
          ])
        ]
      )
  },

  {
    key: 'booking',
    name: 'Room or resource booking',
    description: 'Lookup-driven resource picker with a time slot',
    icon: 'Ticket',
    accentColor: '#69797e',
    headerIcon: 'Ticket',
    build: () =>
      definition(
        {
          formTitle: 'Book a room',
          formDescription: 'Request a room or piece of equipment.',
          accentColor: '#69797e',
          headerIcon: 'Ticket',
          submitButtonText: 'Request booking',
          confirmationTitle: 'Booking requested',
          confirmationMessage: 'We will confirm your booking by email.'
        },
        [
          section('', '', [
            field({ title: 'Your name', type: FieldType.Text, required: true, width: 'half' }),
            field({ title: 'Email address', type: FieldType.Email, required: true, width: 'half' }),
            field({
              title: 'What do you need?',
              type: FieldType.Choice,
              required: true,
              choices: ['Meeting room', 'Desk', 'Equipment', 'Vehicle']
            }),
            field({
              title: 'Date needed',
              type: FieldType.Date,
              required: true,
              width: 'half'
            }),
            field({
              title: 'Start time',
              type: FieldType.Time,
              required: true,
              timeStepMinutes: 15,
              width: 'half'
            }),
            field({
              title: 'End time',
              type: FieldType.Time,
              required: true,
              timeStepMinutes: 15,
              width: 'half'
            }),
            field({
              title: 'How many people?',
              type: FieldType.Number,
              min: 1,
              max: 200,
              width: 'half'
            }),
            field({
              title: 'Anything else we should know?',
              type: FieldType.MultilineText,
              rows: 3
            })
          ])
        ]
      )
  }
];

/**
 * Templates seed illustrative branching with an empty condition list, because a
 * condition has to reference a field id that only exists once the template is
 * built. Wire those up now that the ids are real: each empty rule attaches to
 * the nearest preceding Yes/No or contact-permission question.
 */
export const buildTemplate = (key: string): IFormDefinition => {
  const template = FORM_TEMPLATES.filter((t) => t.key === key)[0] || FORM_TEMPLATES[0];
  const built = template.build();

  built.sections.forEach((sectionEntry) => {
    sectionEntry.fields.forEach((entry, index) => {
      if (!entry.visibleWhen || entry.visibleWhen.conditions.length > 0) {
        return;
      }
      // find the closest earlier Yes/No question to hang the rule on
      let driver: IFormField | undefined;
      for (let i = index - 1; i >= 0; i--) {
        if (sectionEntry.fields[i].type === FieldType.YesNo) {
          driver = sectionEntry.fields[i];
          break;
        }
      }
      if (driver) {
        entry.visibleWhen = {
          match: 'all',
          conditions: [{ fieldId: driver.id, operator: 'equals', value: 'Yes' }]
        };
      } else {
        delete entry.visibleWhen;
      }
    });
  });

  return built;
};

/** Question count, for the template card subtitle. */
export const templateQuestionCount = (template: IFormTemplate): number => {
  const built = template.build();
  let count = 0;
  built.sections.forEach((s) =>
    s.fields.forEach((f) => {
      if (f.type !== FieldType.Content) {
        count++;
      }
    })
  );
  return count;
};

/**
 * English template names / descriptions keyed for translation
 * (Logic_Template_<key>_Name and Logic_Template_<key>_Description).
 */
export const TEMPLATE_MESSAGES: { [key: string]: string } = (() => {
  const bag: { [key: string]: string } = {};
  FORM_TEMPLATES.forEach((t) => {
    bag['Logic_Template_' + t.key + '_Name'] = t.name;
    bag['Logic_Template_' + t.key + '_Description'] = t.description;
  });
  return bag;
})();

/** Localized template name; `messages` is the caller's translated bag. */
export const templateName = (template: IFormTemplate, messages?: { [key: string]: string }): string =>
  (messages && messages['Logic_Template_' + template.key + '_Name']) || template.name;

/** Localized template description. */
export const templateDescription = (template: IFormTemplate, messages?: { [key: string]: string }): string =>
  (messages && messages['Logic_Template_' + template.key + '_Description']) || template.description;
