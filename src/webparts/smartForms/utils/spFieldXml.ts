import { FieldType, IFormField, isInputType } from '../models';
import { effectiveChoices } from './formUtils';

export const FIELD_GROUP = 'Smart Forms';

const escapeXml = (value: string): string =>
  (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

/**
 * Locale ids for the currency symbols offered in the designer. SharePoint
 * renders a Currency column using the column's LCID, so a form using "£" has to
 * provision LCID 2057 or the list view will show dollars regardless of what the
 * form displayed.
 */
const CURRENCY_LCID: { symbol: string; lcid: number; label: string }[] = [
  { symbol: '$', lcid: 1033, label: 'US dollar' },
  { symbol: '£', lcid: 2057, label: 'Pound sterling' },
  { symbol: '€', lcid: 1031, label: 'Euro' },
  { symbol: '¥', lcid: 1041, label: 'Japanese yen' },
  { symbol: 'C$', lcid: 4105, label: 'Canadian dollar' },
  { symbol: 'A$', lcid: 3081, label: 'Australian dollar' },
  { symbol: 'CHF', lcid: 2055, label: 'Swiss franc' },
  { symbol: '₹', lcid: 16393, label: 'Indian rupee' },
  { symbol: 'R', lcid: 7177, label: 'South African rand' },
  { symbol: 'kr', lcid: 1053, label: 'Swedish krona' }
];

export const CURRENCY_OPTIONS = CURRENCY_LCID;

export const lcidForCurrencySymbol = (symbol: string | undefined): number => {
  const match = CURRENCY_LCID.filter((c) => c.symbol === (symbol || '$'))[0];
  return match ? match.lcid : 1033;
};

/**
 * The SharePoint field type backing each Smart Forms field type.
 *
 * Several form types deliberately land on plain text or number columns rather
 * than their "natural" SharePoint equivalent:
 *
 *  - Lookup uses Text, not a real SP Lookup column. A Lookup column stores an
 *    item id, which breaks the moment the source item is renamed or deleted, and
 *    it can't be created against a list the respondent may not be able to read.
 *    Storing the resolved display text keeps responses stable and readable.
 *  - Ranking, Likert and Address use Note, holding an ordered/structured value
 *    that Smart Forms parses back out.
 *  - Signature and FileUpload store filenames in text; the payload itself
 *    becomes a list item attachment.
 */
export const spTypeForField = (field: IFormField): string => {
  switch (field.type) {
    case FieldType.MultilineText:
    case FieldType.RichText:
    case FieldType.Ranking:
    case FieldType.Likert:
    case FieldType.Address:
      return 'Note';
    case FieldType.Number:
    case FieldType.Calculated:
      return field.numberFormat === 'currency' ? 'Currency' : 'Number';
    case FieldType.Rating:
    case FieldType.Slider:
    case FieldType.Scale:
      return 'Number';
    case FieldType.Date:
    case FieldType.Time:
      return 'DateTime';
    case FieldType.Choice:
      return field.allowMultiple ? 'MultiChoice' : 'Choice';
    case FieldType.ImageChoice:
    case FieldType.Lookup:
      return field.allowMultiple ? 'Note' : 'Text';
    case FieldType.YesNo:
    case FieldType.Consent:
      return 'Boolean';
    case FieldType.Hyperlink:
      return 'URL';
    case FieldType.Person:
      return field.allowMultiplePeople ? 'UserMulti' : 'User';
    case FieldType.FileUpload:
    case FieldType.Signature:
      return 'Note';
    default:
      return 'Text';
  }
};

/** Number of decimal places to provision, as SharePoint expects it. */
export const decimalsAttribute = (field: IFormField): string => {
  switch (field.type) {
    case FieldType.Rating:
      return field.allowHalfRating ? '1' : '0';
    case FieldType.Slider:
    case FieldType.Scale:
      // a step of 0.5 needs a decimal place to survive the round trip
      return typeof field.step === 'number' && field.step % 1 !== 0 ? '2' : '0';
    default:
      return field.decimalPlaces === undefined ? 'Automatic' : String(field.decimalPlaces);
  }
};

/**
 * Build the CAML schema XML used to create the SharePoint column for a field.
 * Fields are never marked Required in SharePoint — required-ness is enforced
 * by the form so the OOB list UI and other tools are never blocked by it.
 */
export const buildFieldXml = (field: IFormField): string => {
  const spType = spTypeForField(field);
  const common =
    'ID="{' +
    generateDeterministicGuid(field.internalName) +
    '}" ' +
    'Name="' +
    escapeXml(field.internalName) +
    '" ' +
    'StaticName="' +
    escapeXml(field.internalName) +
    '" ' +
    'DisplayName="' +
    escapeXml(field.title) +
    '" ' +
    'Group="' +
    escapeXml(FIELD_GROUP) +
    '" ' +
    'Required="FALSE"';
  const description = field.description
    ? ' Description="' + escapeXml(field.description) + '"'
    : '';

  switch (spType) {
    case 'Note': {
      const rich = field.type === FieldType.RichText;
      const lines = noteLines(field);
      return (
        '<Field Type="Note" ' +
        common +
        description +
        ' NumLines="' +
        lines +
        '" RichText="' +
        (rich ? 'TRUE' : 'FALSE') +
        '"' +
        (rich ? ' RichTextMode="FullHtml"' : '') +
        ' />'
      );
    }

    case 'Number': {
      // Percentage="TRUE" makes SharePoint store 0..1 and render "45%". Smart
      // Forms keeps the human-facing number (45) in the column so exports, Power
      // BI and the list view all agree with what the respondent typed, and adds
      // the % on display instead.
      return (
        '<Field Type="Number" ' +
        common +
        description +
        ' Decimals="' +
        decimalsAttribute(field) +
        '" Percentage="FALSE"' +
        (typeof field.min === 'number' ? ' Min="' + field.min + '"' : '') +
        (typeof field.max === 'number' ? ' Max="' + field.max + '"' : '') +
        ' />'
      );
    }

    case 'Currency':
      return (
        '<Field Type="Currency" ' +
        common +
        description +
        ' LCID="' +
        lcidForCurrencySymbol(field.currencySymbol) +
        '" Decimals="' +
        (field.decimalPlaces === undefined ? 2 : field.decimalPlaces) +
        '" />'
      );

    case 'DateTime':
      return (
        '<Field Type="DateTime" ' +
        common +
        description +
        ' Format="' +
        (field.type === FieldType.Date && !field.includeTime ? 'DateOnly' : 'DateTime') +
        '" />'
      );

    case 'Choice':
    case 'MultiChoice': {
      const choices = effectiveChoices(field)
        .map((c) => '<CHOICE>' + escapeXml(c) + '</CHOICE>')
        .join('');
      const format = field.choiceDisplay === 'buttons' ? 'RadioButtons' : 'Dropdown';
      // FillInChoice lets the column accept a write-in "Other" answer, which it
      // would otherwise reject as outside the allowed set
      const fillIn = field.allowOther ? ' FillInChoice="TRUE"' : '';
      return (
        '<Field Type="' +
        spType +
        '" ' +
        common +
        description +
        ' Format="' +
        format +
        '"' +
        fillIn +
        '><CHOICES>' +
        choices +
        '</CHOICES></Field>'
      );
    }

    case 'Boolean':
      return '<Field Type="Boolean" ' + common + description + '><Default>0</Default></Field>';

    case 'URL':
      return '<Field Type="URL" ' + common + description + ' Format="Hyperlink" />';

    case 'User':
      return (
        '<Field Type="User" ' +
        common +
        description +
        ' UserSelectionMode="' +
        (field.allowGroups ? 'PeopleAndGroups' : 'PeopleOnly') +
        '" UserSelectionScope="0" />'
      );

    case 'UserMulti':
      return (
        '<Field Type="UserMulti" ' +
        common +
        description +
        ' UserSelectionMode="' +
        (field.allowGroups ? 'PeopleAndGroups' : 'PeopleOnly') +
        '" UserSelectionScope="0" Mult="TRUE" />'
      );

    default: {
      const maxLength =
        field.maxLength && field.maxLength > 0 && field.maxLength <= 255 ? field.maxLength : 255;
      return '<Field Type="Text" ' + common + description + ' MaxLength="' + maxLength + '" />';
    }
  }
};

export const noteLines = (field: IFormField): number => {
  switch (field.type) {
    case FieldType.Ranking:
      return 3;
    case FieldType.Likert:
      return 4;
    case FieldType.Address:
      return 3;
    case FieldType.FileUpload:
    case FieldType.Signature:
      return 2;
    default:
      return field.rows || 6;
  }
};

/** Internal name of the list-level status column. */
export const SF_STATUS_INTERNAL_NAME = 'SFStatus';

/** Internal name of the column recording how long a response took. */
export const SF_DURATION_INTERNAL_NAME = 'SFDurationSeconds';

/**
 * CAML schema XML for the response status column.
 *
 * Indexed="TRUE" matters: every responses query filters on this column, and an
 * unindexed filter throws once the list passes the 5,000-item list view
 * threshold. Indexing it keeps the query legal at any list size.
 */
export const buildStatusFieldXml = (): string => {
  const common =
    'ID="{' +
    generateDeterministicGuid(SF_STATUS_INTERNAL_NAME) +
    '}" ' +
    'Name="' +
    SF_STATUS_INTERNAL_NAME +
    '" ' +
    'StaticName="' +
    SF_STATUS_INTERNAL_NAME +
    '" ' +
    'DisplayName="Response status" ' +
    'Group="' +
    escapeXml(FIELD_GROUP) +
    '" ' +
    'Indexed="TRUE" ' +
    'Required="FALSE"';
  return (
    '<Field Type="Choice" ' +
    common +
    ' Format="Dropdown">' +
    '<Default>Complete</Default>' +
    '<CHOICES><CHOICE>Draft</CHOICE><CHOICE>Complete</CHOICE></CHOICES>' +
    '</Field>'
  );
};

/** CAML for the duration column, used by the dashboard's time-to-complete stat. */
export const buildDurationFieldXml = (): string => {
  const common =
    'ID="{' +
    generateDeterministicGuid(SF_DURATION_INTERNAL_NAME) +
    '}" ' +
    'Name="' +
    SF_DURATION_INTERNAL_NAME +
    '" ' +
    'StaticName="' +
    SF_DURATION_INTERNAL_NAME +
    '" ' +
    'DisplayName="Time to complete (seconds)" ' +
    'Group="' +
    escapeXml(FIELD_GROUP) +
    '" ' +
    'Required="FALSE"';
  return '<Field Type="Number" ' + common + ' Decimals="0" />';
};

/** Columns Smart Forms provisions per list, independent of the form's fields. */
export const SYSTEM_COLUMNS: { internalName: string; xml: () => string; addToView: boolean }[] = [
  { internalName: SF_STATUS_INTERNAL_NAME, xml: buildStatusFieldXml, addToView: false },
  { internalName: SF_DURATION_INTERNAL_NAME, xml: buildDurationFieldXml, addToView: false }
];

/**
 * True when an existing column's SharePoint type still matches what the field
 * definition would provision. A mismatch means the owner changed a question's
 * type after publishing, which SharePoint can't apply in place.
 */
export const typeMatchesExisting = (field: IFormField, existingTypeAsString: string): boolean => {
  const wanted = spTypeForField(field);
  if (wanted === existingTypeAsString) {
    return true;
  }
  // SharePoint reports some types under a different name than the CAML Type
  const equivalents: { [caml: string]: string[] } = {
    Note: ['Note'],
    Text: ['Text'],
    Number: ['Number'],
    Currency: ['Currency'],
    DateTime: ['DateTime'],
    Choice: ['Choice'],
    MultiChoice: ['MultiChoice'],
    Boolean: ['Boolean'],
    URL: ['URL'],
    User: ['User'],
    UserMulti: ['UserMulti', 'User']
  };
  const allowed = equivalents[wanted] || [wanted];
  return allowed.indexOf(existingTypeAsString) !== -1;
};

/** Fields that need a column. Layout-only blocks are skipped. */
export const needsColumn = (field: IFormField): boolean => isInputType(field.type);

/**
 * Stable pseudo-GUID derived from the internal name so re-provisioning the
 * same field never generates a conflicting duplicate id.
 *
 * Uniqueness invariant: internal names are unique within a form definition
 * (generateInternalName enforces it) and field GUIDs only have to be unique
 * within a single list, so distinct columns can never collide here. The hash
 * strength therefore doesn't need to be cryptographic — it only needs to be
 * deterministic, so the same field re-published twice reuses its id instead of
 * failing as a duplicate.
 */
const generateDeterministicGuid = (seed: string): string => {
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < seed.length; i++) {
    h1 = ((h1 ^ seed.charCodeAt(i)) * 0x01000193) >>> 0;
    h2 = ((h2 + seed.charCodeAt(i)) * 0x0100019b) >>> 0;
  }
  const hex = (n: number): string => ('00000000' + n.toString(16)).slice(-8);
  const a = hex(h1);
  const b = hex(h2);
  const c = hex((h1 ^ h2) >>> 0);
  return a + '-' + b.slice(0, 4) + '-4' + b.slice(4, 7) + '-a' + c.slice(0, 3) + '-' + c.slice(3, 8) + a.slice(0, 7);
};
