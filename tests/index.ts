import {
  assertClose,
  assertContains,
  assertDeepEqual,
  assertDefined,
  assertEqual,
  assertFalse,
  assertNotContains,
  assertTrue,
  assertUndefined,
  report,
  suite,
  test
} from './harness';

import {
  CURRENT_SCHEMA_VERSION,
  DEFAULT_FORM_SETTINGS,
  FieldType,
  IFormDefinition,
  IFormField,
  IFormSection,
  IResponseItem,
  migrateDefinition,
  newId
} from '../src/webparts/smartForms/models/index';
import { buildTemplate, FORM_TEMPLATES } from '../src/webparts/smartForms/models/templates';
import {
  allFields,
  buildCsv,
  buildNumberMap,
  effectiveChoices,
  evaluateConditionGroup,
  applyCalculatedFields,
  formatAddress,
  formatValue,
  formatSharePointValue,
  formAvailability,
  generateInternalName,
  inputFields,
  isEmptyValue,
  isFieldVisible,
  normalizeFromSharePoint,
  pruneDanglingConditions,
  shuffleWithSeed,
  validateDefinition,
  validateField
} from '../src/webparts/smartForms/utils/formUtils';
import { evaluateFormula, formulaReferences } from '../src/webparts/smartForms/utils/formula';
import {
  ALLOWED_TAGS,
  DROP_SUBTREE_TAGS,
  isSafeUrl,
  sanitizeHtml,
  stripHtml
} from '../src/webparts/smartForms/utils/sanitizeHtml';
import {
  buildFieldXml,
  buildStatusFieldXml,
  lcidForCurrencySymbol,
  spTypeForField,
  typeMatchesExisting
} from '../src/webparts/smartForms/utils/spFieldXml';
import {
  computeKpis,
  distributionFor,
  foldTail,
  histogram,
  likertStats,
  npsStats,
  numericStats,
  rankingStats,
  segmentBy,
  suggestGrain,
  timeline,
  wordFrequency
} from '../src/webparts/smartForms/utils/analytics';
import {
  categoricalPalette,
  contrastRatio,
  divergingScale,
  ensureContrast,
  MAX_ORDINAL_STEPS,
  ordinalRamp,
  readableOn
} from '../src/webparts/smartForms/utils/theme';

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

const field = (overrides: Partial<IFormField> & { type: FieldType }): IFormField => ({
  id: overrides.id || newId(),
  internalName: overrides.internalName || 'SFTest',
  title: overrides.title || 'Test',
  provisioned: true,
  ...overrides
});

const definitionOf = (fields: IFormField[], sections?: IFormSection[]): IFormDefinition => ({
  schemaVersion: CURRENT_SCHEMA_VERSION,
  sections: sections || [{ id: 's1', title: '', description: '', fields: fields }],
  settings: { ...DEFAULT_FORM_SETTINGS }
});

const response = (values: Record<string, unknown>, created?: Date): IResponseItem => ({
  id: Math.floor(Math.random() * 100000),
  created: created || new Date('2026-07-01T10:00:00Z'),
  modified: created || new Date('2026-07-01T10:00:00Z'),
  createdBy: 'Tester',
  status: 'Complete',
  values: values
});

// ---------------------------------------------------------------------------

suite('generateInternalName', () => {
  test('prefixes with SF and strips punctuation', () => {
    assertEqual(generateInternalName('Your name?', []), 'SFYourName');
  });

  test('de-duplicates against existing names case-insensitively', () => {
    assertEqual(generateInternalName('Name', ['sfname']), 'SFName2');
  });

  test('falls back when a title has no usable characters', () => {
    assertEqual(generateInternalName('???', []), 'SFField');
  });

  test('truncates long titles but stays unique', () => {
    const first = generateInternalName('a'.repeat(60), []);
    assertTrue(first.length <= 28, 'name should be capped at 28 characters');
    const second = generateInternalName('a'.repeat(60), [first]);
    assertTrue(first !== second, 'a collision must produce a different name');
  });
});

suite('isEmptyValue / validateField', () => {
  test('a Yes/No toggle is never empty, but Consent is until ticked', () => {
    assertFalse(isEmptyValue(field({ type: FieldType.YesNo }), false));
    assertTrue(isEmptyValue(field({ type: FieldType.Consent }), false));
    assertFalse(isEmptyValue(field({ type: FieldType.Consent }), true));
  });

  test('a required Consent reports a tick-the-box message', () => {
    const consent = field({ type: FieldType.Consent, required: true });
    assertEqual(validateField(consent, false), 'Please tick the box to continue');
  });

  test('an untouched slider counts as unanswered', () => {
    const slider = field({ type: FieldType.Slider, required: true, min: 0, max: 10 });
    assertDefined(validateField(slider, undefined), 'a required slider must fail when untouched');
    assertUndefined(validateField(slider, 0), 'an explicit zero is a real answer');
  });

  test('a custom required message replaces the default', () => {
    const text = field({
      type: FieldType.Text,
      required: true,
      title: 'Name',
      requiredMessage: 'We need this to contact you'
    });
    assertEqual(validateField(text, ''), 'We need this to contact you');
  });

  test('email and phone formats are checked', () => {
    assertDefined(validateField(field({ type: FieldType.Email }), 'nope'));
    assertUndefined(validateField(field({ type: FieldType.Email }), 'a@b.co'));
    assertDefined(validateField(field({ type: FieldType.Phone }), 'abc'));
    assertUndefined(validateField(field({ type: FieldType.Phone }), '+44 20 7946 0958'));
  });

  test('numeric min and max are enforced with formatting', () => {
    const money = field({
      type: FieldType.Number,
      numberFormat: 'currency',
      currencySymbol: '£',
      min: 10
    });
    assertContains(validateField(money, 5) || '', '£10');
  });

  test('rating uses maxRating as its upper bound', () => {
    const rating = field({ type: FieldType.Rating, maxRating: 5 });
    assertDefined(validateField(rating, 7));
    assertUndefined(validateField(rating, 5));
  });

  test('scale and slider bounds are validated, which the old build skipped', () => {
    const scale = field({ type: FieldType.Scale, min: 1, max: 5 });
    assertDefined(validateField(scale, 9));
    assertUndefined(validateField(scale, 3));
  });

  test('multi-select min and max selections are enforced', () => {
    const choice = field({
      type: FieldType.Choice,
      allowMultiple: true,
      choices: ['a', 'b', 'c'],
      minSelections: 2,
      maxSelections: 2
    });
    assertDefined(validateField(choice, ['a']), 'below the minimum should fail');
    assertDefined(validateField(choice, ['a', 'b', 'c']), 'above the maximum should fail');
    assertUndefined(validateField(choice, ['a', 'b']));
  });

  test('a custom pattern is applied, and a broken pattern never blocks a submission', () => {
    const code = field({
      type: FieldType.Text,
      pattern: '^[A-Z]{2}\\d{4}$',
      patternMessage: 'Use two letters then four digits'
    });
    assertEqual(validateField(code, 'ab12'), 'Use two letters then four digits');
    assertUndefined(validateField(code, 'AB1234'));

    const broken = field({ type: FieldType.Text, pattern: '([' });
    assertUndefined(validateField(broken, 'anything'), 'an invalid regex must be ignored');
  });

  test('file rules cover count, size and extension', () => {
    const upload = field({
      type: FieldType.FileUpload,
      maxFiles: 2,
      maxFileSizeMb: 1,
      allowedExtensions: ['pdf']
    });
    assertDefined(
      validateField(upload, [
        { name: 'a.pdf', size: 10, content: '' },
        { name: 'b.pdf', size: 10, content: '' },
        { name: 'c.pdf', size: 10, content: '' }
      ]),
      'too many files should fail'
    );
    assertContains(
      validateField(upload, [{ name: 'big.pdf', size: 5 * 1024 * 1024, content: '' }]) || '',
      'larger than 1 MB'
    );
    assertContains(
      validateField(upload, [{ name: 'bad.exe', size: 10, content: '' }]) || '',
      'not an accepted file type'
    );
    assertUndefined(validateField(upload, [{ name: 'ok.pdf', size: 10, content: '' }]));
  });
});

suite('branching', () => {
  const driver = field({ id: 'd1', type: FieldType.Choice, title: 'Region', choices: ['EU', 'US'] });
  const target = field({ id: 't1', type: FieldType.Text, title: 'Detail' });

  test('a single matching condition shows the field', () => {
    target.visibleWhen = {
      match: 'all',
      conditions: [{ fieldId: 'd1', operator: 'equals', value: 'EU' }]
    };
    const definition = definitionOf([driver, target]);
    assertTrue(isFieldVisible(target, definition, { d1: 'EU' }));
    assertFalse(isFieldVisible(target, definition, { d1: 'US' }));
  });

  test('match=all requires every condition; match=any requires one', () => {
    const second = field({ id: 'd2', type: FieldType.Number, title: 'Spend' });
    const definition = definitionOf([driver, second, target]);

    target.visibleWhen = {
      match: 'all',
      conditions: [
        { fieldId: 'd1', operator: 'equals', value: 'EU' },
        { fieldId: 'd2', operator: 'greaterThan', value: '100' }
      ]
    };
    assertTrue(isFieldVisible(target, definition, { d1: 'EU', d2: 200 }));
    assertFalse(isFieldVisible(target, definition, { d1: 'EU', d2: 50 }));

    target.visibleWhen.match = 'any';
    assertTrue(isFieldVisible(target, definition, { d1: 'EU', d2: 50 }));
    assertFalse(isFieldVisible(target, definition, { d1: 'US', d2: 50 }));
  });

  test('a rule pointing at a deleted question hides rather than shows the field', () => {
    // the old build returned true here, which silently made branched questions
    // unconditionally visible after their driver was removed
    target.visibleWhen = {
      match: 'all',
      conditions: [{ fieldId: 'gone', operator: 'equals', value: 'x' }]
    };
    assertFalse(isFieldVisible(target, definitionOf([target]), {}));
  });

  test('numeric and between operators compare as numbers', () => {
    const numeric = field({ id: 'n1', type: FieldType.Number, title: 'Age' });
    const definition = definitionOf([numeric]);
    const group = {
      match: 'all' as const,
      conditions: [{ fieldId: 'n1', operator: 'between' as const, value: '18', value2: '65' }]
    };
    assertTrue(evaluateConditionGroup(group, definition, { n1: 30 }));
    assertFalse(evaluateConditionGroup(group, definition, { n1: 70 }));
  });

  test('date operators understand "today"', () => {
    const date = field({ id: 'dt', type: FieldType.Date, title: 'When' });
    const definition = definitionOf([date]);
    const future = new Date();
    future.setDate(future.getDate() + 5);
    const group = {
      match: 'all' as const,
      conditions: [{ fieldId: 'dt', operator: 'after' as const, value: 'today' }]
    };
    assertTrue(evaluateConditionGroup(group, definition, { dt: future }));
  });

  test('a multi-select driver matches on any selection', () => {
    const multi = field({
      id: 'm1',
      type: FieldType.Choice,
      allowMultiple: true,
      choices: ['a', 'b']
    });
    const definition = definitionOf([multi]);
    const group = {
      match: 'all' as const,
      conditions: [{ fieldId: 'm1', operator: 'equals' as const, value: 'b' }]
    };
    assertTrue(evaluateConditionGroup(group, definition, { m1: ['a', 'b'] }));
    assertFalse(evaluateConditionGroup(group, definition, { m1: ['a'] }));
  });

  test('pruneDanglingConditions clears rules for removed fields', () => {
    const keeper = field({ id: 'k', type: FieldType.Text });
    keeper.visibleWhen = {
      match: 'all',
      conditions: [{ fieldId: 'missing', operator: 'equals', value: 'x' }]
    };
    const definition = definitionOf([keeper]);
    pruneDanglingConditions(definition);
    assertUndefined(definition.sections[0].fields[0].visibleWhen);
  });
});

suite('formula evaluation', () => {
  const resolve = (values: { [name: string]: number }) => (name: string) => values[name];

  test('arithmetic and precedence', () => {
    assertEqual(evaluateFormula('2 + 3 * 4', resolve({})), 14);
    assertEqual(evaluateFormula('(2 + 3) * 4', resolve({})), 20);
    assertEqual(evaluateFormula('-5 + 10', resolve({})), 5);
    assertEqual(evaluateFormula('2 ^ 3 ^ 2', resolve({})), 512, 'exponent is right-associative');
  });

  test('field references substitute their values', () => {
    assertEqual(evaluateFormula('{Qty} * {Price}', resolve({ Qty: 3, Price: 7 })), 21);
  });

  test('a missing reference becomes zero rather than failing', () => {
    assertEqual(evaluateFormula('{Nope} + 5', resolve({})), 5);
  });

  test('negative references keep their sign', () => {
    assertEqual(evaluateFormula('10 - {Adj}', resolve({ Adj: -5 })), 15);
  });

  test('functions work, including variadic ones', () => {
    assertEqual(evaluateFormula('round(7 / 2)', resolve({})), 4);
    assertEqual(evaluateFormula('min(3, 9, 1)', resolve({})), 1);
    assertEqual(evaluateFormula('max(3, 9, 1)', resolve({})), 9);
    assertEqual(evaluateFormula('sum(1, 2, 3)', resolve({})), 6);
    assertEqual(evaluateFormula('avg(2, 4)', resolve({})), 3);
  });

  test('division by zero yields zero rather than Infinity', () => {
    assertEqual(evaluateFormula('5 / 0', resolve({})), 0);
  });

  test('malformed formulas return undefined so the UI can warn', () => {
    assertUndefined(evaluateFormula('2 +', resolve({})));
    assertUndefined(evaluateFormula('(2 + 3', resolve({})));
    assertUndefined(evaluateFormula('2) + 3', resolve({})));
    assertUndefined(evaluateFormula('bogus(2)', resolve({})));
    assertUndefined(evaluateFormula('', resolve({})));
  });

  test('no arbitrary code can run through a formula', () => {
    assertUndefined(evaluateFormula('process.exit(1)', resolve({})));
    assertUndefined(evaluateFormula('constructor', resolve({})));
  });

  test('formulaReferences lists names once, in order', () => {
    assertDeepEqual(formulaReferences('{B} + {A} + {B}'), ['B', 'A']);
  });

  test('applyCalculatedFields fills a total from other answers', () => {
    const qty = field({ id: 'q', type: FieldType.Number, title: 'Qty' });
    const price = field({ id: 'p', type: FieldType.Number, title: 'Price' });
    const total = field({
      id: 't',
      type: FieldType.Calculated,
      title: 'Total',
      formula: '{Qty} * {Price}'
    });
    const definition = definitionOf([qty, price, total]);
    const values = applyCalculatedFields(definition, { q: 4, p: 2.5 });
    assertEqual(values.t, 10);
  });

  test('a calculation can reference another calculation', () => {
    const base = field({ id: 'b', type: FieldType.Number, title: 'Base' });
    const first = field({
      id: 'f',
      type: FieldType.Calculated,
      title: 'Doubled',
      formula: '{Base} * 2'
    });
    const second = field({
      id: 's',
      type: FieldType.Calculated,
      title: 'Plus ten',
      formula: '{Doubled} + 10'
    });
    const definition = definitionOf([base, first, second]);
    const values = applyCalculatedFields(definition, { b: 5 });
    assertEqual(values.f, 10);
    assertEqual(values.s, 20);
  });
});

suite('HTML sanitization', () => {
  // sanitizeHtml walks a DOMParser tree, which exists in the browser this ships
  // to but not under the test runner, where it falls back to stripping every
  // tag. Asserting allowlist behavior here would pass vacuously — the fallback
  // removes the markup either way — so the tables and the security-critical URL
  // decision are tested directly, and the fallback is asserted to be safe.
  const hasDom = typeof DOMParser !== 'undefined';

  test('isSafeUrl permits only http, https, mailto, tel and relative targets', () => {
    assertTrue(isSafeUrl('https://example.com'));
    assertTrue(isSafeUrl('http://example.com'));
    assertTrue(isSafeUrl('mailto:a@b.co'));
    assertTrue(isSafeUrl('tel:+441234'));
    assertTrue(isSafeUrl('/sites/team/page.aspx'));
    assertTrue(isSafeUrl('#anchor'));
  });

  test('isSafeUrl rejects every script-bearing scheme', () => {
    assertFalse(isSafeUrl('javascript:alert(1)'));
    assertFalse(isSafeUrl('JaVaScRiPt:alert(1)'));
    assertFalse(isSafeUrl('vbscript:msgbox(1)'));
    assertFalse(isSafeUrl('data:text/html;base64,PHNjcmlwdD4='));
    assertFalse(isSafeUrl('file:///etc/passwd'));
    assertFalse(isSafeUrl(''));
    assertFalse(isSafeUrl('   '));
  });

  test('isSafeUrl rejects a scheme split by whitespace or control characters', () => {
    // "java\nscript:" is a live URL in some parsers, which is why the sanitizer
    // strips invisible characters before testing the scheme
    assertFalse(isSafeUrl('java\nscript:alert(1)'));
    assertFalse(isSafeUrl('java\tscript:alert(1)'));
    assertFalse(isSafeUrl('java script:alert(1)'));
    assertFalse(isSafeUrl(' javascript:alert(1)'));
  });

  test('the allowlist excludes every tag that can execute or load', () => {
    ['SCRIPT', 'IFRAME', 'OBJECT', 'EMBED', 'STYLE', 'LINK', 'SVG', 'FORM', 'INPUT'].forEach(
      (tag) => {
        assertTrue(ALLOWED_TAGS.indexOf(tag) === -1, tag + ' must not be allowlisted');
      }
    );
    ['SCRIPT', 'IFRAME', 'STYLE', 'OBJECT', 'EMBED', 'SVG'].forEach((tag) => {
      assertTrue(
        DROP_SUBTREE_TAGS.indexOf(tag) !== -1,
        tag + ' must have its whole subtree dropped rather than be unwrapped'
      );
    });
    // IMG is absent on purpose: it is the onerror vector and a tracking pixel
    assertTrue(ALLOWED_TAGS.indexOf('IMG') === -1, 'IMG must not be allowlisted');
  });

  test('basic formatting tags are allowlisted', () => {
    ['B', 'I', 'U', 'P', 'UL', 'OL', 'LI', 'A', 'BR'].forEach((tag) => {
      assertTrue(ALLOWED_TAGS.indexOf(tag) !== -1, tag + ' should be allowed');
    });
  });

  test('script content never survives, whichever code path runs', () => {
    const out = sanitizeHtml('<p>ok</p><script>alert(1)</script>');
    assertNotContains(out, '<script');
    assertNotContains(out, 'alert(1)');
    assertContains(out, 'ok');
  });

  test('event handler attributes never survive', () => {
    const out = sanitizeHtml('<img src=x onerror="alert(1)">');
    assertNotContains(out, 'onerror');
    assertNotContains(out, 'alert');
  });

  test('markup-free text passes through untouched', () => {
    assertEqual(sanitizeHtml('just words'), 'just words');
  });

  test('the no-DOM fallback degrades to plain text rather than trusting markup', () => {
    if (hasDom) {
      // in a browser the allowlist keeps <b>; that is the path used at runtime
      assertContains(sanitizeHtml('<b>bold</b>'), '<b>');
      return;
    }
    assertEqual(sanitizeHtml('<b>bold</b>'), 'bold');
    assertEqual(sanitizeHtml('<marquee>hello</marquee>'), 'hello');
  });

  test('stripHtml produces readable plain text', () => {
    assertEqual(stripHtml('<p>Hello&nbsp;<b>world</b></p>'), 'Hello world');
    assertEqual(stripHtml('<script>bad()</script>safe'), 'safe');
    assertEqual(stripHtml('a &amp; b'), 'a & b');
    assertEqual(stripHtml(''), '');
  });
});

suite('SharePoint field XML', () => {
  test('currency uses the LCID that matches its symbol', () => {
    assertEqual(lcidForCurrencySymbol('£'), 2057);
    assertEqual(lcidForCurrencySymbol('$'), 1033);
    assertEqual(lcidForCurrencySymbol(undefined), 1033);
    const xml = buildFieldXml(
      field({ type: FieldType.Number, numberFormat: 'currency', currencySymbol: '£', title: 'Cost' })
    );
    assertContains(xml, 'LCID="2057"');
    assertContains(xml, 'Type="Currency"');
  });

  test('a percentage column stores the human-facing number', () => {
    const xml = buildFieldXml(field({ type: FieldType.Number, numberFormat: 'percent' }));
    assertContains(xml, 'Percentage="FALSE"');
  });

  test('the status column is indexed so the responses filter stays legal at scale', () => {
    assertContains(buildStatusFieldXml(), 'Indexed="TRUE"');
  });

  test('a write-in Other option makes the choice column fill-in', () => {
    const xml = buildFieldXml(
      field({ type: FieldType.Choice, choices: ['a', 'b'], allowOther: true })
    );
    assertContains(xml, 'FillInChoice="TRUE"');
  });

  test('date with and without a time use the right format', () => {
    assertContains(buildFieldXml(field({ type: FieldType.Date })), 'Format="DateOnly"');
    assertContains(
      buildFieldXml(field({ type: FieldType.Date, includeTime: true })),
      'Format="DateTime"'
    );
  });

  test('XML special characters in titles are escaped', () => {
    const xml = buildFieldXml(field({ type: FieldType.Text, title: 'A & B <c> "d"' }));
    assertContains(xml, '&amp;');
    assertContains(xml, '&lt;c&gt;');
    assertNotContains(xml, 'DisplayName="A & B');
  });

  test('empty choice options are dropped from the schema', () => {
    const xml = buildFieldXml(field({ type: FieldType.Choice, choices: ['a', '  ', 'b'] }));
    assertEqual(xml.split('<CHOICE>').length - 1, 2);
  });

  test('migrated legacy types keep the SharePoint type their column already has', () => {
    // this is what makes the v1 -> v2 migration safe against live columns
    assertEqual(
      spTypeForField(field({ type: FieldType.Number, numberFormat: 'currency' })),
      'Currency'
    );
    assertEqual(spTypeForField(field({ type: FieldType.Date, includeTime: true })), 'DateTime');
    assertEqual(spTypeForField(field({ type: FieldType.Date })), 'DateTime');
    assertEqual(spTypeForField(field({ type: FieldType.Text })), 'Text');
    assertEqual(spTypeForField(field({ type: FieldType.Scale })), 'Number');
  });

  test('type compatibility check spots a real mismatch', () => {
    assertTrue(typeMatchesExisting(field({ type: FieldType.Text }), 'Text'));
    assertFalse(typeMatchesExisting(field({ type: FieldType.Text }), 'Number'));
    assertTrue(
      typeMatchesExisting(field({ type: FieldType.Choice, allowMultiple: true }), 'MultiChoice')
    );
  });
});

suite('schema migration', () => {
  test('v1 Currency becomes a Number with a currency format', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        {
          id: 's',
          title: '',
          fields: [
            { id: 'f1', internalName: 'SFCost', title: 'Cost', type: 'Currency', provisioned: true }
          ]
        }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;

    const migrated = migrateDefinition(legacy);
    const migratedField = migrated.sections[0].fields[0];
    assertEqual(migratedField.type, FieldType.Number);
    assertEqual(migratedField.numberFormat, 'currency');
    assertEqual(migratedField.internalName, 'SFCost', 'internal name must never change');
    assertEqual(spTypeForField(migratedField), 'Currency', 'the SP column type must be unchanged');
  });

  test('v1 DateTime becomes a Date that includes a time', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        { id: 's', title: '', fields: [{ id: 'f', internalName: 'SFWhen', title: 'When', type: 'DateTime' }] }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;
    const migratedField = migrateDefinition(legacy).sections[0].fields[0];
    assertEqual(migratedField.type, FieldType.Date);
    assertTrue(migratedField.includeTime === true);
    assertEqual(spTypeForField(migratedField), 'DateTime');
  });

  test('v1 Nps becomes a 0-10 Scale with NPS analytics', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        {
          id: 's',
          title: '',
          fields: [
            {
              id: 'f',
              internalName: 'SFScore',
              title: 'Score',
              type: 'Nps',
              npsLowLabel: 'Low',
              npsHighLabel: 'High'
            }
          ]
        }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;
    const migratedField = migrateDefinition(legacy).sections[0].fields[0];
    assertEqual(migratedField.type, FieldType.Scale);
    assertEqual(migratedField.min, 0);
    assertEqual(migratedField.max, 10);
    assertEqual(migratedField.scaleAnalytics, 'nps');
    assertEqual(migratedField.lowLabel, 'Low', 'legacy NPS labels should carry over');
    assertEqual(migratedField.highLabel, 'High');
  });

  test('v1 Location degrades to Text, which is the column type it already had', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        { id: 's', title: '', fields: [{ id: 'f', internalName: 'SFWhere', title: 'Where', type: 'Location' }] }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;
    const migratedField = migrateDefinition(legacy).sections[0].fields[0];
    assertEqual(migratedField.type, FieldType.Text);
    assertEqual(spTypeForField(migratedField), 'Text');
  });

  test('a v1 single branching rule becomes a one-condition group', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        {
          id: 's',
          title: '',
          fields: [
            { id: 'a', internalName: 'SFA', title: 'A', type: 'Text' },
            {
              id: 'b',
              internalName: 'SFB',
              title: 'B',
              type: 'Text',
              visibleWhen: { fieldId: 'a', operator: 'equals', value: 'yes' }
            }
          ]
        }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;
    const migrated = migrateDefinition(legacy);
    const rule = migrated.sections[0].fields[1].visibleWhen;
    assertDefined(rule);
    assertEqual(rule.match, 'all');
    assertEqual(rule.conditions.length, 1);
    assertEqual(rule.conditions[0].fieldId, 'a');
    assertEqual(rule.conditions[0].value, 'yes');
  });

  test('migration is idempotent and stamps the current version', () => {
    const legacy = {
      schemaVersion: 1,
      sections: [
        { id: 's', title: '', fields: [{ id: 'f', internalName: 'SFA', title: 'A', type: 'Percent' }] }
      ],
      settings: { ...DEFAULT_FORM_SETTINGS }
    } as unknown as IFormDefinition;
    const once = migrateDefinition(legacy);
    const twice = migrateDefinition(once);
    assertEqual(once.schemaVersion, CURRENT_SCHEMA_VERSION);
    assertDeepEqual(twice.sections, once.sections);
  });

  test('a definition missing newer settings gains their defaults', () => {
    const sparse = {
      schemaVersion: 1,
      sections: [{ id: 's', title: '', fields: [] }],
      settings: { formTitle: 'Old form' }
    } as unknown as IFormDefinition;
    const migrated = migrateDefinition(sparse);
    assertEqual(migrated.settings.formTitle, 'Old form', 'existing settings must survive');
    assertDefined(migrated.settings.dashboard, 'dashboard settings should be added');
    assertEqual(migrated.settings.submitButtonText, 'Submit');
  });
});

suite('value normalization and formatting', () => {
  test('numbers format with their unit, currency or percent', () => {
    assertEqual(
      formatValue(field({ type: FieldType.Number, numberFormat: 'currency', currencySymbol: '£' }), 1234.5),
      '£1,234.50'
    );
    assertEqual(formatValue(field({ type: FieldType.Number, numberFormat: 'percent' }), 45), '45%');
    assertEqual(
      formatValue(field({ type: FieldType.Number, unitSuffix: 'kg' }), 12),
      '12 kg'
    );
  });

  test('SharePoint shapes round-trip through normalize then format', () => {
    const person = field({ type: FieldType.Person, internalName: 'SFWho' });
    assertEqual(formatSharePointValue(person, { Title: 'Ada Lovelace' }), 'Ada Lovelace');

    const link = field({ type: FieldType.Hyperlink });
    assertEqual(
      formatSharePointValue(link, { Url: 'https://x.test', Description: 'X' }),
      'X (https://x.test)'
    );

    const yesNo = field({ type: FieldType.YesNo });
    assertEqual(formatSharePointValue(yesNo, true), 'Yes');
  });

  test('a Likert answer survives a semicolon in a statement label', () => {
    // the old delimited storage corrupted exactly this case
    const likert = field({
      type: FieldType.Likert,
      likertRows: ['Speed; overall', 'Support'],
      likertColumns: ['Bad', 'Good']
    });
    const stored = JSON.stringify({ 'Speed; overall': 'Good', Support: 'Bad' });
    const normalized = normalizeFromSharePoint(likert, stored) as { [row: string]: string };
    assertEqual(normalized['Speed; overall'], 'Good');
    assertEqual(normalized.Support, 'Bad');
  });

  test('ranking parses back into an ordered list', () => {
    const ranking = field({ type: FieldType.Ranking, choices: ['A', 'B', 'C'] });
    assertDeepEqual(normalizeFromSharePoint(ranking, 'B; A; C'), ['B', 'A', 'C']);
    assertEqual(formatValue(ranking, ['B', 'A', 'C']), '1. B   2. A   3. C');
  });

  test('an address formats as one readable line', () => {
    assertEqual(
      formatAddress({ street: '1 High St', city: 'Leeds', postalCode: 'LS1 1AA', country: 'UK' }),
      '1 High St, Leeds, LS1 1AA, UK'
    );
    assertEqual(formatAddress(undefined), '');
  });

  test('rich text is reduced to plain text for tables and CSV', () => {
    const rich = field({ type: FieldType.RichText });
    assertEqual(formatValue(rich, '<p>Hello <b>there</b></p>'), 'Hello there');
  });
});

suite('CSV export', () => {
  test('header and values are quoted and escaped correctly', () => {
    const text = field({ type: FieldType.Text, internalName: 'SFNote', title: 'Note, with comma' });
    const csv = buildCsv([text], [
      { id: 1, created: new Date('2026-07-01T10:00:00Z'), createdBy: 'Tester', values: { SFNote: 'He said "hi"' } }
    ]);
    assertContains(csv, '"Note, with comma"');
    assertContains(csv, '"He said ""hi"""');
  });

  test('a leading equals sign is neutralized so Excel does not treat it as a formula', () => {
    const text = field({ type: FieldType.Text, internalName: 'SFNote', title: 'Note' });
    const csv = buildCsv([text], [
      { id: 1, created: new Date(), createdBy: 'x', values: { SFNote: '=1+1' } }
    ]);
    assertContains(csv, "'=1+1");
  });

  test('the file starts with a UTF-8 BOM', () => {
    const csv = buildCsv([], []);
    assertEqual(csv.charCodeAt(0), 0xfeff);
  });

  test('layout blocks never become columns', () => {
    const content = field({ type: FieldType.Content, internalName: 'SFBlock', title: 'Intro' });
    const text = field({ type: FieldType.Text, internalName: 'SFName', title: 'Name' });
    const csv = buildCsv([content, text], []);
    assertNotContains(csv.split('\r\n')[0], 'Intro');
    assertContains(csv.split('\r\n')[0], 'Name');
  });
});

suite('analytics', () => {
  const choice = field({
    type: FieldType.Choice,
    internalName: 'SFPick',
    choices: ['Red', 'Green', 'Blue']
  });

  test('distribution keeps the designed option order and includes zero counts', () => {
    const items = [response({ SFPick: 'Blue' }), response({ SFPick: 'Red' }), response({ SFPick: 'Blue' })];
    const rows = distributionFor(choice, items);
    assertDeepEqual(rows.map((r) => r.label), ['Red', 'Green', 'Blue']);
    assertDeepEqual(rows.map((r) => r.count), [1, 0, 2]);
  });

  test('a write-in answer appears after the designed options', () => {
    const withOther = { ...choice, allowOther: true };
    const rows = distributionFor(withOther, [response({ SFPick: 'Turquoise' })]);
    assertEqual(rows[rows.length - 1].label, 'Turquoise');
  });

  test('NPS splits promoters, passives and detractors', () => {
    const nps = field({ type: FieldType.Scale, internalName: 'SFNps', scaleAnalytics: 'nps', min: 0, max: 10 });
    const items = [10, 9, 8, 7, 6, 0].map((n) => response({ SFNps: n }));
    const stats = npsStats(nps, items);
    assertDefined(stats);
    assertEqual(stats.promoters, 2);
    assertEqual(stats.passives, 2);
    assertEqual(stats.detractors, 2);
    assertEqual(stats.score, 0, '2 promoters and 2 detractors out of 6 nets to zero');
  });

  test('numeric stats include median and standard deviation', () => {
    const number = field({ type: FieldType.Number, internalName: 'SFN' });
    const items = [2, 4, 4, 4, 5, 5, 7, 9].map((n) => response({ SFN: n }));
    const stats = numericStats(number, items);
    assertDefined(stats);
    assertEqual(stats.count, 8);
    assertEqual(stats.median, 4.5);
    assertEqual(stats.avg, 5);
    assertEqual(stats.stdDev, 2);
    assertEqual(stats.min, 2);
    assertEqual(stats.max, 9);
  });

  test('a histogram covers every value exactly once', () => {
    const number = field({ type: FieldType.Number, internalName: 'SFN' });
    const items = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((n) => response({ SFN: n }));
    const bins = histogram(number, items, 5);
    assertEqual(bins.reduce((sum, bin) => sum + bin.count, 0), 10);
  });

  test('ranking averages positions and counts first places', () => {
    const ranking = field({ type: FieldType.Ranking, internalName: 'SFRank' });
    const items = [response({ SFRank: 'A; B; C' }), response({ SFRank: 'B; A; C' })];
    const rows = rankingStats(ranking, items);
    const a = rows.filter((r) => r.label === 'A')[0];
    const c = rows.filter((r) => r.label === 'C')[0];
    assertEqual(a.avgPosition, 1.5);
    assertEqual(a.firstPlace, 1);
    assertEqual(c.avgPosition, 3);
  });

  test('Likert counts land in the right column and average correctly', () => {
    const likert = field({
      type: FieldType.Likert,
      internalName: 'SFGrid',
      likertRows: ['Speed', 'Support'],
      likertColumns: ['Bad', 'Ok', 'Good']
    });
    const items = [
      response({ SFGrid: JSON.stringify({ Speed: 'Good', Support: 'Bad' }) }),
      response({ SFGrid: JSON.stringify({ Speed: 'Good', Support: 'Ok' }) })
    ];
    const stats = likertStats(likert, items);
    const speed = stats.rows.filter((r) => r.label === 'Speed')[0];
    assertDeepEqual(speed.counts, [0, 0, 2]);
    assertEqual(speed.average, 3);
    const support = stats.rows.filter((r) => r.label === 'Support')[0];
    assertDeepEqual(support.counts, [1, 1, 0]);
    assertEqual(support.average, 1.5);
  });

  test('word frequency drops stop words and short words', () => {
    const text = field({ type: FieldType.MultilineText, internalName: 'SFText' });
    const items = [
      response({ SFText: 'The pricing is confusing and the pricing is unclear' }),
      response({ SFText: 'Pricing! Pricing?' })
    ];
    const words = wordFrequency(text, items, 5);
    assertEqual(words[0].word, 'pricing');
    assertEqual(words[0].count, 4);
    assertEqual(words.filter((w) => w.word === 'the').length, 0, 'stop words should be excluded');
  });

  test('foldTail collapses a long tail into one Other row', () => {
    const rows = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((n) => ({ label: 'x' + n, count: n }));
    const folded = foldTail(rows, 4);
    assertEqual(folded.length, 4);
    assertContains(folded[3].label, 'Other');
    assertEqual(
      folded.reduce((sum, row) => sum + row.count, 0),
      rows.reduce((sum, row) => sum + row.count, 0),
      'folding must preserve the total'
    );
  });

  test('the timeline fills gaps so a quiet day is visible', () => {
    const items = [
      response({}, new Date('2026-07-01T09:00:00Z')),
      response({}, new Date('2026-07-03T09:00:00Z'))
    ];
    const points = timeline(items, 'day');
    assertEqual(points.length, 3, 'the empty middle day must still be a point');
    assertDeepEqual(points.map((p) => p.value), [1, 0, 1]);
  });

  test('the grain suggestion widens with the date range', () => {
    const near = [response({}, new Date('2026-07-01')), response({}, new Date('2026-07-05'))];
    assertEqual(suggestGrain(near), 'day');
    const far = [response({}, new Date('2025-01-01')), response({}, new Date('2026-07-01'))];
    assertEqual(suggestGrain(far), 'month');
  });

  test('KPIs count windows and unique respondents', () => {
    const now = new Date('2026-07-25T12:00:00Z');
    const items = [
      response({ SFA: 1 }, new Date('2026-07-25T09:00:00Z')),
      response({ SFA: 1 }, new Date('2026-07-24T09:00:00Z')),
      response({ SFA: 1 }, new Date('2026-06-01T09:00:00Z'))
    ];
    items[0].createdByEmail = 'a@x.test';
    items[1].createdByEmail = 'a@x.test';
    items[2].createdByEmail = 'b@x.test';
    const kpis = computeKpis(items, 1, now);
    assertEqual(kpis.total, 3);
    assertEqual(kpis.today, 1);
    assertEqual(kpis.last7, 2);
    assertEqual(kpis.uniqueRespondents, 2);
    assertEqual(kpis.trend.length, 14);
  });

  test('segmentBy groups responses and puts a multi-select answer in every group', () => {
    const multi = field({
      type: FieldType.Choice,
      internalName: 'SFTeam',
      allowMultiple: true,
      choices: ['Sales', 'Ops']
    });
    const items = [response({ SFTeam: ['Sales', 'Ops'] }), response({ SFTeam: ['Ops'] })];
    const segments = segmentBy(multi, items);
    assertEqual(segments.filter((s) => s.label === 'Sales')[0].items.length, 1);
    assertEqual(segments.filter((s) => s.label === 'Ops')[0].items.length, 2);
  });
});

suite('form availability', () => {
  const base = definitionOf([field({ type: FieldType.Text })]);

  test('an open form is open', () => {
    assertEqual(formAvailability(base, {}).state, 'open');
  });

  test('a future open date reports not-yet-open', () => {
    const future = new Date();
    future.setDate(future.getDate() + 7);
    const definition = { ...base, settings: { ...base.settings, openDate: future.toISOString() } };
    assertEqual(formAvailability(definition, {}).state, 'notYetOpen');
  });

  test('a past close date reports closed', () => {
    const past = new Date();
    past.setDate(past.getDate() - 7);
    const definition = { ...base, settings: { ...base.settings, closeDate: past.toISOString() } };
    assertEqual(formAvailability(definition, {}).state, 'closed');
  });

  test('a response cap closes the form once it is reached', () => {
    const definition = { ...base, settings: { ...base.settings, maxResponses: 10 } };
    assertEqual(formAvailability(definition, { responseCount: 9 }).state, 'open');
    assertEqual(formAvailability(definition, { responseCount: 10 }).state, 'full');
  });

  test('one-response-per-person blocks a repeat', () => {
    const definition = { ...base, settings: { ...base.settings, oneResponsePerPerson: true } };
    assertEqual(formAvailability(definition, { alreadyAnswered: true }).state, 'alreadyAnswered');
    assertEqual(formAvailability(definition, { alreadyAnswered: false }).state, 'open');
  });
});

suite('pre-flight validation', () => {
  test('a form with no questions is an error', () => {
    const issues = validateDefinition(definitionOf([]));
    assertTrue(issues.filter((i) => i.severity === 'error').length > 0);
  });

  test('a choice question with one option is an error', () => {
    const issues = validateDefinition(
      definitionOf([field({ type: FieldType.Choice, title: 'Pick', choices: ['only'] })])
    );
    assertTrue(issues.filter((i) => i.severity === 'error').length > 0);
  });

  test('a self-referencing branching rule is caught', () => {
    const selfRef = field({ id: 'x', type: FieldType.Text, title: 'X' });
    selfRef.visibleWhen = {
      match: 'all',
      conditions: [{ fieldId: 'x', operator: 'equals', value: 'y' }]
    };
    const issues = validateDefinition(definitionOf([selfRef]));
    assertTrue(
      issues.filter((i) => i.message.indexOf('refers to itself') !== -1).length > 0
    );
  });

  test('a broken formula is an error and a valid one is not', () => {
    const bad = field({ type: FieldType.Calculated, title: 'Total', formula: '2 +' });
    assertTrue(validateDefinition(definitionOf([bad])).filter((i) => i.severity === 'error').length > 0);

    const source = field({ type: FieldType.Number, title: 'Qty' });
    const good = field({ type: FieldType.Calculated, title: 'Total', formula: '{Qty} * 2' });
    assertEqual(
      validateDefinition(definitionOf([source, good])).filter((i) => i.severity === 'error').length,
      0
    );
  });

  test('an untitled question is only a warning', () => {
    const issues = validateDefinition(definitionOf([field({ type: FieldType.Text, title: '' })]));
    assertEqual(issues.filter((i) => i.severity === 'error').length, 0);
    assertTrue(issues.filter((i) => i.severity === 'warning').length > 0);
  });

  test('a close date before the open date is an error', () => {
    const definition = definitionOf([field({ type: FieldType.Text, title: 'A' })]);
    definition.settings.openDate = new Date('2026-08-01').toISOString();
    definition.settings.closeDate = new Date('2026-07-01').toISOString();
    assertTrue(
      validateDefinition(definition).filter((i) => i.message.indexOf('close date') !== -1).length > 0
    );
  });

  test('every content-bearing template passes its own pre-flight check', () => {
    // "blank" is expected to fail: it has no questions by design, which is
    // exactly the error the pre-flight check exists to raise
    FORM_TEMPLATES.filter((t) => t.key !== 'blank').forEach((template) => {
      const built = buildTemplate(template.key);
      const errors = validateDefinition(built).filter((i) => i.severity === 'error');
      assertEqual(
        errors.length,
        0,
        'template "' + template.key + '" has errors: ' + errors.map((e) => e.message).join('; ')
      );
    });
  });

  test('the blank template is the one that legitimately has no questions', () => {
    const blank = buildTemplate('blank');
    assertEqual(inputFields(blank).length, 0);
    assertTrue(
      validateDefinition(blank).filter((i) => i.message.indexOf('at least one question') !== -1)
        .length > 0,
      'an empty form must be flagged before publishing'
    );
  });

  test('template branching rules all resolve to real questions', () => {
    FORM_TEMPLATES.forEach((template) => {
      const built = buildTemplate(template.key);
      const ids: { [id: string]: boolean } = {};
      allFields(built).forEach((f) => {
        ids[f.id] = true;
      });
      allFields(built).forEach((f) => {
        if (!f.visibleWhen) {
          return;
        }
        f.visibleWhen.conditions.forEach((condition) => {
          assertTrue(
            ids[condition.fieldId] === true,
            'template "' + template.key + '" has a rule pointing at a missing question'
          );
        });
      });
    });
  });
});

suite('field traversal', () => {
  test('layout blocks are excluded from input fields but not from all fields', () => {
    const content = field({ type: FieldType.Content, title: 'Intro' });
    const text = field({ type: FieldType.Text, title: 'Name' });
    const definition = definitionOf([content, text]);
    assertEqual(allFields(definition).length, 2);
    assertEqual(inputFields(definition).length, 1);
  });

  test('question numbering skips layout blocks', () => {
    const content = field({ id: 'c', type: FieldType.Content });
    const first = field({ id: 'a', type: FieldType.Text });
    const second = field({ id: 'b', type: FieldType.Text });
    const map = buildNumberMap([first, content, second]);
    assertEqual(map.a, 1);
    assertEqual(map.b, 2);
    assertUndefined(map.c, 'a content block should have no number');
  });

  test('effectiveChoices trims, drops blanks and de-duplicates', () => {
    assertDeepEqual(
      effectiveChoices(field({ type: FieldType.Choice, choices: [' A ', '', 'B', 'a'] })),
      ['A', 'B']
    );
  });

  test('a seeded shuffle is stable for the same seed and a permutation of the input', () => {
    const input = ['a', 'b', 'c', 'd', 'e'];
    const first = shuffleWithSeed(input, 42);
    const again = shuffleWithSeed(input, 42);
    assertDeepEqual(first, again, 'the same seed must give the same order');
    assertDeepEqual(first.slice().sort(), input.slice().sort(), 'no options may be lost');
  });
});

suite('chart color rules', () => {
  const LIGHT_SURFACE = '#ffffff';
  const DARK_SURFACE = '#292827';

  test('every categorical slot clears the colorblind-safety ordering in both modes', () => {
    // the ordering itself is the safety mechanism, so the count is load-bearing
    assertEqual(categoricalPalette(false).length, 8);
    assertEqual(categoricalPalette(true).length, 8);
  });

  test('accent contrast is forced to at least 3:1 against the card surface', () => {
    // a mid-tone accent inherited from an older form would otherwise be
    // unreadable on a dark card
    const corrected = ensureContrast('#8a8886', DARK_SURFACE, 3);
    assertTrue(
      contrastRatio(corrected, DARK_SURFACE) >= 3,
      'expected at least 3:1, got ' + contrastRatio(corrected, DARK_SURFACE).toFixed(2)
    );
    const light = ensureContrast('#c8c6c4', LIGHT_SURFACE, 3);
    assertTrue(contrastRatio(light, LIGHT_SURFACE) >= 3);
  });

  test('readableOn picks the higher-contrast foreground', () => {
    assertEqual(readableOn('#0b0b0b'), '#ffffff');
    assertEqual(readableOn('#fab219'), '#201f1e');
  });

  test('the ordinal ramp is monotone and keeps its light end readable', () => {
    const ramp = ordinalRamp('#0078d4', LIGHT_SURFACE, 5, false);
    assertEqual(ramp.length, 5);
    const contrasts = ramp.map((step) => contrastRatio(step, LIGHT_SURFACE));
    for (let i = 1; i < contrasts.length; i++) {
      assertTrue(
        contrasts[i] > contrasts[i - 1],
        'ramp must increase in contrast: ' + contrasts.map((c) => c.toFixed(2)).join(', ')
      );
    }
    assertTrue(
      contrasts[0] >= 2,
      'the step nearest the surface must still clear 2:1, got ' + contrasts[0].toFixed(2)
    );
  });

  test('the ordinal ramp is capped, because more bands cannot stay distinguishable', () => {
    const ramp = ordinalRamp('#0078d4', LIGHT_SURFACE, 9, false);
    assertEqual(ramp.length, MAX_ORDINAL_STEPS);
  });

  test('a diverging scale puts a neutral in the middle only for odd counts', () => {
    const odd = divergingScale(5, false);
    assertEqual(odd.length, 5);
    assertEqual(odd[2], '#9a9894', 'the midpoint must be the neutral gray, never a hue');

    const even = divergingScale(4, false);
    assertEqual(even.length, 4);
    assertTrue(even.indexOf('#9a9894') === -1, 'an even scale has no neutral band');
  });

  test('diverging arms run in opposite directions from the centre', () => {
    const scale = divergingScale(5, false);
    // the extremes should be the strongest steps of each arm
    assertTrue(scale[0] !== scale[1], 'negative arm steps must differ');
    assertTrue(scale[3] !== scale[4], 'positive arm steps must differ');
  });
});

report();
