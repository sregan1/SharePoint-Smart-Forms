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
  suite,
  test
} from './harness';

import {
  CURRENT_SCHEMA_VERSION,
  DEFAULT_FORM_SETTINGS,
  DEFAULT_MESSAGES,
  FieldType,
  IFormDefinition,
  IFormField,
  IResponseItem,
  fieldTypeLabel,
  migrateDefinition,
  msg,
  newId,
  normalizeSettings
} from '../src/webparts/smartForms/models/index';
import { TEMPLATE_MESSAGES, FORM_TEMPLATES, templateName } from '../src/webparts/smartForms/models/templates';
import {
  applyCalculatedFields,
  buildCsv,
  compilePattern,
  computeVisibility,
  conditionIssues,
  evaluateConditionGroup,
  formatSharePointValue,
  formAvailability,
  generateInternalName,
  isFieldVisible,
  normalizeFromSharePoint,
  operatorLabels,
  operatorsForField,
  parseLocalDate,
  parseTimeToMinutes,
  pruneInvalidConditions,
  reservedInternalNames,
  retireDeletedField,
  retireField,
  splitMultiValue,
  validateDefinition,
  validateField,
  validateFieldResult
} from '../src/webparts/smartForms/utils/formUtils';
import { evaluateFormula, plainDecimal } from '../src/webparts/smartForms/utils/formula';
import { isSafeUrl, sanitizeHtml } from '../src/webparts/smartForms/utils/sanitizeHtml';
import {
  answerKeys,
  computeKpis,
  distributionFor,
  foldTail,
  isFieldAnswered,
  matchesRowFilter,
  segmentBy,
  timeline,
  timelineDetailed
} from '../src/webparts/smartForms/utils/analytics';

const field = (overrides: Partial<IFormField> & { type: FieldType }): IFormField => ({
  id: overrides.id || newId(),
  internalName: overrides.internalName || 'SFTest',
  title: overrides.title || 'Test',
  provisioned: true,
  ...overrides
});

const definitionOf = (fields: IFormField[]): IFormDefinition => ({
  schemaVersion: CURRENT_SCHEMA_VERSION,
  sections: [{ id: 's1', title: '', description: '', fields: fields }],
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

suite('retired column names', () => {
  test('a retired name is never reused', () => {
    assertEqual(generateInternalName('Name', [], ['SFName']), 'SFName2');
    assertEqual(generateInternalName('Name', ['SFName2'], ['sfname']), 'SFName3', 'case-insensitive');
  });

  test('old two-argument calls still work', () => {
    assertEqual(generateInternalName('Email', ['SFEmail']), 'SFEmail2');
  });

  test('retireField records once and reservedInternalNames includes it', () => {
    const def = definitionOf([field({ type: FieldType.Text, internalName: 'SFLive' })]);
    retireField(def, 'SFGone');
    retireField(def, 'sfgone');
    assertDeepEqual(def.retiredColumns, ['SFGone']);
    assertEqual(reservedInternalNames(def).indexOf('SFGone') !== -1, true);
    assertEqual(generateInternalName('Gone', reservedInternalNames(def)), 'SFGone2');
  });

  test('retireDeletedField ignores never-published fields', () => {
    const def = definitionOf([]);
    retireDeletedField(def, field({ type: FieldType.Text, internalName: 'SFDraft', provisioned: false }));
    assertUndefined(def.retiredColumns);
    retireDeletedField(def, field({ type: FieldType.Text, internalName: 'SFLive', provisioned: true }));
    assertDeepEqual(def.retiredColumns, ['SFLive']);
  });

  test('retiredColumns survive migration', () => {
    const def = definitionOf([]);
    def.retiredColumns = ['SFOld'];
    assertDeepEqual(migrateDefinition(def).retiredColumns, ['SFOld']);
  });
});

suite('hidden questions in logic', () => {
  const build = (): { def: IFormDefinition; a: IFormField; b: IFormField; c: IFormField } => {
    const c = field({ id: 'c', title: 'C', type: FieldType.Text });
    const a = field({
      id: 'a',
      title: 'A',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'c', operator: 'equals', value: 'y' }] }
    });
    const b = field({
      id: 'b',
      title: 'B',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'a', operator: 'equals', value: 'x' }] }
    });
    // deliberately out of dependency order: B, A, C
    return { def: definitionOf([b, a, c]), a, b, c };
  };

  test('a hidden driver counts as empty, even with a stale answer', () => {
    const { def, b } = build();
    assertFalse(isFieldVisible(b, def, { a: 'x', c: 'n' }), 'A is hidden so its stale x must not show B');
    assertTrue(isFieldVisible(b, def, { a: 'x', c: 'y' }));
  });

  test('hidden driver satisfies an "is not answered" rule', () => {
    const { def } = build();
    const d = field({
      id: 'd',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'a', operator: 'empty' }] }
    });
    def.sections[0].fields.push(d);
    assertTrue(isFieldVisible(d, def, { a: 'stale', c: 'n' }));
  });

  test('computeVisibility resolves the whole chain in one pass', () => {
    const { def } = build();
    assertDeepEqual(computeVisibility(def, { a: 'x', c: 'y' }), { b: true, a: true, c: true });
    assertDeepEqual(computeVisibility(def, { a: 'x', c: 'n' }), { b: false, a: false, c: true });
  });

  test('circular rules terminate', () => {
    const p = field({
      id: 'p',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'q', operator: 'notEmpty' }] }
    });
    const q = field({
      id: 'q',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'p', operator: 'notEmpty' }] }
    });
    const def = definitionOf([p, q]);
    computeVisibility(def, { p: '1', q: '1' });
    assertTrue(true);
  });

  test('calculated fields ignore hidden referenced questions', () => {
    const gate = field({ id: 'g', title: 'Gate', type: FieldType.Text });
    const hidden = field({
      id: 'h',
      title: 'Extra',
      type: FieldType.Number,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'g', operator: 'equals', value: 'show' }] }
    });
    const base = field({ id: 'n', title: 'Base', type: FieldType.Number });
    const total = field({ id: 't', title: 'Total', type: FieldType.Calculated, formula: '{Base} + {Extra}' });
    const def = definitionOf([gate, hidden, base, total]);
    assertEqual(applyCalculatedFields(def, { g: 'no', h: 50, n: 10 })['t'], 10);
    assertEqual(applyCalculatedFields(def, { g: 'show', h: 50, n: 10 })['t'], 60);
  });

  test('chained calculations resolve regardless of order', () => {
    const base = field({ id: 'n', title: 'Base', type: FieldType.Number });
    const c1 = field({ id: 'c1', title: 'C1', type: FieldType.Calculated, formula: '{C2} + 1' });
    const c2 = field({ id: 'c2', title: 'C2', type: FieldType.Calculated, formula: '{C3} + 1' });
    const c3 = field({ id: 'c3', title: 'C3', type: FieldType.Calculated, formula: '{C4} + 1' });
    const c4 = field({ id: 'c4', title: 'C4', type: FieldType.Calculated, formula: '{Base} * 2' });
    const out = applyCalculatedFields(definitionOf([c1, c2, c3, c4, base]), { n: 5 });
    assertEqual(out['c4'], 10);
    assertEqual(out['c1'], 13);
  });

  test('circular calculations stop', () => {
    const c1 = field({ id: 'c1', title: 'C1', type: FieldType.Calculated, formula: '{C2} + 1' });
    const c2 = field({ id: 'c2', title: 'C2', type: FieldType.Calculated, formula: '{C1} + 1' });
    const out = applyCalculatedFields(definitionOf([c1, c2]), {});
    assertDefined(out);
  });
});

suite('date and time conditions', () => {
  const dateField = field({ id: 'd', title: 'Due', type: FieldType.Date });
  const timeField = field({ id: 'tm', title: 'Start', type: FieldType.Time });
  const def = definitionOf([dateField, timeField]);
  const rule = (fieldId: string, operator: string, value?: string, value2?: string) => ({
    match: 'all' as const,
    conditions: [{ fieldId, operator: operator as 'equals', value, value2 }]
  });

  test('parseLocalDate reads a bare date as a LOCAL date', () => {
    const d = parseLocalDate('2026-03-05') as Date;
    assertEqual(d.getFullYear(), 2026);
    assertEqual(d.getMonth(), 2);
    assertEqual(d.getDate(), 5);
    assertEqual(d.getHours(), 0);
    assertUndefined(parseLocalDate('2026-02-31'));
    assertUndefined(parseLocalDate(''));
    assertUndefined(parseLocalDate('not a date'));
    assertDefined(parseLocalDate('today'));
  });

  test('parseTimeToMinutes handles HH:mm and rejects nonsense', () => {
    assertEqual(parseTimeToMinutes('09:30'), 570);
    assertEqual(parseTimeToMinutes('9:05'), 545);
    assertEqual(parseTimeToMinutes('23:59'), 1439);
    assertEqual(parseTimeToMinutes('9:05 PM'), 21 * 60 + 5);
    assertEqual(parseTimeToMinutes('12:00 AM'), 0);
    assertUndefined(parseTimeToMinutes('24:00'));
    assertUndefined(parseTimeToMinutes('10:75'));
    assertUndefined(parseTimeToMinutes('soon'));
  });

  test('equals on a date compares the calendar day, not the instant', () => {
    const lateEvening = new Date(2026, 2, 5, 23, 30);
    assertTrue(evaluateConditionGroup(rule('d', 'equals', '2026-03-05'), def, { d: lateEvening }));
    assertFalse(evaluateConditionGroup(rule('d', 'equals', '2026-03-06'), def, { d: lateEvening }));
    assertTrue(evaluateConditionGroup(rule('d', 'notEquals', '2026-03-06'), def, { d: lateEvening }));
  });

  test('before / after are by day and exclusive of the same day', () => {
    const noon = new Date(2026, 2, 5, 12, 0);
    assertFalse(evaluateConditionGroup(rule('d', 'after', '2026-03-05'), def, { d: noon }));
    assertFalse(evaluateConditionGroup(rule('d', 'before', '2026-03-05'), def, { d: noon }));
    assertTrue(evaluateConditionGroup(rule('d', 'after', '2026-03-04'), def, { d: noon }));
    assertTrue(evaluateConditionGroup(rule('d', 'before', '2026-03-06'), def, { d: noon }));
  });

  test('between is inclusive of both boundary days', () => {
    const first = new Date(2026, 2, 1, 0, 1);
    const last = new Date(2026, 2, 10, 23, 59);
    const r = rule('d', 'between', '2026-03-01', '2026-03-10');
    assertTrue(evaluateConditionGroup(r, def, { d: first }));
    assertTrue(evaluateConditionGroup(r, def, { d: last }));
    assertFalse(evaluateConditionGroup(r, def, { d: new Date(2026, 2, 11, 0, 0) }));
  });

  test('today works as an operand', () => {
    assertTrue(evaluateConditionGroup(rule('d', 'equals', 'today'), def, { d: new Date() }));
  });

  test('time conditions use HH:mm operands', () => {
    const nineThirty = new Date(2000, 0, 1, 9, 30);
    assertTrue(evaluateConditionGroup(rule('tm', 'after', '09:00'), def, { tm: nineThirty }));
    assertFalse(evaluateConditionGroup(rule('tm', 'before', '09:00'), def, { tm: nineThirty }));
    assertTrue(evaluateConditionGroup(rule('tm', 'equals', '09:30'), def, { tm: nineThirty }));
    assertTrue(evaluateConditionGroup(rule('tm', 'between', '09:00', '10:00'), def, { tm: nineThirty }));
    assertFalse(evaluateConditionGroup(rule('tm', 'between', '10:00', '11:00'), def, { tm: nineThirty }));
  });

  test('date and time drivers offer equals / notEquals', () => {
    assertTrue(operatorsForField(dateField).indexOf('equals') !== -1);
    assertTrue(operatorsForField(timeField).indexOf('notEquals') !== -1);
  });
});

suite('formula: exponents and grouping', () => {
  test('tiny and huge substituted values are written as plain decimals', () => {
    assertEqual(plainDecimal(1e-7), '0.0000001');
    assertEqual(plainDecimal(1.5e-7), '0.00000015');
    assertEqual(plainDecimal(1e21), '1000000000000000000000');
    assertEqual(plainDecimal(-2.5e-8), '-0.000000025');
    assertEqual(plainDecimal(12.5), '12.5');
    const tiny = evaluateFormula('{x} * 10000000', () => 1e-7) as number;
    assertClose(tiny, 1, 1e-9);
    assertClose(evaluateFormula('{x} + 1', () => 1e-7) as number, 1.0000001, 1e-12);
    assertClose(evaluateFormula('{x} / 1000000000000000000000', () => 1e21) as number, 1, 1e-9);
  });

  test('a negative tiny value keeps its sign without exponent text', () => {
    assertClose(evaluateFormula('10 + {x}', () => -1e-7) as number, 10 - 1e-7, 1e-12);
  });

  test('the tokenizer accepts exponent numerals', () => {
    assertEqual(evaluateFormula('1e3 + 1', () => 0), 1001);
    assertEqual(evaluateFormula('2.5e-1 * 4', () => 0), 1);
    assertEqual(evaluateFormula('1E2', () => 0), 100);
    assertEqual(evaluateFormula('2e+2', () => 0), 200);
  });

  test('a bare group holding a comma is rejected', () => {
    assertUndefined(evaluateFormula('max((1,2))', () => 0));
    assertEqual(evaluateFormula('max(1,2)', () => 0), 2);
    assertEqual(evaluateFormula('max(1,(2))', () => 0), 2);
    assertEqual(evaluateFormula('sum((1+2),3)', () => 0), 6);
  });
});

suite('sanitizer urls and tag case', () => {
  test('protocol-relative and backslash URLs are rejected', () => {
    assertFalse(isSafeUrl('//evil.example/x'));
    assertFalse(isSafeUrl('  //evil.example'));
    assertFalse(isSafeUrl('/\\evil.example'));
    assertFalse(isSafeUrl('/ /evil.example'));
  });

  test('single-slash paths, fragments, https and mailto are allowed', () => {
    assertTrue(isSafeUrl('/sites/team/page.aspx'));
    assertTrue(isSafeUrl('#section'));
    assertTrue(isSafeUrl('https://contoso.com'));
    assertTrue(isSafeUrl('mailto:a@b.co'));
    assertTrue(isSafeUrl('tel:+15551234'));
  });

  test('script-ish schemes and control-character tricks stay rejected', () => {
    assertFalse(isSafeUrl('javascript:alert(1)'));
    assertFalse(isSafeUrl('java\nscript:alert(1)'));
    assertFalse(isSafeUrl('JaVaScRiPt:alert(1)'));
    assertFalse(isSafeUrl('data:text/html,<b>x</b>'));
    assertFalse(isSafeUrl('vbscript:msgbox(1)'));
    assertFalse(isSafeUrl(''));
  });

  test('svg / math (lowercase tagName in real DOMs) are dropped with their subtree', () => {
    interface IFakeNode {
      nodeType: number;
      tagName?: string;
      textContent?: string;
      childNodes: IFakeNode[];
      attributes: { name: string }[];
      parentNode: IFakeNode | null;
      firstChild: IFakeNode | null;
      removeChild: (c: IFakeNode) => void;
      insertBefore: (n: IFakeNode, ref: IFakeNode) => void;
      removeAttribute: (n: string) => void;
      getAttribute: (n: string) => string | null;
      setAttribute: (n: string, v: string) => void;
      innerHTML?: string;
    }
    const make = (tag: string | undefined, text?: string): IFakeNode => {
      const node: IFakeNode = {
        nodeType: tag ? 1 : 3,
        tagName: tag,
        textContent: text,
        childNodes: [],
        attributes: [],
        parentNode: null,
        firstChild: null,
        removeChild: (c) => {
          node.childNodes = node.childNodes.filter((x) => x !== c);
          node.firstChild = node.childNodes[0] || null;
        },
        insertBefore: (n, ref) => {
          const idx = node.childNodes.indexOf(ref);
          node.childNodes.splice(idx, 0, n);
          n.parentNode = node;
          node.firstChild = node.childNodes[0] || null;
        },
        removeAttribute: () => undefined,
        getAttribute: () => null,
        setAttribute: () => undefined
      };
      return node;
    };
    const add = (parent: IFakeNode, child: IFakeNode): IFakeNode => {
      child.parentNode = parent;
      parent.childNodes.push(child);
      parent.firstChild = parent.childNodes[0];
      return child;
    };
    const serialize = (node: IFakeNode): string =>
      node.nodeType === 3
        ? node.textContent || ''
        : (node.tagName && node.tagName !== 'BODY' ? '<' + node.tagName.toLowerCase() + '>' : '') +
          node.childNodes.map(serialize).join('') +
          (node.tagName && node.tagName !== 'BODY' ? '</' + node.tagName.toLowerCase() + '>' : '');

    const body = make('BODY');
    add(body, make('p')).tagName = 'P';
    const svg = add(body, make('svg')); // lowercase, like real SVG elements
    add(svg, make(undefined, 'evil'));
    add(body, make('math'));
    const g = globalThis as unknown as { DOMParser?: unknown };
    const previous = g.DOMParser;
    g.DOMParser = function () {
      return {
        parseFromString: () => {
          const doc = { body: body } as { body: IFakeNode };
          Object.defineProperty(body, 'innerHTML', { get: () => serialize(body), configurable: true });
          return doc;
        }
      };
    };
    try {
      const html = sanitizeHtml('<p></p><svg>evil</svg><math></math>');
      assertNotContains(html, 'svg');
      assertNotContains(html, 'evil');
      assertNotContains(html, 'math');
      assertContains(html, '<p>');
    } finally {
      g.DOMParser = previous;
    }
  });
});

suite('CSV guards, multi-value and locale', () => {
  const text = field({ id: 'x', title: 'Note', type: FieldType.Text, internalName: 'SFNote' });
  const csvFor = (value: unknown): string => {
    const rows = [{ id: 1, created: new Date('2026-07-01T10:00:00Z'), createdBy: 'T', values: { SFNote: value } }];
    return buildCsv([text], rows).split('\r\n')[1];
  };

  test('real numbers stay numeric', () => {
    assertContains(csvFor('-5'), ',-5');
    assertNotContains(csvFor('-5'), "'-5");
    assertNotContains(csvFor('+12.5'), "'");
    assertNotContains(csvFor('-1,234.50'), "'");
  });

  test('non-numeric text starting with a formula character is guarded', () => {
    assertContains(csvFor('=1+1'), "'=1+1");
    assertContains(csvFor('+cmd|x'), "'+cmd|x");
    assertContains(csvFor('-abc'), "'-abc");
    assertContains(csvFor('@SUM(A1)'), "'@SUM(A1)");
    assertContains(csvFor('\tx'), "'\tx");
    assertContains(csvFor('\rx'), "'");
  });

  test('a lone minus or plus is not a number', () => {
    assertContains(csvFor('-'), "'-");
  });

  test('the submitted column follows the requested locale', () => {
    const created = new Date('2026-07-01T10:00:00Z');
    const rows = [{ id: 1, created: created, createdBy: 'T', values: {} }];
    const german = buildCsv([text], rows, { locale: 'de-DE' });
    const american = buildCsv([text], rows, { locale: 'en-US' });
    assertContains(german, created.toLocaleString('de-DE'));
    assertContains(american, created.toLocaleString('en-US'));
  });

  test('CSV headers come from the message bag', () => {
    const csv = buildCsv([text], [], { messages: { Logic_CsvResponseId: 'ID de respuesta' } });
    assertContains(csv, 'ID de respuesta');
  });

  test('multi-value splitting only breaks on "; " so labels with a bare semicolon survive', () => {
    assertDeepEqual(splitMultiValue('Red; Blue'), ['Red', 'Blue']);
    assertDeepEqual(splitMultiValue('R&D;Ops; Sales'), ['R&D;Ops', 'Sales']);
    const multi = field({ type: FieldType.Choice, allowMultiple: true, choices: ['R&D;Ops', 'Sales'] });
    assertDeepEqual(normalizeFromSharePoint(multi, 'R&D;Ops; Sales'), ['R&D;Ops', 'Sales']);
    assertEqual(formatSharePointValue(multi, 'R&D;Ops; Sales'), 'R&D;Ops, Sales');
  });

  test('conditions on a multi-select match selections, not fragments', () => {
    const multi = field({ id: 'm', type: FieldType.Choice, allowMultiple: true, choices: ['R&D;Ops', 'Sales'] });
    const def = definitionOf([multi]);
    const rule = {
      match: 'all' as const,
      conditions: [{ fieldId: 'm', operator: 'equals' as const, value: 'R&D;Ops' }]
    };
    assertTrue(evaluateConditionGroup(rule, def, { m: ['R&D;Ops', 'Sales'] }));
    assertFalse(evaluateConditionGroup(rule, def, { m: ['Sales'] }));
  });
});

suite('analytics filters, segments and grouping', () => {
  test('rating rows filter on the raw number, not the decorated label', () => {
    const rating = field({ type: FieldType.Rating, internalName: 'SFR', maxRating: 5 });
    const items = [response({ SFR: 4 }), response({ SFR: 4 }), response({ SFR: 2 })];
    const four = distributionFor(rating, items).filter((r) => r.filterValue === '4')[0];
    assertEqual(four.label, '4 ★');
    assertEqual(four.count, 2);
    assertEqual(items.filter((i) => matchesRowFilter(rating, i, four)).length, 2);
  });

  test('localized Yes / No labels keep stable filter values', () => {
    const yn = field({ type: FieldType.YesNo, internalName: 'SFY' });
    const items = [response({ SFY: true }), response({ SFY: false })];
    const rows = distributionFor(yn, items, { Logic_Yes: 'Si', Logic_No: 'No!' });
    assertEqual(rows[0].label, 'Si');
    assertEqual(rows[0].filterValue, 'Yes');
    assertEqual(items.filter((i) => matchesRowFilter(yn, i, rows[0])).length, 1);
  });

  test('an unanswered Yes/No is not counted as No', () => {
    const yn = field({ type: FieldType.YesNo, internalName: 'SFY' });
    const items = [response({ SFY: true }), response({ SFY: null }), response({}), response({ SFY: false })];
    const rows = distributionFor(yn, items);
    assertEqual(rows[0].count, 1);
    assertEqual(rows[1].count, 1);
    assertFalse(isFieldAnswered(yn, null));
    assertTrue(isFieldAnswered(yn, false));
  });

  test('completion rate ignores unanswered Yes/No and non-question keys', () => {
    const yn = field({ type: FieldType.YesNo, internalName: 'SFY' });
    const txt = field({ type: FieldType.Text, internalName: 'SFT' });
    const items = [response({ SFY: null, SFT: 'hi', Junk: 'x' })];
    assertEqual(computeKpis(items, 2, new Date(), [yn, txt]).completionRate, 50);
  });

  test('the folded Other slice filters the folded set', () => {
    const choice = field({ type: FieldType.Choice, internalName: 'SFC', choices: ['A', 'B', 'C', 'D'] });
    const items = [
      response({ SFC: 'A' }),
      response({ SFC: 'A' }),
      response({ SFC: 'A' }),
      response({ SFC: 'B' }),
      response({ SFC: 'B' }),
      response({ SFC: 'C' }),
      response({ SFC: 'D' })
    ];
    const folded = foldTail(distributionFor(choice, items), 3);
    const other = folded[folded.length - 1];
    assertEqual(other.label, 'Other (2)');
    assertUndefined(other.filterValue);
    assertDeepEqual(other.filterValues && other.filterValues.slice().sort(), ['C', 'D']);
    assertEqual(items.filter((i) => matchesRowFilter(choice, i, other)).length, 2);
    assertEqual(foldTail(distributionFor(choice, items), 3, { Logic_OtherSlice: 'Otros ({count})' })[2].label, 'Otros (2)');
  });

  test('multi-select text answers split into their own segments', () => {
    const multi = field({ type: FieldType.Choice, internalName: 'SFM', allowMultiple: true, choices: ['A', 'B'] });
    const items = [response({ SFM: 'A; B' }), response({ SFM: 'A' }), response({ SFM: 'B' })];
    const segments = segmentBy(multi, items);
    const a = segments.filter((s) => s.filterValue === 'A')[0];
    assertEqual(a.items.length, 2);
    assertEqual(items.filter((i) => matchesRowFilter(multi, i, a)).length, 2);
    assertEqual(distributionFor(multi, items)[0].count, 2);
  });

  test('the no-answer segment is filterable and localizable', () => {
    const choice = field({ type: FieldType.Choice, internalName: 'SFC', choices: ['A'] });
    const items = [response({ SFC: 'A' }), response({})];
    const segments = segmentBy(choice, items, { Logic_NoAnswer: '(sin respuesta)' });
    const blank = segments.filter((s) => s.filterValue === '')[0];
    assertEqual(blank.label, '(sin respuesta)');
    assertEqual(items.filter((i) => matchesRowFilter(choice, i, blank)).length, 1);
  });

  test('Date answers group by local date, chronologically', () => {
    const date = field({ type: FieldType.Date, internalName: 'SFD' });
    const items = [
      response({ SFD: new Date(2026, 2, 10, 23, 0).toISOString() }),
      response({ SFD: new Date(2026, 2, 2, 1, 0).toISOString() }),
      response({ SFD: new Date(2026, 2, 10, 1, 0).toISOString() })
    ];
    const rows = distributionFor(date, items);
    assertDeepEqual(rows.map((r) => r.filterValue), ['2026-03-02', '2026-03-10']);
    assertEqual(rows[1].count, 2);
    assertEqual(items.filter((i) => matchesRowFilter(date, i, rows[1])).length, 2);
  });

  test('Time answers group by hour', () => {
    const time = field({ type: FieldType.Time, internalName: 'SFTm' });
    const items = [
      response({ SFTm: new Date(2000, 0, 1, 14, 5).toISOString() }),
      response({ SFTm: new Date(2000, 0, 1, 9, 45).toISOString() }),
      response({ SFTm: new Date(2000, 0, 1, 9, 10).toISOString() })
    ];
    const rows = distributionFor(time, items);
    assertDeepEqual(rows.map((r) => r.filterValue), ['09', '14']);
    assertEqual(rows[0].count, 2);
    assertDeepEqual(answerKeys(time, items[0].values['SFTm']), ['14']);
  });

  test('a long timeline coarsens its grain instead of dropping recent data', () => {
    const items = [
      response({}, new Date(2024, 0, 1)),
      response({}, new Date(2026, 6, 1))
    ];
    const result = timelineDetailed(items, 'day');
    assertTrue(result.grain !== 'day');
    assertFalse(result.truncated);
    assertTrue(result.points.length <= 400);
    assertEqual(result.points[result.points.length - 1].value, 1, 'the newest bucket is kept');
    assertEqual(result.points[0].value, 1);
    assertTrue(timeline(items, 'day').length <= 400);
  });

  test('a short timeline keeps the requested grain', () => {
    const items = [response({}, new Date(2026, 6, 1)), response({}, new Date(2026, 6, 5))];
    const result = timelineDetailed(items, 'day');
    assertEqual(result.grain, 'day');
    assertEqual(result.points.length, 5);
  });

  test('an absurd range is truncated to the newest months and flagged', () => {
    const items = [response({}, new Date(1900, 0, 1)), response({}, new Date(2026, 6, 1))];
    const result = timelineDetailed(items, 'month');
    assertTrue(result.truncated);
    assertEqual(result.points.length, 400);
    assertEqual(result.points[399].value, 1);
  });
});

suite('validation patterns', () => {
  test('a pattern must match the whole answer', () => {
    const f = field({ type: FieldType.Text, pattern: '\\d{3}' });
    assertUndefined(validateField(f, '123'));
    assertDefined(validateField(f, 'abc123def'));
    assertDefined(validateField(f, '1234'));
  });

  test('alternation is anchored as a whole', () => {
    const f = field({ type: FieldType.Text, pattern: 'yes|no' });
    assertUndefined(validateField(f, 'yes'));
    assertUndefined(validateField(f, 'no'));
    assertDefined(validateField(f, 'yesterday'));
    assertDefined(validateField(f, 'piano'));
  });

  test('an invalid pattern never throws or blocks', () => {
    const f = field({ type: FieldType.Text, pattern: '([unclosed' });
    assertUndefined(compilePattern('([unclosed'));
    assertUndefined(validateField(f, 'anything'));
    const issues = validateDefinition(definitionOf([f]));
    assertTrue(issues.filter((i) => i.code === 'Logic_Issue_BadPattern').length === 1);
  });

  test('patternMessage overrides the default', () => {
    const f = field({ type: FieldType.Text, pattern: '\\d+', patternMessage: 'Digits only' });
    assertEqual(validateField(f, 'abc'), 'Digits only');
  });
});

suite('branching rule health', () => {
  test('a choice rule whose option was deleted is broken', () => {
    const driver = field({ id: 'dr', type: FieldType.Choice, choices: ['Red', 'Blue'] });
    const target = field({
      id: 'tg',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'dr', operator: 'equals', value: 'Green' }] }
    });
    const def = definitionOf([driver, target]);
    const issues = conditionIssues(def);
    assertEqual(issues.length, 1);
    assertEqual(issues[0].reason, 'value');
    assertEqual(issues[0].severity, 'broken');
    assertEqual(issues[0].fieldId, 'tg');
    assertEqual(validateDefinition(def).filter((i) => i.code === 'Logic_Issue_RuleMismatch').length, 1);
    pruneInvalidConditions(def);
    assertUndefined(target.visibleWhen);
  });

  test('changing the driver type invalidates the operator', () => {
    const driver = field({ id: 'dr', type: FieldType.Number });
    const target = field({
      id: 'tg',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'dr', operator: 'greaterThan', value: '5' }] }
    });
    const def = definitionOf([driver, target]);
    assertEqual(conditionIssues(def).length, 0);
    driver.type = FieldType.Text;
    const issues = conditionIssues(def);
    assertEqual(issues[0].reason, 'operator');
    pruneInvalidConditions(def);
    assertUndefined(target.visibleWhen);
  });

  test('a rule left with a blank value is incomplete, kept by default, dropped on request', () => {
    const driver = field({ id: 'dr', type: FieldType.Choice, choices: ['Red', 'Blue'] });
    const target = field({
      id: 'tg',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'dr', operator: 'equals', value: '' }] }
    });
    const def = definitionOf([driver, target]);
    assertEqual(conditionIssues(def)[0].severity, 'incomplete');
    pruneInvalidConditions(def);
    assertDefined(target.visibleWhen);
    pruneInvalidConditions(def, { includeIncomplete: true });
    assertUndefined(target.visibleWhen);
  });

  test('unparseable date / time operands are flagged', () => {
    const d = field({ id: 'd', type: FieldType.Date });
    const t = field({ id: 't', type: FieldType.Time });
    const a = field({
      id: 'a',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'd', operator: 'after', value: 'someday' }] }
    });
    const b = field({
      id: 'b',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 't', operator: 'after', value: '25:99' }] }
    });
    const c = field({
      id: 'c',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 't', operator: 'after', value: '09:00' }] }
    });
    const issues = conditionIssues(definitionOf([d, t, a, b, c]));
    assertEqual(issues.filter((i) => i.reason === 'operand').length, 2);
  });

  test('yes/no rules must use Yes or No, and deleted drivers are broken', () => {
    const yn = field({ id: 'yn', type: FieldType.YesNo });
    const a = field({
      id: 'a',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'yn', operator: 'equals', value: 'Maybe' }] }
    });
    const b = field({
      id: 'b',
      type: FieldType.Text,
      visibleWhen: { match: 'all', conditions: [{ fieldId: 'gone', operator: 'equals', value: 'x' }] }
    });
    const reasons = conditionIssues(definitionOf([yn, a, b])).map((i) => i.reason);
    assertDeepEqual(reasons, ['value', 'missingDriver']);
  });

  test('section rules are checked too', () => {
    const def = definitionOf([field({ id: 'q', type: FieldType.Text })]);
    def.sections[0].visibleWhen = { match: 'all', conditions: [{ fieldId: 'nope', operator: 'equals', value: '1' }] };
    const issues = conditionIssues(def);
    assertEqual(issues[0].sectionId, 's1');
    assertUndefined(issues[0].fieldId);
    pruneInvalidConditions(def);
    assertUndefined(def.sections[0].visibleWhen);
  });
});

suite('settings, messages and localization hooks', () => {
  test('new settings default off and normalize', () => {
    assertEqual(DEFAULT_FORM_SETTINGS.allowEdit, false);
    assertEqual(DEFAULT_FORM_SETTINGS.enableApproval, false);
    assertEqual(DEFAULT_FORM_SETTINGS.approvalNotify, false);
    const n = normalizeSettings({ approvalNotify: true });
    assertFalse(n.approvalNotify === true, 'notify needs approval');
    assertTrue(normalizeSettings({ enableApproval: true, approvalNotify: true }).approvalNotify === true);
  });

  test('older saved definitions gain the new settings on load', () => {
    const def = definitionOf([]);
    delete (def.settings as { allowEdit?: boolean }).allowEdit;
    const migrated = migrateDefinition(def);
    assertEqual(migrated.settings.allowEdit, false);
    assertEqual(migrated.settings.enableApproval, false);
  });

  test('approval notify without approval is a warning', () => {
    const def = definitionOf([field({ type: FieldType.Text })]);
    def.settings.approvalNotify = true;
    assertEqual(validateDefinition(def).filter((i) => i.code === 'Logic_Issue_ApprovalNotifyWithoutApproval').length, 1);
  });

  test('every default message key is prefixed and has text', () => {
    Object.keys(DEFAULT_MESSAGES).forEach((key) => {
      assertTrue(key.indexOf('Logic_') === 0, key);
      assertTrue(DEFAULT_MESSAGES[key].length > 0, key);
    });
    assertEqual(fieldTypeLabel(FieldType.Text), 'Short answer');
    assertEqual(fieldTypeLabel(FieldType.Text, { Logic_Type_Text: 'Respuesta corta' }), 'Respuesta corta');
  });

  test('validateField uses a translated bag and keeps English defaults', () => {
    const f = field({ type: FieldType.Text, title: 'Name', required: true });
    assertEqual(validateField(f, ''), 'Name is required');
    assertEqual(validateField(f, '', { Logic_Required: '{label} es obligatorio' }), 'Name es obligatorio');
    const untitled = field({ type: FieldType.Text, title: '', required: true });
    assertEqual(validateField(untitled, ''), 'This question is required');
    const result = validateFieldResult(f, '') as { code: string; params: { label: string } };
    assertEqual(result.code, 'Logic_Required');
    assertEqual(result.params.label, 'Name');
  });

  test('plural-aware selection messages', () => {
    const f = field({ type: FieldType.Choice, allowMultiple: true, minSelections: 1, choices: ['a', 'b'] });
    assertEqual(validateField(f, []), undefined);
    const g = field({ type: FieldType.Choice, allowMultiple: true, maxSelections: 1, choices: ['a', 'b'] });
    assertEqual(validateField(g, ['a', 'b']), 'Choose no more than 1 option');
    const h = field({ type: FieldType.Choice, allowMultiple: true, maxSelections: 2, choices: ['a', 'b', 'c'] });
    assertEqual(validateField(h, ['a', 'b', 'c']), 'Choose no more than 2 options');
  });

  test('operator labels, availability and pre-flight text can be translated', () => {
    assertEqual(operatorLabels()[0].text, 'is');
    assertEqual(operatorLabels({ Logic_Op_equals: 'es' })[0].text, 'es');
    assertEqual(msg('Logic_NotYetOpen', { date: 'X' }), 'This form opens on X.');
    const def = definitionOf([field({ type: FieldType.Text })]);
    def.settings.openDate = '2099-01-01T00:00:00Z';
    const a = formAvailability(def, { messages: { Logic_NotYetOpen: 'Abre el {date}' }, locale: 'en-US' });
    assertContains(a.message, 'Abre el ');
    const empty = validateDefinition(definitionOf([]), { Logic_Issue_NoQuestions: 'Agregue una pregunta.' });
    assertEqual(empty[0].message, 'Agregue una pregunta.');
    assertEqual(empty[0].code, 'Logic_Issue_NoQuestions');
  });

  test('template names and descriptions are keyed for translation', () => {
    const feedback = FORM_TEMPLATES.filter((t) => t.key === 'feedback')[0];
    assertEqual(TEMPLATE_MESSAGES['Logic_Template_feedback_Name'], feedback.name);
    assertEqual(templateName(feedback, { Logic_Template_feedback_Name: 'Opiniones' }), 'Opiniones');
    assertEqual(templateName(feedback), feedback.name);
  });
});
