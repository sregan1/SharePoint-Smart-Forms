import { assertContains, assertNotContains, assertEqual, assertTrue, suite, test } from './harness';
import { FieldType, IFormField } from '../src/webparts/smartForms/models/index';
import {
  APPROVAL_COLUMNS,
  buildApprovalStatusFieldXml,
  buildFieldXml,
  SF_APPROVAL_STATUS_INTERNAL_NAME
} from '../src/webparts/smartForms/utils/spFieldXml';

const make = (partial: Partial<IFormField>): IFormField =>
  ({ id: 'f1', internalName: 'Q1', title: 'Q', ...partial } as IFormField);

suite('SharePoint field XML (provisioning details)', () => {
  test('a Yes/No column has no default so unanswered stays null', () => {
    const xml = buildFieldXml(make({ type: FieldType.YesNo }));
    assertContains(xml, 'Type="Boolean"');
    assertNotContains(xml, '<Default>');
  });

  test('a Number column omits Decimals when automatic', () => {
    assertNotContains(buildFieldXml(make({ type: FieldType.Number })), 'Decimals=');
    assertContains(buildFieldXml(make({ type: FieldType.Number, decimalPlaces: 2 })), 'Decimals="2"');
  });

  test('approval columns are Pending-defaulted and uniquely named', () => {
    assertContains(buildApprovalStatusFieldXml(), '<Default>Pending</Default>');
    assertContains(buildApprovalStatusFieldXml(), '<CHOICE>Rejected</CHOICE>');
    assertEqual(APPROVAL_COLUMNS.length, 3);
    assertEqual(APPROVAL_COLUMNS[0].internalName, SF_APPROVAL_STATUS_INTERNAL_NAME);
    assertTrue(APPROVAL_COLUMNS.every((c) => c.xml().indexOf('Name="' + c.internalName + '"') > 0));
  });
});
