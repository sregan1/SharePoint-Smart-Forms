import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import '@pnp/sp/security';
import '@pnp/sp/sputilities';
import '@pnp/sp/attachments';
import '@pnp/sp/batching';
import { PermissionKind } from '@pnp/sp/security';
import { IList } from '@pnp/sp/lists';
import {
  FieldType,
  IAddressValue,
  IFormDefinition,
  IFormField,
  IFormFile,
  IFormValues,
  IHyperlinkValue,
  IListColumnInfo,
  IListInfo,
  ILikertValue,
  IPersonInfo,
  IResponseItem,
  IResponsePage,
  ResponseStatus
} from '../models';
import {
  buildResponseEmailHtml,
  formatAddress,
  inputFields,
  isFieldVisible,
  MULTI_SEPARATOR,
  normalizeFromSharePoint,
  valueAsComparableString
} from '../utils/formUtils';
import {
  buildFieldXml,
  SF_DURATION_INTERNAL_NAME,
  SF_STATUS_INTERNAL_NAME,
  SYSTEM_COLUMNS,
  typeMatchesExisting
} from '../utils/spFieldXml';
import { logWarning } from '../utils/debug';

/** Options value for AddFieldInternalNameHint — honor the Name attribute. */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;

/** Items fetched per request when paging responses. */
const PAGE_SIZE = 500;

/** Hard ceiling on how many responses the client will hold at once. */
const MAX_RESPONSES = 20000;

/** Cap on concurrent column-creation requests during provisioning. */
const PROVISION_CONCURRENCY = 4;

export interface ISubmitOptions {
  durationSeconds?: number;
  /** promote this existing draft item instead of creating a new one */
  replaceItemId?: number;
}

export interface IDraft {
  id: number;
  values: IFormValues;
}

export interface IProvisionResult {
  definition: IFormDefinition;
  /** columns created on this run */
  created: string[];
  /** fields whose existing column has an incompatible type */
  conflicts: { field: IFormField; existingType: string }[];
}

/** Run async work over a list with a bounded number of requests in flight. */
const mapWithConcurrency = async <TIn, TOut>(
  items: TIn[],
  limit: number,
  worker: (item: TIn, index: number) => Promise<TOut>
): Promise<TOut[]> => {
  const results: TOut[] = new Array(items.length);
  let cursor = 0;
  const runners: Promise<void>[] = [];
  const size = Math.max(1, Math.min(limit, items.length));

  for (let r = 0; r < size; r++) {
    runners.push(
      (async (): Promise<void> => {
        for (;;) {
          const index = cursor++;
          if (index >= items.length) {
            return;
          }
          results[index] = await worker(items[index], index);
        }
      })()
    );
  }
  await Promise.all(runners);
  return results;
};

export class SharePointService {
  constructor(private sp: SPFI) {}

  private list(listId: string): IList {
    return this.sp.web.lists.getById(listId);
  }

  /** Cache of lookup option sets, keyed by list + column + filter. */
  private lookupCache: { [key: string]: string[] } = {};

  // -------------------------------------------------------------------------
  // lists & columns
  // -------------------------------------------------------------------------

  /** Visible generic lists that can back a form. */
  public async getAvailableLists(): Promise<IListInfo[]> {
    const lists = await this.sp.web.lists
      .select('Id', 'Title', 'ItemCount')
      .filter('Hidden eq false and BaseTemplate eq 100')
      .orderBy('Title')();
    return lists.map((l: { Id: string; Title: string; ItemCount: number }) => ({
      id: l.Id,
      title: l.Title,
      itemCount: l.ItemCount
    }));
  }

  public async getListInfo(listId: string): Promise<IListInfo | undefined> {
    try {
      const l = await this.list(listId).select('Id', 'Title', 'ItemCount', 'DefaultViewUrl')();
      return { id: l.Id, title: l.Title, itemCount: l.ItemCount, defaultViewUrl: l.DefaultViewUrl };
    } catch (e) {
      if (this.isNotFoundError(e)) {
        return undefined;
      }
      throw e;
    }
  }

  public async createList(title: string): Promise<IListInfo> {
    const result = await this.sp.web.lists.add(
      title,
      'Created by Smart Forms to store form responses',
      100,
      false
    );
    return { id: result.data.Id, title: result.data.Title, itemCount: 0 };
  }

  /** Columns of a list that can supply lookup option text. */
  public async getListColumns(listId: string): Promise<IListColumnInfo[]> {
    const fields = await this.list(listId)
      .fields.select('InternalName', 'Title', 'TypeAsString', 'Hidden', 'ReadOnlyField')
      .filter('Hidden eq false and ReadOnlyField eq false')();
    const usable = ['Text', 'Note', 'Choice', 'MultiChoice', 'Number', 'Currency', 'Counter', 'DateTime'];
    return fields
      .filter((f: { TypeAsString: string }) => usable.indexOf(f.TypeAsString) !== -1)
      .map((f: { InternalName: string; Title: string; TypeAsString: string }) => ({
        internalName: f.InternalName,
        title: f.Title,
        typeAsString: f.TypeAsString
      }));
  }

  /**
   * Distinct option values for a Lookup field, read from the source list.
   * Cached per list+column+filter for the lifetime of the page: a form with the
   * same lookup on several questions would otherwise refetch per control.
   */
  public async getLookupOptions(
    listId: string,
    column: string,
    filter?: string
  ): Promise<string[]> {
    const key = listId + '|' + column + '|' + (filter || '');
    if (this.lookupCache[key]) {
      return this.lookupCache[key];
    }
    let query = this.list(listId).items.select(column).top(2000).orderBy(column, true);
    if (filter && filter.trim().length > 0) {
      query = query.filter(filter.trim());
    }
    const items = await query();
    const seen: { [value: string]: boolean } = {};
    const values: string[] = [];
    items.forEach((item: Record<string, unknown>) => {
      const raw = item[column];
      if (raw === undefined || raw === null || raw === '') {
        return;
      }
      const text = String(raw).trim();
      if (text.length === 0 || seen[text]) {
        return;
      }
      seen[text] = true;
      values.push(text);
    });
    this.lookupCache[key] = values;
    return values;
  }

  public clearLookupCache(): void {
    this.lookupCache = {};
  }

  /** True when the current user can manage the list (design forms, view responses). */
  public async currentUserIsOwner(listId: string): Promise<boolean> {
    try {
      return await this.list(listId).currentUserHasPermissions(PermissionKind.ManageLists);
    } catch {
      return false;
    }
  }

  /** True when the current user can submit a response. */
  public async currentUserCanRespond(listId: string): Promise<boolean> {
    try {
      return await this.list(listId).currentUserHasPermissions(PermissionKind.AddListItems);
    } catch {
      return false;
    }
  }

  // -------------------------------------------------------------------------
  // provisioning
  // -------------------------------------------------------------------------

  /**
   * Create any SharePoint columns the form needs that don't exist yet, keep
   * choice options in sync, relax the Title column, and surface new columns
   * in the list's default view.
   *
   * Column creation runs with bounded concurrency rather than one await at a
   * time — a 30-question form used to serialize 60+ round trips and take the
   * better part of a minute. Adding a field to the default view is folded into
   * the same worker so each column costs one slot, not two passes.
   */
  public async ensureFields(listId: string, definition: IFormDefinition): Promise<IProvisionResult> {
    const list = this.list(listId);
    const existing: { InternalName: string; TypeAsString: string }[] = await list.fields.select(
      'InternalName',
      'TypeAsString'
    )();
    const existingByName: { [lower: string]: string } = {};
    existing.forEach((f) => {
      existingByName[f.InternalName.toLowerCase()] = f.TypeAsString;
    });

    await this.makeTitleOptional(list);

    const updated: IFormDefinition = JSON.parse(JSON.stringify(definition));
    const fields = inputFields(updated);
    const created: string[] = [];
    const conflicts: { field: IFormField; existingType: string }[] = [];

    const toCreate: IFormField[] = [];
    const toSync: IFormField[] = [];

    fields.forEach((field) => {
      const existingType = existingByName[field.internalName.toLowerCase()];
      if (existingType === undefined) {
        toCreate.push(field);
        return;
      }
      if (!typeMatchesExisting(field, existingType)) {
        // SharePoint can't change a column's type in place; the owner has to
        // rename the question so a fresh column is created
        conflicts.push({ field, existingType });
        return;
      }
      toSync.push(field);
    });

    await mapWithConcurrency(toCreate, PROVISION_CONCURRENCY, async (field) => {
      await list.fields.createFieldAsXml({
        SchemaXml: buildFieldXml(field),
        Options: ADD_FIELD_INTERNAL_NAME_HINT
      });
      await this.tryAddToDefaultView(list, field.internalName);
      created.push(field.internalName);
      return undefined;
    });

    await mapWithConcurrency(toSync, PROVISION_CONCURRENCY, async (field) => {
      await this.syncExistingField(list, field);
      return undefined;
    });

    // system columns: status (indexed, used by every responses query) and duration
    await mapWithConcurrency(SYSTEM_COLUMNS, PROVISION_CONCURRENCY, async (column) => {
      if (existingByName[column.internalName.toLowerCase()] !== undefined) {
        return undefined;
      }
      try {
        await list.fields.createFieldAsXml({
          SchemaXml: column.xml(),
          Options: ADD_FIELD_INTERNAL_NAME_HINT
        });
        created.push(column.internalName);
      } catch (error) {
        // a pre-existing column under a different type shouldn't block publishing
        logWarning('creating system column "' + column.internalName + '" (non-fatal)', error);
      }
      return undefined;
    });

    fields.forEach((field) => {
      const conflicted = conflicts.filter((c) => c.field.id === field.id).length > 0;
      field.provisioned = !conflicted;
    });

    return { definition: updated, created, conflicts };
  }

  /** Push title/description/choice changes to a column that already exists. */
  private async syncExistingField(list: IList, field: IFormField): Promise<void> {
    try {
      const spField = list.fields.getByInternalNameOrTitle(field.internalName);
      const props: Record<string, unknown> = {
        Title: field.title,
        Description: field.description || ''
      };
      if (field.type === FieldType.Choice) {
        props.Choices = (field.choices || []).map((c) => (c || '').trim()).filter((c) => c.length > 0);
        props.FillInChoice = field.allowOther === true;
        await spField.update(props, field.allowMultiple ? 'SP.FieldMultiChoice' : 'SP.FieldChoice');
      } else {
        await spField.update(props);
      }
    } catch (error) {
      // a column someone re-typed or locked shouldn't block publishing the form
      logWarning('syncing existing field "' + field.internalName + '" (non-fatal)', error);
    }
  }

  private async makeTitleOptional(list: IList): Promise<void> {
    try {
      await list.fields.getByInternalNameOrTitle('Title').update({ Required: false });
    } catch (error) {
      // best effort — some lists restrict Title edits
      logWarning('making Title optional (non-fatal)', error);
    }
  }

  private async tryAddToDefaultView(list: IList, internalName: string): Promise<void> {
    try {
      await list.defaultView.fields.add(internalName);
    } catch (error) {
      // view may already contain the field or be read-only
      logWarning('adding "' + internalName + '" to the default view (non-fatal)', error);
    }
  }

  // -------------------------------------------------------------------------
  // submitting
  // -------------------------------------------------------------------------

  private async buildPayload(
    definition: IFormDefinition,
    values: IFormValues
  ): Promise<Record<string, unknown>> {
    const payload: Record<string, unknown> = {};
    const fields = inputFields(definition);

    for (const field of fields) {
      if (!isFieldVisible(field, definition, values)) {
        continue;
      }
      const converted = await this.toSharePointValue(field, values[field.id]);
      if (converted !== undefined) {
        payload[this.payloadKey(field)] = converted;
      }
    }

    payload.Title = this.buildItemTitle(definition, values);
    return payload;
  }

  /**
   * Save a submission. Attachments can only be added after the item exists, so
   * the item is created (or a resumed draft promoted) first and files are
   * uploaded second; a failed upload doesn't roll the response back, since a
   * response with a missing attachment beats no response at all.
   */
  public async submitResponse(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    options?: ISubmitOptions
  ): Promise<number> {
    const payload = await this.buildPayload(definition, values);
    payload[SF_STATUS_INTERNAL_NAME] = 'Complete';
    if (options && typeof options.durationSeconds === 'number') {
      payload[SF_DURATION_INTERNAL_NAME] = options.durationSeconds;
    }

    let itemId: number;
    if (options && typeof options.replaceItemId === 'number') {
      await this.list(listId).items.getById(options.replaceItemId).update(payload);
      itemId = options.replaceItemId;
    } else {
      const result = await this.list(listId).items.add(payload);
      itemId = result.data.Id;
    }

    await this.uploadAttachments(listId, itemId, definition, values);
    return itemId;
  }

  /** Save (or update) a partial response the respondent can come back to. */
  public async saveDraft(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    existingId?: number
  ): Promise<number> {
    const payload = await this.buildPayload(definition, values);
    payload[SF_STATUS_INTERNAL_NAME] = 'Draft';
    if (typeof existingId === 'number') {
      await this.list(listId).items.getById(existingId).update(payload);
      return existingId;
    }
    const result = await this.list(listId).items.add(payload);
    return result.data.Id;
  }

  /** The current user's most recent draft for this list, if any. */
  public async loadDraft(listId: string, definition: IFormDefinition): Promise<IDraft | undefined> {
    try {
      const me = await this.sp.web.currentUser.select('Id')();
      const fields = inputFields(definition).filter((f) => f.provisioned !== false);
      const { selects, expands } = this.selectsFor(fields);
      const items = await this.list(listId)
        .items.select(...selects)
        .expand(...expands)
        .filter(SF_STATUS_INTERNAL_NAME + " eq 'Draft' and AuthorId eq " + me.Id)
        .orderBy('Id', false)
        .top(1)();
      if (items.length === 0) {
        return undefined;
      }
      const item = items[0] as Record<string, unknown>;
      const values: IFormValues = {};
      fields.forEach((field) => {
        const normalized = normalizeFromSharePoint(field, item[field.internalName]);
        if (normalized !== undefined) {
          values[field.id] = normalized;
        }
      });
      return { id: item.Id as number, values };
    } catch {
      return undefined;
    }
  }

  /**
   * Number of completed responses, used by the response-cap check.
   *
   * Pages over ids only. There's no filtered-count endpoint in the REST surface
   * PnPjs exposes, and the list's own ItemCount would include drafts, so an
   * id-only sweep is both correct and cheap.
   */
  public async countResponses(listId: string): Promise<number> {
    try {
      let page = await this.list(listId)
        .items.select('Id')
        .filter(SF_STATUS_INTERNAL_NAME + " ne 'Draft'")
        .top(2000)
        .getPaged();
      let count = (page.results || []).length;
      while (page.hasNext && count < MAX_RESPONSES) {
        page = await page.getNext();
        count += (page.results || []).length;
      }
      return count;
    } catch {
      // a failed count must not close an open form
      return 0;
    }
  }

  /** True when the signed-in user already has a completed response. */
  public async currentUserHasResponded(listId: string): Promise<boolean> {
    try {
      const me = await this.sp.web.currentUser.select('Id')();
      const items = await this.list(listId)
        .items.select('Id')
        .filter(SF_STATUS_INTERNAL_NAME + " ne 'Draft' and AuthorId eq " + me.Id)
        .top(1)();
      return items.length > 0;
    } catch {
      return false;
    }
  }

  /**
   * Upload file and signature answers as list item attachments.
   *
   * Filenames are prefixed with the column's internal name so several upload
   * questions on one form can't collide, and so the responses view can tell
   * which attachment belongs to which question.
   */
  private async uploadAttachments(
    listId: string,
    itemId: number,
    definition: IFormDefinition,
    values: IFormValues
  ): Promise<void> {
    const uploads: { name: string; content: ArrayBuffer }[] = [];

    inputFields(definition).forEach((field) => {
      if (!isFieldVisible(field, definition, values)) {
        return;
      }
      if (field.type === FieldType.FileUpload) {
        const files = (values[field.id] || []) as IFormFile[];
        files.forEach((file) => {
          if (!file.content) {
            return;
          }
          uploads.push({
            name: this.attachmentName(field.internalName, file.name),
            content: base64ToArrayBuffer(file.content)
          });
        });
        return;
      }
      if (field.type === FieldType.Signature) {
        const dataUrl = String(values[field.id] || '');
        const comma = dataUrl.indexOf(',');
        if (dataUrl.indexOf('data:image') !== 0 || comma < 0) {
          return;
        }
        uploads.push({
          name: this.attachmentName(field.internalName, 'signature.png'),
          content: base64ToArrayBuffer(dataUrl.slice(comma + 1))
        });
      }
    });

    if (uploads.length === 0) {
      return;
    }

    const item = this.list(listId).items.getById(itemId);
    for (const upload of uploads) {
      try {
        await item.attachmentFiles.add(upload.name, upload.content);
      } catch (error) {
        // a rejected attachment (size, blocked extension, duplicate on a resumed
        // draft) must not fail an otherwise-valid response
        logWarning('attaching "' + upload.name + '" (non-fatal)', error);
      }
    }
  }

  private attachmentName(internalName: string, fileName: string): string {
    const safe = fileName.replace(/[\\/:*?"<>|#%]/g, '_');
    return internalName + '__' + safe;
  }

  // -------------------------------------------------------------------------
  // notifications
  // -------------------------------------------------------------------------

  /**
   * Email the owner-configured recipients (and optionally the respondent)
   * a formatted copy of a just-submitted response. Uses SharePoint's
   * SendEmail utility — no Power Automate or connectors needed. Recipients
   * must be users in the organization; failures are silent by design so a
   * mail hiccup never breaks a submission.
   */
  public async sendResponseNotifications(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    itemId: number
  ): Promise<void> {
    const settings = definition.settings;
    const formTitle = settings.formTitle || 'Smart Forms';
    const recipients = (settings.notifyEmails || '')
      .split(/[;,]/)
      .map((e) => e.trim())
      .filter((e) => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e));

    const sends: Promise<void>[] = [];

    if (recipients.length > 0) {
      sends.push(
        (async (): Promise<void> => {
          const web = await this.sp.web.select('Url')();
          const itemUrl =
            web.Url +
            '/_layouts/15/listform.aspx?PageType=4&ListId=%7B' +
            listId +
            '%7D&ID=' +
            itemId;
          await this.sp.utility.sendEmail({
            To: recipients,
            Subject: 'New response — ' + formTitle,
            Body: buildResponseEmailHtml(definition, values, {
              heading: 'New response: ' + formTitle,
              intro: 'Someone just submitted a response — ' + new Date().toLocaleString() + '.',
              accentColor: settings.accentColor,
              itemUrl
            }),
            AdditionalHeaders: { 'content-type': 'text/html' }
          });
        })()
      );
    }

    if (settings.respondentReceipt === true) {
      sends.push(
        (async (): Promise<void> => {
          const me = await this.sp.web.currentUser.select('Email')();
          if (!me.Email) {
            return;
          }
          await this.sp.utility.sendEmail({
            To: [me.Email],
            Subject: 'Your response — ' + formTitle,
            Body: buildResponseEmailHtml(definition, values, {
              heading: 'Thanks for your response!',
              intro: 'Here is a copy of what you submitted to "' + formTitle + '".',
              accentColor: settings.accentColor
            }),
            AdditionalHeaders: { 'content-type': 'text/html' }
          });
        })()
      );
    }

    // never let notification failures surface to the person submitting
    await Promise.all(sends.map((p) => p.catch(() => undefined)));
  }

  // -------------------------------------------------------------------------
  // value conversion
  // -------------------------------------------------------------------------

  private payloadKey(field: IFormField): string {
    return field.type === FieldType.Person ? field.internalName + 'Id' : field.internalName;
  }

  private async toSharePointValue(field: IFormField, value: unknown): Promise<unknown> {
    if (value === undefined || value === null || value === '') {
      return undefined;
    }
    switch (field.type) {
      case FieldType.Number:
      case FieldType.Calculated:
      case FieldType.Rating:
      case FieldType.Slider:
      case FieldType.Scale:
        return typeof value === 'number' && !isNaN(value) ? value : undefined;

      case FieldType.YesNo:
      case FieldType.Consent:
        return value === true;

      case FieldType.Date: {
        const d = value as Date;
        if (!(d instanceof Date) || isNaN(d.getTime())) {
          return undefined;
        }
        if (field.includeTime) {
          return d.toISOString();
        }
        // store UTC noon so the calendar date is stable across timezones
        return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0)).toISOString();
      }

      case FieldType.Time: {
        const d = value as Date;
        return d instanceof Date && !isNaN(d.getTime()) ? d.toISOString() : undefined;
      }

      case FieldType.Choice: {
        if (field.allowMultiple) {
          const selections = (Array.isArray(value) ? value : [value]).map((v) => String(v));
          return selections.length > 0 ? selections : undefined;
        }
        return String(value);
      }

      case FieldType.ImageChoice:
      case FieldType.Lookup: {
        if (field.allowMultiple) {
          const selections = (Array.isArray(value) ? value : [value]).map((v) => String(v));
          return selections.length > 0 ? selections.join(MULTI_SEPARATOR) : undefined;
        }
        return Array.isArray(value) ? String((value as unknown[])[0]) : String(value);
      }

      case FieldType.Ranking: {
        // persist the ranked order as readable text: "First; Second; Third"
        const ordered = (Array.isArray(value) ? value : [value]).map((v) => String(v));
        return ordered.length > 0 ? ordered.join(MULTI_SEPARATOR) : undefined;
      }

      case FieldType.Likert: {
        const answers = value as ILikertValue;
        const keys = Object.keys(answers || {}).filter((row) => !!answers[row]);
        // JSON keeps row/column pairs unambiguous even when a label contains a
        // semicolon, which a delimited string can't
        return keys.length > 0 ? JSON.stringify(answers) : undefined;
      }

      case FieldType.Address: {
        const formatted = formatAddress(value as IAddressValue);
        return formatted.length > 0 ? formatted : undefined;
      }

      case FieldType.Hyperlink: {
        const link = value as IHyperlinkValue;
        if (!link.url) {
          return undefined;
        }
        return { Url: link.url.trim(), Description: link.description || link.url.trim() };
      }

      case FieldType.Person: {
        const people = value as IPersonInfo[];
        if (!Array.isArray(people) || people.length === 0) {
          return undefined;
        }
        // resolve every principal in parallel rather than one round trip each
        const ids = await Promise.all(
          people.map(async (person) => {
            const ensured = await this.sp.web.ensureUser(person.loginName);
            return ensured.data.Id as number;
          })
        );
        return field.allowMultiplePeople ? ids : ids[0];
      }

      case FieldType.FileUpload: {
        const files = (value || []) as IFormFile[];
        if (!Array.isArray(files) || files.length === 0) {
          return undefined;
        }
        // the column records the filenames; the bytes become attachments
        return files.map((f) => f.name).join(MULTI_SEPARATOR);
      }

      case FieldType.Signature:
        return String(value).indexOf('data:image') === 0 ? 'Signed' : undefined;

      default:
        return String(value);
    }
  }

  private buildItemTitle(definition: IFormDefinition, values: IFormValues): string {
    const fields = inputFields(definition);
    for (const field of fields) {
      if (field.type === FieldType.Text || field.type === FieldType.Email) {
        const text = valueAsComparableString(field, values[field.id]).trim();
        if (text.length > 0) {
          return text.slice(0, 255);
        }
      }
    }
    return 'Response — ' + new Date().toLocaleString();
  }

  // -------------------------------------------------------------------------
  // reading responses
  // -------------------------------------------------------------------------

  private selectsFor(fields: IFormField[]): { selects: string[]; expands: string[] } {
    const selects: string[] = [
      'Id',
      'Created',
      'Modified',
      'Author/Title',
      'Author/EMail',
      'AttachmentFiles/FileName',
      SF_STATUS_INTERNAL_NAME,
      SF_DURATION_INTERNAL_NAME
    ];
    const expands: string[] = ['Author', 'AttachmentFiles'];

    fields.forEach((field) => {
      if (field.type === FieldType.Person) {
        selects.push(field.internalName + '/Title');
        selects.push(field.internalName + '/EMail');
        expands.push(field.internalName);
      } else {
        selects.push(field.internalName);
      }
    });
    return { selects, expands };
  }

  private toResponseItem(item: Record<string, unknown>, fields: IFormField[]): IResponseItem {
    const values: Record<string, unknown> = {};
    fields.forEach((field) => {
      values[field.internalName] = item[field.internalName];
    });
    const author = item.Author as { Title?: string; EMail?: string };
    const status = item[SF_STATUS_INTERNAL_NAME] as ResponseStatus | undefined;
    const duration = item[SF_DURATION_INTERNAL_NAME];
    const attachments = (item.AttachmentFiles || []) as { FileName?: string }[];
    return {
      id: item.Id as number,
      created: new Date(String(item.Created)),
      modified: new Date(String(item.Modified || item.Created)),
      createdBy: (author && author.Title) || '',
      createdByEmail: author ? author.EMail : undefined,
      status: status === 'Draft' ? 'Draft' : 'Complete',
      durationSeconds: typeof duration === 'number' ? duration : undefined,
      attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
      values
    };
  }

  /**
   * Load submitted items shaped for the responses views.
   *
   * Pages through the list rather than taking a single `top(2000)` slice, and
   * filters with `ne 'Draft'` rather than `eq 'Complete'` so items created
   * directly in the list — which have no status value at all — still appear
   * instead of silently vanishing from the owner's results.
   */
  public async getResponses(
    listId: string,
    definition: IFormDefinition,
    options?: { maxItems?: number; includeDrafts?: boolean }
  ): Promise<IResponsePage> {
    const fields = inputFields(definition).filter((f) => f.provisioned !== false);
    const { selects, expands } = this.selectsFor(fields);
    const ceiling = Math.min(options && options.maxItems ? options.maxItems : MAX_RESPONSES, MAX_RESPONSES);
    const includeDrafts = options && options.includeDrafts === true;

    let query = this.list(listId)
      .items.select(...selects)
      .expand(...expands)
      .orderBy('Id', false)
      .top(PAGE_SIZE);
    if (!includeDrafts) {
      query = query.filter(SF_STATUS_INTERNAL_NAME + " ne 'Draft'");
    }

    const items: IResponseItem[] = [];
    let page = await query.getPaged();
    for (;;) {
      (page.results || []).forEach((raw: Record<string, unknown>) => {
        if (items.length < ceiling) {
          items.push(this.toResponseItem(raw, fields));
        }
      });
      if (!page.hasNext || items.length >= ceiling) {
        break;
      }
      page = await page.getNext();
    }

    return {
      items,
      total: items.length,
      hasMore: page.hasNext === true && items.length >= ceiling
    };
  }

  public async deleteResponse(listId: string, itemId: number): Promise<void> {
    await this.list(listId).items.getById(itemId).delete();
  }

  /** Delete several responses in one batched request. */
  public async deleteResponses(listId: string, itemIds: number[]): Promise<void> {
    if (itemIds.length === 0) {
      return;
    }
    if (itemIds.length === 1) {
      await this.deleteResponse(listId, itemIds[0]);
      return;
    }
    const [batchedSp, execute] = this.sp.batched();
    const list = batchedSp.web.lists.getById(listId);
    itemIds.forEach((id) => {
      // deliberately not awaited: batched calls are queued, then executed once
      void list.items.getById(id).delete();
    });
    await execute();
  }

  /** Absolute URLs of an item's attachments, grouped by question. */
  public async getResponseAttachments(
    listId: string,
    itemId: number
  ): Promise<{ fileName: string; url: string; internalName: string }[]> {
    try {
      const files = await this.list(listId)
        .items.getById(itemId)
        .attachmentFiles.select('FileName', 'ServerRelativeUrl')();
      return files.map((f: { FileName: string; ServerRelativeUrl: string }) => {
        const separator = f.FileName.indexOf('__');
        return {
          fileName: separator > 0 ? f.FileName.slice(separator + 2) : f.FileName,
          internalName: separator > 0 ? f.FileName.slice(0, separator) : '',
          url: f.ServerRelativeUrl
        };
      });
    } catch {
      return [];
    }
  }

  // -------------------------------------------------------------------------
  // people
  // -------------------------------------------------------------------------

  /** People picker search across the organization. */
  public async searchPeople(query: string, includeGroups?: boolean): Promise<IPersonInfo[]> {
    if (!query || query.trim().length < 2) {
      return [];
    }
    const results = await this.sp.profiles.clientPeoplePickerSearchUser({
      AllowEmailAddresses: true,
      AllowMultipleEntities: false,
      AllUrlZones: false,
      MaximumEntitySuggestions: 15,
      PrincipalSource: 15, // all sources
      // 1 = users, 4 = security groups, 8 = SharePoint groups; 13 covers all
      PrincipalType: includeGroups ? 13 : 1,
      QueryString: query.trim()
    });
    return results.map((r) => {
      // PrincipalType arrives as a string ("User", "SharePointGroup",
      // "SecurityGroup") and isn't on the published EntityData type
      const entity = (r.EntityData || {}) as { Email?: string; PrincipalType?: string };
      return {
        loginName: r.Key,
        displayName: r.DisplayText,
        email: entity.Email,
        isGroup: /group/i.test(entity.PrincipalType || '')
      };
    });
  }

  // -------------------------------------------------------------------------
  // errors
  // -------------------------------------------------------------------------

  private isNotFoundError(e: unknown): boolean {
    const error = e as { status?: number; message?: string };
    return (
      (error && error.status === 404) ||
      (error && typeof error.message === 'string' && error.message.indexOf('404') !== -1)
    );
  }

  /** True when a create-list call failed because the name is taken. */
  public isDuplicateListNameError(e: unknown): boolean {
    const error = e as { message?: string };
    return !!(
      error &&
      typeof error.message === 'string' &&
      (error.message.indexOf('already exists') !== -1 || error.message.indexOf('-2130575342') !== -1)
    );
  }
}

/** Decode base64 into the ArrayBuffer the attachment API expects. */
const base64ToArrayBuffer = (base64: string): ArrayBuffer => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
};
