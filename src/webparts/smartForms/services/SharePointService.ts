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
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/batching';
import { PermissionKind } from '@pnp/sp/security';
import { PrincipalType } from '@pnp/sp/types';
import { IList } from '@pnp/sp/lists';
import {
  ChoiceFieldFormatType,
  DateTimeFieldFormatType,
  FieldUserSelectionMode,
  IFieldAddResult,
  UrlFieldFormatType
} from '@pnp/sp/fields';
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
  effectiveChoices,
  formatAddress,
  inputFields,
  isFieldVisible,
  MULTI_SEPARATOR,
  normalizeFromSharePoint,
  valueAsComparableString
} from '../utils/formUtils';
import {
  APPROVAL_COLUMNS,
  buildFieldXml,
  decimalsAttribute,
  FIELD_GROUP,
  lcidForCurrencySymbol,
  noteLines,
  SF_APPROVAL_COMMENT_INTERNAL_NAME,
  SF_APPROVAL_STATUS_INTERNAL_NAME,
  SF_REVIEWED_BY_INTERNAL_NAME,
  SF_DURATION_INTERNAL_NAME,
  SF_STATUS_INTERNAL_NAME,
  spTypeForField,
  SYSTEM_COLUMNS,
  typeMatchesExisting
} from '../utils/spFieldXml';
import { logWarning } from '../utils/debug';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../utils/localeUtils';

/** Options value for AddFieldInternalNameHint — honor the Name attribute. */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;

/** Items fetched per request when paging responses. */
const PAGE_SIZE = 500;

/** Hard ceiling on how many responses the client will hold at once. */
const MAX_RESPONSES = 20000;

/** Cap on concurrent column-creation requests during provisioning. */
const PROVISION_CONCURRENCY = 4;

/**
 * Hidden list that stores each form's definition JSON, keyed by web part
 * instance id. This is what makes a form's questions durable independent of
 * the page: SPFx only writes web part properties to the page itself on save
 * or publish, but a response can be provisioned and shared long before that,
 * so the definition needs somewhere to live the moment it's edited.
 */
const CONFIG_LIST_TITLE = 'Smart Forms Configuration';
const CONFIG_DEFINITION_FIELD = 'SFDefinitionJson';

export interface ISubmitOptions {
  durationSeconds?: number;
  /** promote this existing draft item instead of creating a new one */
  replaceItemId?: number;
  /** BCP 47 locale for the default item title date (defaults to the browser's) */
  locale?: string;
}

export interface IDraft {
  id: number;
  values: IFormValues;
  /** titles of file/signature questions whose stored file could not be restored; the user must re-provide them */
  missingFiles: string[];
  /** field ids matching missingFiles */
  missingFileFieldIds: string[];
}

/** Result of a save that also uploads attachments. */
export interface ISubmitResult {
  id: number;
  /** original names of files (or "Signature") that could not be attached */
  failedAttachments: string[];
}

export type ApprovalStatus = 'Pending' | 'Approved' | 'Rejected';

/** A response item plus approval-workflow fields (all undefined when approval is not provisioned). */
export interface IResponseItemEx extends IResponseItem {
  approvalStatus?: ApprovalStatus;
  approvalComment?: string;
  reviewedBy?: string;
}

export interface IResponsePageEx extends IResponsePage {
  items: IResponseItemEx[];
}

export interface IFormVersionInfo {
  versionId: number;
  label: string;
  created: string;
  createdBy: string;
}

export interface ILookupOptionsResult {
  values: string[];
  /** true when the source list had more rows than were read */
  truncated: boolean;
}

/** Defaults for the new Service_ strings, used until the localized bundle carries them. */
const SERVICE_STRING_DEFAULTS: { [key: string]: string } = {
  Service_Item_DefaultTitle: 'Response — {date}',
  Service_Email_ApprovalSubject: 'Your response was {status} — {title}',
  Service_Email_ApprovalHeading: 'Your response was {status}',
  Service_Email_ApprovalIntro: 'A reviewer marked your response to "{title}" as {status}.',
  Service_Email_ApprovalCommentLabel: 'Reviewer comment',
  Service_Approval_Approved: 'approved',
  Service_Approval_Rejected: 'rejected',
  Service_Approval_Pending: 'pending',
  Service_Lookup_InvalidFilter: 'The lookup filter is not valid.',
  Service_Edit_NotAllowed: 'This form does not allow editing a submitted response.',
  Service_Edit_NotYours: 'You can only edit your own response.',
  Service_Signature_AttachmentLabel: 'Signature'
};

const svc = (key: string, params?: { [k: string]: string | number }): string => {
  const localized = (strings as unknown as { [k: string]: string | undefined })[key];
  const template = localized || SERVICE_STRING_DEFAULTS[key] || key;
  return params ? formatString(template, params) : template;
};

/** Max lookup rows read from a source list. */
const MAX_LOOKUP_ITEMS = 10000;

/** Person/lookup joins allowed per request (SharePoint caps a query at 12). */
const MAX_JOINS_PER_REQUEST = 8;

/** Rough budget for the characters in one $select, keeping the URL under limits. */
const SELECT_CHAR_BUDGET = 1400;

/** Items per delete batch. */
const BATCH_SIZE = 100;

/** Id window walked when a filter trips the list view threshold. */
const ID_WINDOW = 4000;

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
  private lookupCache: { [key: string]: ILookupOptionsResult } = {};

  private backfilled: { [listId: string]: Promise<void> } = {};
  private approvalEnsured: { [listId: string]: Promise<void> } = {};

  /** In-flight (or completed) attempt to ensure the config list exists, so concurrent
   *  saves don't race to create it twice; cleared on failure so the next call retries. */
  private configListEnsured: Promise<void> | undefined;

  // -------------------------------------------------------------------------
  // form definition (durable, independent of page save)
  // -------------------------------------------------------------------------

  private escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private async ensureConfigList(): Promise<void> {
    if (!this.configListEnsured) {
      this.configListEnsured = (async () => {
        const result = await this.sp.web.lists.ensure(
          CONFIG_LIST_TITLE,
          'Stores Smart Forms form definitions, independent of page save. Created and managed by Smart Forms — do not delete.',
          100,
          false,
          { Hidden: true, EnableAttachments: false, OnQuickLaunch: false }
        );
        // Check for the column directly rather than trusting result.created: if a
        // previous run created the list but the field creation below failed (or
        // the list was left over from some other partial run), created would be
        // false forever after and the column would never get another chance.
        const existingFields: { InternalName: string }[] = await result.list.fields.select(
          'InternalName'
        )();
        const hasDefinitionField = existingFields.some(
          (f) => f.InternalName === CONFIG_DEFINITION_FIELD
        );
        if (!hasDefinitionField) {
          await result.list.fields.createFieldAsXml({
            SchemaXml:
              '<Field Type="Note" DisplayName="Definition" StaticName="' +
              CONFIG_DEFINITION_FIELD +
              '" Name="' +
              CONFIG_DEFINITION_FIELD +
              '" Group="Smart Forms" Required="FALSE" NumLines="6" RichText="FALSE" />',
            Options: ADD_FIELD_INTERNAL_NAME_HINT
          });
        }
        await this.tuneConfigList(result.list);
      })().catch((error) => {
        // let the next call retry rather than caching a permanent failure
        this.configListEnsured = undefined;
        throw error;
      });
    }
    return this.configListEnsured;
  }

  /**
   * Best-effort hardening of the configuration list: keep version history (so
   * an earlier definition can be restored) and make the list read-only for site
   * members, so a form's definition can't be altered by everyone who can open
   * the site. Owners / full control keep write, as does the current user (who
   * is by definition managing the form). Failures never block.
   */
  private async tuneConfigList(list: IList): Promise<void> {
    try {
      await list.update({ EnableVersioning: true, MajorVersionLimit: 50 });
    } catch (error) {
      logWarning('enabling versioning on the configuration list (non-fatal)', error);
    }
    try {
      const info: { HasUniqueRoleAssignments: boolean } = await list.select('HasUniqueRoleAssignments')();
      if (info.HasUniqueRoleAssignments) {
        return;
      }
      const web = this.sp.web;
      const [memberGroup, readDefs, editDefs, contributeDefs, me] = await Promise.all([
        web.associatedMemberGroup.select('Id')(),
        web.roleDefinitions.filter('RoleTypeKind eq 2').select('Id')(),
        web.roleDefinitions.filter('RoleTypeKind eq 6').select('Id')(),
        web.roleDefinitions.filter('RoleTypeKind eq 3').select('Id')(),
        web.currentUser.select('Id')()
      ]);
      await list.breakRoleInheritance(true, false);
      // the person managing the form keeps write before members lose it
      if (editDefs.length > 0) {
        await list.roleAssignments.add(me.Id, editDefs[0].Id);
      }
      if (readDefs.length > 0) {
        await list.roleAssignments.add(memberGroup.Id, readDefs[0].Id);
      }
      for (const def of [...editDefs, ...contributeDefs]) {
        try {
          await list.roleAssignments.remove(memberGroup.Id, def.Id);
        } catch {
          // the group may not hold this binding
        }
      }
    } catch (error) {
      logWarning('restricting the configuration list to read-only for members (non-fatal)', error);
    }
  }

  /**
   * Owner path: create the configuration list if needed (and tighten it).
   * Non-owners never call this; they only read.
   */
  public async ensureFormConfigStore(): Promise<void> {
    await this.ensureConfigList();
  }

  /** Save (or create) a form's definition, keyed by its web part instance id. */
  public async saveFormDefinition(instanceId: string, definitionJson: string): Promise<void> {
    await this.ensureConfigList();
    const list = this.sp.web.lists.getByTitle(CONFIG_LIST_TITLE);
    const key = this.escapeODataString(instanceId);
    const existing = await list.items.select('Id').filter("Title eq '" + key + "'").top(1)();
    const payload = { Title: instanceId, [CONFIG_DEFINITION_FIELD]: definitionJson };
    if (existing.length > 0) {
      await list.items.getById(existing[0].Id).update(payload);
    } else {
      await list.items.add(payload);
    }
  }

  /**
   * The durably-stored definition for a web part instance, if one has been saved.
   * Read-only by default: it never creates the configuration list or changes
   * anything, so it is safe for any viewer. The owner path passes
   * `createIfMissing` to provision the list (and tighten it) first.
   */
  public async loadFormDefinition(
    instanceId: string,
    createIfMissing?: boolean
  ): Promise<string | undefined> {
    try {
      if (createIfMissing === true) {
        try {
          await this.ensureConfigList();
        } catch (error) {
          logWarning('ensuring the configuration list (non-fatal)', error);
        }
      }
      const list = this.sp.web.lists.getByTitle(CONFIG_LIST_TITLE);
      const key = this.escapeODataString(instanceId);
      const items = await list.items
        .select(CONFIG_DEFINITION_FIELD)
        .filter("Title eq '" + key + "'")
        .top(1)();
      if (items.length === 0) {
        return undefined;
      }
      return (items[0] as Record<string, unknown>)[CONFIG_DEFINITION_FIELD] as string;
    } catch {
      return undefined;
    }
  }

  private async configItemId(instanceId: string): Promise<number | undefined> {
    const list = this.sp.web.lists.getByTitle(CONFIG_LIST_TITLE);
    const key = this.escapeODataString(instanceId);
    const items = await list.items.select('Id').filter("Title eq '" + key + "'").top(1)();
    return items.length > 0 ? (items[0].Id as number) : undefined;
  }

  /** Saved versions of a form's definition, newest first. Empty when unavailable. */
  public async listFormVersions(instanceId: string): Promise<IFormVersionInfo[]> {
    try {
      const id = await this.configItemId(instanceId);
      if (id === undefined) {
        return [];
      }
      const rows: Record<string, unknown>[] = await this.sp.web.lists
        .getByTitle(CONFIG_LIST_TITLE)
        .items.getById(id)
        .versions();
      const nameOf = (value: unknown): string => {
        const v = value as { LookupValue?: string; Title?: string; Email?: string } | string | undefined;
        if (!v) {
          return '';
        }
        return typeof v === 'string' ? v : v.LookupValue || v.Title || v.Email || '';
      };
      return rows
        .map((r) => ({
          versionId: Number(r.VersionId),
          label: String(r.VersionLabel || ''),
          created: String(r.Created || ''),
          createdBy: nameOf(r.CreatedBy) || nameOf(r.Editor)
        }))
        .filter((v) => !isNaN(v.versionId))
        .sort((a, b) => b.versionId - a.versionId);
    } catch {
      return [];
    }
  }

  /** The definition JSON stored in one saved version (restore = save it again). */
  public async getFormVersion(instanceId: string, versionId: number): Promise<string> {
    const id = await this.configItemId(instanceId);
    if (id === undefined) {
      throw new Error('Form definition not found');
    }
    const row: Record<string, unknown> = await this.sp.web.lists
      .getByTitle(CONFIG_LIST_TITLE)
      .items.getById(id)
      .versions.getById(versionId)();
    return String(row[CONFIG_DEFINITION_FIELD] || '');
  }

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
  public async getLookupOptions(listId: string, column: string, filter?: string): Promise<string[]> {
    return (await this.getLookupOptionsEx(listId, column, filter)).values;
  }

  /**
   * As getLookupOptions, but also says whether the source list was larger than
   * what was read (up to 10,000 rows, read in pages). The column name and filter
   * are validated: the filter is spliced into the request URL, so anything that
   * could start another query option or break out of a quoted string is rejected.
   */
  public async getLookupOptionsEx(
    listId: string,
    column: string,
    filter?: string
  ): Promise<ILookupOptionsResult> {
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(column || '')) {
      throw new Error(svc('Service_Lookup_InvalidFilter'));
    }
    const cleanFilter = (filter || '').trim();
    if (cleanFilter.length > 0) {
      const quotes = cleanFilter.split("'").length - 1;
      if (cleanFilter.length > 500 || /[&#\r\n]/.test(cleanFilter) || quotes % 2 !== 0) {
        throw new Error(svc('Service_Lookup_InvalidFilter'));
      }
    }
    const key = listId + '|' + column + '|' + cleanFilter;
    if (this.lookupCache[key]) {
      return this.lookupCache[key];
    }
    let query = this.list(listId).items.select(column).top(2000).orderBy(column, true);
    if (cleanFilter.length > 0) {
      query = query.filter(cleanFilter);
    }
    const seen: { [value: string]: boolean } = {};
    const values: string[] = [];
    let read = 0;
    let truncated = false;
    let page = await query.getPaged();
    for (;;) {
      (page.results || []).forEach((item: Record<string, unknown>) => {
        read++;
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
      if (!page.hasNext) {
        break;
      }
      if (read >= MAX_LOOKUP_ITEMS) {
        truncated = true;
        break;
      }
      page = await page.getNext();
    }
    const result = { values, truncated };
    this.lookupCache[key] = result;
    return result;
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
  public async ensureFields(
    listId: string,
    definition: IFormDefinition,
    enableApproval?: boolean
  ): Promise<IProvisionResult> {
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

    await mapWithConcurrency(toCreate, PROVISION_CONCURRENCY, async (field): Promise<undefined> => {
      const actualInternalName = await this.createField(list, field);
      if (actualInternalName && actualInternalName.toLowerCase() !== field.internalName.toLowerCase()) {
        field.internalName = actualInternalName;
      }
      await this.tryAddToDefaultView(list, field.internalName);
      created.push(field.internalName);
      return undefined;
    });

    await mapWithConcurrency(toSync, PROVISION_CONCURRENCY, async (field): Promise<undefined> => {
      await this.syncExistingField(list, field);
      return undefined;
    });

    // system columns: status (indexed, used by every responses query) and duration
    await this.createSystemColumns(list, SYSTEM_COLUMNS, existingByName, created);
    if (enableApproval === true) {
      await this.createSystemColumns(list, APPROVAL_COLUMNS, existingByName, created);
    }
    // the author filter (edit my response, one-per-person check) needs an index
    // to stay legal past the 5,000-item threshold
    try {
      await list.fields.getByInternalNameOrTitle('Author').update({ Indexed: true });
    } catch (error) {
      logWarning('indexing the Author column (non-fatal)', error);
    }

    fields.forEach((field) => {
      const conflicted = conflicts.filter((c) => c.field.id === field.id).length > 0;
      field.provisioned = !conflicted;
    });

    return { definition: updated, created, conflicts };
  }

  private async createSystemColumns(
    list: IList,
    columns: { internalName: string; xml: () => string }[],
    existingByName: { [lower: string]: string },
    created: string[]
  ): Promise<void> {
    await mapWithConcurrency(columns, PROVISION_CONCURRENCY, async (column): Promise<undefined> => {
      if (existingByName[column.internalName.toLowerCase()] !== undefined) {
        return undefined;
      }
      try {
        await list.fields.createFieldAsXml({
          SchemaXml: column.xml(),
          Options: ADD_FIELD_INTERNAL_NAME_HINT
        });
        created.push(column.internalName);
      } catch {
        try {
          await this.createSystemFieldTyped(list, column.internalName);
          created.push(column.internalName);
        } catch (typedError) {
          // a pre-existing column under a different type shouldn't block publishing
          logWarning('creating system column "' + column.internalName + '" (non-fatal)', typedError);
        }
      }
      return undefined;
    });
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

  /**
   * Creates one form question's column. Sites with custom script disabled reject
   * `createFieldAsXml` outright — CAML can carry a JSLink, so SharePoint blocks it
   * wholesale on those sites and returns a garbled, mistranslated error that has
   * nothing to do with permissions. The plain JSON field-creation endpoints can't
   * carry a JSLink and aren't blocked, so they're tried as a fallback.
   *
   * That endpoint can't set the internal name directly: SharePoint derives it from
   * the title given at creation time. Creating under the field's own (already
   * clean) internal name as that seed, then renaming to the human-readable
   * question text afterward, gets the same result — the internal name is fixed at
   * creation, but the display name (Title) can always change later.
   *
   * Returns the internal name actually assigned, if the fallback path had to run
   * (undefined when the normal CAML path worked, meaning nothing changed).
   */
  private async createField(list: IList, field: IFormField): Promise<string | undefined> {
    try {
      await list.fields.createFieldAsXml({
        SchemaXml: buildFieldXml(field),
        Options: ADD_FIELD_INTERNAL_NAME_HINT
      });
      return undefined;
    } catch (error) {
      logWarning(
        'createFieldAsXml rejected for "' + field.internalName + '", retrying via the typed field API (non-fatal)',
        error
      );
      return this.createFieldTyped(list, field);
    }
  }

  /** JSON (non-CAML) field creation, used when a site blocks createFieldAsXml. */
  private async createFieldTyped(list: IList, field: IFormField): Promise<string> {
    const spType = spTypeForField(field);
    const seedName = field.internalName;
    let added: IFieldAddResult;

    switch (spType) {
      case 'Note':
        added = await list.fields.addMultilineText(seedName, {
          NumberOfLines: noteLines(field),
          RichText: field.type === FieldType.RichText,
          RestrictedMode: false,
          AppendOnly: false,
          AllowHyperlink: true
        });
        break;

      case 'Number':
        added = await list.fields.addNumber(seedName, {
          MinimumValue: field.min,
          MaximumValue: field.max
        });
        break;

      case 'Currency':
        added = await list.fields.addCurrency(seedName, {
          CurrencyLocaleId: lcidForCurrencySymbol(field.currencySymbol)
        });
        break;

      case 'DateTime':
        added = await list.fields.addDateTime(seedName, {
          DisplayFormat:
            field.type === FieldType.Date && !field.includeTime
              ? DateTimeFieldFormatType.DateOnly
              : DateTimeFieldFormatType.DateTime
        });
        break;

      case 'Choice':
        added = await list.fields.addChoice(seedName, {
          Choices: effectiveChoices(field),
          EditFormat:
            field.choiceDisplay === 'buttons' ? ChoiceFieldFormatType.RadioButtons : ChoiceFieldFormatType.Dropdown,
          FillInChoice: field.allowOther === true
        });
        break;

      case 'MultiChoice':
        added = await list.fields.addMultiChoice(seedName, {
          Choices: effectiveChoices(field),
          FillInChoice: field.allowOther === true
        });
        break;

      case 'Boolean':
        added = await list.fields.addBoolean(seedName);
        break;

      case 'URL':
        added = await list.fields.addUrl(seedName, { DisplayFormat: UrlFieldFormatType.Hyperlink });
        break;

      case 'User': {
        const userProps: { SelectionMode: FieldUserSelectionMode; Mult?: boolean } = {
          SelectionMode: field.allowGroups ? FieldUserSelectionMode.PeopleAndGroups : FieldUserSelectionMode.PeopleOnly
        };
        added = await list.fields.addUser(seedName, userProps);
        break;
      }

      case 'UserMulti': {
        const userProps: { SelectionMode: FieldUserSelectionMode; Mult?: boolean } = {
          SelectionMode: field.allowGroups ? FieldUserSelectionMode.PeopleAndGroups : FieldUserSelectionMode.PeopleOnly,
          Mult: true
        };
        added = await list.fields.addUser(seedName, userProps);
        break;
      }

      default: {
        const maxLength = field.maxLength && field.maxLength > 0 && field.maxLength <= 255 ? field.maxLength : 255;
        added = await list.fields.addText(seedName, { MaxLength: maxLength });
        break;
      }
    }

    const props: Record<string, unknown> = {
      Title: field.title || seedName,
      Group: FIELD_GROUP,
      Description: field.description || ''
    };
    if (spType === 'Number') {
      const decimals = decimalsAttribute(field);
      if (decimals !== 'Automatic') {
        props.Decimals = Number(decimals);
      }
    } else if (spType === 'Currency') {
      props.Decimals = field.decimalPlaces === undefined ? 2 : field.decimalPlaces;
    }
    await added.field.update(props);

    return added.data.InternalName || seedName;
  }

  /** JSON fallback for the SFStatus/SFDurationSeconds system columns. */
  private async createSystemFieldTyped(list: IList, internalName: string): Promise<void> {
    if (internalName === SF_STATUS_INTERNAL_NAME) {
      const added = await list.fields.addChoice(internalName, {
        Choices: ['Draft', 'Complete'],
        Indexed: true
      });
      await added.field.update({ Title: 'Response status', Group: FIELD_GROUP, DefaultValue: 'Complete' });
    } else if (internalName === SF_DURATION_INTERNAL_NAME) {
      const added = await list.fields.addNumber(internalName, {});
      await added.field.update({ Title: 'Time to complete (seconds)', Group: FIELD_GROUP, Decimals: 0 });
    } else if (internalName === SF_APPROVAL_STATUS_INTERNAL_NAME) {
      const added = await list.fields.addChoice(internalName, { Choices: ['Pending', 'Approved', 'Rejected'] });
      await added.field.update({ Title: 'Approval status', Group: FIELD_GROUP, DefaultValue: 'Pending' });
    } else if (internalName === SF_APPROVAL_COMMENT_INTERNAL_NAME) {
      const added = await list.fields.addMultilineText(internalName, { NumberOfLines: 3, RichText: false });
      await added.field.update({ Title: 'Approval comment', Group: FIELD_GROUP });
    } else if (internalName === SF_REVIEWED_BY_INTERNAL_NAME) {
      const added = await list.fields.addText(internalName, { MaxLength: 255 });
      await added.field.update({ Title: 'Reviewed by', Group: FIELD_GROUP });
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
    values: IFormValues,
    locale?: string,
    clearMissing?: boolean
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

    if (clearMissing === true) {
      // updating an existing item (draft, edit): anything not in the payload must
      // be blanked, or a stale answer from an earlier save would survive
      fields.forEach((field) => {
        if (field.provisioned === false) {
          return;
        }
        const key = this.payloadKey(field);
        if (payload[key] === undefined) {
          payload[key] = this.emptyValueFor(field);
        }
      });
    }

    payload.Title = this.buildItemTitle(definition, values, locale);
    return payload;
  }

  private emptyValueFor(field: IFormField): unknown {
    const multi =
      (field.type === FieldType.Choice && field.allowMultiple === true) ||
      (field.type === FieldType.Person && field.allowMultiplePeople === true);
    return multi ? [] : null;
  }

  /**
   * Save a submission. Attachments can only be added after the item exists, so
   * the item is created (or a resumed draft promoted) first and files are
   * uploaded second; a failed upload doesn't roll the response back, since a
   * response with a missing attachment beats no response at all. Use
   * submitResponseDetailed to learn which attachments failed.
   */
  public async submitResponse(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    options?: ISubmitOptions
  ): Promise<number> {
    return (await this.submitResponseDetailed(listId, definition, values, options)).id;
  }

  /** As submitResponse, but reports attachments that could not be saved. */
  public async submitResponseDetailed(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    options?: ISubmitOptions
  ): Promise<ISubmitResult> {
    const existing = options && typeof options.replaceItemId === 'number';
    const payload = await this.buildPayload(definition, values, options && options.locale, existing);
    payload[SF_STATUS_INTERNAL_NAME] = 'Complete';
    if (options && typeof options.durationSeconds === 'number') {
      payload[SF_DURATION_INTERNAL_NAME] = options.durationSeconds;
    }

    let itemId: number;
    if (existing) {
      itemId = options.replaceItemId as number;
      await this.list(listId).items.getById(itemId).update(payload);
    } else {
      const result = await this.list(listId).items.add(payload);
      itemId = result.data.Id;
    }

    const failed = await this.uploadAttachments(listId, itemId, definition, values, existing === true);
    return { id: itemId, failedAttachments: failed };
  }

  /** Save (or update) a partial response the respondent can come back to. */
  public async saveDraft(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    existingId?: number,
    locale?: string
  ): Promise<number> {
    return (await this.saveDraftDetailed(listId, definition, values, existingId, locale)).id;
  }

  /**
   * As saveDraft, but reports attachments that could not be saved. Files and
   * signatures are stored as attachments, exactly as on submit, so a resumed
   * draft gets them back.
   */
  public async saveDraftDetailed(
    listId: string,
    definition: IFormDefinition,
    values: IFormValues,
    existingId?: number,
    locale?: string
  ): Promise<ISubmitResult> {
    const existing = typeof existingId === 'number';
    const payload = await this.buildPayload(definition, values, locale, existing);
    payload[SF_STATUS_INTERNAL_NAME] = 'Draft';
    let itemId: number;
    if (existing) {
      itemId = existingId as number;
      await this.list(listId).items.getById(itemId).update(payload);
    } else {
      const result = await this.list(listId).items.add(payload);
      itemId = result.data.Id;
    }
    const failed = await this.uploadAttachments(listId, itemId, definition, values, existing);
    return { id: itemId, failedAttachments: failed };
  }

  /**
   * The current user's most recent draft for this list, if any. File and
   * signature answers are returned only when their attachment is really stored;
   * otherwise they are left empty and listed in `missingFiles` so the form can
   * ask the user to provide them again.
   */
  public async loadDraft(listId: string, definition: IFormDefinition): Promise<IDraft | undefined> {
    try {
      const me = await this.sp.web.currentUser.select('Id')();
      const fields = inputFields(definition).filter((f) => f.provisioned !== false);
      const plan = this.planSelects(fields, false);
      const items = await this.readItems(
        listId,
        plan,
        SF_STATUS_INTERNAL_NAME + " eq 'Draft' and AuthorId eq " + me.Id,
        1
      );
      if (items.length === 0) {
        return undefined;
      }
      return await this.itemToDraft(listId, items[0], fields);
    } catch {
      return undefined;
    }
  }

  /**
   * Values of one of the current user's own responses, shaped for the form
   * (edit-my-response). Same file/signature rules as loadDraft.
   */
  public async loadResponseForEdit(
    listId: string,
    definition: IFormDefinition,
    itemId: number
  ): Promise<IDraft | undefined> {
    try {
      const fields = inputFields(definition).filter((f) => f.provisioned !== false);
      const plan = this.planSelects(fields, false);
      const items = await this.readItems(listId, plan, 'Id eq ' + itemId, 1);
      return items.length === 0 ? undefined : await this.itemToDraft(listId, items[0], fields);
    } catch {
      return undefined;
    }
  }

  private async itemToDraft(
    listId: string,
    item: Record<string, unknown>,
    fields: IFormField[]
  ): Promise<IDraft> {
    const values: IFormValues = {};
    const missing: IFormField[] = [];
    const fileFields: IFormField[] = [];
    fields.forEach((field) => {
      if (field.type === FieldType.FileUpload || field.type === FieldType.Signature) {
        fileFields.push(field);
        return;
      }
      const normalized = normalizeFromSharePoint(field, item[field.internalName]);
      if (normalized !== undefined) {
        values[field.id] = normalized;
      }
    });

    if (fileFields.length > 0) {
      const itemId = item.Id as number;
      const attachments = await this.listAttachments(listId, itemId);
      for (const field of fileFields) {
        const claimed = item[field.internalName];
        const wasAnswered = claimed !== undefined && claimed !== null && String(claimed).length > 0;
        const prefix = (field.internalName + '__').toLowerCase();
        const mine = attachments.filter((a) => a.FileName.toLowerCase().indexOf(prefix) === 0);
        const restored: IFormFile[] = [];
        for (const att of mine) {
          try {
            const buffer = await this.sp.web.getFileByServerRelativePath(att.ServerRelativeUrl).getBuffer();
            restored.push({
              name: att.FileName.slice(prefix.length),
              size: buffer.byteLength,
              content: arrayBufferToBase64(buffer)
            });
          } catch (error) {
            logWarning('restoring attachment "' + att.FileName + '" (non-fatal)', error);
          }
        }
        if (field.type === FieldType.FileUpload) {
          if (restored.length > 0) {
            values[field.id] = restored;
          }
          const expected = wasAnswered ? String(claimed).split(';').filter((n) => n.trim().length > 0).length : 0;
          if (wasAnswered && restored.length < expected) {
            missing.push(field);
          }
        } else if (restored.length > 0) {
          values[field.id] = 'data:image/png;base64,' + restored[0].content;
        } else if (wasAnswered) {
          missing.push(field);
        }
      }
    }

    return {
      id: item.Id as number,
      values,
      missingFiles: missing.map((f) => f.title),
      missingFileFieldIds: missing.map((f) => f.id)
    };
  }

  private async listAttachments(
    listId: string,
    itemId: number
  ): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
    try {
      return await this.list(listId)
        .items.getById(itemId)
        .attachmentFiles.select('FileName', 'ServerRelativeUrl')();
    } catch {
      return [];
    }
  }

  /** True when an error looks like SharePoint's list view threshold (5,000 items). */
  private isThresholdError(e: unknown): boolean {
    const error = e as { message?: string; status?: number };
    return !!(error && typeof error.message === 'string' && /threshold|5000|too many items/i.test(error.message));
  }

  /**
   * Number of completed responses, used by the response-cap check.
   *
   * Reads the list's own ItemCount and subtracts drafts (drafts are an indexed
   * `eq` query, legal at any size), rather than downloading every id. Returns
   * undefined when the count can't be determined; callers must treat that as
   * UNKNOWN (do not close the form, do not claim zero).
   */
  public async countResponses(listId: string): Promise<number | undefined> {
    try {
      const list = this.list(listId);
      const info: { ItemCount: number } = await list.select('ItemCount')();
      let drafts = 0;
      let page = await list.items
        .select('Id')
        .filter(SF_STATUS_INTERNAL_NAME + " eq 'Draft'")
        .top(2000)
        .getPaged();
      drafts += (page.results || []).length;
      while (page.hasNext && drafts < MAX_RESPONSES) {
        page = await page.getNext();
        drafts += (page.results || []).length;
      }
      return Math.max(0, info.ItemCount - drafts);
    } catch (error) {
      logWarning('counting responses (unknown)', error);
      return undefined;
    }
  }

  /**
   * True when the signed-in user already has a completed response; false when
   * they don't; undefined when that could not be determined (treat as unknown,
   * never as "no"). Falls back to walking Id ranges when the filter trips the
   * list view threshold.
   */
  public async currentUserHasResponded(listId: string): Promise<boolean | undefined> {
    let me: { Id: number };
    try {
      me = await this.sp.web.currentUser.select('Id')();
    } catch {
      return undefined;
    }
    const filter = SF_STATUS_INTERNAL_NAME + " eq 'Complete' and AuthorId eq " + me.Id;
    try {
      const items = await this.list(listId).items.select('Id').filter(filter).top(1)();
      return items.length > 0;
    } catch (error) {
      if (!this.isThresholdError(error)) {
        logWarning('checking for an existing response (unknown)', error);
        return undefined;
      }
    }
    try {
      const found = await this.walkIdRanges(listId, filter, 1, ['Id']);
      return found.length > 0;
    } catch (error) {
      logWarning('walking id ranges for an existing response (unknown)', error);
      return undefined;
    }
  }

  /** Newest-first search across descending Id windows, each small enough to stay under the threshold. */
  private async walkIdRanges(
    listId: string,
    filter: string,
    limit: number,
    selects: string[],
    expands?: string[]
  ): Promise<Record<string, unknown>[]> {
    const list = this.list(listId);
    const top = await list.items.select('Id').orderBy('Id', false).top(1)();
    if (top.length === 0) {
      return [];
    }
    let hi = top[0].Id as number;
    const found: Record<string, unknown>[] = [];
    for (let i = 0; i < 100 && hi > 0 && found.length < limit; i++) {
      const lo = Math.max(0, hi - ID_WINDOW);
      let query = list.items.select(...selects);
      if (expands && expands.length > 0) {
        query = query.expand(...expands);
      }
      const rows = await query
        .filter('Id gt ' + lo + ' and Id le ' + hi + ' and ' + filter)
        .orderBy('Id', false)
        .top(limit - found.length)();
      rows.forEach((r: Record<string, unknown>) => found.push(r));
      hi = lo;
    }
    return found;
  }

  /**
   * Upload file and signature answers as list item attachments.
   *
   * Filenames are prefixed with the column's internal name so several upload
   * questions on one form can't collide, and so the responses view can tell
   * which attachment belongs to which question. On an existing item (resumed
   * draft, edit) a question's earlier attachments are replaced when it has new
   * content, or removed when the answer was cleared. Names are made unique
   * case-insensitively. Returns the names that failed.
   */
  private async uploadAttachments(
    listId: string,
    itemId: number,
    definition: IFormDefinition,
    values: IFormValues,
    itemExisted: boolean
  ): Promise<string[]> {
    interface IPlan {
      prefix: string;
      replace: boolean;
      uploads: { display: string; name: string; content: ArrayBuffer }[];
    }
    const plans: IPlan[] = [];

    inputFields(definition).forEach((field) => {
      if (field.type !== FieldType.FileUpload && field.type !== FieldType.Signature) {
        return;
      }
      const visible = isFieldVisible(field, definition, values);
      const plan: IPlan = { prefix: field.internalName + '__', replace: false, uploads: [] };
      if (field.type === FieldType.FileUpload) {
        const files = visible && Array.isArray(values[field.id]) ? (values[field.id] as IFormFile[]) : [];
        const withContent = files.filter((f) => !!f.content);
        withContent.forEach((file) => {
          plan.uploads.push({
            display: file.name,
            name: this.attachmentName(field.internalName, file.name),
            content: base64ToArrayBuffer(file.content)
          });
        });
        // names-only entries mean "keep what's stored"
        plan.replace = withContent.length > 0 || files.length === 0;
      } else {
        const dataUrl = visible ? String(values[field.id] || '') : '';
        const comma = dataUrl.indexOf(',');
        if (dataUrl.indexOf('data:image') === 0 && comma >= 0) {
          plan.uploads.push({
            display: svc('Service_Signature_AttachmentLabel'),
            name: this.attachmentName(field.internalName, 'signature.png'),
            content: base64ToArrayBuffer(dataUrl.slice(comma + 1))
          });
          plan.replace = true;
        } else {
          plan.replace = dataUrl.length === 0;
        }
      }
      plans.push(plan);
    });

    const anyWork = plans.some((p) => p.uploads.length > 0 || (itemExisted && p.replace));
    if (!anyWork) {
      return [];
    }

    const item = this.list(listId).items.getById(itemId);
    const taken: { [lower: string]: boolean } = {};
    if (itemExisted) {
      const existing = await this.listAttachments(listId, itemId);
      for (const att of existing) {
        const lower = att.FileName.toLowerCase();
        const owner = plans.filter((p) => lower.indexOf(p.prefix.toLowerCase()) === 0)[0];
        if (owner && owner.replace) {
          try {
            await item.attachmentFiles.getByName(att.FileName).delete();
            continue;
          } catch (error) {
            logWarning('removing old attachment "' + att.FileName + '" (non-fatal)', error);
          }
        }
        taken[lower] = true;
      }
    }

    const failed: string[] = [];
    for (const plan of plans) {
      for (const upload of plan.uploads) {
        const name = uniqueAttachmentName(upload.name, taken);
        try {
          await item.attachmentFiles.add(name, upload.content);
        } catch (error) {
          // a rejected attachment (size, blocked extension) must not fail an
          // otherwise-valid response, but the caller has to be told
          logWarning('attaching "' + name + '" (non-fatal)', error);
          failed.push(upload.display);
        }
      }
    }
    return failed;
  }

  private attachmentName(internalName: string, fileName: string): string {
    const safe = fileName.replace(/[\\/:*?"<>|#%]/g, '_');
    return internalName + '__' + safe;
  }

  // -------------------------------------------------------------------------
  // edit my response
  // -------------------------------------------------------------------------

  /**
   * The current user's latest completed responses, newest first (default cap 1).
   * Values are raw, keyed by column internal name, like getResponses.
   */
  public async getMyResponses(
    listId: string,
    definition: IFormDefinition,
    limit?: number,
    includeApproval?: boolean
  ): Promise<IResponseItemEx[]> {
    const fields = inputFields(definition).filter((f) => f.provisioned !== false);
    const me = await this.sp.web.currentUser.select('Id')();
    const plan = this.planSelects(fields, includeApproval === true);
    const raw = await this.readItems(
      listId,
      plan,
      SF_STATUS_INTERNAL_NAME + " eq 'Complete' and AuthorId eq " + me.Id,
      Math.max(1, limit || 1)
    );
    return raw.map((r) => this.toResponseItem(r, fields));
  }

  /**
   * Edit a submitted response. The caller passes the form's `allowEdit` flag; it
   * is enforced here too. Only the author (or a list owner) may edit. Attachments
   * are replaced/cleared per question exactly as for a resumed draft, and the
   * status and duration are left alone.
   */
  public async updateResponse(
    listId: string,
    definition: IFormDefinition,
    itemId: number,
    values: IFormValues,
    allowEdit: boolean,
    options?: { locale?: string }
  ): Promise<ISubmitResult> {
    if (allowEdit !== true) {
      throw new Error(svc('Service_Edit_NotAllowed'));
    }
    const item = this.list(listId).items.getById(itemId);
    const [me, current] = await Promise.all([
      this.sp.web.currentUser.select('Id')(),
      item.select('AuthorId')()
    ]);
    if (current.AuthorId !== me.Id && !(await this.currentUserIsOwner(listId))) {
      throw new Error(svc('Service_Edit_NotYours'));
    }
    const payload = await this.buildPayload(definition, values, options && options.locale, true);
    await item.update(payload);
    const failed = await this.uploadAttachments(listId, itemId, definition, values, true);
    return { id: itemId, failedAttachments: failed };
  }

  // -------------------------------------------------------------------------
  // approval workflow
  // -------------------------------------------------------------------------

  /** Create the approval columns if they're missing. Idempotent; cached per list. */
  public async ensureApprovalColumns(listId: string): Promise<void> {
    if (!this.approvalEnsured[listId]) {
      this.approvalEnsured[listId] = (async () => {
        const list = this.list(listId);
        const existing: { InternalName: string }[] = await list.fields.select('InternalName')();
        const names: { [lower: string]: string } = {};
        existing.forEach((f) => {
          names[f.InternalName.toLowerCase()] = f.InternalName;
        });
        await this.createSystemColumns(list, APPROVAL_COLUMNS, names, []);
      })().catch((error) => {
        delete this.approvalEnsured[listId];
        throw error;
      });
    }
    return this.approvalEnsured[listId];
  }

  /**
   * Record a reviewer's decision on a response. Provisions the approval columns
   * on first use. When `notifyRespondent` is set, the respondent is emailed the
   * outcome (best-effort; a mail failure never fails the decision).
   */
  public async setResponseApproval(
    listId: string,
    definition: IFormDefinition,
    itemId: number,
    status: ApprovalStatus,
    comment?: string,
    options?: { notifyRespondent?: boolean; locale?: string }
  ): Promise<void> {
    await this.ensureApprovalColumns(listId);
    const me = await this.sp.web.currentUser.select('Title')();
    const item = this.list(listId).items.getById(itemId);
    await item.update({
      [SF_APPROVAL_STATUS_INTERNAL_NAME]: status,
      [SF_APPROVAL_COMMENT_INTERNAL_NAME]: comment || '',
      [SF_REVIEWED_BY_INTERNAL_NAME]: status === 'Pending' ? '' : me.Title
    });

    if (options && options.notifyRespondent === true && status !== 'Pending') {
      try {
        const author = await item.select('Author/EMail')
          .expand('Author')() as { Author?: { EMail?: string } };
        const email = author.Author && author.Author.EMail;
        if (!email) {
          return;
        }
        const title = definition.settings.formTitle || strings.Service_DefaultFormTitle;
        const statusText = svc('Service_Approval_' + status);
        const heading = svc('Service_Email_ApprovalHeading', { status: statusText });
        const intro = svc('Service_Email_ApprovalIntro', { title, status: statusText });
        const body =
          '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#323130">' +
          '<h2 style="margin:0 0 8px">' + escapeHtml(heading) + '</h2>' +
          '<p>' + escapeHtml(intro) + '</p>' +
          (comment
            ? '<p><strong>' + escapeHtml(svc('Service_Email_ApprovalCommentLabel')) + ':</strong><br/>' +
              escapeHtml(comment).replace(/\n/g, '<br/>') + '</p>'
            : '') +
          '</div>';
        await this.sp.utility.sendEmail({
          To: [email],
          Subject: svc('Service_Email_ApprovalSubject', { title, status: statusText }),
          Body: body,
          AdditionalHeaders: { 'content-type': 'text/html' }
        });
      } catch (error) {
        logWarning('emailing the approval decision (non-fatal)', error);
      }
    }
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
    itemId: number,
    locale?: string
  ): Promise<void> {
    const settings = definition.settings;
    const formTitle = settings.formTitle || strings.Service_DefaultFormTitle;
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
            Subject: formatString(strings.Service_Email_NewResponseSubject, { title: formTitle }),
            Body: buildResponseEmailHtml(definition, values, {
              heading: formatString(strings.Service_Email_NewResponseHeading, { title: formTitle }),
              intro: formatString(strings.Service_Email_NewResponseIntro, { date: new Date().toLocaleString(locale) }),
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
            Subject: formatString(strings.Service_Email_ReceiptSubject, { title: formTitle }),
            Body: buildResponseEmailHtml(definition, values, {
              heading: strings.Service_Email_ReceiptHeading,
              intro: formatString(strings.Service_Email_ReceiptIntro, { title: formTitle }),
              accentColor: settings.accentColor
            }),
            AdditionalHeaders: { 'content-type': 'text/html' }
          });
        })()
      );
    }

    // never let notification failures surface to the person submitting
    await Promise.all(sends.map((p) => p.catch((): undefined => undefined)));
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

  private buildItemTitle(definition: IFormDefinition, values: IFormValues, locale?: string): string {
    const fields = inputFields(definition);
    for (const field of fields) {
      if (field.type === FieldType.Text || field.type === FieldType.Email) {
        const text = valueAsComparableString(field, values[field.id]).trim();
        if (text.length > 0) {
          return text.slice(0, 255);
        }
      }
    }
    return svc('Service_Item_DefaultTitle', { date: new Date().toLocaleString(locale) });
  }

  // -------------------------------------------------------------------------
  // reading responses
  // -------------------------------------------------------------------------

  /**
   * Split the columns to read across requests. SharePoint rejects a query with
   * more than 12 lookup joins (each Person column is one, plus Author), and long
   * $select lists risk the URL length limit. The first group carries the item
   * basics; extra groups carry only Id plus their columns and are merged by Id.
   */
  private planSelects(fields: IFormField[], includeApproval: boolean): ISelectPlan {
    const base: string[] = [
      'Id',
      'Created',
      'Modified',
      'Author/Title',
      'Author/EMail',
      'AttachmentFiles/FileName',
      SF_STATUS_INTERNAL_NAME,
      SF_DURATION_INTERNAL_NAME
    ];
    if (includeApproval) {
      base.push(
        SF_APPROVAL_STATUS_INTERNAL_NAME,
        SF_APPROVAL_COMMENT_INTERNAL_NAME,
        SF_REVIEWED_BY_INTERNAL_NAME
      );
    }
    const groups: ISelectGroup[] = [
      { selects: base, expands: ['Author', 'AttachmentFiles'], joins: 1, length: base.join(',').length }
    ];

    fields.forEach((field) => {
      const isPerson = field.type === FieldType.Person;
      const add = isPerson
        ? [field.internalName + '/Title', field.internalName + '/EMail']
        : [field.internalName];
      const cost = add.join(',').length + 1;
      let group = groups[groups.length - 1];
      if ((isPerson && group.joins + 1 > MAX_JOINS_PER_REQUEST) || group.length + cost > SELECT_CHAR_BUDGET) {
        group = { selects: ['Id'], expands: [], joins: 0, length: 3 };
        groups.push(group);
      }
      group.selects.push(...add);
      group.length += cost;
      if (isPerson) {
        group.expands.push(field.internalName);
        group.joins++;
      }
    });

    return { primary: groups[0], extra: groups.slice(1) };
  }

  /** Fetch the secondary column groups for an Id window and merge them into the primary rows. */
  private async mergeExtraGroups(
    listId: string,
    plan: ISelectPlan,
    rows: Record<string, unknown>[]
  ): Promise<void> {
    if (plan.extra.length === 0 || rows.length === 0) {
      return;
    }
    let minId = Number.MAX_VALUE;
    let maxId = 0;
    const byId: { [id: number]: Record<string, unknown> } = {};
    rows.forEach((r) => {
      const id = r.Id as number;
      byId[id] = r;
      minId = Math.min(minId, id);
      maxId = Math.max(maxId, id);
    });
    await mapWithConcurrency(plan.extra, 2, async (group): Promise<undefined> => {
      let query = this.list(listId).items.select(...group.selects);
      if (group.expands.length > 0) {
        query = query.expand(...group.expands);
      }
      let page = await query
        .filter('Id ge ' + minId + ' and Id le ' + maxId)
        .orderBy('Id', false)
        .top(PAGE_SIZE)
        .getPaged();
      for (;;) {
        (page.results || []).forEach((extra: Record<string, unknown>) => {
          const target = byId[extra.Id as number];
          if (target) {
            Object.keys(extra).forEach((k) => {
              if (k !== 'Id' && k !== '__metadata') {
                target[k] = extra[k];
              }
            });
          }
        });
        if (!page.hasNext) {
          break;
        }
        page = await page.getNext();
      }
      return undefined;
    });
  }

  /** Newest-first items matching a filter (small result sets), with all column groups merged. */
  private async readItems(
    listId: string,
    plan: ISelectPlan,
    filter: string,
    limit: number
  ): Promise<Record<string, unknown>[]> {
    let rows: Record<string, unknown>[];
    try {
      const query = this.list(listId)
        .items.select(...plan.primary.selects)
        .expand(...plan.primary.expands)
        .filter(filter)
        .orderBy('Id', false)
        .top(limit);
      rows = await query();
    } catch (error) {
      if (!this.isThresholdError(error)) {
        throw error;
      }
      rows = await this.walkIdRanges(listId, filter, limit, plan.primary.selects, plan.primary.expands);
    }
    await this.mergeExtraGroups(listId, plan, rows);
    return rows;
  }

  private toResponseItem(item: Record<string, unknown>, fields: IFormField[]): IResponseItemEx {
    const values: Record<string, unknown> = {};
    fields.forEach((field) => {
      values[field.internalName] = item[field.internalName];
    });
    const author = item.Author as { Title?: string; EMail?: string };
    const status = item[SF_STATUS_INTERNAL_NAME] as ResponseStatus | undefined;
    const duration = item[SF_DURATION_INTERNAL_NAME];
    const attachments = (item.AttachmentFiles || []) as { FileName?: string }[];
    const approval = item[SF_APPROVAL_STATUS_INTERNAL_NAME];
    const hasApprovalColumn = SF_APPROVAL_STATUS_INTERNAL_NAME in item;
    return {
      id: item.Id as number,
      created: new Date(String(item.Created)),
      modified: new Date(String(item.Modified || item.Created)),
      createdBy: (author && author.Title) || '',
      createdByEmail: author ? author.EMail : undefined,
      status: status === 'Draft' ? 'Draft' : 'Complete',
      durationSeconds: typeof duration === 'number' ? duration : undefined,
      attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
      values,
      approvalStatus: hasApprovalColumn
        ? approval === 'Approved' || approval === 'Rejected'
          ? (approval as ApprovalStatus)
          : 'Pending'
        : undefined,
      approvalComment: hasApprovalColumn ? String(item[SF_APPROVAL_COMMENT_INTERNAL_NAME] || '') : undefined,
      reviewedBy: hasApprovalColumn ? String(item[SF_REVIEWED_BY_INTERNAL_NAME] || '') : undefined
    };
  }

  /**
   * Owner-only, best-effort: items created before the status column existed (or
   * directly in the list) have no status. Mark them Complete so the indexed
   * `eq 'Complete'` filter, which stays legal past 5,000 items, still finds
   * them. Batched; runs once per list per page load.
   */
  private backfillStatus(listId: string): Promise<void> {
    if (!this.backfilled[listId]) {
      this.backfilled[listId] = (async (): Promise<void> => {
        try {
          if (!(await this.currentUserIsOwner(listId))) {
            return;
          }
          const list = this.list(listId);
          for (let round = 0; round < 10; round++) {
            const rows = await list.items
              .select('Id')
              .filter(SF_STATUS_INTERNAL_NAME + ' eq null')
              .orderBy('Id', true)
              .top(500)();
            if (rows.length === 0) {
              return;
            }
            for (let i = 0; i < rows.length; i += BATCH_SIZE) {
              const chunk = rows.slice(i, i + BATCH_SIZE);
              const [batchedSp, execute] = this.sp.batched();
              const batchedList = batchedSp.web.lists.getById(listId);
              const pending = chunk.map((r: { Id: number }) =>
                // bNewDocumentUpdate keeps Modified / Editor as they were
                batchedList.items
                  .getById(r.Id)
                  .validateUpdateListItem([{ FieldName: SF_STATUS_INTERNAL_NAME, FieldValue: 'Complete' }], true)
              );
              await execute();
              await Promise.all(pending);
            }
            if (rows.length < 500) {
              return;
            }
          }
        } catch (error) {
          logWarning('backfilling response status (non-fatal)', error);
        }
      })();
    }
    return this.backfilled[listId];
  }

  /**
   * Load submitted items shaped for the responses views.
   *
   * Pages through the list rather than taking a single `top(2000)` slice and
   * filters on `eq 'Complete'`, which uses the status column's index and so
   * stays legal past the 5,000-item list view threshold (`ne` cannot use the
   * index). Items with no status at all (created directly in the list) are
   * backfilled to Complete first when the caller owns the list.
   *
   * Person columns and long column lists are read in extra requests and merged,
   * to stay under SharePoint's 12-join and URL-length limits.
   * `includeApproval` also returns approvalStatus/comment/reviewedBy.
   */
  public async getResponses(
    listId: string,
    definition: IFormDefinition,
    options?: { maxItems?: number; includeDrafts?: boolean; includeApproval?: boolean }
  ): Promise<IResponsePageEx> {
    if (options && options.includeApproval === true) {
      try {
        return await this.getResponsesCore(listId, definition, options, true);
      } catch (error) {
        // approval columns not provisioned (yet): read without them
        logWarning('reading responses with approval columns (retrying without)', error);
      }
    }
    return this.getResponsesCore(listId, definition, options, false);
  }

  private async getResponsesCore(
    listId: string,
    definition: IFormDefinition,
    options: { maxItems?: number; includeDrafts?: boolean } | undefined,
    includeApproval: boolean
  ): Promise<IResponsePageEx> {
    const fields = inputFields(definition).filter((f) => f.provisioned !== false);
    const plan = this.planSelects(fields, includeApproval);
    const ceiling = Math.min(options && options.maxItems ? options.maxItems : MAX_RESPONSES, MAX_RESPONSES);
    const includeDrafts = options && options.includeDrafts === true;

    await this.backfillStatus(listId);

    let query = this.list(listId)
      .items.select(...plan.primary.selects)
      .expand(...plan.primary.expands)
      .orderBy('Id', false)
      .top(PAGE_SIZE);
    if (!includeDrafts) {
      query = query.filter(SF_STATUS_INTERNAL_NAME + " eq 'Complete'");
    }

    const items: IResponseItemEx[] = [];
    let page = await query.getPaged();
    for (;;) {
      const raws: Record<string, unknown>[] = page.results || [];
      await this.mergeExtraGroups(listId, plan, raws);
      raws.forEach((raw) => {
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

  /** Delete several responses in batched requests of at most 100. */
  public async deleteResponses(listId: string, itemIds: number[]): Promise<void> {
    if (itemIds.length === 0) {
      return;
    }
    if (itemIds.length === 1) {
      await this.deleteResponse(listId, itemIds[0]);
      return;
    }
    for (let i = 0; i < itemIds.length; i += BATCH_SIZE) {
      const [batchedSp, execute] = this.sp.batched();
      const list = batchedSp.web.lists.getById(listId);
      const pending = itemIds.slice(i, i + BATCH_SIZE).map((id) => list.items.getById(id).delete());
      await execute();
      await Promise.all(pending);
    }
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
      PrincipalType: (includeGroups ? 13 : 1) as PrincipalType,
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

interface ISelectGroup {
  selects: string[];
  expands: string[];
  joins: number;
  length: number;
}

interface ISelectPlan {
  primary: ISelectGroup;
  extra: ISelectGroup[];
}

/** Make an attachment name unique (case-insensitively) against the names already taken. */
const uniqueAttachmentName = (name: string, taken: { [lower: string]: boolean }): string => {
  let candidate = name;
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  for (let n = 2; taken[candidate.toLowerCase()]; n++) {
    candidate = stem + ' (' + n + ')' + ext;
  }
  taken[candidate.toLowerCase()] = true;
  return candidate;
};

const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes.subarray(i, i + 0x8000)));
  }
  return btoa(binary);
};

const escapeHtml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
