import { sanitizeFileName, withCopySuffix } from './uploadUtils';
import { config } from './config';
import { mergeTagIds, tagIds, type Tag, type TagIds } from './tagging';

export interface Folder {
  Name: string;
  ServerRelativeUrl: string;
}

export interface PreviewTag extends Tag {
  description: string;
}

interface Page<T> {
  value: T[];
  'odata.nextLink'?: string;
  '@odata.nextLink'?: string;
}

export interface IUploadResult {
  originalName: string;
  uploadedName?: string;
  serverRelativeUrl?: string;
  ok: boolean;
  batchIdSet: boolean;
  error?: string;
}

export interface IPreflightResult {
  folderServerRelativeUrl: string;
  batchFieldInternalName?: string;
  error?: string;
  warning?: string;
}

type TokenProvider = () => Promise<string>;

const MAX_NAME_ATTEMPTS: number = 25;
const MAX_RETRIES: number = 4;

function odataLiteral(value: string): string {
  return encodeURIComponent(`'${value.replace(/'/g, "''")}'`);
}

async function readError(response: Response): Promise<string> {
  try {
    const body = await response.json();
    const msg = (body.error ?? body['odata.error'])?.message;
    if (typeof msg === 'string') { return msg; }
    if (msg?.value) { return msg.value; }
  } catch { /* non-JSON body */ }
  if (response.status === 403) { return 'Access denied. You need Contribute permission on the upload library.'; }
  return `${response.status} ${response.statusText}`;
}

const sleep = (ms: number): Promise<void> => new Promise(r => setTimeout(r, ms));

export class SharePointService {
  private readonly _webServerRelativeUrl: string;
  private readonly _siteUrl: string;
  private readonly _getToken: TokenProvider;

  public constructor(siteUrl: string, getToken: TokenProvider) {
    this._siteUrl = siteUrl.replace(/\/+$/, '');
    this._getToken = getToken;
    this._webServerRelativeUrl = new URL(this._siteUrl).pathname.replace(/\/+$/, '') || '/';
  }

  /** Authenticated fetch with retry on throttling (429/503) honoring Retry-After. */
  private async _fetch(path: string, init: RequestInit = {}): Promise<Response> {
    for (let attempt = 0; ; attempt++) {
      const token: string = await this._getToken();
      const res: Response = await fetch(`${this._siteUrl}${path}`, {
        ...init,
        headers: { Accept: 'application/json;odata=nometadata', Authorization: `Bearer ${token}`, ...(init.headers || {}) }
      });
      if ((res.status === 429 || res.status === 503) && attempt < MAX_RETRIES) {
        const retryAfter: number = Number(res.headers.get('Retry-After')) || 2 ** attempt;
        await sleep(retryAfter * 1000);
        continue;
      }
      return res;
    }
  }

  public resolveFolderUrl(folderPath: string): string {
    const trimmed: string = (folderPath || '').trim().replace(/\/+$/, '');
    if (trimmed.startsWith('/')) { return trimmed; }
    const web: string = this._webServerRelativeUrl === '/' ? '' : this._webServerRelativeUrl;
    return `${web}/${trimmed}`;
  }

  public assertLibraryFolder(folderUrl: string): void {
    const root = this.resolveFolderUrl(config.libraryPath).toLowerCase();
    const path = folderUrl.toLowerCase();
    if (path !== root && !path.startsWith(`${root}/`)) {
      throw new Error('Choose a folder inside the configured Documents library.');
    }
  }

  private async _pages<T>(path: string): Promise<T[]> {
    const records: T[] = [];
    let next: string | undefined = path;
    while (next) {
      const res = await this._fetch(next);
      if (!res.ok) { throw new Error(await readError(res)); }
      const page: Page<T> = await res.json();
      if (!Array.isArray(page.value)) { throw new Error('SharePoint returned an invalid list response.'); }
      records.push(...page.value);
      const link: string | undefined = page['odata.nextLink'] ?? page['@odata.nextLink'];
      next = undefined;
      if (link) {
        const url = new URL(link, this._siteUrl);
        if (url.origin !== new URL(this._siteUrl).origin ||
            !url.pathname.startsWith(`${this._webServerRelativeUrl}/_api/`)) {
          throw new Error('SharePoint returned a pagination link outside the configured site.');
        }
        next = `${url.pathname.slice(this._webServerRelativeUrl.length)}${url.search}`;
      }
    }
    return records;
  }

  public async listFolders(folderUrl: string): Promise<Folder[]> {
    this.assertLibraryFolder(folderUrl);
    const folders = await this._pages<Folder>(
      `/_api/web/GetFolderByServerRelativePath(decodedurl=@f)/Folders` +
      `?$select=Name,ServerRelativeUrl&$top=200&@f=${odataLiteral(folderUrl)}`);
    return folders.filter(folder => folder.Name !== 'Forms')
      .sort((a, b) => a.Name.localeCompare(b.Name));
  }

  public async loadTags(): Promise<Tag[]> {
    const fields = await this._pages<{ InternalName: string; TypeAsString: string; LookupList?: string }>(
      `/_api/web/lists(guid'${config.libraryId}')/fields?$select=InternalName,TypeAsString,LookupList`);
    for (const name of ['LocationTag', 'UseTag', 'SubjectTag', 'ItemTags']) {
      const field = fields.find(f => f.InternalName === name);
      if (!field || field.TypeAsString !== 'LookupMulti' ||
          field.LookupList?.replace(/[{}]/g, '').toLowerCase() !== config.tagsListId.toLowerCase()) {
        throw new Error(`Tagging requires the ${name} multi-lookup field linked to the Tags list.`);
      }

    }
    const tags = await this._pages<Tag>(
      `/_api/web/lists(guid'${config.tagsListId}')/items?$select=Id,Tag,TagType&$top=200`);
    if (tags.some(tag => !Number.isInteger(tag.Id) || !tag.Tag?.trim() || !tag.TagType?.trim())) {
      throw new Error('The Tags list contains a missing tag name or category. Correct it before adding tags.');
    }
    return tags.sort((a, b) => a.TagType.localeCompare(b.TagType) || a.Tag.localeCompare(b.Tag));
  }

  public async loadPreviewTags(): Promise<PreviewTag[]> {
    const rows = await this._pages<{
      Id: number; Title: string; field_1: string; field_2?: string | null; field_3: boolean;
    }>(
      `/_api/web/lists(guid'${config.previewTagsListId}')/items` +
      '?$select=Id,Title,field_1,field_2,field_3&$filter=field_3 eq 1&$top=200');
    const active = rows.filter(row => row.field_3 === true);
    if (active.some(row => !Number.isInteger(row.Id) ||
        typeof row.Title !== 'string' || !row.Title.trim() ||
        typeof row.field_1 !== 'string' || !row.field_1.trim())) {
      throw new Error('An active tag is missing its category or tag name. Correct the tag list and retry.');
    }
    return active.map(row => ({
      Id: row.Id, TagType: row.Title.trim(), Tag: row.field_1.trim(),
      description: row.field_2?.trim() || ''
    })).sort((a, b) => a.TagType.localeCompare(b.TagType) || a.Tag.localeCompare(b.Tag));
  }

  public async addTags(fileUrl: string, tags: readonly Tag[]): Promise<void> {
    this.assertLibraryFolder(fileUrl.substring(0, fileUrl.lastIndexOf('/')));
    const path = `/_api/web/GetFileByServerRelativePath(decodedurl=@p)/ListItemAllFields?@p=${odataLiteral(fileUrl)}`;
    const res = await this._fetch(path, { headers: { Accept: 'application/json;odata=verbose' } });
    if (!res.ok) { throw new Error(await readError(res)); }
    const body: { d?: { __metadata?: { etag?: string; type?: string }; [key: string]: unknown } } = await res.json();
    const item = body.d;
    const etag = item?.__metadata?.etag;
    const itemType = item?.__metadata?.type;
    if (!item || !etag || !itemType) { throw new Error('SharePoint did not return a version and type for this photo. Retry saving tags.'); }
    const readIds = (field: string): number[] => {
      const value = item[`${field}Id`];
      if (!value || typeof value !== 'object' || !('results' in value) ||
          !Array.isArray(value.results) || value.results.some(id => !Number.isInteger(id))) {
        throw new Error(`SharePoint returned invalid values for ${field}.`);
      }
      return value.results;
    };
    const existing: TagIds = {
      LocationTag: readIds('LocationTag'), UseTag: readIds('UseTag'),
      SubjectTag: readIds('SubjectTag'), ItemTags: readIds('ItemTags')
    };
    const merged = mergeTagIds(existing, tagIds(tags));
    const saved = await this._fetch(path, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'MERGE', 'IF-MATCH': etag
      },
      body: JSON.stringify({
        __metadata: { type: itemType },
        LocationTagId: { results: merged.LocationTag }, UseTagId: { results: merged.UseTag },
        SubjectTagId: { results: merged.SubjectTag }, ItemTagsId: { results: merged.ItemTags }
      })
    });
    if (!saved.ok) {
      if (saved.status === 412) { throw new Error('This photo changed while saving. Retry to preserve the latest tags.'); }
      throw new Error(await readError(saved));
    }
  }

  /** Verifies the folder exists and that the batch column exists on its library. */
  public async preflight(folderPath: string, batchFieldName: string): Promise<IPreflightResult> {
    const folderUrl: string = this.resolveFolderUrl(folderPath);
    this.assertLibraryFolder(folderUrl);
    const result: IPreflightResult = { folderServerRelativeUrl: folderUrl };
    const folderApi: string = '/_api/web/GetFolderByServerRelativePath(decodedurl=@f)';

    const folderRes: Response = await this._fetch(`${folderApi}?$select=Exists,ServerRelativeUrl&@f=${odataLiteral(folderUrl)}`);
    if (!folderRes.ok) {
      result.error = `Upload folder "${folderUrl}" could not be opened: ${await readError(folderRes)}`;
      return result;
    }
    const folder: { Exists: boolean; ServerRelativeUrl: string } = await folderRes.json();
    if (!folder.Exists) {
      result.error = `Upload folder "${folderUrl}" does not exist.`;
      return result;
    }
    result.folderServerRelativeUrl = folder.ServerRelativeUrl;

    if (!batchFieldName) { return result; }
    const fieldRes: Response = await this._fetch(
      `/_api/web/lists(guid'${config.libraryId}')/Fields/GetByInternalNameOrTitle(${odataLiteral(batchFieldName)})` +
      `?$select=InternalName,TypeAsString`);
    if (fieldRes.ok) {
      const field: { InternalName: string; TypeAsString: string } = await fieldRes.json();
      if (field.TypeAsString !== 'Text') {
        result.error = `Column "${batchFieldName}" must be a single line of text.`;
      } else {
        result.batchFieldInternalName = field.InternalName;
      }
    } else {
      result.error = `Cannot verify required column "${batchFieldName}": ${await readError(fieldRes)}`;
    }
    return result;
  }

  public async uploadFile(folderServerRelativeUrl: string, file: File, batchId: string, batchFieldInternalName?: string): Promise<IUploadResult> {
    this.assertLibraryFolder(folderServerRelativeUrl);
    if (!batchFieldInternalName) { throw new Error('The required Batch ID column has not been verified.'); }
    const result: IUploadResult = { originalName: file.name, ok: false, batchIdSet: false };
    const baseName: string = sanitizeFileName(file.name);

    for (let attempt = 1; attempt <= MAX_NAME_ATTEMPTS; attempt++) {
      const name: string = attempt === 1 ? baseName : withCopySuffix(baseName, attempt);
      const res: Response = await this._fetch(
        `/_api/web/GetFolderByServerRelativePath(decodedurl=@f)/Files/AddUsingPath(decodedurl=@n,overwrite=false)` +
        `?@f=${odataLiteral(folderServerRelativeUrl)}&@n=${odataLiteral(name)}`,
        { method: 'POST', body: file });
      if (res.ok) {
        const uploaded: { ServerRelativeUrl: string; Name: string } = await res.json();
        result.ok = true;
        result.uploadedName = uploaded.Name;
        result.serverRelativeUrl = uploaded.ServerRelativeUrl;
        break;
      }
      const message: string = await readError(res);
      if (!/already exists/i.test(message)) {
        result.error = message;
        return result;
      }
    }
    if (!result.ok) {
      result.error = 'Could not find a unique file name.';
      return result;
    }

    if (batchFieldInternalName && result.serverRelativeUrl) {
      result.error = await this._setBatchId(result.serverRelativeUrl, batchFieldInternalName, batchId);
      result.batchIdSet = !result.error;
    }
    return result;
  }

  private async _setBatchId(fileServerRelativeUrl: string, fieldName: string, batchId: string): Promise<string | undefined> {
    const res: Response = await this._fetch(
      `/_api/web/GetFileByServerRelativePath(decodedurl=@p)/ListItemAllFields/ValidateUpdateListItem?@p=${odataLiteral(fileServerRelativeUrl)}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({ formValues: [{ FieldName: fieldName, FieldValue: batchId }], bNewDocumentUpdate: true })
      });
    if (!res.ok) { return `Uploaded, but Batch ID not set: ${await readError(res)}`; }
    const body: { value?: { HasException: boolean; ErrorMessage?: string }[] } = await res.json();
    if (!Array.isArray(body.value) || body.value.length !== 1) {
      return 'Uploaded, but SharePoint did not confirm the Batch ID. Retry setting the Batch ID.';
    }
    const failed = body.value.find(v => v.HasException);
    return failed ? `Uploaded, but Batch ID not set: ${failed.ErrorMessage}` : undefined;
  }

  public async retryBatchId(fileUrl: string, batchId: string, field: string): Promise<void> {
    this.assertLibraryFolder(fileUrl.substring(0, fileUrl.lastIndexOf('/')));
    const error = await this._setBatchId(fileUrl, field, batchId);
    if (error) { throw new Error(error); }
  }
}
