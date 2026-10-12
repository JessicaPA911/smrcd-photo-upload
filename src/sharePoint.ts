import { sanitizeFileName, withCopySuffix } from './uploadUtils';

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

  /** Verifies the folder exists and that the batch column exists on its library. */
  public async preflight(folderPath: string, batchFieldName: string): Promise<IPreflightResult> {
    const folderUrl: string = this.resolveFolderUrl(folderPath);
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
      `${folderApi}/ListItemAllFields/ParentList/Fields/GetByInternalNameOrTitle(${odataLiteral(batchFieldName)})` +
      `?$select=InternalName&@f=${odataLiteral(folder.ServerRelativeUrl)}`);
    if (fieldRes.ok) {
      result.batchFieldInternalName = (await fieldRes.json()).InternalName;
    } else {
      result.warning = `Column "${batchFieldName}" was not found on the library. Photos will upload, but will not be tagged with the Batch ID.`;
    }
    return result;
  }

  public async uploadFile(folderServerRelativeUrl: string, file: File, batchId: string, batchFieldInternalName?: string): Promise<IUploadResult> {
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
    const failed = (body.value || []).find(v => v.HasException);
    return failed ? `Uploaded, but Batch ID not set: ${failed.ErrorMessage}` : undefined;
  }
}
