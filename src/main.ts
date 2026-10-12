import './style.css';
import { config } from './config';
import { ensureSignedIn, getSharePointToken, signOut } from './auth';
import { SharePointService, type IPreflightResult, type IUploadResult } from './sharePoint';
import { appendParams, createBatchId, isAllowedReturnUrl, isImageFile, sanitizeBatchId } from './uploadUtils';

interface ILaunchContext {
  returnUrl?: string;
  batchId?: string;
  upn?: string;
}

const CONTEXT_KEY: string = 'smrcd-upload-context';

const app: HTMLElement = document.getElementById('app')!;
const service: SharePointService = new SharePointService(config.siteUrl, getSharePointToken);

let userName: string = '';
let batchId: string = '';
let returnUrl: string = '';
let preflight: IPreflightResult | undefined;
let selected: File[] = [];
let rejected: string[] = [];
const results: IUploadResult[] = [];
let uploading: boolean = false;
let progressDone: number = 0;
let progressTotal: number = 0;
let previewUrls: string[] = [];

/** Query params are lost across the sign-in redirect, so they travel in MSAL `state` and sessionStorage. */
function readLaunchContext(): ILaunchContext {
  const q: URLSearchParams = new URLSearchParams(window.location.search);
  if (q.has('returnUrl') || q.has('batchId') || q.has('upn')) {
    const ctx: ILaunchContext = {
      returnUrl: q.get('returnUrl') ?? undefined,
      batchId: q.get('batchId') ?? undefined,
      upn: q.get('upn') ?? undefined
    };
    sessionStorage.setItem(CONTEXT_KEY, JSON.stringify(ctx));
    return ctx;
  }
  try {
    return JSON.parse(sessionStorage.getItem(CONTEXT_KEY) || '{}');
  } catch {
    return {};
  }
}

async function start(): Promise<void> {
  let ctx: ILaunchContext = readLaunchContext();
  const auth = await ensureSignedIn(JSON.stringify(ctx), ctx.upn);
  if (!auth) { return; } // navigating to sign-in

  if (auth.state) {
    try { ctx = { ...JSON.parse(auth.state), ...ctx }; } catch { /* ignore malformed state */ }
  }
  if (window.location.search || window.location.hash) {
    history.replaceState(undefined, '', window.location.pathname);
  }

  userName = auth.account.name || auth.account.username;
  batchId = sanitizeBatchId(ctx.batchId) || createBatchId();
  ctx.batchId = batchId;
  sessionStorage.setItem(CONTEXT_KEY, JSON.stringify(ctx));

  const requested: string = ctx.returnUrl || '';
  returnUrl = isAllowedReturnUrl(requested, config.defaultReturnUrl) ? requested : config.defaultReturnUrl;

  try {
    preflight = await service.preflight(config.folderPath, config.batchFieldName);
  } catch (e) {
    preflight = { folderServerRelativeUrl: '', error: (e as Error).message };
  }
  render();
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, children: (Node | string)[] = []): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), props);
  children.forEach(c => node.append(c));
  return node;
}

const plural = (n: number): string => `${n} photo${n === 1 ? '' : 's'}`;

function render(): void {
  revokePreviews();
  app.replaceChildren();
  const done: IUploadResult[] = results.filter(r => r.ok);
  const failed: IUploadResult[] = results.filter(r => !r.ok);
  const blocked: boolean = !!preflight?.error;

  app.append(
    el('h1', { className: 'title', textContent: 'Upload photos' }),
    el('div', { className: 'batch' }, ['Batch ID: ', el('strong', { textContent: batchId })]),
    el('div', { className: 'user' }, [`Signed in as ${userName} · `, el('button', { type: 'button', className: 'inlineLink', textContent: 'Switch account', onclick: () => { signOut().catch(console.error); } })])
  );
  if (preflight?.error) { app.append(el('div', { className: 'error', role: 'alert', textContent: preflight.error })); }
  if (preflight?.warning) { app.append(el('div', { className: 'warning', role: 'status', textContent: preflight.warning })); }

  const fileInput = el('input', { type: 'file', accept: 'image/*', multiple: true, className: 'hiddenInput', tabIndex: -1 });
  fileInput.setAttribute('aria-hidden', 'true');
  fileInput.addEventListener('change', () => onFilesChosen(fileInput.files));
  app.append(fileInput, el('button', {
    type: 'button', className: 'primary', disabled: blocked || uploading,
    textContent: done.length ? 'Add more photos to this batch' : 'Choose photos',
    onclick: () => fileInput.click()
  }));

  if (selected.length && !uploading) {
    const bytes: number = selected.reduce((s, f) => s + f.size, 0);
    const thumbs = el('ul', { className: 'thumbs' });
    selected.forEach((file, index) => {
      const url: string = URL.createObjectURL(file);
      previewUrls.push(url);
      const remove = el('button', { type: 'button', className: 'remove', textContent: '×', onclick: () => { selected.splice(index, 1); render(); } });
      remove.setAttribute('aria-label', `Remove ${file.name}`);
      thumbs.append(el('li', {}, [el('img', { src: url, alt: file.name, loading: 'lazy' }), remove]));
    });
    app.append(
      el('p', { className: 'summary', textContent: `${plural(selected.length)} selected (${(bytes / 1048576).toFixed(1)} MB)` }),
      thumbs,
      el('button', { type: 'button', className: 'primary', disabled: blocked, textContent: `Upload ${plural(selected.length)}`, onclick: () => { upload().catch(console.error); } }),
      el('button', { type: 'button', className: 'secondary', textContent: 'Clear selection', onclick: () => { selected = []; rejected = []; render(); } })
    );
  }

  if (rejected.length) {
    app.append(el('div', { className: 'warning', role: 'status', textContent: `Skipped (not photos): ${rejected.join(', ')}` }));
  }

  if (uploading) {
    const pct: number = progressTotal ? Math.round((progressDone / progressTotal) * 100) : 0;
    const bar = el('div', { className: 'progressBar' }, [el('div', { className: 'progressFill' })]);
    bar.setAttribute('role', 'progressbar');
    bar.setAttribute('aria-valuemin', '0');
    bar.setAttribute('aria-valuemax', '100');
    bar.setAttribute('aria-valuenow', String(pct));
    (bar.firstChild as HTMLElement).style.width = `${pct}%`;
    app.append(bar, el('p', { className: 'summary', textContent: `Uploading ${progressDone} / ${progressTotal}…` }));
  }

  if (results.length && !uploading) {
    app.append(el('p', { className: failed.length ? 'success partial' : 'success', textContent: `${plural(done.length)} uploaded to batch ${batchId}.` }));
    const problems: IUploadResult[] = results.filter(r => !r.ok || r.error);
    if (problems.length) {
      app.append(el('ul', { className: 'failures' }, problems.map(r => el('li', { textContent: `${r.originalName}: ${r.error}` }))));
    }
  }

  const footer = el('div', { className: 'footer' });
  if (done.length && !uploading && returnUrl) {
    footer.append(el('button', { type: 'button', className: 'primary', textContent: 'Return to app', onclick: () => navigateBack(false) }));
  }
  if (!uploading && returnUrl) {
    footer.append(el('button', { type: 'button', className: 'link', textContent: done.length ? 'Return without sending batch' : 'Cancel', onclick: () => navigateBack(true) }));
  }
  app.append(footer);
}

function onFilesChosen(files: FileList | null): void {
  if (!files) { return; }
  rejected = [];
  Array.from(files).forEach(f => (isImageFile(f) ? selected.push(f) : rejected.push(f.name)));
  render();
}

async function upload(): Promise<void> {
  if (!preflight || preflight.error || !selected.length) { return; }
  const queue: File[] = selected.slice();
  selected = [];
  uploading = true;
  progressDone = 0;
  progressTotal = queue.length;
  render();

  const worker = async (): Promise<void> => {
    let file: File | undefined;
    while ((file = queue.shift()) !== undefined) {
      let result: IUploadResult;
      try {
        result = await service.uploadFile(preflight!.folderServerRelativeUrl, file, batchId, preflight!.batchFieldInternalName);
      } catch (e) {
        result = { originalName: file.name, ok: false, batchIdSet: false, error: (e as Error).message };
      }
      results.push(result);
      progressDone++;
      render();
    }
  };
  await Promise.all(Array.from({ length: Math.min(config.uploadConcurrency, queue.length) }, worker));

  uploading = false;
  render();
  if (config.autoReturn && results.every(r => r.ok && !r.error)) { navigateBack(false); }
}

function navigateBack(cancelled: boolean): void {
  if (!returnUrl) { return; }
  const uploaded: number = results.filter(r => r.ok).length;
  const params: Record<string, string> = cancelled || !uploaded
    ? { UploadStatus: 'Cancelled' }
    : { [config.returnParamName]: batchId, UploadCount: String(uploaded), UploadStatus: 'Complete' };
  sessionStorage.removeItem(CONTEXT_KEY);
  window.location.href = appendParams(returnUrl, params);
}

function revokePreviews(): void {
  previewUrls.forEach(u => URL.revokeObjectURL(u));
  previewUrls = [];
}

window.addEventListener('beforeunload', e => {
  if (uploading) { e.preventDefault(); }
});

start().catch(e => {
  app.replaceChildren(el('div', { className: 'error', role: 'alert', textContent: `Sign-in or startup failed: ${(e as Error).message}` }));
  console.error(e);
});
