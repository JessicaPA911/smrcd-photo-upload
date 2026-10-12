import './style.css';
import { config } from './config';
import { ensureSignedIn, getSharePointToken, signOut } from './auth';
import { SharePointService, type Folder, type IPreflightResult, type IUploadResult } from './sharePoint';
import { appendParams, createBatchId, isAllowedReturnUrl, isImageFile, sanitizeBatchId } from './uploadUtils';
import type { Tag } from './tagging';

interface LaunchContext { returnUrl?: string; batchId?: string; upn?: string }
interface Photo {
  file: File;
  preview: string;
  result?: IUploadResult;
  selected: boolean;
  tags: Set<number>;
  savedTags: Set<number>;
  tagError?: string;
}

const CONTEXT_KEY = 'smrcd-upload-context';
const app = document.getElementById('app');
if (!app) { throw new Error('The upload page is missing its app container.'); }
const root: HTMLElement = app;
const service = new SharePointService(config.siteUrl, getSharePointToken);
const libraryRoot = service.resolveFolderUrl(config.libraryPath);
let browsePath = libraryRoot;
let folders: Folder[] = [];
let destination: IPreflightResult | undefined;
let userName = '';
let batchId = '';
let returnUrl = '';
let busy = false;
let message = '';
let error = '';
let tagsError = '';
let tags: Tag[] = [];
let tagging = false;
let addTags = false;
let rejected: string[] = [];
const photos: Photo[] = [];
const chosenTags = new Set<number>();

function errorMessage(e: unknown): string { return e instanceof Error ? e.message : String(e); }

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, children: (Node | string)[] = []
): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), props);
  children.forEach(child => node.append(child));
  return node;
}

function button(text: string, action: () => void | Promise<void>, className = 'secondary', disabled = busy): HTMLButtonElement {
  return el('button', {
    type: 'button', textContent: text, className, disabled,
    onclick: () => { Promise.resolve().then(action).catch(showError); }
  });
}

function showError(e: unknown): void { error = errorMessage(e); console.error(e); render(); }
function label(text: string, input: HTMLInputElement): HTMLLabelElement {
  return el('label', { className: 'option' }, [input, text]);
}
function ready(photo: Photo): boolean { return !!photo.result?.ok && photo.result.batchIdSet; }
function dirty(photo: Photo): boolean {
  return photo.tags.size !== photo.savedTags.size || [...photo.tags].some(id => !photo.savedTags.has(id));
}

function readContext(): LaunchContext {
  const query = new URLSearchParams(location.search);
  const stored: LaunchContext = JSON.parse(sessionStorage.getItem(CONTEXT_KEY) || '{}');
  const ctx = query.has('returnUrl') || query.has('batchId') || query.has('upn') ? {
    returnUrl: query.get('returnUrl') ?? undefined,
    batchId: query.get('batchId') ?? undefined, upn: query.get('upn') ?? undefined
  } : stored;
  sessionStorage.setItem(CONTEXT_KEY, JSON.stringify(ctx));
  return ctx;
}

async function start(): Promise<void> {
  let ctx = readContext();
  const auth = await ensureSignedIn(JSON.stringify(ctx), ctx.upn);
  if (!auth) { return; }
  if (auth.state) { ctx = { ...JSON.parse(auth.state), ...ctx }; }
  history.replaceState(undefined, '', location.pathname);
  userName = auth.account.name || auth.account.username;
  batchId = sanitizeBatchId(ctx.batchId) || createBatchId();
  sessionStorage.setItem(CONTEXT_KEY, JSON.stringify({ ...ctx, batchId }));
  returnUrl = ctx.returnUrl && isAllowedReturnUrl(ctx.returnUrl, config.defaultReturnUrl)
    ? ctx.returnUrl : config.defaultReturnUrl;
  await getSharePointToken(true);
  render();
  await browse(libraryRoot);
}

async function browse(path: string): Promise<void> {
  busy = true; error = ''; message = 'Loading folders...'; render();
  try {
    const found = await service.listFolders(path);
    folders = found; browsePath = path;
  } catch (e) { showError(e); }
  finally { busy = false; message = ''; render(); }
}

async function chooseDestination(): Promise<void> {
  busy = true; error = ''; message = 'Checking upload destination...'; render();
  try { destination = await service.preflight(browsePath, config.batchFieldName); }
  catch (e) { destination = undefined; showError(e); }
  finally { busy = false; message = ''; render(); }
}

async function loadTags(): Promise<void> {
  busy = true; tagsError = ''; message = 'Loading tag choices...'; render();
  try {
    tags = await service.loadTags();
    if (!tags.length) { throw new Error('The Tags list is empty. Add tag choices in SharePoint first.'); }
  } catch (e) { tagsError = errorMessage(e); console.error(e); }
  finally { busy = false; message = ''; render(); }
}

function render(): void {
  root.replaceChildren(
    el('h1', { className: 'title', textContent: tagging ? 'Add tags to your photos' : 'Upload photos' }),
    el('div', { className: 'batch' }, ['Batch ID: ', el('strong', { textContent: batchId })]),
    el('div', { className: 'user' }, [
      `Signed in as ${userName} · `,
      button('Switch account', async () => {
        if (photos.length && !confirm('Switching accounts will discard this page session. Uploaded files stay in SharePoint. Continue?')) { return; }
        sessionStorage.removeItem(CONTEXT_KEY);
        await signOut();
      }, 'inlineLink')
    ])
  );
  if (error) { root.append(el('div', { className: 'error', role: 'alert', textContent: error })); }
  if (destination?.error) { root.append(el('div', { className: 'error', role: 'alert', textContent: destination.error })); }
  if (message) { root.append(el('p', { role: 'status', className: 'summary', textContent: message })); }

  const uploaded = photos.some(photo => photo.result?.ok);
  if (!uploaded && !tagging) { renderFolders(); }
  if (destination && !destination.error) {
    root.append(el('p', { className: 'destination', textContent: `Upload destination: ${destination.folderServerRelativeUrl}` }));
  }
  if (!uploaded && !tagging) {
    const input = el('input', { type: 'file', accept: 'image/*', multiple: true, className: 'hiddenInput', tabIndex: -1 });
    input.setAttribute('aria-label', 'Choose photos from your device');
    input.onchange = () => {
      rejected = [];
      Array.from(input.files || []).forEach(file => {
        if (!isImageFile(file)) { rejected.push(file.name); return; }
        photos.push({ file, preview: URL.createObjectURL(file), selected: false, tags: new Set(), savedTags: new Set() });
      });
      render();
    };
    root.append(input, button('Choose photos', () => input.click(), 'primary'));
    const mode = el('fieldset', { className: 'mode' }, [el('legend', { textContent: 'After uploading' })]);
    for (const [text, value] of [['No tags needed', false], ['Add tags on this page', true]] as const) {
      const radio = el('input', { type: 'radio', name: 'tagMode', checked: addTags === value, disabled: busy });
      radio.onchange = () => { addTags = value; render(); };
      mode.append(label(text, radio));
    }
    root.append(mode);
  }
  if (rejected.length) {
    root.append(el('div', { className: 'warning', role: 'status', textContent: `Skipped (not photos): ${rejected.join(', ')}` }));
  }
  renderPhotos();

  if (photos.length && !tagging) {
    const pending = photos.filter(photo => !ready(photo)).length;
    if (pending) {
      root.append(button(uploaded ? `Retry ${pending} incomplete photos` : `Upload ${photos.length} photos`,
        upload, 'primary', busy || !destination || !!destination.error));
    } else {
      if (addTags) { root.append(button('Continue to add tags', beginTagging, 'primary')); }
      else { root.append(button('Return to app Home', () => navigateBack(false), 'primary')); }
    }
  }
  if (tagging) { renderTagPanel(); }
  if (!photos.some(photo => photo.result)) {
    root.append(button('Cancel and return to app', () => navigateBack(true), 'link'));
  } else if (!busy && photos.some(photo => !ready(photo))) {
    root.append(button('Return with incomplete upload', () => {
      if (confirm('Some photos or Batch IDs failed. Return now without completing them?')) { navigateBack(false); }
    }, 'link'));
  }
}

function renderFolders(): void {
  const panel = el('section', { className: 'folderPanel' }, [el('h2', { textContent: 'Choose a SharePoint folder' })]);
  const crumbs = el('nav', { className: 'breadcrumbs' });
  crumbs.setAttribute('aria-label', 'SharePoint folder breadcrumb');
  crumbs.append(button('Documents', () => browse(libraryRoot), 'crumb'));
  let path = libraryRoot;
  for (const segment of browsePath.slice(libraryRoot.length).split('/').filter(Boolean)) {
    path += `/${segment}`;
    const target = path;
    crumbs.append(' / ', button(segment, () => browse(target), 'crumb'));
  }
  const list = el('ul', { className: 'folders' });
  folders.forEach(folder => list.append(el('li', {}, [
    button(folder.Name, () => browse(folder.ServerRelativeUrl), 'folder')
  ])));
  panel.append(crumbs, list);
  if (!folders.length && !busy) { panel.append(el('p', { textContent: 'No subfolders in this folder.' })); }
  panel.append(button('Use this folder', chooseDestination, 'secondary'));
  root.append(panel);
}

function renderPhotos(): void {
  const list = el('ul', { className: 'photos' });
  photos.forEach((photo, index) => {
    const image = el('img', { src: photo.preview, alt: photo.file.name, loading: 'lazy' });
    image.onerror = () => {
      image.replaceWith(el('div', { className: 'previewUnavailable', textContent: 'Preview unavailable on this device' }));
    };
    const card = el('li', { className: 'photoCard' }, [image, el('strong', { textContent: photo.file.name })]);
    if (!photo.result && !tagging && !busy) {
      card.append(button('Remove', () => {
        URL.revokeObjectURL(photo.preview); photos.splice(index, 1); render();
      }, 'smallButton'));
    }
    if (photo.result) {
      card.append(el('p', {
        className: ready(photo) ? 'success' : 'error',
        textContent: ready(photo) ? 'Uploaded - Batch ID saved' :
          photo.result.error || (photo.result.ok ? 'Uploaded, but Batch ID not saved.' : 'Upload failed.')
      }));
    }
    if (tagging && ready(photo)) {
      const check = el('input', { type: 'checkbox', checked: photo.selected, disabled: busy });
      check.onchange = () => { photo.selected = check.checked; render(); };
      card.append(label(`Select ${photo.file.name}`, check));
      const names = tags.filter(tag => photo.tags.has(tag.Id)).map(tag => tag.Tag);
      card.append(el('p', { textContent: names.length ? names.join(', ') : 'No tags selected' }));
      if (photo.tagError) { card.append(el('p', { className: 'error', role: 'alert', textContent: photo.tagError })); }
      else if (dirty(photo)) { card.append(el('p', { className: 'warning', textContent: 'Tags not yet saved' })); }
      else if (photo.savedTags.size) { card.append(el('p', { className: 'success', textContent: 'Tags saved' })); }
    }
    list.append(card);
  });
  root.append(list);
}

function renderTagPanel(): void {
  const panel = el('section', { className: 'tagPanel' }, [el('h2', { textContent: 'Choose tags' })]);
  if (tagsError) {
    panel.append(el('div', { className: 'error', role: 'alert', textContent: tagsError }),
      button('Retry loading tags', loadTags));
  }
  for (const category of [...new Set(tags.map(tag => tag.TagType))]) {
    const group = el('fieldset', { className: 'tagGroup' }, [el('legend', { textContent: category })]);
    tags.filter(tag => tag.TagType === category).forEach(tag => {
      const check = el('input', { type: 'checkbox', checked: chosenTags.has(tag.Id), disabled: busy });
      check.onchange = () => {
        if (check.checked) { chosenTags.add(tag.Id); } else { chosenTags.delete(tag.Id); }
        const disabled = busy || !chosenTags.size;
        allButton.disabled = disabled;
        selectedButton.disabled = disabled || !photos.some(photo => ready(photo) && photo.selected);
      };
      group.append(label(tag.Tag, check));
    });
    panel.append(group);
  }
  const apply = (all: boolean): void => {
    photos.filter(photo => ready(photo) && (all || photo.selected)).forEach(photo => {
      chosenTags.forEach(id => photo.tags.add(id));
    });
    render();
  };
  const allButton = button('Add chosen tags to all photos', () => apply(true), 'secondary', busy || !chosenTags.size || !!tagsError);
  const selectedButton = button('Add chosen tags to selected photos', () => apply(false), 'secondary',
    busy || !chosenTags.size || !!tagsError || !photos.some(photo => ready(photo) && photo.selected));
  panel.append(
    el('p', { textContent: 'Select one photo for individual tagging, or several for bulk tagging. Tags are added without removing existing tags.' }),
    allButton, selectedButton,
    button('Select all photos', () => { photos.forEach(photo => { photo.selected = ready(photo); }); render(); }, 'smallButton'),
    button('Clear photo selection', () => { photos.forEach(photo => { photo.selected = false; }); render(); }, 'smallButton')
  );
  const unsaved = photos.some(dirty);
  panel.append(button('Save tags', saveTags, 'primary', busy || !unsaved || !!tagsError));
  panel.append(button('Done - return to app Home', () => navigateBack(false), 'primary',
    busy || unsaved || photos.some(photo => !ready(photo) || !!photo.tagError)));
  if (unsaved) {
    panel.append(button('Discard unsaved tag additions', () => {
      if (!confirm('Discard tag additions that have not been saved? Uploaded photos and saved tags will be kept.')) { return; }
      photos.forEach(photo => { photo.tags = new Set(photo.savedTags); photo.tagError = undefined; });
      render();
    }, 'link'));
  }
  root.append(panel);
}

async function beginTagging(): Promise<void> {
  tagging = true; render();
  if (!tags.length) { await loadTags(); }
}

async function upload(): Promise<void> {
  if (!destination?.batchFieldInternalName || destination.error) { throw new Error('Choose and verify an upload destination first.'); }
  const verified = destination;
  busy = true; error = ''; render();
  const queue = photos.filter(photo => !ready(photo));
  const total = queue.length;
  let completed = 0;
  try {
    const worker = async (): Promise<void> => {
      let photo: Photo | undefined;
      while ((photo = queue.shift())) {
        try {
          if (photo.result?.ok && photo.result.serverRelativeUrl) {
            await service.retryBatchId(photo.result.serverRelativeUrl, batchId, verified.batchFieldInternalName!);
            photo.result.batchIdSet = true; photo.result.error = undefined;
          } else {
            photo.result = await service.uploadFile(verified.folderServerRelativeUrl, photo.file, batchId, verified.batchFieldInternalName);
          }
        } catch (e) {
          if (photo.result?.ok) { photo.result.error = errorMessage(e); }
          else { photo.result = { originalName: photo.file.name, ok: false, batchIdSet: false, error: errorMessage(e) }; }
        }
        message = `Uploaded/checked ${++completed} of ${total} photos...`; render();
      }
    };
    await Promise.all(Array.from({ length: Math.min(config.uploadConcurrency, total) }, worker));
  } finally { busy = false; message = ''; render(); }
  if (photos.every(ready)) {
    if (addTags) { await beginTagging(); } else { navigateBack(false); }
  }
}

async function saveTags(): Promise<void> {
  busy = true; error = ''; render();
  try {
    for (const photo of photos.filter(photo => ready(photo) && dirty(photo))) {
      try {
        if (!photo.result?.serverRelativeUrl) { throw new Error('The uploaded photo URL is missing.'); }
        await service.addTags(photo.result.serverRelativeUrl, tags.filter(tag => photo.tags.has(tag.Id)));
        photo.savedTags = new Set(photo.tags); photo.tagError = undefined;
      } catch (e) { photo.tagError = errorMessage(e); console.error(e); }
      render();
    }
  } finally { busy = false; render(); }
}

function navigateBack(cancelled: boolean): void {
  const uploaded = photos.filter(photo => photo.result?.ok).length;
  const partial = photos.some(photo => !ready(photo) || dirty(photo) || photo.tagError);
  const params: Record<string, string> = {
    ReturnTo: 'Home', UploadStatus: cancelled ? 'Cancelled' : partial ? 'Partial' : 'Complete',
    UploadCount: String(uploaded), TagStatus: photos.some(photo => photo.savedTags.size) ? 'Saved' : 'NotRequested',
    FolderPath: `${(destination?.folderServerRelativeUrl || libraryRoot).slice(new URL(config.siteUrl).pathname.length + 1)}/`
  };
  if (uploaded) { params[config.returnParamName] = batchId; }
  sessionStorage.removeItem(CONTEXT_KEY);
  location.href = appendParams(returnUrl, params);
}

window.addEventListener('beforeunload', event => {
  if (busy || photos.some(dirty) || (photos.some(photo => photo.result) && photos.some(photo => !ready(photo)))) {
    event.preventDefault();
  }
});

start().catch(showError);
