const IMAGE_EXTENSIONS: string[] = ['jpg', 'jpeg', 'png', 'gif', 'heic', 'heif', 'webp', 'bmp', 'tif', 'tiff'];

const BATCH_ID_PATTERN: RegExp = /^[A-Za-z0-9-]{4,64}$/;

/** Human-readable, sortable, collision-resistant batch ID, e.g. B-20261012-093623-K7Q2 */
export function createBatchId(now: Date = new Date()): string {
  const pad = (n: number): string => (n < 10 ? '0' : '') + n;
  const stamp: string =
    `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  const alphabet: string = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const rnd = new Uint32Array(4);
  crypto.getRandomValues(rnd);
  let suffix: string = '';
  rnd.forEach(v => { suffix += alphabet[v % alphabet.length]; });
  return `B-${stamp}-${suffix}`;
}

/** Accepts a caller-supplied batch ID only if it is a safe token; otherwise returns undefined. */
export function sanitizeBatchId(value: string | undefined): string | undefined {
  if (!value) { return undefined; }
  const trimmed: string = value.trim();
  return BATCH_ID_PATTERN.test(trimmed) ? trimmed : undefined;
}

export function getExtension(fileName: string): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx >= 0 ? fileName.substring(idx + 1).toLowerCase() : '';
}

/** Some platforms (e.g. HEIC on Windows) report an empty MIME type, so fall back to the extension. */
export function isImageFile(file: File): boolean {
  if (file.type && file.type.toLowerCase().indexOf('image/') === 0) { return true; }
  return !file.type && IMAGE_EXTENSIONS.indexOf(getExtension(file.name)) >= 0;
}

/** Strips characters SharePoint does not allow in file names. */
export function sanitizeFileName(fileName: string): string {
  let cleaned: string = fileName.replace(/["*:<>?/\\|#%]/g, '_').replace(/\s+/g, ' ').trim();
  cleaned = cleaned.replace(/^[.~ ]+/, '').replace(/[. ]+$/, '');
  return cleaned || 'photo.jpg';
}

/** photo.jpg + 2 => photo (2).jpg */
export function withCopySuffix(fileName: string, n: number): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx > 0
    ? `${fileName.substring(0, idx)} (${n})${fileName.substring(idx)}`
    : `${fileName} (${n})`;
}

/**
 * Only allows returning to Power Apps (web player or the mobile ms-apps: deep link),
 * or the configured default URL, to prevent this page being used as an open redirect.
 */
export function isAllowedReturnUrl(url: string, defaultReturnUrl: string): boolean {
  if (!url) { return false; }
  if (defaultReturnUrl && url.toLowerCase() === defaultReturnUrl.toLowerCase()) { return true; }
  if (url.toLowerCase().indexOf('ms-apps://') === 0) { return true; }
  try {
    const parsed: URL = new URL(url);
    const host: string = parsed.hostname.toLowerCase();
    return parsed.protocol === 'https:' &&
      (host === 'apps.powerapps.com' || host.endsWith('.powerapps.com') ||
       host === 'apps.gov.powerapps.us' || host.endsWith('.powerapps.us'));
  } catch {
    return false;
  }
}

/** Appends/replaces query params on a URL, including non-http schemes like ms-apps:/// */
export function appendParams(url: string, params: { [key: string]: string }): string {
  const hashIdx: number = url.indexOf('#');
  const hash: string = hashIdx >= 0 ? url.substring(hashIdx) : '';
  const base: string = hashIdx >= 0 ? url.substring(0, hashIdx) : url;
  const qIdx: number = base.indexOf('?');
  const path: string = qIdx >= 0 ? base.substring(0, qIdx) : base;
  const existing: string[] = qIdx >= 0 ? base.substring(qIdx + 1).split('&').filter(p => !!p) : [];
  const keys: string[] = Object.keys(params).map(k => k.toLowerCase());
  const kept: string[] = existing.filter(p => keys.indexOf(decodeURIComponent(p.split('=')[0]).toLowerCase()) < 0);
  const added: string[] = Object.keys(params).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`);
  return `${path}?${kept.concat(added).join('&')}${hash}`;
}
