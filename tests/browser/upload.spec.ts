import { test, expect, type Page } from '@playwright/test';

interface Write { upload?: string; batch?: unknown; tags?: object; etag?: string }

async function mockSite(page: Page, options: { failBatch?: boolean; failTagSave?: boolean } = {}): Promise<Write[]> {
  const writes: Write[] = [];
  await page.route(/\/src\/auth\.ts(?:\?.*)?$/, route => route.fulfill({
    contentType: 'application/javascript',
    body: `export async function ensureSignedIn(){return {account:{name:"Test user",username:"test@example.com"}}}
      export async function getSharePointToken(){return "mock"}
      export async function signOut(){}`
  }));
  await page.route('https://login.microsoftonline.com/**', route => route.abort());
  await page.route('https://shanescows.sharepoint.com/**', async route => {
    const req = route.request();
    const url = new URL(req.url());
    const path = url.pathname;
    const folder = '/sites/JessicasTestingSite/Shared Documents/Photos';
    let body: object;
    if (path.endsWith('/Folders')) {
      body = { value: url.searchParams.get('@f')?.includes('/Photos') ? [] : [{ Name: 'Photos', ServerRelativeUrl: folder }] };
    } else if (path.includes('/Fields/GetBy')) {
      body = { InternalName: 'BatchID', TypeAsString: 'Text' };
    } else if (path.endsWith('/fields')) {
      body = { value: ['LocationTag', 'UseTag', 'SubjectTag', 'ItemTags'].map(InternalName => ({
        InternalName, TypeAsString: 'LookupMulti', LookupList: '{431acb4f-c6f5-4479-8f98-558da4c6cf8d}'
      })) };
    } else if (path.endsWith('/items')) {
      body = { value: [{ Id: 1, Tag: 'Coast', TagType: 'Location' }, { Id: 2, Tag: 'Bird', TagType: 'Subject' }] };
    } else if (path.includes('/Files/AddUsingPath')) {
      const name = url.searchParams.get('@n')!.replace(/^'|'$/g, '');
      writes.push({ upload: name });
      body = { Name: name, ServerRelativeUrl: `${folder}/${name}` };
    } else if (path.includes('ValidateUpdateListItem')) {
      writes.push({ batch: req.postDataJSON() });
      body = { value: [{ HasException: options.failBatch ?? false, ErrorMessage: options.failBatch ? 'Batch write failed' : '' }] };
    } else if (path.endsWith('/ListItemAllFields')) {
      if (req.method() === 'POST') {
        writes.push({ tags: req.postDataJSON(), etag: req.headers()['if-match'] });
        if (options.failTagSave) {
          return route.fulfill({ status: 403, contentType: 'application/json', body: JSON.stringify({ error: { message: 'Tag save denied' } }) });
        }
        return route.fulfill({ status: 204 });
      }
      body = { d: { __metadata: { etag: '"1"', type: 'SP.Data.DocumentsItem' },
        LocationTagId: { results: [] }, UseTagId: { results: [] },
        SubjectTagId: { results: [] }, ItemTagsId: { results: [] } } };
    } else {
      body = { Exists: true, ServerRelativeUrl: folder };
    }
    await route.fulfill({ contentType: 'application/json', body: JSON.stringify(body) });
  });
  await page.route('https://apps.powerapps.com/**', route => route.fulfill({
    contentType: 'text/html', body: '<h1>Returned to app</h1>'
  }));
  await page.goto('./?upn=test@example.com');
  await page.getByRole('button', { name: 'Photos', exact: true }).click();
  await page.getByRole('button', { name: 'Use this folder', exact: true }).click();
  return writes;
}

const file = (name: string) => ({ name, mimeType: 'image/jpeg', buffer: Buffer.from([255, 216, 255]) });

test('upload, tag all, tag a single photo, save, and return Home with Batch ID', async ({ page }) => {
  const writes = await mockSite(page);
  await page.locator('input[type=file]').setInputFiles([file('one.jpg'), file('two.jpg')]);
  await page.getByLabel('Add tags on this page').check();
  await page.getByRole('button', { name: 'Upload 2 photos', exact: true }).click();
  await page.getByLabel('Coast', { exact: true }).check();
  await page.getByRole('button', { name: 'Add chosen tags to all photos', exact: true }).click();
  await page.getByLabel('Select one.jpg', { exact: true }).check();
  await page.getByLabel('Coast', { exact: true }).uncheck();
  await page.getByLabel('Bird', { exact: true }).check();
  await page.getByRole('button', { name: 'Add chosen tags to selected photos', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Done - return to app Home' })).toBeDisabled();
  await page.getByRole('button', { name: 'Save tags', exact: true }).click();
  await expect(page.getByText('Tags saved', { exact: true })).toHaveCount(2);
  await page.getByRole('button', { name: 'Done - return to app Home' }).click();
  await page.waitForURL('https://apps.powerapps.com/**');
  const returned = new URL(page.url()).searchParams;
  expect(returned.get('ReturnTo')).toBe('Home');
  expect(returned.get('BatchId')).toMatch(/^B-/);
  expect(returned.get('FolderPath')).toBe('Shared Documents/Photos/');
  expect(returned.get('TagStatus')).toBe('Saved');
  const tags = writes.filter(write => write.tags).map(write => write.tags);
  expect(tags).toEqual([
    { __metadata: { type: 'SP.Data.DocumentsItem' }, LocationTagId: { results: [1] },
      UseTagId: { results: [] }, SubjectTagId: { results: [2] }, ItemTagsId: { results: [1, 2] } },
    { __metadata: { type: 'SP.Data.DocumentsItem' }, LocationTagId: { results: [1] },
      UseTagId: { results: [] }, SubjectTagId: { results: [] }, ItemTagsId: { results: [1] } }
  ]);
});

test('no-tags upload returns Home automatically, still stamping Batch ID', async ({ page }) => {
  const writes = await mockSite(page);
  await page.locator('input[type=file]').setInputFiles(file('one.jpg'));
  await page.getByRole('button', { name: 'Upload 1 photos', exact: true }).click();
  await page.waitForURL('https://apps.powerapps.com/**');
  expect(writes.filter(write => write.batch)).toHaveLength(1);
  expect(writes.filter(write => write.tags)).toHaveLength(0);
  expect(new URL(page.url()).searchParams.get('TagStatus')).toBe('NotRequested');
  expect(new URL(page.url()).searchParams.get('UploadStatus')).toBe('Complete');
});

test('Batch ID failure prevents automatic return; retry only stamps metadata', async ({ page }) => {
  const options = { failBatch: true };
  const writes = await mockSite(page, options);
  await page.locator('input[type=file]').setInputFiles(file('one.jpg'));
  await page.getByRole('button', { name: 'Upload 1 photos', exact: true }).click();
  await expect(page.getByText(/Batch write failed/)).toBeVisible();
  expect(page.url()).toContain('127.0.0.1');
  options.failBatch = false;
  await page.getByRole('button', { name: 'Retry 1 incomplete photos' }).click();
  await page.waitForURL('https://apps.powerapps.com/**');
  expect(writes.filter(write => write.upload)).toHaveLength(1);
  expect(writes.filter(write => write.batch)).toHaveLength(2);
});

test('failed tag save remains on page with explicit error and retry', async ({ page }) => {
  const options = { failTagSave: true };
  await mockSite(page, options);
  await page.locator('input[type=file]').setInputFiles(file('one.jpg'));
  await page.getByLabel('Add tags on this page').check();
  await page.getByRole('button', { name: 'Upload 1 photos', exact: true }).click();
  await page.getByLabel('Coast', { exact: true }).check();
  await page.getByRole('button', { name: 'Add chosen tags to all photos' }).click();
  await page.getByRole('button', { name: 'Save tags', exact: true }).click();
  await expect(page.getByText('Tag save denied', { exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Done - return to app Home' })).toBeDisabled();
  options.failTagSave = false;
  await page.getByRole('button', { name: 'Save tags', exact: true }).click();
  await expect(page.getByText('Tags saved', { exact: true })).toBeVisible();
});

test('mobile folder and tagging controls fit a phone viewport', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await mockSite(page);
  await page.locator('input[type=file]').setInputFiles(file('one.jpg'));
  await page.getByLabel('Add tags on this page').check();
  await page.getByRole('button', { name: 'Upload 1 photos', exact: true }).click();
  await expect(page.getByLabel('Coast', { exact: true })).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});
