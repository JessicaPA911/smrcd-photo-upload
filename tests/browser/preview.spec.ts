import { test, expect } from '@playwright/test';

test('new-list preview groups categories, filters inactive tags, and selects without any writes', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.route(/\/src\/auth\.ts(?:\?.*)?$/, route => route.fulfill({
    contentType: 'application/javascript',
    body: `export async function ensureSignedIn(){return {account:{name:"Preview user"}}}
      export async function getSharePointToken(){return "mock"}
      export async function signOut(){}`
  }));
  const requests: { url: string; method: string }[] = [];
  await page.route('https://shanescows.sharepoint.com/**', route => {
    requests.push({ url: route.request().url(), method: route.request().method() });
    return route.fulfill({ contentType: 'application/json', body: JSON.stringify({ value: [
      { Id: 1, Title: 'Location', field_1: 'Coast', field_2: 'Where the photo was taken', field_3: true },
      { Id: 2, Title: 'People', field_1: 'Staff', field_3: true },
      { Id: 3, Title: 'Practices', field_1: 'Planting', field_3: true },
      { Id: 4, Title: 'Program Area', field_1: 'Restoration', field_3: true },
      { Id: 5, Title: 'Subject', field_1: 'Birds', field_3: true },
      { Id: 6, Title: 'Use', field_1: 'Reporting', field_3: true },
      { Id: 7, Title: 'Use', field_1: 'Inactive tag', field_3: false }
    ] }) });
  });
  await page.route('https://login.microsoftonline.com/**', route => route.abort());
  await page.goto('./?previewTags=1');
  await expect(page.getByRole('heading', { name: 'Add tags to your photos' })).toBeVisible();
  await expect(page.getByText(/Preview only:/)).toBeVisible();
  await expect(page.getByRole('checkbox')).toHaveCount(6);
  await expect(page.getByText('Inactive tag')).toHaveCount(0);
  await page.getByRole('checkbox', { name: 'Location: Coast' }).check();
  await page.getByRole('checkbox', { name: 'People: Staff' }).check();
  await expect(page.getByRole('heading', { name: 'Selected tags (2)' })).toBeVisible();
  await page.getByLabel('Find a tag').fill('practices');
  await expect(page.getByRole('checkbox')).toHaveCount(1);
  await expect(page.getByRole('checkbox', { name: 'Practices: Planting' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'Selected tags (2)' })).toBeVisible();
  await page.getByRole('button', { name: 'Remove Location: Coast' }).click();
  await expect(page.getByRole('heading', { name: 'Selected tags (1)' })).toBeVisible();
  await page.getByRole('button', { name: 'Clear selections' }).click();
  await expect(page.getByRole('heading', { name: 'Selected tags (0)' })).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  expect(requests).toHaveLength(1);
  expect(requests[0].method).toBe('GET');
  expect(requests[0].url).toContain('e5873870-d3fc-4ab8-96b4-e3d19a51f63e');
  await expect(page.getByRole('button', { name: 'Save tags' })).toHaveCount(0);
  expect(page.url()).toContain('previewTags=1');
});

test('preview shows read errors and provides a retry instead of sample-data fallback', async ({ page }) => {
  await page.route(/\/src\/auth\.ts(?:\?.*)?$/, route => route.fulfill({
    contentType: 'application/javascript',
    body: `export async function ensureSignedIn(){return {account:{name:"Preview user"}}}
      export async function getSharePointToken(){return "mock"}
      export async function signOut(){}`
  }));
  let denied = true;
  await page.route('https://shanescows.sharepoint.com/**', route => route.fulfill({
    status: denied ? 403 : 200, contentType: 'application/json',
    body: JSON.stringify(denied ? { error: { message: 'Cannot read tag list' } } : { value: [] })
  }));
  await page.goto('./?previewTags=1');
  await expect(page.getByRole('alert')).toHaveText('Cannot read tag list');
  denied = false;
  await page.getByRole('button', { name: 'Retry loading tags' }).click();
  await expect(page.getByText('No active tags are available in this list.')).toBeVisible();
});
