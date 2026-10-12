import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ts from 'typescript';

const modules = new Map();
async function moduleUrl(name) {
  if (modules.has(name)) { return modules.get(name); }
  const source = await readFile(new URL(`../src/${name}.ts`, import.meta.url), 'utf8');
  let { outputText } = ts.transpileModule(source, {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext }
  });
  const imports = [...outputText.matchAll(/from '(\.\/([^']+))'/g)];
  for (const [, specifier, dependency] of imports) {
    outputText = outputText.replace(`from '${specifier}'`, `from '${await moduleUrl(dependency)}'`);
  }
  const url = `data:text/javascript;base64,${Buffer.from(outputText).toString('base64')}`;
  modules.set(name, url);
  return url;
}
const { SharePointService } = await import(await moduleUrl('sharePoint'));
const { config } = await import(await moduleUrl('config'));
const { tagIds, mergeTagIds } = await import(await moduleUrl('tagging'));
const { isImageFile, appendParams } = await import(await moduleUrl('uploadUtils'));
const service = () => new SharePointService(config.siteUrl, async () => 'test-token');
const folder = '/sites/JessicasTestingSite/Shared Documents/Photos';
const json = (body, status = 200, headers = {}) => new Response(JSON.stringify(body), {
  status, headers: { 'Content-Type': 'application/json', ...headers }
});
const originalFetch = globalThis.fetch;
test.afterEach(() => { globalThis.fetch = originalFetch; });

test('only photos pass file selection, including HEIC without a MIME type', () => {
  assert.equal(isImageFile(new File(['x'], 'a.heic')), true);
  assert.equal(isImageFile(new File(['x'], 'a.txt', { type: 'text/plain' })), false);
  assert.equal(isImageFile(new File(['x'], 'a.jpg', { type: 'text/plain' })), false);
});
test('tagging matches all four test-app lookup fields and adds without losing existing tags', () => {
  const added = tagIds([
    { Id: 1, Tag: 'Coast', TagType: 'Location' }, { Id: 2, Tag: 'Bird', TagType: 'Subject' },
    { Id: 3, Tag: 'Report', TagType: 'Use' }
  ]);
  assert.deepEqual(added, { LocationTag: [1], UseTag: [3], SubjectTag: [2], ItemTags: [1, 2, 3] });
  assert.deepEqual(mergeTagIds({ LocationTag: [8], UseTag: [], SubjectTag: [2], ItemTags: [8, 2] }, added),
    { LocationTag: [8, 1], UseTag: [3], SubjectTag: [2], ItemTags: [8, 2, 1, 3] });
});
test('folder boundary rejects siblings and unrelated libraries', () => {
  assert.doesNotThrow(() => service().assertLibraryFolder(folder));
  assert.throws(() => service().assertLibraryFolder('/sites/JessicasTestingSite/Shared Documents Other'));
  assert.throws(() => service().assertLibraryFolder('/sites/other/Shared Documents'));
});
test('folder browser follows pagination, sorts children, and hides Forms', async () => {
  let calls = 0;
  globalThis.fetch = async () => ++calls === 1 ?
    json({ value: [{ Name: 'Z', ServerRelativeUrl: `${folder}/Z` }, { Name: 'Forms', ServerRelativeUrl: `${folder}/Forms` }],
      'odata.nextLink': `${config.siteUrl}/_api/web/folders?next=1` }) :
    json({ value: [{ Name: 'A', ServerRelativeUrl: `${folder}/A` }] });
  assert.deepEqual((await service().listFolders(folder)).map(f => f.Name), ['A', 'Z']);
  assert.equal(calls, 2);
});
test('pagination cannot send bearer tokens to another site', async () => {
  globalThis.fetch = async () => json({ value: [], 'odata.nextLink': 'https://example.com/_api/items' });
  await assert.rejects(service().listFolders(folder), /outside the configured site/);
});
test('root-folder preflight checks BatchID on the library, not root ListItemAllFields', async () => {
  const urls = [];
  globalThis.fetch = async url => {
    urls.push(url);
    return urls.length === 1 ? json({ Exists: true, ServerRelativeUrl: '/sites/JessicasTestingSite/Shared Documents' })
      : json({ InternalName: 'BatchID', TypeAsString: 'Text' });
  };
  const result = await service().preflight(config.libraryPath, 'BatchID');
  assert.equal(result.batchFieldInternalName, 'BatchID');
  assert.match(urls[1], /lists\(guid/);
  assert.doesNotMatch(urls[1], /ListItemAllFields/);
});
test('missing BatchID blocks uploads rather than silently uploading untracked photos', async () => {
  globalThis.fetch = async url => url.includes('/Fields/') ? json({ error: { message: 'Missing column' } }, 404)
    : json({ Exists: true, ServerRelativeUrl: folder });
  assert.match((await service().preflight(folder, 'BatchID')).error, /required column/);
  await assert.rejects(service().uploadFile(folder, new File(['x'], 'a.jpg'), 'batch'), /not been verified/);
});
test('upload conflict renames the file, then stamps BatchID', async () => {
  const urls = [];
  globalThis.fetch = async (url, init) => {
    urls.push(url);
    if (urls.length === 1) { return json({ error: { message: 'File already exists' } }, 409); }
    if (urls.length === 2) { return json({ Name: 'a (2).jpg', ServerRelativeUrl: `${folder}/a (2).jpg` }); }
    assert.deepEqual(JSON.parse(init.body).formValues, [{ FieldName: 'BatchID', FieldValue: 'B-test' }]);
    return json({ value: [{ HasException: false }] });
  };
  const result = await service().uploadFile(folder, new File(['x'], 'a.jpg'), 'B-test', 'BatchID');
  assert.equal(result.batchIdSet, true);
  assert.equal(result.uploadedName, 'a (2).jpg');
});
test('unconfirmed BatchID is an error; retry only writes metadata', async () => {
  const urls = [];
  globalThis.fetch = async url => {
    urls.push(url);
    return json({ value: [] });
  };
  await assert.rejects(service().retryBatchId(`${folder}/a.jpg`, 'B-test', 'BatchID'), /did not confirm/);
  assert.equal(urls.length, 1);
  assert.match(urls[0], /ValidateUpdateListItem/);
});
test('saving tags preserves current lookups and uses an ETag to avoid lost edits', async () => {
  let calls = 0;
  globalThis.fetch = async (url, init) => {
    calls++;
    if (calls === 1) {
      assert.match(init.headers.Accept, /verbose/);
      return json({ d: { __metadata: { etag: '"4"', type: 'SP.Data.DocumentsItem' },
        LocationTagId: { results: [9] }, UseTagId: { results: [] }, SubjectTagId: { results: [] }, ItemTagsId: { results: [9] } } });
    }
    assert.equal(init.headers['IF-MATCH'], '"4"');
    assert.deepEqual(JSON.parse(init.body), {
      __metadata: { type: 'SP.Data.DocumentsItem' },
      LocationTagId: { results: [9, 1] }, UseTagId: { results: [] },
      SubjectTagId: { results: [] }, ItemTagsId: { results: [9, 1] }
    });
    return new Response(null, { status: 204 });
  };
  await service().addTags(`${folder}/a.jpg`, [{ Id: 1, Tag: 'Coast', TagType: 'Location' }]);
  assert.equal(calls, 2);
});
test('concurrent tag edit and access denial are surfaced', async () => {
  globalThis.fetch = async (url, init) => init.method === 'POST' ? json({}, 412) :
    json({ d: { __metadata: { etag: '"4"', type: 'SP.Data.DocumentsItem' },
      LocationTagId: { results: [] }, UseTagId: { results: [] },
      SubjectTagId: { results: [] }, ItemTagsId: { results: [] } } });
  await assert.rejects(service().addTags(`${folder}/a.jpg`, []), /changed while saving/);
  globalThis.fetch = async () => json({ error: { message: 'Access denied' } }, 403);
  await assert.rejects(service().listFolders(folder), /Access denied/);
});
test('return parameters preserve native-player restart and BatchId', () => {
  const returned = new URL(appendParams('ms-apps:///providers/Microsoft.PowerApps/apps/test?restartApp=true',
    { BatchId: 'B-test', ReturnTo: 'Home', FolderPath: 'Shared Documents/Photos/' }));
  assert.equal(returned.searchParams.get('restartApp'), 'true');
  assert.equal(returned.searchParams.get('BatchId'), 'B-test');
  assert.equal(returned.searchParams.get('ReturnTo'), 'Home');
});
