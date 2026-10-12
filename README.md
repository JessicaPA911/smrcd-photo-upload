# Photo upload and tagging (mock)

Live page: https://jessicapa911.github.io/smrcd-photo-upload/

## New tag-list design preview

Open https://jessicapa911.github.io/smrcd-photo-upload/?previewTags=1 to preview
**Photo Management Tags (2)**. Sign-in and read access to the list are required.
The page reads active items at runtime using the verified internal fields:
`Title` (category), `field_1` (tag), `field_2` (description), and `field_3` (Active).
Tag-list contents are not bundled into the public GitHub repository.

The preview groups choices by category, supports search and multiple selections,
and shows selected category/tag pairs. It performs no uploads, metadata writes,
or canvas navigation. Selections disappear on reload. The existing uploader below
continues using its original Tags list and lookup fields; no new-list integration
or migration has been applied.

A static page launched by a Power Apps canvas app. Users sign in, browse folders in
the configured SharePoint Documents library, choose an upload destination, and
multi-select photos from their devices. Every uploaded photo receives `BatchID`.
No files are stored in GitHub. The page calls SharePoint REST using the signed-in
user's delegated permissions; it introduces no premium Power Apps connector or flow.

## User workflow

1. Browse Documents using folder buttons and breadcrumbs. Click **Use this folder**.
2. Choose photos. Non-images are rejected; mobile HEIC files are supported, although
   a device may not support their thumbnail previews.
3. Choose **No tags needed** or **Add tags on this page** and upload.
4. No tags: return to app Home automatically after all files and Batch IDs succeed.
5. Add tags: photos upload first, then the page reads the private SharePoint Tags
   list. Choose categorized tags and apply them to all photos or selected photos.
   Selecting one photo supports individual tagging. Click **Save tags**, then
   **Done - return to app Home**.

Tags are additions: existing tags are preserved. Failed uploads, Batch ID updates,
and tag saves remain visible and retryable. A Batch ID retry never uploads the file
again. Duplicate names are renamed rather than overwritten. An explicit partial
return reports `UploadStatus=Partial`, not Complete.

## Configuration and permissions

Edit `src/config.ts` for tenant/client IDs, site, library path/ID, tag-list ID, batch
field name, and default return URL. These are public identifiers, not secrets.
Folder browsing is limited to the configured library; no folder creation is offered.

Users need read access to the Tags list and permission to upload/update metadata
on their selected folder. SharePoint enforces those permissions on every request.
The SPA uses MSAL redirect sign-in with delegated SharePoint `AllSites.Write`.

Current mock schema, verified against the test app and SharePoint:

| List/library | Fields |
|---|---|
| Documents | `BatchID` (single-line text, required by uploader) |
| Documents | `LocationTag`, `UseTag`, `SubjectTag`, `ItemTags` (multi-lookups to Tags) |
| Tags | `Tag` (text), `TagType` (choice: Location, Use, Subject) |

Tag saves write all four lookup-ID fields, mirroring the test app. The San Mateo
app has a different schema/indexing workflow; it needs a separate integration
before deployment there.

## Launch and return contract

Launch parameters: `returnUrl` (Power Apps web or native-player URL), `upn`
(optional login hint), `batchId` (optional existing batch). Launch context is
retained across initial sign-in redirects.

Return parameters:

| Parameter | Value |
|---|---|
| `ReturnTo` | `Home` |
| `BatchId` | Present when files were uploaded, even with no user-selected tags |
| `FolderPath` | Site-relative destination, e.g. `Shared Documents/Photos/` |
| `UploadCount` | Number of files uploaded |
| `UploadStatus` | `Complete`, `Partial`, or `Cancelled` |
| `TagStatus` | `Saved` or `NotRequested` |

`ReturnTo=Home` takes precedence over any older "BatchId means open tagging"
logic in a receiving app. Home in the test app is **TagPhotos**, its first screen.
The updated test-app draft refreshes Documents and selects the returned folder;
the existing Upload button launches this page.

Native-player return URLs include `restartApp=true` so launch parameters reload.
The launch formula uses `Host.OSType`; Safari on iOS also reports iOS and therefore
gets a native-player return URL. Actual device behavior still needs a signed-in test.

## Development and checks

`npm ci`, then `npm run dev`: http://localhost:5173/smrcd-photo-upload/

- `npm test`: upload helpers and mocked REST service contracts.
- `npm run test:ui`: isolated Edge browser checks for folder browsing, no-tags
  return, bulk/individual tags, failures/retries, and phone-width layout.
- `npm run build`: TypeScript and production bundle.

Browser tests mock sign-in and SharePoint; they do not replace a signed-in upload
test. GitHub Actions builds and deploys pushes to main.
