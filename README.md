# SMRCD Photo Upload (mock)

Static page (GitHub Pages) that lets a Power Apps canvas app user multi-select photos and upload them to a SharePoint folder, tagging each with a Batch ID, then returns to the app with `?BatchId=...`.

- Auth: MSAL.js (redirect) with delegated SharePoint `AllSites.Write`; uses the signed-in user's own SharePoint permissions.
- Config: `src/config.ts` (tenant, client ID, site, folder, column, default return URL).
- Query params: `returnUrl` (Power Apps URL only), `batchId` (optional, reuse a batch), `upn` (optional login hint).
- Dev: `npm install`, `npm run dev`, then open http://localhost:5173/smrcd-photo-upload/
