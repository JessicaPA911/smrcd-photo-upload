// Public, non-secret configuration. Change these values when moving to another tenant/site.
export const config = {
  tenantId: '9416a1d9-2b7f-4847-8223-18c2057145e3',
  clientId: 'e4f604b6-396c-4d0d-8a46-dea352d13220',
  /** Absolute URL of the SharePoint site that holds the upload folder. */
  siteUrl: 'https://shanescows.sharepoint.com/sites/JessicasTestingSite',
  /** Site-relative or server-relative path of the holding folder. */
  folderPath: 'Shared Documents/Upload Repository',
  /** Internal or display name of the text column that receives the Batch ID. Empty = don't tag. */
  batchFieldName: 'BatchID',
  /** Used when the canvas app doesn't pass ?returnUrl= */
  defaultReturnUrl:
    'https://apps.powerapps.com/play/e/9b58e755-0372-e61d-ba6e-bdf027ce651a/a/49da2fb8-8ca9-4039-bfef-05ebafaadc94?tenantId=9416a1d9-2b7f-4847-8223-18c2057145e3',
  returnParamName: 'BatchId',
  /** Return to the app automatically when every photo uploads successfully. */
  autoReturn: false,
  uploadConcurrency: 3
};

export const sharePointScope: string = `${new URL(config.siteUrl).origin}/AllSites.Write`;
