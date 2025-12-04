export interface IQrCodeWebPartConfig {
  siteUrl: string;
  listName: string;
  powerAutomateUrl: string;
}

export const QrCodeWebPartConfig: IQrCodeWebPartConfig = {
  siteUrl: 'https://tecq8.sharepoint.com/sites/DigitalBusinessCard',
  listName: 'DigitalBusinessCards',
  powerAutomateUrl: 'https://0d8a5cb67bd747f6b7d0905b392268.d1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b89862b78ac4460c92927cde489001f7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7E7z2lqp6CU1uGgh9eOcZ_74K-852GZQc4lqvyo0Z2k'
};
