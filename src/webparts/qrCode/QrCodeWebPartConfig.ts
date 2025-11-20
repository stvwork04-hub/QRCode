export interface IQrCodeWebPartConfig {
  siteUrl: string;
  listName: string;
  powerAutomateUrl: string;
}

export const QrCodeWebPartConfig: IQrCodeWebPartConfig = {
  siteUrl: 'https://tecq8.sharepoint.com/sites/DigitalBusinessCard',
  listName: 'DigitalBusinessCards',
  powerAutomateUrl: 'https://0d8a5cb67bd747f6b7d0905b392268.d1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e015ead33fae4e3fb92520c44abcd344/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iwE0bpVAv5yngCMksvxC9BGS5kyC6frA8tuIH5Eh3JE'
};
