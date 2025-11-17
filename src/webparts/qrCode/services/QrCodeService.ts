import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { QrCodeWebPartConfig } from '../QrCodeWebPartConfig';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeService {
  private spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
  }

  public async getUserItem(email: string): Promise<IQRCodeItem | null> {
    try {
      // Try exact match first (case-sensitive)
      let listUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items?$filter=Email eq '${encodeURIComponent(email)}'&$select=Id,Title,FirstName,LastName,Email,PhoneNumber,Company,JobTitle,QRCodeURL,ContactID,AttachmentFiles&$expand=AttachmentFiles`;
      
      console.log('Fetching data from:', listUrl);
      console.log('User email:', email);
      
      let response: SPHttpClientResponse = await this.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      let data = await response.json();
      console.log('Response data:', data);
      
      // If no exact match, try case-insensitive search
      if (!data.value || data.value.length === 0) {
        console.log('No exact match found, trying case-insensitive search...');
        
        // Get all items and filter client-side
        listUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items?$select=Id,Title,FirstName,LastName,Email,PhoneNumber,Company,JobTitle,QRCodeURL,ContactID,AttachmentFiles&$expand=AttachmentFiles`;
        
        response = await this.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1
        );

        data = await response.json();
        
        if (data.value && data.value.length > 0) {
          // Filter client-side with case-insensitive comparison
          const matchedItem = data.value.find((item: any) => 
            item.Email && item.Email.toLowerCase() === email.toLowerCase()
          );
          
          if (matchedItem) {
            console.log('Found matching item (case-insensitive):', matchedItem);
            return matchedItem;
          }
        }
      } else {
        return data.value[0];
      }
      
      return null;
    } catch (error) {
      console.error('Error loading data:', error);
      throw error;
    }
  }

  public async getAttachments(itemId: number): Promise<any[]> {
    try {
      const attachmentsUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})/AttachmentFiles`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        attachmentsUrl,
        SPHttpClient.configurations.v1
      );

      const data = await response.json();
      return data.value || [];
    } catch (error) {
      console.error('Error fetching attachments:', error);
      throw error;
    }
  }

  public async updateItem(itemId: number, title: string, firstName: string, lastName: string, phoneNumber: string, company: string, jobTitle: string): Promise<boolean> {
    try {
      const updateUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})`;
      
      const body = JSON.stringify({
        Title: title,
        FirstName: firstName,
        LastName: lastName,
        PhoneNumber: phoneNumber,
        Company: company,
        JobTitle: jobTitle
      });

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error updating item:', error);
      throw error;
    }
  }

  public async requestQRCodeGeneration(itemId: number): Promise<boolean> {
    try {
      const updateUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})`;
      
      const body = JSON.stringify({
        GenerateQRCode: true
      });

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error requesting QR code generation:', error);
      throw error;
    }
  }

  public downloadAttachment(attachment: any): void {
    const fileUrl = `https://tecq8.sharepoint.com/${attachment.ServerRelativeUrl}`;
    console.log('Downloading file from:', fileUrl);
    
    const link = document.createElement('a');
    link.href = fileUrl;
    link.download = attachment.FileName;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}
