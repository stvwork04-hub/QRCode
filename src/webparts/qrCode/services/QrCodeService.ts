import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { QrCodeWebPartConfig } from '../QrCodeWebPartConfig';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeService {
  private spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
  }

  public async getUserItem(email: string): Promise<IQRCodeItem | undefined> {
    try {
      // Try exact match first (case-sensitive)
      let listUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items?$filter=Email eq '${encodeURIComponent(email)}'&$select=Id,Title,FirstName,LastName,Email,PhoneNumber,Company,JobTitle,MobilePhone,Instagram,Facebook,Gmail,OtherPhone,AttachmentFiles&$expand=AttachmentFiles`;
      
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
        listUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items?$select=Id,Title,FirstName,LastName,Email,PhoneNumber,Company,JobTitle,MobilePhone,Instagram,Facebook,Gmail,OtherPhone,AttachmentFiles&$expand=AttachmentFiles`;
        
        response = await this.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1
        );

        data = await response.json();
        
        if (data.value && data.value.length > 0) {
          // Filter client-side with case-insensitive comparison
          const matchedItem = data.value.find((item: { Email?: string }) => 
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
      
      return undefined;
    } catch (error) {
      console.error('Error loading data:', error);
      throw error;
    }
  }

  public async getAttachments(itemId: number): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
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

  public async updateItem(itemId: number, updatedFields: Partial<IQRCodeItem>): Promise<boolean> {
    try {
      console.log('🔧 DEBUG: updateItem called with:', { itemId, updatedFields });
      const updateUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})`;
      
      const body = JSON.stringify(updatedFields);
      console.log('🔧 DEBUG: Request body:', body);

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
      console.log('🔧 DEBUG: Starting QR Code generation request via HTTP endpoint');
      console.log('🔧 DEBUG: Item ID:', itemId);
      
      // Fetch the item details to get the Name
      const item = await this.getItemById(itemId);
      
      if (!item) {
        throw new Error('Failed to fetch item details');
      }
      
      const name = `${item.FirstName || ''} ${item.LastName || ''}`.trim();
      console.log('🔧 DEBUG: User Name:', name);
      
      const payload = {
        ListID: itemId.toString()
      };
      
      console.log('🔧 DEBUG: itemId type:', typeof itemId, 'value:', itemId);
      console.log('🔧 DEBUG: listIdNumber type:', typeof itemId, 'value:', itemId);
      console.log('🔧 DEBUG: Payload:', JSON.stringify(payload, null, 2));
      console.log('🔧 DEBUG: Power Automate URL:', QrCodeWebPartConfig.powerAutomateUrl);
      
      try {
        // Make HTTP POST request to Power Automate endpoint
        const response = await fetch(QrCodeWebPartConfig.powerAutomateUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(payload)
        });
        
        console.log('🔧 DEBUG: Response received');
        console.log('🔧 DEBUG: Response status:', response.status);
        console.log('🔧 DEBUG: Response statusText:', response.statusText);
        console.log('🔧 DEBUG: Response ok:', response.ok);
        
        // Try to get response body for debugging
        let responseText = '';
        try {
          responseText = await response.text();
          console.log('🔧 DEBUG: Response body:', responseText);
        } catch {
          console.log('🔧 DEBUG: Could not read response body');
        }
        
        if (response.ok || response.status === 202) {
          console.log('✅ SUCCESS: QR Code generation request sent successfully!');
          return true;
        } else {
          console.error('❌ ERROR: Response status:', response.status);
          console.error('❌ ERROR: Response statusText:', response.statusText);
          console.error('❌ ERROR: Response body:', responseText);
          
          // Provide more helpful error message
          let errorMessage = `Failed to request QR Code generation: ${response.status} ${response.statusText}`;
          if (responseText) {
            try {
              const errorJson = JSON.parse(responseText);
              if (errorJson.error && errorJson.error.message) {
                errorMessage += ` - ${errorJson.error.message}`;
              }
            } catch {
              errorMessage += ` - ${responseText}`;
            }
          }
          
          throw new Error(errorMessage);
        }
      } catch (fetchError) {
        console.error('💥 FETCH ERROR:', fetchError);
        throw fetchError;
      }
    } catch (error) {
      console.error('💥 EXCEPTION: Error in requestQRCodeGeneration:', error);
      console.error('💥 EXCEPTION: Error message:', error.message);
      throw error;
    }
  }

  public async getItemById(itemId: number): Promise<IQRCodeItem | undefined> {
    try {
      const url = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})?$select=Id,FirstName,LastName,PhoneNumber,MobilePhone,Instagram,Facebook,Gmail,OtherPhone,Company,JobTitle,Email`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        return data;
      }
      return undefined;
    } catch (error) {
      console.error('Error fetching item by ID:', error);
      throw error;
    }
  }

  public generateVCardData(item: IQRCodeItem): string {
    // Generate complete vCard 3.0 format with all fields
    let vCard = 'BEGIN:VCARD\n';
    vCard += 'VERSION:3.0\n';
    
    // Full name
    if (item.FirstName || item.LastName) {
      vCard += `FN:${(item.FirstName || '').trim()} ${(item.LastName || '').trim()}`.trim() + '\n';
      vCard += `N:${(item.LastName || '').trim()};${(item.FirstName || '').trim()};;;\n`;
    }
    
    // Phone numbers
    if (item.PhoneNumber) {
      vCard += `TEL;TYPE=WORK,VOICE:${item.PhoneNumber}\n`;
    }
    
    if (item.MobilePhone) {
      vCard += `TEL;TYPE=CELL:${item.MobilePhone}\n`;
    }
    
    if (item.OtherPhone) {
      vCard += `TEL;TYPE=HOME,VOICE:${item.OtherPhone}\n`;
    }
    
    // Social media as phone entries (for better contact app integration)
    if (item.Facebook) {
      vCard += `TEL;TYPE=X-FACEBOOK:${item.Facebook}\n`;
    }
    
    if (item.Instagram) {
      vCard += `TEL;TYPE=X-INSTAGRAM:${item.Instagram}\n`;
    }
    
    // Email
    if (item.Gmail) {
      vCard += `EMAIL;TYPE=INTERNET:${item.Gmail}\n`;
    }
    
    // Social media URLs
    if (item.Instagram) {
      const instagramUrl = item.Instagram.indexOf('http') === 0 ? item.Instagram : `https://instagram.com/${item.Instagram.replace('@', '')}`;
      vCard += `URL;TYPE=Instagram:${instagramUrl}\n`;
    }
    
    if (item.Facebook) {
      const facebookUrl = item.Facebook.indexOf('http') === 0 ? item.Facebook : `https://facebook.com/${item.Facebook}`;
      vCard += `URL;TYPE=Facebook:${facebookUrl}\n`;
    }
    
    // Organization (you can add this field if needed)
    // vCard += `ORG:Your Company Name\n`;
    
    // Notes with comprehensive social media and contact info
    const socialInfo = [];
    if (item.Instagram) {
      const instagramHandle = item.Instagram.indexOf('@') === 0 ? item.Instagram : `@${item.Instagram}`;
      socialInfo.push(`Instagram: ${instagramHandle}`);
    }
    if (item.Facebook) {
      socialInfo.push(`Facebook: ${item.Facebook}`);
    }
    if (item.MobilePhone) {
      socialInfo.push(`Mobile: ${item.MobilePhone}`);
    }
    if (item.OtherPhone) {
      socialInfo.push(`Other Phone: ${item.OtherPhone}`);
    }
    
    if (socialInfo.length > 0) {
      vCard += `NOTE:Contact Info - ${socialInfo.join(' | ')}\n`;
    }
    
    vCard += 'END:VCARD';
    
    return vCard;
  }



  public downloadAttachment(attachment: { FileName: string; ServerRelativeUrl: string }): void {
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
