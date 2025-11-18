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

  public async updateItem(itemId: number, updatedFields: Partial<IQRCodeItem>): Promise<boolean> {
    try {
      const updateUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})`;
      
      const body = JSON.stringify(updatedFields);

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
      console.log('🔧 DEBUG: Starting QR Code generation request');
      console.log('🔧 DEBUG: Item ID:', itemId);
      console.log('🔧 DEBUG: Site URL:', QrCodeWebPartConfig.siteUrl);
      console.log('🔧 DEBUG: List Name:', QrCodeWebPartConfig.listName);
      
      const updateUrl = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})`;
      console.log('🔧 DEBUG: Update URL:', updateUrl);
      
      const updateData = {
        GenerateQRCode: true
      };
      console.log('🔧 DEBUG: Update data:', updateData);
      
      const bodyString = JSON.stringify(updateData);
      console.log('🔧 DEBUG: Body string:', bodyString);

      const headers = {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      };
      console.log('🔧 DEBUG: Headers:', headers);

      console.log('🔧 DEBUG: Making POST request...');
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: headers,
          body: bodyString
        }
      );

      console.log('🔧 DEBUG: Response received');
      console.log('🔧 DEBUG: Response status:', response.status);
      console.log('🔧 DEBUG: Response statusText:', response.statusText);
      console.log('🔧 DEBUG: Response ok:', response.ok);

      if (response.ok) {
        console.log('✅ SUCCESS: GenerateQRCode field updated successfully!');
        return true;
      } else {
        console.log('❌ ERROR: Response not OK, getting error details...');
        const errorText = await response.text();
        console.error('❌ ERROR: Response status:', response.status);
        console.error('❌ ERROR: Response statusText:', response.statusText);
        console.error('❌ ERROR: Error details:', errorText);
        
        // Try to parse error details
        try {
          const errorJson = JSON.parse(errorText);
          console.error('❌ ERROR: Parsed error:', errorJson);
        } catch (parseError) {
          console.error('❌ ERROR: Could not parse error JSON');
        }
        
        throw new Error(`Failed to update GenerateQRCode: ${response.status} ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      console.error('💥 EXCEPTION: Error in requestQRCodeGeneration:', error);
      console.error('💥 EXCEPTION: Error type:', typeof error);
      console.error('💥 EXCEPTION: Error message:', error.message);
      console.error('💥 EXCEPTION: Error stack:', error.stack);
      throw error;
    }
  }

  public async getItemById(itemId: number): Promise<IQRCodeItem | null> {
    try {
      const url = `${QrCodeWebPartConfig.siteUrl}/_api/web/lists/getbytitle('${QrCodeWebPartConfig.listName}')/items(${itemId})?$select=Id,FirstName,LastName,PhoneNumber,MobilePhone,Instagram,Facebook,Gmail,OtherPhone`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        return data;
      }
      return null;
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
