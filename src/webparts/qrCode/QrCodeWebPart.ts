import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './QrCodeWebPart.module.scss';
import * as strings from 'QrCodeWebPartStrings';
import { IQRCodeItem } from './models/IQRCodeItem';
import { QrCodeService } from './services/QrCodeService';
import { QrCodeHomeView } from './views/QrCodeHomeView';
import { QrCodeEditView } from './views/QrCodeEditView';

export interface IQrCodeWebPartProps {
  description: string;
}

export default class QrCodeWebPart extends BaseClientSideWebPart<IQrCodeWebPartProps> {

  private _currentUserEmail: string = '';
  private _userItem: IQRCodeItem | undefined = undefined;
  private _currentView: 'home' | 'edit' = 'home';
  private _qrCodeService: QrCodeService;

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.qrCode} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <div class="${styles.profileImageContainer}" id="profileImageContainer">
          <div class="${styles.profileImagePlaceholder}">
            <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
              <circle cx="12" cy="7" r="4"></circle>
            </svg>
          </div>
        </div>
      </div>
      <div id="contentContainer">
        <div id="loadingMessage">Loading your information...</div>
      </div>
    </section>`;

    this._loadUserData().catch(console.error);
    this._loadUserProfilePhoto().catch(console.error);
  }

  private _renderView(): void {
    if (this._currentView === 'home') {
      this._renderHomePage();
    } else {
      this._renderEditPage().catch(console.error);
    }
  }

  private _renderHomePage(): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container || !this._userItem) return;

    this._checkForAttachment().catch(console.error);
  }

  private async _checkForAttachment(): Promise<void> {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container || !this._userItem) return;

    try {
      const attachments = await this._qrCodeService.getAttachments(this._userItem.Id);
      
      if (attachments.length > 0) {
        const attachment = attachments[0];
        const fileUrl = `https://tecq8.sharepoint.com/${attachment.ServerRelativeUrl}`;
        this._renderQRCodeView(fileUrl);
      } else {
        this._renderGeneratePrompt();
      }
    } catch (error) {
      console.error('Error checking attachments:', error);
      container.innerHTML = `<div style="color: red;">Error loading QR Code: ${error}</div>`;
    }
  }

  private _renderQRCodeView(qrCodeUrl: string): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container || !this._userItem) return;

    QrCodeHomeView.renderQRCodeView(container, qrCodeUrl, this._userItem);
    QrCodeHomeView.attachHomePageHandlers(
      this.domElement,
      () => this._switchToEditView(),
      () => { this._downloadQRCode().catch(console.error); }
    );
  }

  private _renderGeneratePrompt(): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container) return;

    QrCodeHomeView.renderGeneratePrompt(container);
    QrCodeHomeView.attachGeneratePromptHandlers(
      this.domElement,
      () => this._switchToEditView()
    );
  }

  private _switchToEditView(): void {
    this._currentView = 'edit';
    this._renderView();
  }

  private _switchToHomeView(): void {
    this._currentView = 'home';
    this._renderView();
  }

  private async _renderEditPage(): Promise<void> {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container || !this._userItem) return;

    try {
      const attachments = await this._qrCodeService.getAttachments(this._userItem.Id);
      const hasAttachment = attachments.length > 0;

      QrCodeEditView.renderForm(container, this._userItem, hasAttachment);
      QrCodeEditView.attachFormHandlers(
        this.domElement,
        async (formData) => {
          if (!this._userItem) return;
          
          console.log('🔧 DEBUG: Save and Generate QR Code - Starting...');
          
          // Step 1: Save the data to SharePoint
          console.log('🔧 DEBUG: Step 1 - Saving data to SharePoint');
          await this._qrCodeService.updateItem(this._userItem.Id, formData);
          
          // Update the local user item with the new data
          this._userItem.PhoneNumber = formData.PhoneNumber;
          this._userItem.MobilePhone = formData.MobilePhone;
          if (formData.Instagram !== undefined) this._userItem.Instagram = formData.Instagram;
          if (formData.Facebook !== undefined) this._userItem.Facebook = formData.Facebook;
          if (formData.Gmail !== undefined) this._userItem.Gmail = formData.Gmail;
          if (formData.OtherPhone !== undefined) this._userItem.OtherPhone = formData.OtherPhone;
          
          console.log('🔧 DEBUG: Step 2 - Data saved, now generating QR Code');
          
          // Step 2: Generate QR Code
          await this._generateQRCode();
          
          console.log('🔧 DEBUG: Save and Generate QR Code - Complete!');
        },
        () => this._switchToHomeView()
      );
    } catch (error) {
      console.error('Error rendering edit page:', error);
      container.innerHTML = `<div style="color: red;">Error loading form: ${error}</div>`;
    }
  }

  private async _loadUserData(): Promise<void> {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container) return;

    try {
      this._userItem = await this._qrCodeService.getUserItem(this._currentUserEmail);
      
      if (this._userItem) {
        console.log('User item found:', this._userItem);
        this._renderView();
      } else {
        console.log('No records found for email:', this._currentUserEmail);
        this._renderNoRecordMessage();
      }
    } catch (error) {
      console.error('Error loading data:', error);
      container.innerHTML = `<div style="color: red;">Error loading data: ${error}</div>`;
    }
  }

  private async _loadUserProfilePhoto(): Promise<void> {
    setTimeout(async () => {
      try {
        const profileImageContainer = this.domElement.querySelector('#profileImageContainer');
        if (!profileImageContainer) {
          console.log('Profile image container not found');
          return;
        }

        // Get current user's email
        const email = this.context.pageContext.user.email;
        console.log('Loading profile photo for:', email);
        
        // Create the SharePoint user photo URL
        const photoUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${encodeURIComponent(email)}`;
        
        console.log('Attempting to load photo from:', photoUrl);
        
        // Create image element
        const img = document.createElement('img');
        img.className = styles.profileImage;
        img.alt = 'Profile Photo';
        img.style.display = 'none'; // Hide initially
        
        img.onload = () => {
          console.log('Profile photo loaded successfully');
          profileImageContainer.innerHTML = '';
          profileImageContainer.appendChild(img);
          img.style.display = 'block';
        };
        
        img.onerror = () => {
          console.log('Profile photo failed to load, keeping placeholder');
          // Keep the existing placeholder
        };
        
        // Start loading the image
        img.src = photoUrl;
        
      } catch (error) {
        console.error('Error loading profile photo:', error);
        // Keep the placeholder on error
      }
    }, 1000); // Delay to ensure DOM is ready
  }

  private _renderNoRecordMessage(): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container) return;

    QrCodeHomeView.renderNoRecordMessage(container, this._currentUserEmail);
  }

  private async _generateQRCode(): Promise<void> {
    console.log('🎯 DEBUG: _generateQRCode method called');
    
    if (!this._userItem) {
      console.log('🎯 ERROR: No user item available');
      return;
    }

    console.log('🎯 DEBUG: User item ID:', this._userItem.Id);

    // Try to find the generate button (could be #generateButton from home or #generateQRButton from edit page)
    const generateButton = (this.domElement.querySelector('#generateButton') || this.domElement.querySelector('#generateQRButton')) as HTMLButtonElement;
    const saveMessage = this.domElement.querySelector('#saveMessage');
    const successMessage = (this.domElement.querySelector('#successMessage') || this.domElement.querySelector('#saveSuccessMessage')) as HTMLElement;
    
    console.log('🎯 DEBUG: generateButton found:', !!generateButton);
    console.log('🎯 DEBUG: saveMessage found:', !!saveMessage);
    console.log('🎯 DEBUG: successMessage found:', !!successMessage);
    
    if (!generateButton) {
      console.log('🎯 ERROR: Generate button not found');
      return;
    }

    try {
      console.log('🎯 DEBUG: Starting QR code generation process');
      generateButton.disabled = true;
      
      // Show loading message
      if (saveMessage) {
        saveMessage.innerHTML = '<span style="color: blue;">Requesting QR Code...</span>';
      }
      
      // Hide success message initially
      if (successMessage) {
        successMessage.style.display = 'none';
      }

      console.log('🎯 DEBUG: About to call requestQRCodeGeneration with ID:', this._userItem.Id);
      await this._qrCodeService.requestQRCodeGeneration(this._userItem.Id);
      console.log('🎯 DEBUG: requestQRCodeGeneration completed successfully');

      // Clear loading message
      if (saveMessage) {
        saveMessage.innerHTML = '';
      }
      
      // Show success message below buttons
      if (successMessage) {
        successMessage.innerHTML = 'QR Code generation requested successfully! Please close this page to view/download the QR image.';
        successMessage.style.display = 'block';
        console.log('🎯 DEBUG: Success message displayed');
        
        // Hide the message after 8 seconds
        setTimeout(() => {
          successMessage.style.display = 'none';
          console.log('🎯 DEBUG: Success message hidden after timeout');
        }, 8000);
      }
      
    } catch (error) {
      console.error('💥 ERROR in _generateQRCode:', error);
      console.error('💥 ERROR type:', typeof error);
      console.error('💥 ERROR message:', error.message);
      
      if (saveMessage) {
        saveMessage.innerHTML = `<span style="color: red;">Error requesting QR Code: ${error.message || error}</span>`;
      }
      
      // Hide success message on error
      if (successMessage) {
        successMessage.style.display = 'none';
      }
    } finally {
      generateButton.disabled = false;
      console.log('🎯 DEBUG: Generate button re-enabled');
    }
  }

  private async _downloadQRCode(): Promise<void> {
    if (!this._userItem) return;

    try {
      const attachments = await this._qrCodeService.getAttachments(this._userItem.Id);
      
      if (attachments.length > 0) {
        await this._downloadCompressedQRCode(attachments[0]);
      } else {
        alert('No QR Code attachment found for this record.');
      }
    } catch (error) {
      console.error('Error downloading QR Code:', error);
      alert(`Error downloading QR Code: ${error}`);
    }
  }

  private async _downloadCompressedQRCode(attachment: { ServerRelativeUrl: string; FileName: string }): Promise<void> {
    try {
      const fileUrl = `https://tecq8.sharepoint.com/${attachment.ServerRelativeUrl}`;
      await this._compressImageToPNG(fileUrl);
    } catch (error) {
      console.error('Error processing QR code:', error);
      // Fall back to original download method
      this._qrCodeService.downloadAttachment(attachment);
    }
  }

  private async _compressImageToPNG(imageUrl: string): Promise<void> {
    // Create a new image to load the QR code
    const img = new Image();
    img.crossOrigin = 'anonymous';
    
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error('Failed to load image'));
      img.src = imageUrl;
    });

      // Create canvas for compression
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      
      if (!ctx) {
        throw new Error('Canvas context not available');
      }

      // Calculate proper dimensions maintaining aspect ratio
      const maxSize = 512; // Higher resolution for better quality
      let canvasWidth = img.naturalWidth || img.width;
      let canvasHeight = img.naturalHeight || img.height;
      
      // Scale down if too large, maintaining aspect ratio
      if (canvasWidth > maxSize || canvasHeight > maxSize) {
        const scale = Math.min(maxSize / canvasWidth, maxSize / canvasHeight);
        canvasWidth = Math.round(canvasWidth * scale);
        canvasHeight = Math.round(canvasHeight * scale);
      }
      
      // Ensure minimum size for QR codes
      const minSize = 256;
      if (canvasWidth < minSize && canvasHeight < minSize) {
        const scale = Math.max(minSize / canvasWidth, minSize / canvasHeight);
        canvasWidth = Math.round(canvasWidth * scale);
        canvasHeight = Math.round(canvasHeight * scale);
      }
      
      canvas.width = canvasWidth;
      canvas.height = canvasHeight;
      
      // Set white background
      ctx.fillStyle = 'white';
      ctx.fillRect(0, 0, canvasWidth, canvasHeight);

      // Draw image on canvas maintaining aspect ratio
      ctx.drawImage(img, 0, 0, canvasWidth, canvasHeight);

      // Convert to PNG blob
      canvas.toBlob((blob) => {
        if (blob) {
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = `QR_Code_${this._userItem?.FirstName || 'User'}.png`;
          link.style.display = 'none';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
        }
      }, 'image/png', 0.9);
  }

  protected onInit(): Promise<void> {
    this._currentUserEmail = this.context.pageContext.user.email;
    this._qrCodeService = new QrCodeService(this.context.spHttpClient);
    return Promise.resolve();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
