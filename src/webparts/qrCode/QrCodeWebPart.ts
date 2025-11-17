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
  private _userItem: IQRCodeItem | null = null;
  private _currentView: 'home' | 'edit' = 'home';
  private _qrCodeService: QrCodeService;

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.qrCode} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <h2>Your QR Code</h2>
      </div>
      <div id="contentContainer">
        <div id="loadingMessage">Loading your information...</div>
      </div>
    </section>`;

    this._loadUserData();
  }

  private _renderView(): void {
    if (this._currentView === 'home') {
      this._renderHomePage();
    } else {
      this._renderEditPage();
    }
  }

  private _renderHomePage(): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container || !this._userItem) return;

    void this._checkForAttachment();
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
      () => void this._downloadQRCode()
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
        async (title, firstName, lastName, phoneNumber, company, jobTitle) => {
          if (!this._userItem) return;
          await this._qrCodeService.updateItem(this._userItem.Id, title, firstName, lastName, phoneNumber, company, jobTitle);
        },
        () => this._switchToHomeView(),
        async () => await this._generateQRCode(),
        () => void this._downloadQRCode()
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

  private _renderNoRecordMessage(): void {
    const container = this.domElement.querySelector('#contentContainer');
    if (!container) return;

    QrCodeHomeView.renderNoRecordMessage(container, this._currentUserEmail);
  }

  private async _generateQRCode(): Promise<void> {
    if (!this._userItem) return;

    const generateButton = this.domElement.querySelector('#generateQRButton') as HTMLButtonElement;
    const saveMessage = this.domElement.querySelector('#saveMessage');
    
    if (!generateButton || !saveMessage) return;

    try {
      generateButton.disabled = true;
      saveMessage.innerHTML = '<span style="color: blue;">Requesting QR Code...</span>';

      await this._qrCodeService.requestQRCodeGeneration(this._userItem.Id);

      saveMessage.innerHTML = '<span style="color: green;">✓ Thank you, your QR code will be received via email</span>';
      
      // Keep the message visible
      setTimeout(() => {
        saveMessage.innerHTML = '';
      }, 5000);
    } catch (error) {
      console.error('Error generating QR Code:', error);
      saveMessage.innerHTML = `<span style="color: red;">Error requesting QR Code: ${error}</span>`;
    } finally {
      generateButton.disabled = false;
    }
  }

  private async _downloadQRCode(): Promise<void> {
    if (!this._userItem) return;

    try {
      const attachments = await this._qrCodeService.getAttachments(this._userItem.Id);
      
      if (attachments.length > 0) {
        this._qrCodeService.downloadAttachment(attachments[0]);
      } else {
        alert('No QR Code attachment found for this record.');
      }
    } catch (error) {
      console.error('Error downloading QR Code:', error);
      alert(`Error downloading QR Code: ${error}`);
    }
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
