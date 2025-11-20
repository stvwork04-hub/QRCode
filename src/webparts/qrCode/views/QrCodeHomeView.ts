import { escape } from '@microsoft/sp-lodash-subset';
import styles from '../QrCodeWebPart.module.scss';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeHomeView {
  
  public static renderQRCodeView(container: Element, qrCodeUrl: string, userItem: IQRCodeItem): void {
    container.innerHTML = `
      <div class="${styles.qrCodeDisplay}">
        <div class="${styles.qrCodeImageContainer}">
          <img src="${qrCodeUrl}" alt="QR Code" class="${styles.qrCodeImage}" />
        </div>
        <div class="${styles.homeButtons}">
          <button type="button" id="editButton" class="${styles.iconButton}" title="Edit Details">
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
              <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
            </svg>
            <span class="${styles.buttonLabel}">Edit Details</span>
          </button>
          <button type="button" id="downloadQRButtonHome" class="${styles.iconButton}" title="Download QR Code">
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            <span class="${styles.buttonLabel}">Download</span>
          </button>
        </div>
      </div>
    `;
  }

  public static renderGeneratePrompt(container: Element): void {
    container.innerHTML = `
      <div class="${styles.generatePrompt}">
        <div class="${styles.promptIcon}">
          <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <rect x="3" y="3" width="7" height="7"></rect>
            <rect x="14" y="3" width="7" height="7"></rect>
            <rect x="14" y="14" width="7" height="7"></rect>
            <rect x="3" y="14" width="7" height="7"></rect>
          </svg>
        </div>
        <p>Would you like to generate your QR Code?</p>
        <button type="button" id="generateButton" class="${styles.iconButton}" title="Verify details and Generate QR Code">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M9 12l2 2 4-4"></path>
            <path d="M21 12c0 4.97-4.03 9-9 9s-9-4.03-9-9 4.03-9 9-9c2.12 0 4.07.74 5.61 1.98"></path>
          </svg>
          <span class="${styles.buttonLabel}">Verify details and Generate QR Code</span>
        </button>
      </div>
    `;
  }

  public static renderNoRecordMessage(container: Element, email: string): void {
    container.innerHTML = `
      <div class="${styles.noRecord}">
        <p>No record found for your email: <strong>${escape(email)}</strong></p>
        <p>Please contact your administrator to create a record for you in the DigitalBusinessCards list.</p>
      </div>
    `;
  }

  public static attachHomePageHandlers(
    domElement: HTMLElement, 
    onEdit: () => void, 
    onDownload: () => void
  ): void {
    const editButton = domElement.querySelector('#editButton');
    if (editButton) {
      editButton.addEventListener('click', onEdit);
    }

    const downloadButton = domElement.querySelector('#downloadQRButtonHome');
    if (downloadButton) {
      downloadButton.addEventListener('click', onDownload);
    }
  }

  public static attachGeneratePromptHandlers(
    domElement: HTMLElement, 
    onGenerate: () => void
  ): void {
    const generateButton = domElement.querySelector('#generateButton');
    if (generateButton) {
      generateButton.addEventListener('click', onGenerate);
    }
  }
}
