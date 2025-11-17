import { escape } from '@microsoft/sp-lodash-subset';
import styles from '../QrCodeWebPart.module.scss';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeEditView {
  
  public static renderForm(container: Element, userItem: IQRCodeItem, hasAttachment: boolean): void {
    container.innerHTML = `
      <div class="${styles.formContainer}">
        <form id="qrCodeForm">
          <div class="${styles.formField}">
            <label for="title">Title: *</label>
            <input type="text" id="title" name="title" value="${escape(userItem.Title || '')}" required />
          </div>
          
          <div class="${styles.formField}">
            <label for="firstName">First Name: *</label>
            <input type="text" id="firstName" name="firstName" value="${escape(userItem.FirstName || '')}" required />
          </div>
          
          <div class="${styles.formField}">
            <label for="lastName">Last Name: *</label>
            <input type="text" id="lastName" name="lastName" value="${escape(userItem.LastName || '')}" required />
          </div>
          
          <div class="${styles.formField}">
            <label for="email">Email: *</label>
            <input type="email" id="email" name="email" value="${escape(userItem.Email || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="phoneNumber">Phone Number: *</label>
            <input type="tel" id="phoneNumber" name="phoneNumber" value="${escape(userItem.PhoneNumber || '')}" required />
          </div>
          
          <div class="${styles.formField}">
            <label for="company">Company:</label>
            <input type="text" id="company" name="company" value="${escape(userItem.Company || '')}" />
          </div>
          
          <div class="${styles.formField}">
            <label for="jobTitle">Job Title:</label>
            <input type="text" id="jobTitle" name="jobTitle" value="${escape(userItem.JobTitle || '')}" />
          </div>
          
          <div class="${styles.formField} ${styles.buttonGroup}">
            <button type="submit" id="saveButton" class="${styles.iconButton}" title="Save">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path>
                <polyline points="17 21 17 13 7 13 7 21"></polyline>
                <polyline points="7 3 7 8 15 8"></polyline>
              </svg>
              <span class="${styles.buttonLabel}">Save</span>
            </button>
            <button type="button" id="generateQRButton" class="${styles.iconButton}" title="Generate QR Code">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <rect x="3" y="3" width="7" height="7"></rect>
                <rect x="14" y="3" width="7" height="7"></rect>
                <rect x="14" y="14" width="7" height="7"></rect>
                <rect x="3" y="14" width="7" height="7"></rect>
              </svg>
              <span class="${styles.buttonLabel}">Generate</span>
            </button>
            ${hasAttachment ? `
            <button type="button" id="downloadQRButton" class="${styles.iconButton}" title="Download QR Code">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              <span class="${styles.buttonLabel}">Download</span>
            </button>
            ` : ''}
            <button type="button" id="closeButton" class="${styles.iconButton}" title="Close">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
              <span class="${styles.buttonLabel}">Close</span>
            </button>
            <span id="saveMessage" style="margin-left: 10px;"></span>
          </div>
        </form>
      </div>
    `;
  }

  public static attachFormHandlers(
    domElement: HTMLElement,
    onSave: (title: string, firstName: string, lastName: string, phoneNumber: string, company: string, jobTitle: string) => Promise<void>,
    onClose: () => void,
    onGenerate: () => Promise<void>,
    onDownload: () => void
  ): void {
    const form = domElement.querySelector('#qrCodeForm') as HTMLFormElement;
    if (!form) return;

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const saveMessage = domElement.querySelector('#saveMessage');
      const saveButton = domElement.querySelector('#saveButton') as HTMLButtonElement;
      
      if (!saveMessage || !saveButton) return;

      try {
        saveButton.disabled = true;
        saveMessage.innerHTML = 'Saving...';

        const titleInput = domElement.querySelector('#title') as HTMLInputElement;
        const firstNameInput = domElement.querySelector('#firstName') as HTMLInputElement;
        const lastNameInput = domElement.querySelector('#lastName') as HTMLInputElement;
        const phoneNumberInput = domElement.querySelector('#phoneNumber') as HTMLInputElement;
        const companyInput = domElement.querySelector('#company') as HTMLInputElement;
        const jobTitleInput = domElement.querySelector('#jobTitle') as HTMLInputElement;

        await onSave(titleInput.value, firstNameInput.value, lastNameInput.value, phoneNumberInput.value, companyInput.value, jobTitleInput.value);

        saveMessage.innerHTML = '<span style="color: green;">✓ Saved successfully!</span>';
        setTimeout(() => {
          saveMessage.innerHTML = '';
        }, 3000);
      } catch (error) {
        saveMessage.innerHTML = `<span style="color: red;">Error saving: ${error}</span>`;
      } finally {
        saveButton.disabled = false;
      }
    });

    const closeButton = domElement.querySelector('#closeButton');
    if (closeButton) {
      closeButton.addEventListener('click', onClose);
    }

    const generateQRButton = domElement.querySelector('#generateQRButton');
    if (generateQRButton) {
      generateQRButton.addEventListener('click', () => {
        void onGenerate();
      });
    }

    const downloadQRButton = domElement.querySelector('#downloadQRButton');
    if (downloadQRButton) {
      downloadQRButton.addEventListener('click', onDownload);
    }
  }
}
