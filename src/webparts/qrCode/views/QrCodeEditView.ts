import { escape } from '@microsoft/sp-lodash-subset';
import styles from '../QrCodeWebPart.module.scss';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeEditView {
  
  public static renderForm(container: Element, userItem: IQRCodeItem, hasAttachment: boolean): void {
    container.innerHTML = `
      <div class="${styles.formContainer}">
        <form id="qrCodeForm">
          <div class="${styles.formField}">
            <label for="firstName">First Name:</label>
            <input type="text" id="firstName" name="firstName" value="${escape(userItem.FirstName || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="lastName">Last Name:</label>
            <input type="text" id="lastName" name="lastName" value="${escape(userItem.LastName || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="email">Email:</label>
            <input type="email" id="email" name="email" value="${escape(userItem.Email || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="company">Company:</label>
            <input type="text" id="company" name="company" value="${escape(userItem.Company || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="jobTitle">Job Title:</label>
            <input type="text" id="jobTitle" name="jobTitle" value="${escape(userItem.JobTitle || '')}" readonly />
          </div>
          
          <div class="${styles.formField}">
            <label for="phoneNumber">Phone Number: *</label>
            <input type="tel" id="phoneNumber" name="phoneNumber" value="${escape(userItem.PhoneNumber || '')}" required pattern="[0-9\\s-()\\+]*" inputmode="numeric" title="Please enter only numbers and phone formatting characters" />
          </div>
          
          <div class="${styles.formField}">
            <label for="mobilePhone">Mobile Phone:</label>
            <input type="tel" id="mobilePhone" name="mobilePhone" value="${escape(userItem.MobilePhone || '')}" pattern="[0-9\\s-()\\+]*" inputmode="numeric" title="Please enter only numbers and phone formatting characters" />
          </div>
          
          <div class="${styles.formField}">
            <label for="otherPhone">Other Phone:</label>
            <input type="tel" id="otherPhone" name="otherPhone" value="${escape(userItem.OtherPhone || '')}" pattern="[0-9\\s-()\\+]*" inputmode="numeric" title="Please enter only numbers and phone formatting characters" />
          </div>
          
          <div class="${styles.formField}">
            <label for="instagram">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="url(#instagram-gradient)" style="margin-right: 8px; vertical-align: middle;">
                <defs>
                  <radialGradient id="instagram-gradient" cx="0.5" cy="1.2" r="1.5">
                    <stop offset="0" stop-color="#FED576"/>
                    <stop offset="0.263" stop-color="#F47133"/>
                    <stop offset="0.609" stop-color="#BC3081"/>
                    <stop offset="1" stop-color="#4C63D2"/>
                  </radialGradient>
                </defs>
                <rect x="2" y="2" width="20" height="20" rx="5" ry="5"></rect>
                <path d="m16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z" fill="white"></path>
                <circle cx="17.5" cy="6.5" r="1" fill="white"></circle>
              </svg>
              Instagram:
            </label>
            <input type="text" id="instagram" name="instagram" value="${escape(userItem.Instagram || '')}" />
          </div>
          
          <div class="${styles.formField}">
            <label for="facebook">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="#1877F2" stroke="none" style="margin-right: 8px; vertical-align: middle;">
                <path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/>
              </svg>
              Facebook:
            </label>
            <input type="text" id="facebook" name="facebook" value="${escape(userItem.Facebook || '')}" />
          </div>
          
          <div class="${styles.formField}">
            <label for="gmail">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" style="margin-right: 8px; vertical-align: middle;">
                <path d="M24 5.457v13.909c0 .904-.732 1.636-1.636 1.636h-3.819V11.73L12 16.64l-6.545-4.91v9.273H1.636A1.636 1.636 0 0 1 0 19.366V5.457c0-.904.732-1.636 1.636-1.636h.273L12 10.845l10.091-7.024h.273c.904 0 1.636.732 1.636 1.636Z" fill="#EA4335"/>
                <path d="M0 5.457V19.366c0 .904.732 1.636 1.636 1.636h3.819V11.73L12 16.64l6.545-4.91v9.273h3.819A1.636 1.636 0 0 0 24 19.366V5.457c0-.904-.732-1.636-1.636-1.636H22.364L12 10.845 1.636 3.821H1.636C.732 3.821 0 4.553 0 5.457Z" fill="#34A853"/>
                <path d="M18.545 11.73V21.002H22.364A1.636 1.636 0 0 0 24 19.366V5.457c0-.904-.732-1.636-1.636-1.636H22.364L18.545 7.731v4Z" fill="#FBBC04"/>
                <path d="M5.455 11.73V21.002H1.636A1.636 1.636 0 0 1 0 19.366V5.457c0-.904.732-1.636 1.636-1.636H1.636L5.455 7.731v4Z" fill="#EA4335"/>
              </svg>
              Gmail:
            </label>
            <input type="email" id="gmail" name="gmail" value="${escape(userItem.Gmail || '')}" />
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
              <span class="${styles.buttonLabel}">Generate QR Code</span>
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
          <div class="${styles.formField}" id="successMessage" style="grid-column: 1 / -1; text-align: center; margin-top: 1rem; padding: 1.25rem 2rem; background: #d4edda; border: 1px solid #c3e6cb; border-radius: 12px; color: #155724; font-weight: 600; font-size: 1rem; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1); display: none;">
            Thank you! Your QR code will be sent to your email shortly.
          </div>
          <div class="${styles.formField}" id="saveSuccessMessage" style="grid-column: 1 / -1; text-align: center; margin-top: 1rem; padding: 1.25rem 2rem; background: #d4edda; border: 1px solid #c3e6cb; border-radius: 12px; color: #155724; font-weight: 600; font-size: 1rem; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1); display: none;">
            ✓ Saved successfully!
          </div>
        </form>
      </div>
    `;
  }

  public static attachFormHandlers(
    domElement: HTMLElement,
    onSave: (formData: { PhoneNumber: string; MobilePhone?: string; Instagram?: string; Facebook?: string; Gmail?: string; OtherPhone?: string; }) => Promise<void>,
    onClose: () => void,
    onGenerate: () => Promise<void>,
    onDownload: () => void
  ): void {
    const form = domElement.querySelector('#qrCodeForm') as HTMLFormElement;
    if (!form) return;

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      
    const saveMessage = domElement.querySelector('#saveMessage');
    const saveSuccessMessage = domElement.querySelector('#saveSuccessMessage') as HTMLElement;
    const saveButton = domElement.querySelector('#saveButton') as HTMLButtonElement;

    if (!saveMessage || !saveButton || !saveSuccessMessage) return;      try {
        saveButton.disabled = true;
        saveMessage.innerHTML = 'Saving...';

        const phoneNumberInput = domElement.querySelector('#phoneNumber') as HTMLInputElement;
        const mobilePhoneInput = domElement.querySelector('#mobilePhone') as HTMLInputElement;
        const instagramInput = domElement.querySelector('#instagram') as HTMLInputElement;
        const facebookInput = domElement.querySelector('#facebook') as HTMLInputElement;
        const gmailInput = domElement.querySelector('#gmail') as HTMLInputElement;
        const otherPhoneInput = domElement.querySelector('#otherPhone') as HTMLInputElement;

        const formData = {
          PhoneNumber: phoneNumberInput.value,
          MobilePhone: mobilePhoneInput.value || undefined,
          Instagram: instagramInput.value || undefined,
          Facebook: facebookInput.value || undefined,
          Gmail: gmailInput.value || undefined,
          OtherPhone: otherPhoneInput.value || undefined
        };

        await onSave(formData);

        // Clear any previous messages
        saveMessage.innerHTML = '';
        
        // Show centered success message
        const messageElement = saveSuccessMessage;
        messageElement.style.display = 'block';
        setTimeout(() => {
          messageElement.style.display = 'none';
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
      console.log('🔥 DEBUG: Generate QR Code button found and event listener being attached');
      generateQRButton.addEventListener('click', (event) => {
        console.log('🔥 DEBUG: Generate QR Code button CLICKED!');
        console.log('🔥 DEBUG: Event:', event);
        event.preventDefault();
        event.stopPropagation();
        console.log('🔥 DEBUG: About to call onGenerate function');
        onGenerate().catch((error) => {
          console.error('🔥 ERROR: onGenerate failed:', error);
        });
      });
    } else {
      console.log('🔥 ERROR: Generate QR Code button NOT found in DOM');
    }

    const downloadQRButton = domElement.querySelector('#downloadQRButton');
    if (downloadQRButton) {
      downloadQRButton.addEventListener('click', onDownload);
    }

    // Add strict numeric-only validation to phone fields
    const phoneFieldIds = ['phoneNumber', 'mobilePhone', 'otherPhone'];
    
    phoneFieldIds.forEach(fieldId => {
      const field = domElement.querySelector(`#${fieldId}`) as HTMLInputElement;
      if (field) {
        // Strict input validation - only allow numbers and basic phone formatting
        const validatePhoneInput = (input: HTMLInputElement): void => {
          const value = input.value;
          // Allow only digits, spaces, hyphens, parentheses, and plus sign
          const filteredValue = value.replace(/[^0-9\s-()+]/g, '');
          if (value !== filteredValue) {
            input.value = filteredValue;
            // Trigger change event to update any bound data
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }
        };

        // Real-time input filtering
        field.addEventListener('input', () => {
          validatePhoneInput(field);
        });

        // Prevent non-numeric key presses
        field.addEventListener('keypress', (e) => {
          const char = String.fromCharCode(e.which || e.keyCode);
          const allowedChars = /[0-9\s-()+]/;
          
          // Allow control keys (backspace, delete, etc.)
          if (e.which === 0 || e.which === 8 || e.which === 9 || e.which === 13 || e.which === 27) {
            return true;
          }
          
          if (!allowedChars.test(char)) {
            e.preventDefault();
            return false;
          }
        });

        // Handle paste events
        field.addEventListener('paste', (e) => {
          e.preventDefault();
          const clipboardData = e.clipboardData || (window as any).clipboardData;
          const pastedData = clipboardData.getData('text');
          const filteredData = pastedData.replace(/[^0-9\s-()+]/g, '');
          
          // Insert filtered data at cursor position
          const start = field.selectionStart || 0;
          const end = field.selectionEnd || 0;
          const currentValue = field.value;
          field.value = currentValue.substring(0, start) + filteredData + currentValue.substring(end);
          
          // Set cursor position after pasted content
          const newPosition = start + filteredData.length;
          field.setSelectionRange(newPosition, newPosition);
          
          // Trigger change event
          field.dispatchEvent(new Event('input', { bubbles: true }));
        });

        // Additional validation on blur to ensure clean data
        field.addEventListener('blur', () => {
          validatePhoneInput(field);
        });
      }
    });
  }
}
