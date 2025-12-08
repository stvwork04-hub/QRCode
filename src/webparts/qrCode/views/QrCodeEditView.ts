import { escape } from '@microsoft/sp-lodash-subset';
import styles from '../QrCodeWebPart.module.scss';
import { IQRCodeItem } from '../models/IQRCodeItem';

export class QrCodeEditView {
  
  public static renderForm(container: Element, userItem: IQRCodeItem, hasAttachment: boolean): void {
    container.innerHTML = `
      <div class="${styles.formContainer}">
        <h2 style="text-align: center; margin-top: 0; margin-bottom: 2.5rem; color: maroon;">My Digital Business Card Details</h2>
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
              <label for="phoneNumber">Work Phone <span style="color: red;">*</span>:</label>
            <div style="display: flex; align-items: center;">
              <span style="padding: 8px 12px; background: #f5f5f5; border: 1px solid #ddd; border-right: none; border-radius: 4px 0 0 4px; color: #666;">+965</span>
              <input type="tel" id="phoneNumber" name="phoneNumber" value="${escape((userItem.PhoneNumber || '').replace(/^\+965\s*/, ''))}" placeholder="12345678" maxlength="8" inputmode="numeric" title="Enter exactly 8 digits" style="border-radius: 0 4px 4px 0; margin: 0;" />
            </div>
            <div id="phoneNumberError" style="color: red; font-size: 0.875rem; margin-top: 4px; display: none; font-style: italic;"></div>
          </div>
          
          <div class="${styles.formField}">
            <label for="mobilePhone">Mobile Phone:</label>
            <div style="display: flex; align-items: center;">
              <span style="padding: 8px 12px; background: #f5f5f5; border: 1px solid #ddd; border-right: none; border-radius: 4px 0 0 4px; color: #666;">+965</span>
              <input type="tel" id="mobilePhone" name="mobilePhone" value="${escape((userItem.MobilePhone || '').replace(/^\+965\s*/, ''))}" placeholder="12345678" maxlength="8" inputmode="numeric" title="Enter exactly 8 digits" style="border-radius: 0 4px 4px 0; margin: 0;" />
            </div>
            <div id="mobilePhoneError" style="color: red; font-size: 0.875rem; margin-top: 4px; display: none; font-style: italic;"></div>
          </div>
          
          <div class="${styles.formField}">
            <label for="otherPhone">Other Phone:</label>
            <input type="tel" id="otherPhone" name="otherPhone" value="${escape(userItem.OtherPhone || '')}" placeholder="Any phone number" />
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
            <button type="submit" id="generateQRButton" class="${styles.iconButton}" title="Save & Generate QR Code">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <rect x="3" y="3" width="7" height="7"></rect>
                <rect x="14" y="3" width="7" height="7"></rect>
                <rect x="14" y="14" width="7" height="7"></rect>
                <rect x="3" y="14" width="7" height="7"></rect>
              </svg>
              <span class="${styles.buttonLabel}">Generate QR Code</span>
            </button>
            <button type="button" id="closeButton" class="${styles.iconButton}" title="Close">
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
              <span class="${styles.buttonLabel}">Close</span>
            </button>
            <span id="generateMessage" style="margin-left: 10px;"></span>
          </div>
          <div class="${styles.formField}" id="successMessage" style="grid-column: 1 / -1; text-align: center; margin-top: 1rem; padding: 1.25rem 2rem; background: #d4edda; border: 1px solid #c3e6cb; border-radius: 12px; color: #155724; font-weight: 600; font-size: 1rem; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1); display: none;">
            ✅ Data saved and QR code generated! Check your email shortly.
          </div>
        </form>
      </div>
    `;
  }

  public static attachFormHandlers(
    domElement: HTMLElement,
    onSaveAndGenerate: (formData: { PhoneNumber: string; MobilePhone?: string; Instagram?: string; Facebook?: string; Gmail?: string; OtherPhone?: string; }) => Promise<void>,
    onClose: () => void
  ): void {
    const form = domElement.querySelector('#qrCodeForm') as HTMLFormElement;
    if (!form) return;

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const generateMessage = domElement.querySelector('#generateMessage');
      const successMessage = domElement.querySelector('#successMessage') as HTMLElement;
      const generateButton = domElement.querySelector('#generateQRButton') as HTMLButtonElement;

      if (!generateMessage || !generateButton || !successMessage) return;
      
      try {
        generateButton.disabled = true;
        generateMessage.innerHTML = 'Saving and generating QR code...';

        const phoneNumberInput = domElement.querySelector('#phoneNumber') as HTMLInputElement;
        const mobilePhoneInput = domElement.querySelector('#mobilePhone') as HTMLInputElement;
        const instagramInput = domElement.querySelector('#instagram') as HTMLInputElement;
        const facebookInput = domElement.querySelector('#facebook') as HTMLInputElement;
        const gmailInput = domElement.querySelector('#gmail') as HTMLInputElement;
        const otherPhoneInput = domElement.querySelector('#otherPhone') as HTMLInputElement;

        // Validate phone numbers with different rules for each field
        let hasValidationError = false;
        const phoneErrorElement = domElement.querySelector('#phoneNumberError') as HTMLElement;
        const mobileErrorElement = domElement.querySelector('#mobilePhoneError') as HTMLElement;
        const otherErrorElement = domElement.querySelector('#otherPhoneError') as HTMLElement;
        
        // Clear previous errors
        [phoneErrorElement, mobileErrorElement, otherErrorElement].forEach(el => {
          if (el) el.style.display = 'none';
        });
        
        // Validate Phone Number (mandatory, exactly 8 digits, +965 prefix added automatically)
          const phoneValue = phoneNumberInput.value.trim();
          if (!phoneValue) {
            if (phoneErrorElement) {
              phoneErrorElement.textContent = 'Work Phone is mandatory.';
            phoneErrorElement.style.display = 'block';
          }
          hasValidationError = true;
          phoneNumberInput.focus();
        } else if (!/^\d{8}$/.test(phoneValue)) {
          if (phoneErrorElement) {
            phoneErrorElement.textContent = 'Enter a valid 8 digit number.';
            phoneErrorElement.style.display = 'block';
          }
          hasValidationError = true;
          phoneNumberInput.focus();
        }
        
        // Validate Mobile Phone (optional, exactly 8 digits if filled, +965 prefix added automatically)
        const mobileValue = mobilePhoneInput.value.trim();
        if (mobileValue && !/^\d{8}$/.test(mobileValue)) {
          if (mobileErrorElement) {
            mobileErrorElement.textContent = 'Enter a valid 8 digit number.';
            mobileErrorElement.style.display = 'block';
          }
          hasValidationError = true;
          if (!hasValidationError) mobilePhoneInput.focus();
        }
        
        // Other Phone field has no validation
        
        if (hasValidationError) {
          generateButton.disabled = false;
          generateMessage.innerHTML = '';
          return;
        }

        const formData = {
          PhoneNumber: phoneNumberInput.value ? `+965 ${phoneNumberInput.value}` : '',
          MobilePhone: mobilePhoneInput.value ? `+965 ${mobilePhoneInput.value}` : '',
          Instagram: instagramInput.value || '',
          Facebook: facebookInput.value || '',
          Gmail: gmailInput.value || '',
          OtherPhone: otherPhoneInput.value || ''
        };
        
        console.log('🔧 DEBUG: Form data being saved and QR code being generated:', formData);

        // Call the combined save and generate function
        await onSaveAndGenerate(formData);

        // Clear any previous messages
        generateMessage.innerHTML = '';
        
        // Show centered success message
        if (successMessage) {
          successMessage.style.display = 'block';
          setTimeout(() => {
            if (successMessage) {
              successMessage.style.display = 'none';
            }
          }, 5000);
        }
      } catch (error) {
        generateMessage.innerHTML = `<span style="color: red;">Error: ${error}</span>`;
      } finally {
        generateButton.disabled = false;
      }
    });

    const closeButton = domElement.querySelector('#closeButton');
    if (closeButton) {
      closeButton.addEventListener('click', onClose);
    }

    // Note: Generate QR Button is now handled by the form submit event above

    // Add validation to phone fields
    const phoneFieldIds = ['phoneNumber', 'mobilePhone'];
    const otherPhoneField = domElement.querySelector('#otherPhone') as HTMLInputElement;
    
    // Other Phone field - allow only digits and special characters, no alphabets
    if (otherPhoneField) {
      otherPhoneField.addEventListener('input', () => {
        const value = otherPhoneField.value;
        // Remove any alphabetic characters (a-z, A-Z), keep digits and special characters
        const cleanValue = value.replace(/[a-zA-Z]/g, '');
        
        if (value !== cleanValue) {
          const cursorPos = otherPhoneField.selectionStart || 0;
          otherPhoneField.value = cleanValue;
          // Try to maintain cursor position
          setTimeout(() => {
            otherPhoneField.setSelectionRange(Math.min(cursorPos, cleanValue.length), Math.min(cursorPos, cleanValue.length));
          }, 0);
        }
      });
      
      // Also prevent alphabetic input on keypress
      otherPhoneField.addEventListener('keypress', (e) => {
        // Allow control keys
        if (e.ctrlKey || e.metaKey) {
          return;
        }
        
        const char = e.key;
        // Block alphabetic characters
        if (/[a-zA-Z]/.test(char)) {
          e.preventDefault();
        }
      });
    }
    
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
          const clipboardData = e.clipboardData || (window as Window & { clipboardData?: DataTransfer }).clipboardData;
          if (!clipboardData) return;
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
