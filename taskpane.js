// Contact Extractor Pro - Main Application Logic for Outlook
// Matching Gmail Add-on Functionality

// Global variables
let currentContacts = [];
let currentEditingContact = null;
let threadCache = {};

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Hide loading screen and show main container
        document.getElementById('loadingScreen').style.display = 'none';
        document.getElementById('mainContainer').style.display = 'block';
        
        // Initialize the application
        initializeApp();
    }
});

// Initialize application
function initializeApp() {
    try {
        console.log('Contact Extractor Pro - Initializing...');
        
        // Load and process current email
        loadCurrentEmail();
        
        // Setup event listeners
        setupEventListeners();
        
        console.log('Contact Extractor Pro - Initialized successfully');
    } catch (error) {
        console.error('Initialization error:', error);
        showMessage('Failed to initialize Contact Extractor Pro', 'error');
    }
}

// Setup event listeners
function setupEventListeners() {
    // Manual signature input
    const extractButton = document.getElementById('extractButton');
    if (extractButton) {
        extractButton.addEventListener('click', processManualSignature);
    }
    
    // Refresh button
    const refreshButton = document.getElementById('refreshButton');
    if (refreshButton) {
        refreshButton.addEventListener('click', refreshExtraction);
    }
    
    // Modal close on background click
    const modalOverlay = document.getElementById('editModal');
    if (modalOverlay) {
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                closeEditModal();
            }
        });
    }
    
    // Keyboard shortcuts
    document.addEventListener('keydown', handleKeyboardShortcuts);
}

// Load current email content
function loadCurrentEmail() {
    try {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Text,
            { asyncContext: "getCurrentEmailBody" },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailBody = result.value;
                    const senderInfo = {
                        name: Office.context.mailbox.item.sender?.displayName || '',
                        email: Office.context.mailbox.item.sender?.emailAddress || ''
                    };
                    
                    // Process email body for signatures
                    processEmailContent(emailBody, senderInfo);
                } else {
                    console.error('Failed to get email body:', result.error);
                    showMessage('Failed to load email content', 'error');
                    showNoSignaturesSection();
                }
            }
        );
    } catch (error) {
        console.error('Error loading email:', error);
        showMessage('Error loading email content', 'error');
        showNoSignaturesSection();
    }
}

// Process email content for signatures
function processEmailContent(emailBody, senderInfo) {
    try {
        console.log('Processing email content...');
        
        // Extract signatures using same logic as Gmail add-on
        const signatures = extractSignaturesFromMessage(emailBody);
        const contacts = [];
        const seenEmails = new Set();
        
        // Process each signature
        signatures.forEach(sigText => {
            const contact = parseSignature(sigText, senderInfo.name + ' <' + senderInfo.email + '>');
            if (contact.email && !seenEmails.has(contact.email.toLowerCase())) {
                seenEmails.add(contact.email.toLowerCase());
                contacts.push(contact);
            }
        });
        
        currentContacts = contacts;
        
        // Update UI
        if (contacts.length === 0) {
            showNoSignaturesSection();
        } else {
            showContactsSection(contacts);
        }
        
        console.log(`Found ${contacts.length} contact signatures`);
    } catch (error) {
        console.error('Error processing email content:', error);
        showMessage('Error processing email content', 'error');
        showNoSignaturesSection();
    }
}

// Extract signatures from message body (same logic as Gmail add-on)
function extractSignaturesFromMessage(body) {
    if (!body || typeof body !== 'string') return [];
    
    // Clean the body
    let cleaned = body
        .replace(/\[.*?image.*?\]/gi, '')
        .replace(/\r?\n{2,}/g, '\n')
        .trim();
    
    const signatures = [];
    
    // Split by common email separators
    const messageBlocks = cleaned.split(/\nOn .* wrote:|\n---|\n______|\nFrom:|\nSent:|\n> |\n\*|\nReply|\nForward/i);
    
    messageBlocks.forEach(block => {
        const lines = block.split(/\n/).map(l => l.trim()).filter(l => l.length > 0);
        
        // Filter for English content only
        const englishLines = lines.filter(line => isEnglishText(line));
        
        if (englishLines.length < CONFIG.MIN_ENGLISH_LINES) return;
        
        // Look for signature indicators
        let sigStartIndex = -1;
        for (let i = englishLines.length - 1; i >= 0; i--) {
            const line = englishLines[i].toLowerCase();
            if (CONFIG.SIGNATURE_INDICATORS.some(indicator => line.includes(indicator))) {
                sigStartIndex = i + 1;
                break;
            }
        }
        
        // If no indicator found, take last 10 lines as potential signature
        if (sigStartIndex === -1) {
            sigStartIndex = Math.max(0, englishLines.length - CONFIG.MAX_SIGNATURE_LINES);
        }
        
        const signatureLines = englishLines.slice(sigStartIndex);
        if (signatureLines.length > 2) {
            signatures.push(signatureLines.join('\n'));
        }
    });
    
    return signatures.filter(sig => sig.length > CONFIG.MIN_SIGNATURE_LENGTH);
}

// Check if text is primarily English (same logic as Gmail add-on)
function isEnglishText(text) {
    if (!text || text.length === 0) return false;
    
    // Remove special characters, numbers, and punctuation for analysis
    const textForAnalysis = text.replace(/[^a-zA-Z\u0600-\u06FF\u4e00-\u9fff\u3400-\u4dbf]/g, '');
    
    if (textForAnalysis.length === 0) return true; // Allow lines with only numbers/punctuation
    
    // Count English characters (basic Latin)
    const englishChars = (textForAnalysis.match(/[a-zA-Z]/g) || []).length;
    const totalChars = textForAnalysis.length;
    
    // Consider it English if more than threshold are Latin characters
    return (englishChars / totalChars) > CONFIG.ENGLISH_THRESHOLD;
}

// Parse signature (same logic as Gmail add-on)
function parseSignature(signature, fromHeader = '') {
    if (!signature) return {};
    
    const rawSignature = signature;
    
    // Clean signature
    signature = signature
        .replace(/\[.*?image.*?\]/gi, '')
        .replace(/\*+/g, '')
        .replace(/\r?\n{2,}/g, '\n')
        .trim();
    
    const lines = signature.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    
    let contact = {
        rawSignature,
        company: '',
        sector: '',
        industry: '',
        name: '',
        title: '',
        email: '',
        phone: '',
        phone2: '',
        website: '',
        linkedin: '',
        address: '',
        notes: ''
    };
    
    // Extract email
    for (const line of lines) {
        const emailMatch = line.match(CONFIG.PATTERNS.EMAIL);
        if (emailMatch) {
            contact.email = emailMatch[0].toLowerCase();
            break;
        }
    }
    
    // Extract phone numbers
    const phoneMatches = [];
    for (const line of lines) {
        let match;
        while ((match = CONFIG.PATTERNS.PHONE.exec(line)) !== null) {
            phoneMatches.push(normalizePhone(match[0]));
        }
    }
    if (phoneMatches.length > 0) contact.phone = phoneMatches[0];
    if (phoneMatches.length > 1) contact.phone2 = phoneMatches[1];
    
    // Extract URLs - STRICTLY company websites only
    for (const line of lines) {
        let match;
        while ((match = CONFIG.PATTERNS.URL.exec(line)) !== null) {
            let url = match[0].replace(/[>\]]/g, '');
            if (url.startsWith('www.')) {
                url = 'https://' + url;
            }
            
            // Check if it's an excluded domain
            const isExcludedDomain = CONFIG.EXCLUDED_DOMAINS.some(domain => 
                url.toLowerCase().includes(domain));
            
            if (isExcludedDomain) {
                // Handle LinkedIn separately
                if (/linkedin/i.test(url)) {
                    contact.linkedin = url;
                }
                continue;
            }
            
            // Only add legitimate company websites
            if (!contact.website && isCompanyWebsite(url)) {
                contact.website = url;
            }
        }
    }
    
    // Extract name (English names only)
    for (const line of lines) {
        if (CONFIG.NAME_PATTERNS.some(pattern => pattern.test(line)) && 
            isEnglishText(line) &&
            !line.includes('@') && 
            !CONFIG.PATTERNS.PHONE.test(line) && 
            !CONFIG.PATTERNS.URL.test(line)) {
            contact.name = line;
            break;
        }
    }
    
    // Fallback: extract name from email header
    if (!contact.name && fromHeader) {
        const headerMatch = fromHeader.match(/^"?([^"<>]+)"?\s*</);
        if (headerMatch) {
            const headerName = headerMatch[1].trim();
            if (isEnglishText(headerName)) {
                contact.name = headerName;
            }
        }
    }
    
    // Extract title (English only)
    if (contact.name) {
        const nameIndex = lines.indexOf(contact.name);
        for (let i = nameIndex + 1; i < lines.length; i++) {
            const line = lines[i];
            if (line !== contact.email && line !== contact.phone && 
                !CONFIG.PATTERNS.URL.test(line) && 
                !CONFIG.PATTERNS.PHONE.test(line) && 
                isEnglishText(line)) {
                contact.title = line;
                break;
            }
        }
    } else {
        // Look for title-like patterns
        for (const line of lines) {
            if (CONFIG.TITLE_PATTERNS.some(pattern => pattern.test(line)) &&
                isEnglishText(line) &&
                !line.includes('@') && 
                !CONFIG.PATTERNS.PHONE.test(line) && 
                !CONFIG.PATTERNS.URL.test(line)) {
                contact.title = line;
                break;
            }
        }
    }
    
    // Extract company from email domain if not found
    if (!contact.company && contact.email) {
        const domain = contact.email.split('@')[1];
        if (domain && !CONFIG.COMMON_EMAIL_PROVIDERS.some(provider => domain.includes(provider))) {
            contact.company = toTitleCase(domain.split('.')[0]);
        }
    }
    
    // Extract address
    const addressLines = lines.filter(line => {
        const isOtherField = [contact.name, contact.title, contact.email, contact.phone].includes(line);
        const hasAddressKeywords = CONFIG.ADDRESS_KEYWORDS.test(line);
        const hasNumbers = /\d/.test(line);
        const isNotUrl = !CONFIG.PATTERNS.URL.test(line);
        const isEnglish = isEnglishText(line);
        const isNotExcludedDomain = !CONFIG.EXCLUDED_DOMAINS.some(domain => 
            line.toLowerCase().includes(domain));
        
        return !isOtherField && (hasAddressKeywords || hasNumbers) && 
               isNotUrl && isEnglish && isNotExcludedDomain;
    });
    
    contact.address = addressLines.join(', ');
    
    return contact;
}

// Check if URL is a legitimate company website
function isCompanyWebsite(url) {
    try {
        const urlObj = new URL(url.startsWith('http') ? url : 'https://' + url);
        const hostname = urlObj.hostname.toLowerCase();
        
        // Skip excluded domains
        if (CONFIG.EXCLUDED_DOMAINS.some(domain => hostname.includes(domain))) {
            return false;
        }
        
        // Skip non-company domains
        if (CONFIG.NON_COMPANY_DOMAINS.some(domain => hostname.includes(domain))) {
            return false;
        }
        
        // Must have proper domain structure
        const domainParts = hostname.split('.');
        if (domainParts.length < 2) return false;
        
        return true;
    } catch (e) {
        return false;
    }
}

// Normalize phone number
function normalizePhone(number) {
    if (!number) return '';
    number = number.replace(/[^\d\+]/g, '');
    if (!number.startsWith('+')) {
        number = CONFIG.DEFAULT_COUNTRY_CODE + number;
    }
    return number;
}

// Convert to title case
function toTitleCase(str) {
    if (!str) return '';
    return str.replace(/\b\w+/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
}

// Show contacts section in UI
function showContactsSection(contacts) {
    // Hide no signatures section
    document.getElementById('noSignaturesSection').style.display = 'none';
    
    // Show and update stats section
    const statsSection = document.getElementById('statsSection');
    const statsText = document.getElementById('statsText');
    
    const count = contacts.length;
    statsText.textContent = count === 1 ? 'FOUND 1 SIGNATURE' : `FOUND ${count} SIGNATURES`;
    statsSection.style.display = 'block';
    
    // Render contact cards
    const contactsContainer = document.getElementById('contactsContainer');
    contactsContainer.innerHTML = '';
    
    contacts.forEach((contact, index) => {
        const contactCard = createContactCard(contact, index);
        contactsContainer.appendChild(contactCard);
    });
    
    // Show action buttons
    document.getElementById('actionButtons').style.display = 'block';
    
    // Add animations
    setTimeout(() => {
        statsSection.classList.add('fade-in');
        contactsContainer.classList.add('slide-up');
    }, 100);
}

// Show no signatures section
function showNoSignaturesSection() {
    document.getElementById('statsSection').style.display = 'none';
    document.getElementById('contactsContainer').innerHTML = '';
    document.getElementById('actionButtons').style.display = 'none';
    document.getElementById('noSignaturesSection').style.display = 'block';
}

// Create contact card element
function createContactCard(contact, index) {
    const card = document.createElement('div');
    card.className = 'contact-card fade-in';
    card.style.animationDelay = `${index * 0.1}s`;
    
    const contactNumber = `CONTACT ${(index + 1).toString().padStart(2, '0')}`;
    
    card.innerHTML = `
        <div class="contact-header">
            <div class="contact-number">${contactNumber}</div>
            ${contact.name ? `<div class="contact-name">${contact.name}</div>` : ''}
            ${contact.title ? `<div class="contact-title">${contact.title}</div>` : ''}
            ${contact.company ? `<div class="contact-company">${contact.company}</div>` : ''}
        </div>
        
        <div class="contact-details">
            ${contact.email ? `
                <div class="contact-detail-item">
                    <div class="detail-label">EMAIL</div>
                    <div class="detail-value">${contact.email}</div>
                </div>
            ` : ''}
            
            ${contact.phone ? `
                <div class="contact-detail-item">
                    <div class="detail-label">PHONE</div>
                    <div class="detail-value">${contact.phone}</div>
                </div>
            ` : ''}
            
            ${contact.phone2 ? `
                <div class="contact-detail-item">
                    <div class="detail-label">PHONE 2</div>
                    <div class="detail-value">${contact.phone2}</div>
                </div>
            ` : ''}
            
            ${contact.website ? `
                <div class="contact-detail-item">
                    <div class="detail-label">WEBSITE</div>
                    <div class="detail-value"><a href="${contact.website}" target="_blank">${contact.website}</a></div>
                </div>
            ` : ''}
            
            ${contact.address ? `
                <div class="contact-detail-item">
                    <div class="detail-label">LOCATION</div>
                    <div class="detail-value">
                        <a href="https://maps.google.com/?q=${encodeURIComponent(contact.address)}" target="_blank">
                            ${contact.address.length > 40 ? contact.address.substring(0, 40) + '...' : contact.address}
                        </a>
                    </div>
                </div>
            ` : ''}
        </div>
        
        ${contact.rawSignature ? `
            <div class="signature-preview">
                <div class="preview-label">SIGNATURE PREVIEW</div>
                <div class="preview-text">${cleanSignatureForDisplay(contact.rawSignature).substring(0, 120)}${contact.rawSignature.length > 120 ? '...' : ''}</div>
            </div>
        ` : ''}
        
        <button class="contact-action-button" onclick="showEditForm(${index})">
            <span class="button-text">EDIT & SAVE</span>
        </button>
        
        <div class="contact-separator"></div>
    `;
    
    return card;
}

// Clean signature for display
function cleanSignatureForDisplay(signature) {
    if (!signature) return '';
    return signature
        .replace(/\[.*?image.*?\]/gi, '')
        .replace(/\*+/g, '')
        .replace(/\r?\n{2,}/g, '\n')
        .replace(/\n/g, '<br>')
        .trim();
}

// Show edit form modal
function showEditForm(contactIndex) {
    const contact = currentContacts[contactIndex];
    if (!contact) return;
    
    currentEditingContact = contact;
    
    // Populate form fields
    populateEditForm(contact);
    
    // Show modal
    const modal = document.getElementById('editModal');
    modal.style.display = 'flex';
    modal.classList.add('fade-in');
    
    // Focus first input
    setTimeout(() => {
        const firstInput = modal.querySelector('input, select, textarea');
        if (firstInput) firstInput.focus();
    }, 300);
}

// Populate edit form with contact data
function populateEditForm(contact) {
    document.getElementById('company').value = contact.company || '';
    document.getElementById('sector').value = contact.sector || '';
    document.getElementById('industry').value = contact.industry || '';
    document.getElementById('name').value = contact.name || '';
    document.getElementById('title').value = contact.title || '';
    document.getElementById('email').value = contact.email || '';
    document.getElementById('phone').value = contact.phone || '';
    document.getElementById('phone2').value = contact.phone2 || '';
    document.getElementById('website').value = contact.website || '';
    document.getElementById('linkedin').value = contact.linkedin || '';
    document.getElementById('address').value = contact.address || '';
    document.getElementById('notes').value = contact.notes || '';
}

// Close edit modal
function closeEditModal() {
    const modal = document.getElementById('editModal');
    modal.style.display = 'none';
    modal.classList.remove('fade-in');
    currentEditingContact = null;
}

// Save contact
function saveContact() {
    try {
        if (!currentEditingContact) {
            showMessage('No contact selected for saving', 'error');
            return;
        }
        
        // Get form data
        const formData = {
            company: document.getElementById('company').value.trim(),
            sector: document.getElementById('sector').value,
            industry: document.getElementById('industry').value,
            name: document.getElementById('name').value.trim(),
            title: document.getElementById('title').value.trim(),
            email: document.getElementById('email').value.trim(),
            phone: document.getElementById('phone').value.trim(),
            phone2: document.getElementById('phone2').value.trim(),
            website: document.getElementById('website').value.trim(),
            linkedin: document.getElementById('linkedin').value.trim(),
            address: document.getElementById('address').value.trim(),
            notes: document.getElementById('notes').value.trim()
        };
        
        // Validate required fields
        if (!formData.email && !formData.name) {
            showMessage('Please provide at least an email or name', 'error');
            return;
        }
        
        // Show saving state
        const saveButton = document.getElementById('saveButton');
        const originalText = saveButton.innerHTML;
        saveButton.innerHTML = '<span class="button-text">SAVING...</span>';
        saveButton.disabled = true;
        
        // Save to Google Sheets via Google Apps Script
        saveContactToSheet(formData)
            .then(() => {
                showMessage(`Contact ${formData.name || formData.email} saved successfully!`, 'success');
                closeEditModal();
                
                // Send notification email
                sendNotificationEmail(formData, 'added');
                
            })
            .catch((error) => {
                console.error('Save error:', error);
                showMessage('Failed to save contact. Please try again.', 'error');
            })
            .finally(() => {
                saveButton.innerHTML = originalText;
                saveButton.disabled = false;
            });
        
    } catch (error) {
        console.error('Error saving contact:', error);
        showMessage('Error saving contact', 'error');
    }
}

// Save contact to Google Sheets
function saveContactToSheet(contact) {
    return new Promise((resolve, reject) => {
        const data = {
            action: 'saveContact',
            contact: contact,
            addedBy: Office.context.mailbox.userProfile.emailAddress,
            timestamp: new Date().toISOString()
        };
        
        fetch(CONFIG.GOOGLE_APPS_SCRIPT_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data)
        })
        .then(response => response.json())
        .then(result => {
            if (result.success) {
                resolve(result);
            } else {
                reject(new Error(result.error || 'Failed to save contact'));
            }
        })
        .catch(error => {
            console.error('Network error:', error);
            reject(error);
        });
    });
}

// Send notification email
function sendNotificationEmail(contact, action) {
    const data = {
        action: 'sendNotification',
        contact: contact,
        actionType: action,
        addedBy: Office.context.mailbox.userProfile.emailAddress,
        timestamp: new Date().toISOString()
    };
    
    fetch(CONFIG.GOOGLE_APPS_SCRIPT_URL, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
    })
    .catch(error => {
        console.error('Failed to send notification:', error);
    });
}

// Process manual signature input
function processManualSignature() {
    const input = document.getElementById('manualSignatureInput');
    const sigText = input.value.trim();
    
    if (!sigText) {
        showMessage('Please paste a signature before extracting', 'warning');
        return;
    }
    
    try {
        // Show processing state
        const button = document.getElementById('extractButton');
        const originalText = button.innerHTML;
        button.innerHTML = '<span class="button-text">PROCESSING...</span>';
        button.disabled = true;
        
        // Parse the manual signature
        const contact = parseSignature(sigText);
        
        if (!contact.email && !contact.name) {
            showMessage('No contact information found in the signature', 'warning');
        } else {
            // Show edit form with the parsed contact
            currentContacts = [contact];
            currentEditingContact = contact;
            populateEditForm(contact);
            showEditForm(0);
        }
        
        // Reset button
        button.innerHTML = originalText;
        button.disabled = false;
        
        // Clear input
        input.value = '';
        
    } catch (error) {
        console.error('Error processing manual signature:', error);
        showMessage('Error processing signature', 'error');
    }
}

// Refresh extraction
function refreshExtraction() {
    try {
        showMessage('Refreshing contact extraction...', 'success');
        
        // Clear current data
        currentContacts = [];
        threadCache = {};
        
        // Reload email content
        loadCurrentEmail();
        
    } catch (error) {
        console.error('Error refreshing:', error);
        showMessage('Error refreshing extraction', 'error');
    }
}

// Handle keyboard shortcuts
function handleKeyboardShortcuts(event) {
    // ESC to close modal
    if (event.key === 'Escape') {
        const modal = document.getElementById('editModal');
        if (modal.style.display === 'flex') {
            closeEditModal();
        }
    }
    
    // Ctrl/Cmd + Enter to save when modal is open
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        const modal = document.getElementById('editModal');
        if (modal.style.display === 'flex') {
            saveContact();
        }
    }
    
    // Ctrl/Cmd + R to refresh
    if ((event.ctrlKey || event.metaKey) && event.key === 'r') {
        event.preventDefault();
        refreshExtraction();
    }
}

// Show message to user
function showMessage(message, type = 'success') {
    const container = document.getElementById('messageContainer');
    const content = document.getElementById('messageContent');
    
    // Clear existing classes
    content.className = 'message-content';
    
    // Add type-specific class
    content.classList.add(`message-${type}`);
    content.textContent = message;
    
    // Show message
    container.style.display = 'block';
    container.classList.add('fade-in');
    
    // Auto-hide after duration
    setTimeout(() => {
        container.style.display = 'none';
        container.classList.remove('fade-in');
    }, CONFIG.MESSAGE_DISPLAY_DURATION);
}

// Error handler for unhandled promises
window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    showMessage('An unexpected error occurred', 'error');
});

// Global error handler
window.onerror = (message, source, lineno, colno, error) => {
    console.error('Global error:', message, 'at', source, lineno, colno, error);
    showMessage('An error occurred in the application', 'error');
    return true;
};

console.log('Contact Extractor Pro - Outlook Add-in Loaded');
