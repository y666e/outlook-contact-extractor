// Contact Extractor Pro - Configuration for Outlook Add-in
// Your personal backend configuration

const CONFIG = {
    // Google Apps Script Backend URL (your personal deployment)
    GOOGLE_APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbynyzJJPGweQkMafIcOMr8e2hIY4QY8nDGKtLeAOeMZRuuvT7F-00lfp-ju-v8yTjVbrA/exec',
    
    // Google Sheets Integration (same as Gmail add-on)
    SHEET_ID: '1RVI_GoWmpzuUjK5cl-ZimMDtuIxdpTr2iCIqtDVjTFw',
    SHEET_URL: 'https://docs.google.com/spreadsheets/d/1RVI_GoWmpzuUjK5cl-ZimMDtuIxdpTr2iCIqtDVjTFw',
    
    // Team Configuration
    TEAM_EMAILS: ['yyoussef@ud.ac.ae'],
    
    // Phone Number Settings
    DEFAULT_COUNTRY_CODE: '+971',
    
    // Processing Settings
    MAX_SIGNATURES_PER_EMAIL: 20,
    CACHE_DURATION: 300000, // 5 minutes
    
    // Excluded domains - same as Gmail add-on
    EXCLUDED_DOMAINS: [
        // Meeting platforms
        'calendly.com', 'cal.com', 'acuityscheduling.com', 'booking.com', 
        'scheduleonce.com', 'doodle.com', 'zoom.us', 'teams.microsoft.com', 
        'meet.google.com', 'goto.com', 'webex.com', 'skype.com',
        
        // Social media platforms
        'facebook.com', 'instagram.com', 'twitter.com', 'x.com', 
        'linkedin.com', 'youtube.com', 'tiktok.com', 'snapchat.com', 
        'pinterest.com', 'whatsapp.com', 'telegram.org', 'discord.com', 
        'reddit.com',
        
        // Other non-company sites
        'bit.ly', 'tinyurl.com', 'short.link', 'ow.ly'
    ],
    
    // Dropdown Options - same as Gmail add-on
    SECTOR_OPTIONS: [
        'Government',
        'Semi-government', 
        'Private'
    ],
    
    INDUSTRY_OPTIONS: [
        'Accounting, Finance & Banking',
        'Agriculture, Food & Beverage',
        'Arts, Media & Entertainment',
        'Construction, Real Estate & Infrastructure',
        'Consulting, Legal & Professional Services',
        'Consumer Goods, Retail & Hospitality',
        'Education, Training & Non-Profit',
        'Energy, Utilities & Environmental Services',
        'Government, Public Policy & International Affairs',
        'Healthcare, Pharmaceuticals & Life Sciences',
        'Human Resources & Career Services',
        'Information Technology, Software & Digital Services',
        'Logistics, Transportation & Supply Chain',
        'Manufacturing & Industrial Engineering',
        'Marketing, Advertising & Communications',
        'Science, Research & Innovation',
        'Security, Defense & Aerospace',
        'Sports, Recreation & Wellness'
    ],
    
    // UI Settings
    ANIMATION_DURATION: 300,
    MESSAGE_DISPLAY_DURATION: 5000,
    
    // Processing Rules
    ENGLISH_THRESHOLD: 0.7, // 70% Latin characters for English detection
    MIN_SIGNATURE_LENGTH: 20,
    MIN_ENGLISH_LINES: 2,
    MAX_SIGNATURE_LINES: 10,
    
    // Address Keywords for Detection
    ADDRESS_KEYWORDS: /\b(street|st|road|rd|avenue|ave|tower|building|floor|suite|po box|p\.o\.|dubai|uae|abu dhabi|sharjah|singapore|city|district)\b/i,
    
    // Name Patterns for English Detection
    NAME_PATTERNS: [
        /^[A-Z][a-z]+ [A-Z][a-z]+( [A-Z][a-z]+)*$/, // Standard Western names
        /^[A-Z]\. [A-Z][a-z]+$/, // Initial + surname
        /^[A-Z][a-z]+ [A-Z]\.$/, // First name + Initial
        /^[A-Za-z]+ [A-Za-z]+ [A-Za-z]+$/ // More flexible 3-word names
    ],
    
    // Title Patterns for Job Title Detection
    TITLE_PATTERNS: [
        /\b(manager|director|president|ceo|cfo|coo|vice|senior|junior|lead|head|chief|officer|analyst|coordinator|specialist|engineer|developer|designer|consultant|advisor)\b/i
    ],
    
    // Common Email Providers (to exclude from company name extraction)
    COMMON_EMAIL_PROVIDERS: [
        'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'icloud.com'
    ],
    
    // Non-company domains for website validation
    NON_COMPANY_DOMAINS: [
        'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'icloud.com',
        'google.com', 'microsoft.com', 'apple.com', 'amazon.com', 'dropbox.com'
    ],
    
    // Email Signature Indicators
    SIGNATURE_INDICATORS: [
        'regards', 'cheers', 'thanks', 'sincerely', 'best', 
        'kind regards', 'yours', 'respectfully'
    ],
    
    // Regex Patterns
    PATTERNS: {
        EMAIL: /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i,
        PHONE: /(\+?\d[\d\s\-\(\)]{6,}\d)/g,
        URL: /https?:\/\/\S+|www\.\S+/ig
    }
};

// Export configuration for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CONFIG;
}
