// Google Sheets Configuration
// Copy this file to config.js and fill in your actual values
// DO NOT commit config.js to version control!

const CONFIG = {
    // ═══════════════════════════════════════════════════════════════
    // CHOOSE YOUR METHOD (pick one):
    // ═══════════════════════════════════════════════════════════════
    
    // 'csv'      - Easiest! Just publish your sheet (2 min setup, FREE)
    // 'api_key'  - More control, requires Google Cloud setup (FREE)
    // 'oauth'    - Most secure, for private sheets (FREE)
    
    AUTH_METHOD: 'csv',
    
    // ═══════════════════════════════════════════════════════════════
    // METHOD 1: CSV (Recommended for getting started)
    // ═══════════════════════════════════════════════════════════════
    // 1. In Google Sheets: File → Share → Publish to web
    // 2. Select "Comma-separated values (.csv)"
    // 3. Click Publish and copy the URL
    // 4. Paste it below:
    
    CSV_URL: 'https://docs.google.com/spreadsheets/d/e/YOUR_SHEET_ID/pub?output=csv',
    
    // ═══════════════════════════════════════════════════════════════
    // METHOD 2: API Key (for more control)
    // ═══════════════════════════════════════════════════════════════
    // See GOOGLE_SHEETS_SETUP.md for detailed instructions
    
    SPREADSHEET_ID: 'YOUR_GOOGLE_SHEET_ID',
    SHEET_NAME: 'Payments',
    RANGE: 'A:D',
    API_KEY: 'YOUR_GOOGLE_SHEETS_API_KEY',
    
    // ═══════════════════════════════════════════════════════════════
    // METHOD 3: OAuth 2.0 (for private sheets)
    // ═══════════════════════════════════════════════════════════════
    // See GOOGLE_SHEETS_SETUP.md for detailed instructions
    
    CLIENT_ID: 'YOUR_OAUTH_CLIENT_ID',
    DISCOVERY_DOCS: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
    SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly'
};
