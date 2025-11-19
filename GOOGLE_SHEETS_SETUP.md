# Google Sheets API Setup Guide

This guide will walk you through setting up the Google Sheets API for your Payment Analytics Dashboard.

## Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click on the project dropdown at the top
3. Click "New Project"
4. Name it "Payment Analytics Dashboard"
5. Click "Create"

## Step 2: Enable Google Sheets API

1. In the Google Cloud Console, go to **APIs & Services** > **Library**
2. Search for "Google Sheets API"
3. Click on it and press **Enable**

## Step 3: Create API Credentials

### Option A: API Key (Simpler, for public read-only sheets)

1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **API Key**
3. Copy the API key
4. Click **Restrict Key** (recommended)
5. Under "API restrictions", select "Restrict key"
6. Choose "Google Sheets API" from the dropdown
7. Click **Save**

**Important**: With an API key, your Google Sheet must be publicly accessible (anyone with the link can view).

### Option B: OAuth 2.0 (More secure, for private sheets)

1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **OAuth client ID**
3. If prompted, configure the OAuth consent screen:
   - Choose "External" user type
   - Fill in app name: "Payment Analytics Dashboard"
   - Add your email as developer contact
   - Click "Save and Continue" through the scopes and test users
4. Back in Credentials, click **Create Credentials** > **OAuth client ID**
5. Choose "Web application"
6. Add authorized JavaScript origins:
   - `http://localhost:8000` (for local testing)
   - Your production domain (e.g., `https://yourdomain.com`)
7. Click **Create**
8. Copy the **Client ID** (you'll need this)

## Step 4: Prepare Your Google Sheet

1. Open your Google Sheet or create a new one
2. Structure your data with headers in the first row. Example:

   | Date       | Volume  | Success Rate | Avg Transaction |
   |------------|---------|--------------|-----------------|
   | 2024-01-01 | 15000   | 95.5         | 75.50           |
   | 2024-02-01 | 18000   | 96.2         | 80.25           |
   | 2024-03-01 | 22000   | 94.8         | 85.00           |

3. Note your **Sheet Name** (the tab name at the bottom, e.g., "Payments")
4. Get your **Spreadsheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
   ```

### If using API Key (Option A):
5. Click **Share** button
6. Change "Restricted" to "Anyone with the link"
7. Set permission to "Viewer"
8. Click **Done**

### If using OAuth 2.0 (Option B):
5. No need to make the sheet public
6. Just ensure the Google account you'll authenticate with has access

## Step 5: Update Your Dashboard Configuration

### For API Key Method:

1. Open `app.js`
2. Replace the configuration at the top:
   ```javascript
   const SPREADSHEET_ID = 'your_actual_spreadsheet_id_here';
   const SHEET_NAME = 'Payments'; // Your sheet tab name
   const API_KEY = 'your_api_key_here';
   const RANGE = 'A:D'; // Adjust based on your columns
   ```

### For OAuth 2.0 Method:

1. Open `app.js`
2. Update the configuration:
   ```javascript
   const SPREADSHEET_ID = 'your_actual_spreadsheet_id_here';
   const SHEET_NAME = 'Payments';
   const CLIENT_ID = 'your_oauth_client_id_here';
   const RANGE = 'A:D';
   ```

## Step 6: Test Your Setup

1. Open `index.html` in a web browser, or
2. Run a local server:
   ```bash
   python3 -m http.server 8000
   # or
   npx http-server -p 8000
   ```
3. Navigate to `http://localhost:8000`
4. Check the browser console for any errors

## Troubleshooting

### "API key not valid" error
- Make sure you've enabled the Google Sheets API
- Check that your API key restrictions allow the Google Sheets API
- Verify the API key is copied correctly

### "The caller does not have permission" error
- Ensure your Google Sheet is shared publicly (for API key method)
- Check that the Spreadsheet ID is correct

### "Invalid range" error
- Verify your SHEET_NAME matches the tab name exactly (case-sensitive)
- Ensure the RANGE covers your data columns

### CORS errors
- You must serve the HTML file through a web server, not open it directly as a file
- Use `python3 -m http.server` or similar

## Column Mapping

Your Google Sheet columns will be automatically mapped based on the headers. The dashboard expects:

- **Date/Month**: Date or period identifier
- **Volume/Amount**: Transaction volume or amount
- **Success Rate/Success**: Success percentage
- **Avg Transaction**: Average transaction value

The code automatically converts header names to lowercase and replaces spaces with underscores.

## Security Best Practices

1. **Never commit API keys to version control**
   - Add `config.js` to `.gitignore`
   - Use environment variables for production

2. **Restrict your API key**
   - Limit to Google Sheets API only
   - Add HTTP referrer restrictions for production

3. **For sensitive data**
   - Use OAuth 2.0 instead of API keys
   - Keep sheets private
   - Implement proper authentication

## Next Steps

Once configured, you can:
- Customize the charts in `app.js`
- Modify the table columns in `index.html`
- Add more visualizations
- Deploy to a hosting service (Netlify, Vercel, GitHub Pages)
