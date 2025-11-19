# Simple Google Sheets Setup (No Google Cloud Required!)

## Method 1: Publish to Web (Easiest - 2 Minutes Setup)

This method requires **NO Google Cloud Project** and **NO API Keys**!

### Steps:

1. **Open your Google Sheet**
   - Create or open your spreadsheet with payment data

2. **Publish to Web**
   - Click `File` → `Share` → `Publish to web`
   - In the dialog:
     - Choose the specific sheet/tab you want to publish
     - Select "Comma-separated values (.csv)" format
     - Click "Publish"
     - Copy the published URL

3. **Update your dashboard**
   - Open `app.js`
   - Find the `CONFIG` section
   - Set `AUTH_METHOD: 'csv'`
   - Set `CSV_URL` to your published URL

4. **Done!** Your dashboard will now pull data directly from Google Sheets.

### Example Configuration:

```javascript
const CONFIG = {
    AUTH_METHOD: 'csv',
    CSV_URL: 'https://docs.google.com/spreadsheets/d/e/YOUR_SHEET_ID/pub?output=csv',
    // No API keys needed!
};
```

### Pros:
- ✅ Zero setup time
- ✅ No Google Cloud account needed
- ✅ No API keys to manage
- ✅ Completely free
- ✅ Auto-updates when you change the sheet

### Cons:
- ⚠️ Data is publicly accessible (anyone with the URL can view it)
- ⚠️ Not suitable for sensitive/private data

---

## Method 2: Google Sheets API (Still Free!)

If you need private data, the Google Sheets API is **still completely free**. You don't pay anything:

### What's Free:
- Creating a Google Cloud Project: **FREE**
- Google Sheets API usage: **FREE** (up to very high limits)
- API Keys: **FREE**
- OAuth credentials: **FREE**

### What Could Cost Money (but you won't use):
- Google Cloud Storage, Compute Engine, etc. (we're not using these)
- You only pay if you use paid services like servers or databases

### Free Tier Limits:
- **500 requests per 100 seconds per project**
- **100 requests per 100 seconds per user**
- This is more than enough for a dashboard that refreshes every few minutes!

**Bottom Line**: Using Google Sheets API costs $0. You never need to enter a credit card.

---

## Recommendation

**For non-sensitive data**: Use Method 1 (Publish to Web)
- Fastest setup
- Zero configuration
- Perfect for executive dashboards

**For sensitive/private data**: Use Method 2 (Google Sheets API)
- Still free
- More secure
- Follow the `GOOGLE_SHEETS_SETUP.md` guide
- No credit card required

---

## Quick Start with CSV Method

1. Publish your Google Sheet to web as CSV
2. Copy the published URL
3. Update `app.js` with the CSV URL
4. Open `index.html` in a browser
5. Done!

Your dashboard will automatically fetch fresh data from Google Sheets every time it loads.
