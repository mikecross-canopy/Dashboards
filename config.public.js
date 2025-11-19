// Public configuration for GitHub Pages
// Safe to commit (no secrets). Adjust AUTH_METHOD and values as needed.
window.CONFIG = {
  // Use 'oauth' (private sheets) or 'csv' (published sheet)
  AUTH_METHOD: 'oauth',

  // CSV (optional): publish your sheet and paste the CSV URL here if you switch to CSV
  CSV_URL: 'https://docs.google.com/spreadsheets/d/e/YOUR_PUBLISHED_CSV_URL/pub?output=csv',

  // Shared settings used by both root and adm dashboards
  SPREADSHEET_ID: '1XhNpvY1SYsvszBugJeD-gQKar9OwU4lLLiF5cTkAk6I',

  // OAuth (private sheets)
  CLIENT_ID: '630950450890-lofitb7ofs2q6ae3olqv1jis88uhfjeg.apps.googleusercontent.com',
  DISCOVERY_DOCS: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly'
};
