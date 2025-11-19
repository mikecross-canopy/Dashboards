# Payment Analytics Dashboard 📊

A beautiful, responsive dashboard for visualizing payment data from Google Sheets using Chart.js.

![Dashboard Preview](https://img.shields.io/badge/Status-Ready-green)

## ✨ Features

- 📈 **Beautiful Charts**: Line, bar, doughnut, and more using Chart.js
- 🔄 **Real-time Data**: Automatically pulls data from Google Sheets
- 📱 **Responsive Design**: Works perfectly on desktop, tablet, and mobile
- 🎨 **Modern UI**: Clean, professional interface with Tailwind CSS
- 🚀 **Easy Setup**: Get started in 2 minutes with zero cost

## 🚀 Quick Start (2 Minutes)

### Option 1: CSV Method (Easiest - No API Keys!)

1. **Prepare your Google Sheet**
   - Create a sheet with headers in the first row
   - Example: `Date`, `Volume`, `Success Rate`, `Avg Transaction`

2. **Publish to Web**
   - In Google Sheets: `File` → `Share` → `Publish to web`
   - Select "Comma-separated values (.csv)"
   - Click "Publish" and copy the URL

3. **Configure the Dashboard**
   - Open `app.js`
   - Find the `CONFIG` section (around line 5)
   - Paste your CSV URL:
     ```javascript
     const CONFIG = {
         AUTH_METHOD: 'csv',
         CSV_URL: 'YOUR_PUBLISHED_CSV_URL_HERE',
         // ... rest of config
     };
     ```

4. **Open the Dashboard**
   - Open `index.html` in your browser, or
   - Run a local server:
     ```bash
     python3 -m http.server 8000
     # Then visit http://localhost:8000
     ```

**Done!** 🎉 Your dashboard is now live and pulling data from Google Sheets.

## 📖 Documentation

- **[SIMPLE_SETUP.md](SIMPLE_SETUP.md)** - Quick setup guide (CSV method)
- **[GOOGLE_SHEETS_SETUP.md](GOOGLE_SHEETS_SETUP.md)** - Advanced setup with Google Sheets API

## 🎯 Setup Methods Comparison

| Method | Setup Time | Cost | Security | Best For |
|--------|-----------|------|----------|----------|
| **CSV** | 2 min | Free | Public data | Quick demos, non-sensitive data |
| **API Key** | 10 min | Free | Public sheets | More control, still simple |
| **OAuth 2.0** | 15 min | Free | Private sheets | Sensitive/private data |

## 📊 Google Sheet Format

Your Google Sheet should have headers in the first row:

| Date       | Volume  | Success Rate | Avg Transaction |
|------------|---------|--------------|-----------------|
| 2024-01-01 | 15000   | 95.5         | 75.50           |
| 2024-02-01 | 18000   | 96.2         | 80.25           |
| 2024-03-01 | 22000   | 94.8         | 85.00           |

The dashboard automatically adapts to your column names.

## 🛠️ Customization

### Change Chart Types
Edit the chart configurations in `app.js` (around line 115+):
```javascript
volumeChart = new Chart(volumeCtx, {
    type: 'line', // Change to 'bar', 'pie', etc.
    // ...
});
```

### Modify Colors
Update the color schemes in the chart data sections:
```javascript
backgroundColor: 'rgba(79, 70, 229, 0.8)', // Your custom color
```

### Add More Charts
1. Add a canvas element in `index.html`
2. Create a new chart instance in `renderCharts()` function
3. Process your data accordingly

## 💰 Cost

**Everything is 100% FREE:**
- ✅ Google Sheets API: Free (generous limits)
- ✅ Chart.js: Free and open source
- ✅ No server costs (runs in browser)
- ✅ No credit card required

## 🔒 Security Notes

**CSV Method:**
- Data is publicly accessible to anyone with the URL
- Don't use for sensitive information
- Perfect for executive dashboards and reports

**API Key Method:**
- Sheet must be publicly viewable
- API key can be restricted to specific domains

**OAuth Method:**
- Most secure option
- Works with private sheets
- User authentication required

## 🚀 Deployment

Deploy to any static hosting service:

### Netlify
```bash
# Drag and drop your folder to netlify.com
```

### GitHub Pages
```bash
git init
git add .
git commit -m "Initial commit"
git push origin main
# Enable GitHub Pages in repository settings
```

### Vercel
```bash
vercel deploy
```

## 🐛 Troubleshooting

### "Failed to fetch CSV"
- Ensure your sheet is published to web
- Check that the CSV URL is correct
- Verify the sheet has data

### Charts not showing
- Check browser console for errors
- Ensure Chart.js is loaded (check Network tab)
- Verify data format matches expected structure

### CORS errors
- Don't open HTML file directly (file://)
- Use a local server: `python3 -m http.server`

## 📝 License

MIT License - feel free to use for personal or commercial projects!

## 🤝 Contributing

Suggestions and improvements welcome! This is a simple, standalone project designed to be easily customizable.

## 📧 Support

Check the documentation files:
- `SIMPLE_SETUP.md` - Quick start guide
- `GOOGLE_SHEETS_SETUP.md` - Detailed API setup

---

**Built with ❤️ for Payment Account Executive Managers**
