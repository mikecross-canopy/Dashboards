# Sample Google Sheet Data Template

Copy this data into your Google Sheet to get started quickly.

## Basic Template

Create a new Google Sheet with these columns and sample data:

### Sheet Structure

| Date       | Volume | Success Rate | Avg Transaction |
|------------|--------|--------------|-----------------|
| Jan 2024   | 15000  | 95.5         | 75.50           |
| Feb 2024   | 18000  | 96.2         | 80.25           |
| Mar 2024   | 22000  | 94.8         | 85.00           |
| Apr 2024   | 19500  | 97.1         | 78.90           |
| May 2024   | 24000  | 95.9         | 82.15           |
| Jun 2024   | 26500  | 96.5         | 88.40           |
| Jul 2024   | 28000  | 94.2         | 91.25           |
| Aug 2024   | 25500  | 96.8         | 86.70           |
| Sep 2024   | 27000  | 95.3         | 89.50           |
| Oct 2024   | 29500  | 97.4         | 93.80           |
| Nov 2024   | 31000  | 96.1         | 95.20           |
| Dec 2024   | 33500  | 95.7         | 98.50           |

## Extended Template (More Metrics)

For a more comprehensive dashboard, you can add additional columns:

| Date       | Volume | Success Rate | Avg Transaction | Failed Trans | Refunds | New Customers | Region       |
|------------|--------|--------------|-----------------|--------------|---------|---------------|--------------|
| Jan 2024   | 15000  | 95.5         | 75.50           | 675          | 120     | 450           | North America|
| Feb 2024   | 18000  | 96.2         | 80.25           | 684          | 135     | 520           | North America|
| Mar 2024   | 22000  | 94.8         | 85.00           | 1144         | 165     | 680           | Europe       |
| Apr 2024   | 19500  | 97.1         | 78.90           | 566          | 145     | 590           | Asia         |
| May 2024   | 24000  | 95.9         | 82.15           | 984          | 180     | 720           | North America|
| Jun 2024   | 26500  | 96.5         | 88.40           | 928          | 195     | 795           | Europe       |

## Column Descriptions

- **Date**: The period for the data (can be daily, weekly, or monthly)
- **Volume**: Total transaction amount in dollars
- **Success Rate**: Percentage of successful transactions (0-100)
- **Avg Transaction**: Average transaction value in dollars
- **Failed Trans**: Number of failed transactions (optional)
- **Refunds**: Number of refunds processed (optional)
- **New Customers**: Number of new customers acquired (optional)
- **Region**: Geographic region (optional)

## Tips for Your Data

1. **Consistent Headers**: Keep header names consistent across updates
2. **Date Format**: Use a consistent date format (e.g., "Jan 2024" or "2024-01-01")
3. **Numbers Only**: Don't include currency symbols or commas in number columns
4. **Percentages**: Enter as numbers (e.g., 95.5 for 95.5%, not "95.5%")
5. **No Empty Rows**: Avoid empty rows between data

## How to Use This Template

1. **Create a new Google Sheet**
2. **Copy the table above** (including headers)
3. **Paste into your sheet** starting at cell A1
4. **Customize the data** with your actual payment metrics
5. **Follow the setup guide** in SIMPLE_SETUP.md to connect it to your dashboard

## Example Google Sheet URL

After publishing, your CSV URL will look like:
```
https://docs.google.com/spreadsheets/d/e/2PACX-1vT...abc123.../pub?output=csv
```

Copy this entire URL and paste it into the `CONFIG.CSV_URL` in `app.js`.

## Data Update Frequency

The dashboard fetches fresh data every time it loads. To see updates:
1. Edit your Google Sheet
2. Wait a few seconds for Google to process
3. Refresh your dashboard

**Note**: Published sheets may have a slight delay (usually 1-2 minutes) before changes appear.
