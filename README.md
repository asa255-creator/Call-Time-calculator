# Call Time Calculator 📊

A Google Apps Script tool that automatically scans your sent emails for call time reports and calculates totals for all your fundraising metrics.

## Features ✨

- 🔍 **Smart Email Scanning**: Automatically searches your sent emails to specific recipients
- 📅 **Flexible Date Ranges**: Search emails from the last 7 days up to 2 years
- 📊 **Automatic Metric Extraction**: Parses and sums up:
  - Session hours
  - Soft pledges ($)
  - Hard pledges ($)
  - Estimated pledges ($)
  - Number of pledges
  - Number of calls
  - Number of pickups
- 📈 **Calculated Analytics**: Automatically computes:
  - Pickup rate (%)
  - Average pledge amount
  - Calls per hour
- 🎨 **Beautiful Results Display**: Clean, formatted output in your Google Sheet
- 🚀 **Easy to Use**: Simple dropdown menu interface
- 🗓️ **Calendar Alignment**: Pulls "Call Time" events from your Google Calendars and aligns daily scheduled vs actual hours

## Setup Instructions 🛠️

### Step 1: Create a Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it something like "Call Time Reports"

### Step 2: Open Apps Script Editor

1. In your Google Sheet, click **Extensions** → **Apps Script**
2. This will open the Apps Script editor in a new tab

### Step 3: Add the Code Files

1. **Delete the default code** in `Code.gs`
2. **Copy the contents** of `Code.gs` from this repository
3. **Paste it** into the Apps Script editor

4. **Add the HTML file**:
   - Click the **+** button next to "Files"
   - Select **HTML**
   - Name it `EmailScanner` (without .html extension)
   - Copy the contents of `EmailScanner.html` from this repository
   - Paste it into the new HTML file

5. **Save the project** (Ctrl+S or Cmd+S)
6. **Name your project** (e.g., "Call Time Calculator")

### Step 4: Initial Authorization

1. Go back to your Google Sheet and **refresh the page**
2. You should see a new menu called **"Call Time Scanner"** appear (wait a few seconds)
3. Click **Call Time Scanner** → **Scan Sent Emails**
4. The first time, you'll see an authorization dialog:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** → **Go to [Your Project Name] (unsafe)**
   - Click **Allow**

This authorization is needed so the script can read your Gmail sent folder and Google Calendar.

## How to Use 📖

### Scanning Emails

1. Open your Google Sheet
2. Click **Call Time Scanner** → **Scan Sent Emails**
3. A sidebar will appear on the right
4. Fill in the form:
   - **Recipient Email Address**: Enter the email address you send reports to (e.g., `dan@example.com`)
   - **How Far Back to Search**: Select a date range from the dropdown
5. Click **🔍 Scan Emails**
6. Wait for the scan to complete (you'll see a loading indicator)
7. Results will appear in your spreadsheet!

### Understanding the Results

The script will create a formatted report showing:

**Search Parameters**
- Email address searched
- Date range
- Number of emails found

**Totals**
- Total Session Hours
- Actual Session Hours (Attended)
- Calendar Scheduled Hours
- No Attendance Hours
- Scheduled vs Actual (Calendar - Actual)
- Total Soft Pledges
- Total Hard Pledges
- Total Estimated Pledges
- Total Number of Pledges
- Total Number of Calls
- Total Number of Pickups

**Calculated Metrics**
- Pickup Rate (percentage)
- Average Pledge Amount
- Calls Per Hour

**Calendar Alignment**
- Daily scheduled hours from Google Calendar events with "Call Time" in the title across your calendars
- Daily actual hours from email reports
- Declined hours (when you or another attendee marked "No")

### Clearing Results

To clear the current results:
1. Click **Call Time Scanner** → **Clear Results**

## Email Format Requirements 📧

Your emails should contain lines with these patterns:

```
Session length: 2 hours (2hrs)
Total in soft pledges: $250
Total in hard pledges: $0
Total estimated pledges: $250
Total number of pledges: 3
Number of calls: 20
Number of pickups: 4
Attendance: yes
```

If an email includes `Attendance: no`, the session will count toward scheduled hours but not actual hours.

The script is flexible and will find these patterns even if there's other text around them.

## Troubleshooting 🔧

### "Call Time Scanner" menu doesn't appear
- Try refreshing your Google Sheet
- Make sure you saved the Apps Script project
- Wait up to 30 seconds after opening the sheet

### "Authorization required" errors
- Follow the authorization steps in Step 4 above
- Make sure you're using the same Google account for both Sheets and Gmail

### No emails found
- Check that the email address is correct
- Try a longer date range
- Verify that you sent emails to that address from your Gmail account
- Make sure your emails contain the required format

### Missing metrics in results
- Check that your emails follow the format requirements
- The script looks for specific patterns (case-insensitive)
- Even if some metrics are missing from some emails, the script will still process available data

## Privacy & Security 🔒

- This script only runs in your Google account
- It only reads emails from your **sent folder**
- No data is sent to external servers
- All processing happens within Google's infrastructure
- You can review the code to see exactly what it does

## Customization 🎨

You can customize the script by editing `Code.gs`:

- **Add more date ranges**: Edit the dropdown options in `EmailScanner.html`
- **Change the output format**: Modify the `displayResults()` function
- **Add more metrics**: Update the `parseEmailMetrics()` function with new patterns
- **Export to CSV**: Add a new menu item with an export function

## Support & Contributions 💬

If you encounter issues or have suggestions:
- Check the troubleshooting section above
- Review the code comments for detailed explanations
- Feel free to modify the code for your specific needs

## License

This project is open source and available for personal and commercial use.

---

**Made for tracking fundraising call time metrics efficiently!** 🎯
