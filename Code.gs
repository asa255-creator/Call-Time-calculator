/**
 * Call Time Calculator - Google Apps Script
 * Scans sent emails and calculates call time metrics
 */

/**
 * Creates custom menu when the spreadsheet is opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Call Time Scanner')
    .addItem('Scan Sent Emails', 'showEmailScannerDialog')
    .addItem('Clear Results', 'clearResults')
    .addToUi();
}

/**
 * Shows the email scanner dialog
 */
function showEmailScannerDialog() {
  var html = HtmlService.createHtmlOutputFromFile('EmailScanner')
    .setWidth(400)
    .setHeight(350)
    .setTitle('Scan Sent Emails');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Main function to scan emails and calculate totals
 * @param {string} emailAddress - The recipient email address to filter by
 * @param {string} dateRange - How far back to search (e.g., "7d", "30d", "90d", "1y")
 */
function scanSentEmails(emailAddress, dateRange) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Build the Gmail search query
    var searchQuery = 'to:' + emailAddress + ' in:sent after:' + calculateDateString(dateRange);

    // Search for matching emails
    var threads = GmailApp.search(searchQuery);
    var totalEmails = 0;

    // Initialize totals
    var totals = {
      sessionHours: 0,
      softPledges: 0,
      hardPledges: 0,
      estimatedPledges: 0,
      numberOfPledges: 0,
      numberOfCalls: 0,
      numberOfPickups: 0
    };

    // Process each thread
    threads.forEach(function(thread) {
      var messages = thread.getMessages();
      messages.forEach(function(message) {
        var body = message.getPlainBody();
        var metrics = parseEmailMetrics(body);

        if (metrics) {
          totalEmails++;
          totals.sessionHours += metrics.sessionHours || 0;
          totals.softPledges += metrics.softPledges || 0;
          totals.hardPledges += metrics.hardPledges || 0;
          totals.estimatedPledges += metrics.estimatedPledges || 0;
          totals.numberOfPledges += metrics.numberOfPledges || 0;
          totals.numberOfCalls += metrics.numberOfCalls || 0;
          totals.numberOfPickups += metrics.numberOfPickups || 0;
        }
      });
    });

    // Display results
    displayResults(sheet, emailAddress, dateRange, totalEmails, totals);

    return {
      success: true,
      message: 'Successfully scanned ' + totalEmails + ' emails',
      totals: totals
    };

  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

/**
 * Parses email body to extract call time metrics
 * @param {string} emailBody - The plain text body of the email
 * @return {Object} Extracted metrics
 */
function parseEmailMetrics(emailBody) {
  var metrics = {};

  // Pattern for "Session length: X hours" or "Session length: X hrs"
  var sessionMatch = emailBody.match(/Session length:\s*(\d+(?:\.\d+)?)\s*(?:hours?|hrs?)/i);
  if (sessionMatch) {
    metrics.sessionHours = parseFloat(sessionMatch[1]);
  }

  // Pattern for "Total in soft pledges: $XXX"
  var softPledgeMatch = emailBody.match(/Total in soft pledges:\s*\$(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
  if (softPledgeMatch) {
    metrics.softPledges = parseFloat(softPledgeMatch[1].replace(/,/g, ''));
  }

  // Pattern for "Total in hard pledges: $XXX"
  var hardPledgeMatch = emailBody.match(/Total in hard pledges:\s*\$(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
  if (hardPledgeMatch) {
    metrics.hardPledges = parseFloat(hardPledgeMatch[1].replace(/,/g, ''));
  }

  // Pattern for "Total estimated pledges: $XXX"
  var estimatedPledgeMatch = emailBody.match(/Total estimated pledges:\s*\$(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
  if (estimatedPledgeMatch) {
    metrics.estimatedPledges = parseFloat(estimatedPledgeMatch[1].replace(/,/g, ''));
  }

  // Pattern for "Total number of pledges: X"
  var pledgeCountMatch = emailBody.match(/Total number of pledges:\s*(\d+)/i);
  if (pledgeCountMatch) {
    metrics.numberOfPledges = parseInt(pledgeCountMatch[1]);
  }

  // Pattern for "Number of calls: X"
  var callsMatch = emailBody.match(/Number of calls:\s*(\d+)/i);
  if (callsMatch) {
    metrics.numberOfCalls = parseInt(callsMatch[1]);
  }

  // Pattern for "Number of pickups: X"
  var pickupsMatch = emailBody.match(/Number of pickups:\s*(\d+)/i);
  if (pickupsMatch) {
    metrics.numberOfPickups = parseInt(pickupsMatch[1]);
  }

  // Only return metrics if at least one field was found
  if (Object.keys(metrics).length > 0) {
    return metrics;
  }

  return null;
}

/**
 * Calculates the date string for Gmail search based on the date range
 * @param {string} dateRange - Date range (e.g., "7d", "30d", "90d", "1y")
 * @return {string} Date in YYYY/MM/DD format
 */
function calculateDateString(dateRange) {
  var today = new Date();
  var targetDate = new Date();

  // Parse the date range
  var value = parseInt(dateRange);
  var unit = dateRange.slice(-1).toLowerCase();

  if (unit === 'd') {
    targetDate.setDate(today.getDate() - value);
  } else if (unit === 'm') {
    targetDate.setMonth(today.getMonth() - value);
  } else if (unit === 'y') {
    targetDate.setFullYear(today.getFullYear() - value);
  }

  // Format as YYYY/MM/DD
  var year = targetDate.getFullYear();
  var month = String(targetDate.getMonth() + 1).padStart(2, '0');
  var day = String(targetDate.getDate()).padStart(2, '0');

  return year + '/' + month + '/' + day;
}

/**
 * Displays results in the spreadsheet
 * @param {Sheet} sheet - The active sheet
 * @param {string} emailAddress - The searched email address
 * @param {string} dateRange - The date range searched
 * @param {number} emailCount - Number of emails found
 * @param {Object} totals - Calculated totals
 */
function displayResults(sheet, emailAddress, dateRange, emailCount, totals) {
  // Clear previous results (start from row 1)
  sheet.clear();

  // Set up headers with styling
  var headerRange = sheet.getRange('A1:B1');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // Display search parameters
  sheet.getRange('A1').setValue('Call Time Calculator Results');
  sheet.getRange('A1:B1').merge();

  var row = 3;
  sheet.getRange('A' + row).setValue('Search Parameters').setFontWeight('bold');
  row++;
  sheet.getRange('A' + row).setValue('Email Address:');
  sheet.getRange('B' + row).setValue(emailAddress);
  row++;
  sheet.getRange('A' + row).setValue('Date Range:');
  sheet.getRange('B' + row).setValue(dateRange);
  row++;
  sheet.getRange('A' + row).setValue('Emails Found:');
  sheet.getRange('B' + row).setValue(emailCount);

  // Add spacing
  row += 2;

  // Display totals
  sheet.getRange('A' + row).setValue('TOTALS').setFontWeight('bold').setFontSize(12);
  row++;

  sheet.getRange('A' + row).setValue('Total Session Hours:');
  sheet.getRange('B' + row).setValue(totals.sessionHours);
  row++;

  sheet.getRange('A' + row).setValue('Total Soft Pledges:');
  sheet.getRange('B' + row).setValue('$' + totals.softPledges.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
  row++;

  sheet.getRange('A' + row).setValue('Total Hard Pledges:');
  sheet.getRange('B' + row).setValue('$' + totals.hardPledges.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
  row++;

  sheet.getRange('A' + row).setValue('Total Estimated Pledges:');
  sheet.getRange('B' + row).setValue('$' + totals.estimatedPledges.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
  row++;

  sheet.getRange('A' + row).setValue('Total Number of Pledges:');
  sheet.getRange('B' + row).setValue(totals.numberOfPledges);
  row++;

  sheet.getRange('A' + row).setValue('Total Number of Calls:');
  sheet.getRange('B' + row).setValue(totals.numberOfCalls);
  row++;

  sheet.getRange('A' + row).setValue('Total Number of Pickups:');
  sheet.getRange('B' + row).setValue(totals.numberOfPickups);
  row++;

  // Calculate some useful metrics
  row++;
  sheet.getRange('A' + row).setValue('CALCULATED METRICS').setFontWeight('bold').setFontSize(12);
  row++;

  var pickupRate = totals.numberOfCalls > 0 ? (totals.numberOfPickups / totals.numberOfCalls * 100) : 0;
  sheet.getRange('A' + row).setValue('Pickup Rate:');
  sheet.getRange('B' + row).setValue(pickupRate.toFixed(2) + '%');
  row++;

  var avgPledgeAmount = totals.numberOfPledges > 0 ? (totals.estimatedPledges / totals.numberOfPledges) : 0;
  sheet.getRange('A' + row).setValue('Avg Pledge Amount:');
  sheet.getRange('B' + row).setValue('$' + avgPledgeAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
  row++;

  var callsPerHour = totals.sessionHours > 0 ? (totals.numberOfCalls / totals.sessionHours) : 0;
  sheet.getRange('A' + row).setValue('Calls Per Hour:');
  sheet.getRange('B' + row).setValue(callsPerHour.toFixed(2));

  // Auto-resize columns
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  // Add timestamp
  row += 2;
  sheet.getRange('A' + row).setValue('Last Updated:');
  sheet.getRange('B' + row).setValue(new Date().toLocaleString());
}

/**
 * Clears all results from the sheet
 */
function clearResults() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  SpreadsheetApp.getUi().alert('Results cleared!');
}
