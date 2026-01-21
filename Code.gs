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
    .addItem('Show Sample Email', 'showSampleEmail')
    .addItem('Clear Results', 'clearResults')
    .addToUi();
}

/**
 * Shows a sample of the most recent email for debugging
 */
function showSampleEmail() {
  var ui = SpreadsheetApp.getUi();

  // Prompt for email address
  var response = ui.prompt('Enter the recipient email address to check:');

  if (response.getSelectedButton() == ui.Button.OK) {
    var emailAddress = response.getResponseText();
    var searchQuery = 'to:' + emailAddress + ' in:sent';
    var threads = GmailApp.search(searchQuery, 0, 1);

    if (threads.length > 0) {
      var message = threads[0].getMessages()[0];
      var body = message.getPlainBody();
      var subject = message.getSubject();
      var date = message.getDate();

      // Show first 2000 characters in a dialog
      var preview = 'Subject: ' + subject + '\n';
      preview += 'Date: ' + date + '\n';
      preview += '---\n';
      preview += body.substring(0, 2000);

      if (body.length > 2000) {
        preview += '\n\n... (truncated, see full text in logs)';
      }

      // Log the full email
      Logger.log('=== FULL EMAIL CONTENT ===');
      Logger.log('Subject: ' + subject);
      Logger.log('Date: ' + date);
      Logger.log('Body:\n' + body);
      Logger.log('=== END EMAIL ===');

      ui.alert('Most Recent Email Preview', preview, ui.ButtonSet.OK);
      ui.alert('Full email logged', 'The complete email content has been logged. Go to Extensions > Apps Script, then View > Executions to see the logs.', ui.ButtonSet.OK);
    } else {
      ui.alert('No emails found to: ' + emailAddress);
    }
  }
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
    var emailsWithMetrics = 0;

    // Initialize totals
    var totals = {
      sessionHours: 0,
      scheduledHours: 0,
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
        totalEmails++;
        var body = message.getPlainBody();

        // Log the email for debugging
        Logger.log('Processing email from: ' + message.getDate());
        Logger.log('Email subject: ' + message.getSubject());
        Logger.log('First 1000 chars: ' + body.substring(0, 1000));

        var metrics = parseEmailMetrics(body);

        if (metrics) {
          emailsWithMetrics++;
          totals.sessionHours += metrics.sessionHours || 0;
          totals.scheduledHours += metrics.scheduledHours || 0;
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
    displayResults(sheet, emailAddress, dateRange, totalEmails, emailsWithMetrics, totals);

    return {
      success: true,
      message: 'Found ' + totalEmails + ' emails (' + emailsWithMetrics + ' with metrics)',
      totals: totals
    };

  } catch (error) {
    Logger.log('Error: ' + error.toString());
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
  var normalizedBody = emailBody.replace(/\u00a0/g, ' ');

  var extractDecimal = function(body, regex) {
    var match = body.match(regex);
    if (!match) {
      return null;
    }
    return parseFloat(match[1].replace(/,/g, ''));
  };

  var extractInteger = function(body, regex) {
    var match = body.match(regex);
    if (!match) {
      return null;
    }
    return parseInt(match[1].replace(/,/g, ''), 10);
  };

  // Pattern for "Session length: X hours" or "Session length: X hrs"
  // Now handles: "2 hours", "2.5 hours", "2hrs", "2 hours (2hrs)", etc.
  var sessionLineMatch = normalizedBody.match(/Session length:\s*([^\n\r]+)/i);
  if (sessionLineMatch) {
    var sessionLine = sessionLineMatch[1];
    var completedHours = extractDecimal(sessionLine, /(\d+(?:\.\d+)?)/);
    if (completedHours !== null) {
      metrics.sessionHours = completedHours;
    }
    var scheduledHours = extractDecimal(sessionLine, /\((?:[^)]*?)(\d+(?:\.\d+)?)/);
    if (scheduledHours !== null) {
      metrics.scheduledHours = scheduledHours;
    }
  } else {
    var sessionHours = extractDecimal(normalizedBody, /Session length:\s*(\d+(?:\.\d+)?)\s*(?:hours?|hrs?|hr)\b/i);
    if (sessionHours !== null) {
      metrics.sessionHours = sessionHours;
    }
  }

  // Pattern for "Total in soft pledges: $XXX"
  // Now handles variations like "Soft pledges: $250", "$250*", "$250* some note", etc.
  var softPledges = extractDecimal(normalizedBody, /(?:Total\s+(?:in\s+)?)?soft pledges:[^0-9$]*\$?(\d+(?:,\d{3})*(?:\.\d+)?)/i);
  if (softPledges !== null) {
    metrics.softPledges = softPledges;
  }

  // Pattern for "Total in hard pledges: $XXX"
  var hardPledges = extractDecimal(normalizedBody, /(?:Total\s+(?:in\s+)?)?hard pledges:[^0-9$]*\$?(\d+(?:,\d{3})*(?:\.\d+)?)/i);
  if (hardPledges !== null) {
    metrics.hardPledges = hardPledges;
  }

  // Pattern for "Total estimated pledges: $XXX"
  var estimatedPledges = extractDecimal(normalizedBody, /(?:Total\s+)?estimated pledges:[^0-9$]*\$?(\d+(?:,\d{3})*(?:\.\d+)?)/i);
  if (estimatedPledges !== null) {
    metrics.estimatedPledges = estimatedPledges;
  }

  // Pattern for "Total number of pledges: X"
  // Now handles asterisks and text after: "3*", "3* two of these...", etc.
  var pledgeCount = extractInteger(normalizedBody, /(?:Total\s+)?number of pledges:[^0-9]*(\d+(?:,\d{3})*)/i);
  if (pledgeCount !== null) {
    metrics.numberOfPledges = pledgeCount;
  }

  // Pattern for "Number of calls: X"
  var calls = extractInteger(normalizedBody, /(?:Number of calls|Calls):[^0-9]*(\d+(?:,\d{3})*)/i);
  if (calls !== null) {
    metrics.numberOfCalls = calls;
  }

  // Pattern for "Number of pickups: X"
  var pickups = extractInteger(normalizedBody, /(?:Number of pickups|Pickups):[^0-9]*(\d+(?:,\d{3})*)/i);
  if (pickups !== null) {
    metrics.numberOfPickups = pickups;
  }

  // Debug logging - add email body snippet if no metrics found
  if (Object.keys(metrics).length === 0) {
    Logger.log('No metrics found in email. First 500 chars: ' + emailBody.substring(0, 500));
  } else {
    Logger.log('Extracted metrics: ' + JSON.stringify(metrics));
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
 * @param {number} emailsWithMetrics - Number of emails with extractable metrics
 * @param {Object} totals - Calculated totals
 */
function displayResults(sheet, emailAddress, dateRange, emailCount, emailsWithMetrics, totals) {
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
  row++;
  sheet.getRange('A' + row).setValue('Emails With Metrics:');
  sheet.getRange('B' + row).setValue(emailsWithMetrics);
  if (emailsWithMetrics < emailCount) {
    sheet.getRange('B' + row).setNote('Some emails were found but did not contain recognizable metrics. Check the Apps Script logs (View > Executions) to see email content.');
  }

  // Add spacing
  row += 2;

  // Display totals
  sheet.getRange('A' + row).setValue('TOTALS').setFontWeight('bold').setFontSize(12);
  row++;

  sheet.getRange('A' + row).setValue('Total Session Hours:');
  sheet.getRange('B' + row).setValue(totals.sessionHours).setNumberFormat('0.00');
  row++;

  sheet.getRange('A' + row).setValue('Total Scheduled Hours:');
  sheet.getRange('B' + row).setValue(totals.scheduledHours).setNumberFormat('0.00');
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
