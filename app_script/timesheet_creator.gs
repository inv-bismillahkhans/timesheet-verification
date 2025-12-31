// Configuration
const EMPLOYEES = [
  "Jinu T J",
  "Bismillakhan S",
  "Midhun",
  "Aravind",
  "Akhil Mohan",
  "Akash T K",
  "Shinoj",
];

const STATUS_OPTIONS = ["In progress", "Completed", "On hold"];

const ACTIVITY_OPTIONS = [
  "Development",
  "System design",
  "Unit testing",
  "Testing",
  "Bug fix(QA)",
  "Bug fix(Customer)",
  "Verification testing(Customer)",
  "Release",
  "Maintenance Tasks",
  "Meeting",
];

const HOLIDAY_DATES = [
            "2026-01-01",
            "2026-01-26",
            "2026-03-20",
            "2026-04-03",
            "2026-04-15",
            "2026-05-01",
            "2026-08-15",
            "2026-08-25",
            "2026-08-26",
            "2026-09-04",
            "2026-10-02",
            "2026-10-20",
            "2026-12-25"
        ];

// Colors
const HEADER_BG = "#FEF2CB"; // Light yellow
const WEEKEND_BG = "#FF9999"; // Light Red 1

// Destination folder ID (optional - leave empty for root Drive)
const DESTINATION_FOLDER_ID =
  PropertiesService.getScriptProperties().getProperty("DESTINATION_FOLDER_ID");
const WEBHOOK_URL =
  PropertiesService.getScriptProperties().getProperty("WEBHOOK_URL");

// Test mode env var: set Script Property "TEST" or "TEST_MODE" to "true", "1" or "yes" to enable
const TEST_ENV = (function () {
  const props = PropertiesService.getScriptProperties();
  const raw = (
    props.getProperty("TEST") || props.getProperty("TEST_MODE") || ""
  ).toString();
  const normalized = raw.trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
})();

/**
 * Main function - Creates the monthly timesheet
 * You can customize the month and year parameters
 * Added optional testMode flag: when true, the function will NOT create any
 * Drive/Spreadsheet files and will only log the actions it would perform.
 */
function createMonthlyTimesheet(monthName = null, year = null, testMode = false) {
  // If the caller did not pass the testMode argument (arguments.length < 3),
  // use the TEST_ENV script property. If the caller passed an explicit value,
  // use that value.
  if (arguments.length < 3) {
    testMode = TEST_ENV;
  } else {
    testMode = !!testMode;
  }

  // Use current month/year if not specified
  if (typeof monthName !== "string" || !monthName) {
    monthName = null;
  }
  if (typeof year !== "number" || !year) {
    year = null;
  }
  const now = new Date();
  if (!monthName) {
    monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  }
  if (!year) {
    year = now.getFullYear();
  }

  Logger.log("Creating timesheet for " + monthName + " " + year + (testMode ? " (TEST MODE)" : ""));

  // Create the workbook
  const workbookName = "IZ_work_report_" + monthName;

  // Get dates for the month
  const dates = getDatesForMonth(monthName, year);
  Logger.log("Generated " + dates.length + " dates");

  if (testMode) {
    // In test mode we do not create any files or modify Drive. Just log intended actions.
    const fakeId = "TEST_SPREADSHEET_ID_" + Utilities.getUuid();
    const url = "https://docs.google.com/spreadsheets/d/" + fakeId;

    Logger.log("TEST MODE: Would create workbook: " + workbookName + " (id=" + fakeId + ")");
    if (DESTINATION_FOLDER_ID) {
      Logger.log("TEST MODE: Would move file to folder ID: " + DESTINATION_FOLDER_ID);
    } else {
      Logger.log("TEST MODE: Would leave file in root Drive");
    }

    // Log the sheets that would be created
    EMPLOYEES.forEach(function (empName) {
      Logger.log("TEST MODE: Would create sheet for employee: " + empName);
    });

    Logger.log("TEST MODE: Would delete default sheet and finalize workbook");
    Logger.log("TEST MODE: Workbook: " + workbookName);
    Logger.log("TEST MODE: Sheets: " + EMPLOYEES.length + " employee tabs");
    Logger.log("TEST MODE: URL: " + url);

    // Update config JSON (no Drive operations in test mode)
    updateConfigJson(fakeId, monthName, year, true);

    // Send (or simulate) notification
    sendChatNotification(monthName, year, url, true);

    return url;
  }

  // Move to destination folder if specified
  const file = DriveApp.getFileById(ss.getId());
  if (DESTINATION_FOLDER_ID) {
    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    file.moveTo(folder);
    Logger.log("Moved to folder: " + folder.getName());
  }

  // Share the file with employee emails from Script Properties
  const emailProp =
    PropertiesService.getScriptProperties().getProperty("EMPLOYEE_EMAILS");
  if (emailProp) {
    const emailList = emailProp
      .split(",")
      .map(function (e) {
        return e.trim();
      })
      .filter(function (e) {
        return e.length > 0;
      });
    emailList.forEach(function (email) {
      try {
        file.addEditor(email);
        Logger.log("Shared with: " + email);
      } catch (err) {
        Logger.log("Failed to share with " + email + ": " + err);
      }
    });
  } else {
    Logger.log("No EMPLOYEE_EMAILS property set in Script Properties.");
  }

  // Get dates for the month
  // const dates = getDatesForMonth(monthName, year); // already obtained above
  Logger.log("Generated " + dates.length + " dates");

  // Delete default "Sheet1"
  const defaultSheet = ss.getSheets()[0];

  // Create employee sheets
  EMPLOYEES.forEach(function (empName) {
    Logger.log("Creating sheet: " + empName);
    setupEmployeeSheet(ss, empName, dates);
  });

  // Delete the default sheet
  ss.deleteSheet(defaultSheet);

  // Open the spreadsheet
  const url = ss.getUrl();
  Logger.log("====================================");
  Logger.log("✓ Timesheet created successfully!");
  Logger.log("====================================");
  Logger.log("Workbook: " + workbookName);
  Logger.log("Sheets: " + EMPLOYEES.length + " employee tabs");
  Logger.log("URL: " + url);
  Logger.log("====================================");

  // Update the config JSON file with new spreadsheet ID
  updateConfigJson(ss.getId(), monthName, year);

  // Send notification to Google Chat
  sendChatNotification(monthName, year, url);

  // Return URL for programmatic use
  return url;
}

/**
 * Updates or creates a JSON config file with spreadsheet IDs by year-month
 * Format: {"2025-December": "sheet_id", "2025-November": "sheet_id", ...}
 * This allows scripts to read the sheet ID for any month
 */
function updateConfigJson(spreadsheetId, monthName, year, testMode = false) {
  const configFileName = "timesheet_config.json";

  // Create key as "year-month" format
  const configKey = year + "-" + monthName;

  try {
    // Build the new mapping entry
    const newEntry = {};
    newEntry[configKey] = spreadsheetId;

    const configContent = JSON.stringify(newEntry, null, 2);

    if (testMode) {
      Logger.log("TEST MODE: Would update config file '" + configFileName + "' in folder: " + (DESTINATION_FOLDER_ID || 'root'));
      Logger.log("TEST MODE: Config content:\n" + configContent);
      return;
    }

    // Determine folder to use
    const folder = DESTINATION_FOLDER_ID
      ? DriveApp.getFolderById(DESTINATION_FOLDER_ID)
      : DriveApp.getRootFolder();

    // Try to find existing config file
    const existingFiles = folder.getFilesByName(configFileName);
    let configFile = null;

    if (existingFiles.hasNext()) {
      configFile = existingFiles.next();
      try {
        const existingContent = configFile.getBlob().getDataAsString();
        const parsed = JSON.parse(existingContent || '{}');

        // Build a merged mapping that keeps only string-valued entries (assumed to be sheet IDs)
        const merged = {};
        for (const k in parsed) {
          if (Object.prototype.hasOwnProperty.call(parsed, k)) {
            if (typeof parsed[k] === 'string') {
              merged[k] = parsed[k];
            }
          }
        }

        // Overwrite/add the new entry
        merged[configKey] = spreadsheetId;

        // Write back only the mapping (no metadata fields)
        configFile.setContent(JSON.stringify(merged, null, 2));
        Logger.log("✓ Updated existing config file: " + configFile.getId());
      } catch (e) {
        Logger.log("Warning: Could not parse existing config, overwriting with new mapping");
        // Overwrite with minimal mapping
        configFile.setContent(JSON.stringify(newEntry, null, 2));
      }
    } else {
      // Create new file containing the minimal mapping
      configFile = folder.createFile(configFileName, JSON.stringify(newEntry, null, 2), MimeType.PLAIN_TEXT);
    }

    Logger.log("Config file URL: " + configFile.getUrl());
    Logger.log("Config file ID: " + configFile.getId());

    // Store the config file ID in Script Properties for easy access
    try {
      PropertiesService.getScriptProperties().setProperty("CONFIG_FILE_ID", configFile.getId());
    } catch (e) {
      Logger.log("Warning: Could not set CONFIG_FILE_ID script property: " + e);
    }
  } catch (err) {
    Logger.log("ERROR updating config file: " + err);
  }
}

/**
 * Generate list of dates for the given month
 */
function getDatesForMonth(monthName, year) {
  // Parse month name to month number
  const monthMap = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11,
    Jan: 0,
    Feb: 1,
    Mar: 2,
    Apr: 3,
    Jun: 5,
    Jul: 6,
    Aug: 7,
    Sep: 8,
    Oct: 9,
    Nov: 10,
    Dec: 11,
  };

  const monthNum = monthMap[monthName];
  if (monthNum === undefined) {
    throw new Error("Invalid month name: " + monthName);
  }

  // Get number of days in month
  const lastDay = new Date(year, monthNum + 1, 0).getDate();

  // Get month abbreviation
  const monthAbbr = Utilities.formatDate(
    new Date(year, monthNum, 1),
    Session.getScriptTimeZone(),
    "MMM"
  );

  const dates = [];
  for (let day = 1; day <= lastDay; day++) {
    const date = new Date(year, monthNum, day);
    const dayOfWeek = date.getDay(); // 0=Sunday, 6=Saturday
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const dateStr = day + "-" + monthAbbr;

    dates.push({
      date: dateStr,
      isWeekend: isWeekend,
      dateObj: date, // include Date object for holiday checks
    });
  }

  return dates;
}

// Helper: returns true if HOLIDAY_DATES is defined and contains the given date
function isHoliday(dateObj) {
  try {
    if (typeof HOLIDAY_DATES === 'undefined' || !Array.isArray(HOLIDAY_DATES)) {
      return false;
    }

    const tz = Session.getScriptTimeZone();
    // Candidate formats to match common holiday formats
    const candidates = [
      Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'), // 2025-12-25
      Utilities.formatDate(dateObj, tz, 'd-MMM'), // 25-Dec
      Utilities.formatDate(dateObj, tz, 'dd-MMM'), // 25-Dec (leading zero if any)
      Utilities.formatDate(dateObj, tz, 'd-MMMM'), // 25-December
      Utilities.formatDate(dateObj, tz, 'dd-MMMM'),
    ].map(function(s){ return (s||'').toString().trim().toLowerCase(); });

    // Normalize holiday entries and check
    for (let i = 0; i < HOLIDAY_DATES.length; i++) {
      const raw = (HOLIDAY_DATES[i] || '').toString().trim().toLowerCase();
      if (!raw) continue;
      if (candidates.indexOf(raw) !== -1) return true;

      // Also allow matching just day/month numeric like '12-25' or '25-12'
      // Normalize numeric forms in candidate as 'M-D' and 'D-M'
      const numeric1 = Utilities.formatDate(dateObj, tz, 'M-d'); // 12-25
      const numeric2 = Utilities.formatDate(dateObj, tz, 'd-M'); // 25-12
      if (raw === numeric1.toLowerCase() || raw === numeric2.toLowerCase()) return true;
    }

    return false;
  } catch (e) {
    // On any unexpected error, don't treat as holiday
    Logger.log('isHoliday error: ' + e);
    return false;
  }
}

/**
 * Create and format an individual employee timesheet
 */
function setupEmployeeSheet(ss, empName, dates) {
  // Create new sheet
  const sheet = ss.insertSheet(empName);

  // Set column widths
  sheet.setColumnWidth(1, 30); // A
  sheet.setColumnWidth(2, 80); // B - Date
  sheet.setColumnWidth(3, 400); // C - Module/Area
  sheet.setColumnWidth(4, 400); // D - Task details
  sheet.setColumnWidth(5, 100); // E - Status
  sheet.setColumnWidth(6, 150); // F - Activity Type
  sheet.setColumnWidth(7, 70); // G - Start
  sheet.setColumnWidth(8, 70); // H - End
  sheet.setColumnWidth(9, 70); // I - Task
  sheet.setColumnWidth(10, 70); // J - Total
  sheet.setColumnWidth(11, 420); // K - Remarks

  // Set Calibri font for the whole sheet
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily("Calibri");

  // Row 2: Title
  sheet.getRange("B2:J2").merge();
  sheet
    .getRange("B2")
    .setValue("Interview Zero- " + empName)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(12);

  // Row 4: Headers with proper structure
  // Main headers row
  const mainHeaders = [
    "",
    "Date",
    "Module/Area",
    "Task details/Ticket number",
    "Status",
    "Activity Type",
    "",
    "",
    "",
    "",
    "Remarks",
  ];
  sheet
    .getRange(4, 1, 1, mainHeaders.length)
    .setValues([mainHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge cells for "Time" header (columns G-H)
  sheet.getRange(4, 7, 1, 2).merge();
  sheet.getRange(4, 7).setValue("Time");

  // Merge cells for "Duration" header (columns I-J)
  sheet.getRange(4, 9, 1, 2).merge();
  sheet.getRange(4, 9).setValue("Duration");

  // Row 5: Sub-headers for Time and Duration
  const subHeaders = [
    "",
    "",
    "",
    "",
    "",
    "",
    "Start",
    "End",
    "Task",
    "Total",
    "",
  ];
  sheet
    .getRange(5, 1, 1, subHeaders.length)
    .setValues([subHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge Date, Module/Area, Task details, Status, Activity Type, and Remarks vertically (rows 4-5)
  sheet.getRange(4, 2, 2, 1).merge(); // Date
  sheet.getRange(4, 3, 2, 1).merge(); // Module/Area
  sheet.getRange(4, 4, 2, 1).merge(); // Task details
  sheet.getRange(4, 5, 2, 1).merge(); // Status
  sheet.getRange(4, 6, 2, 1).merge(); // Activity Type
  sheet.getRange(4, 11, 2, 1).merge(); // Remarks

  // Data rows: 2 rows per date starting from row 6 (after headers)
  let currentRow = 6;

  dates.forEach(function (dateInfo) {
    const dateStr = dateInfo.date;
    const isWeekend = dateInfo.isWeekend;
    const dateObj = dateInfo.dateObj; // available from getDatesForMonth

    // Merge date cells vertically (2 rows) and center align
    const dateRange = sheet.getRange(currentRow, 2, 2, 1);
    dateRange.merge();
    dateRange
      .setValue(dateStr)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Add task duration formulas for both rows
    for (let offset = 0; offset < 2; offset++) {
      const row = currentRow + offset;
      const taskFormula =
        "=IF(AND(G" +
        row +
        "<TIME(12,30,0),H" +
        row +
        ">TIME(13,0,0)),H" +
        row +
        "-G" +
        row +
        "-TIME(0,30,0),H" +
        row +
        "-G" +
        row +
        ")";
      sheet.getRange(row, 9).setFormula(taskFormula);
    }

    // Add total formula in first row only (sums both task rows) and merge vertically
    const totalFormula = "=SUM(I" + currentRow + ":I" + (currentRow + 1) + ")";
    const totalRange = sheet.getRange(currentRow, 10, 2, 1);
    totalRange.merge();
    totalRange
      .setFormula(totalFormula)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setFontSize(11);

    // Format Start and End as clock time
    sheet.getRange(currentRow, 7, 2, 2).setNumberFormat("hh:mm");

    // Format Task and Total as duration hh:mm
    sheet.getRange(currentRow, 9, 2, 2).setNumberFormat("hh:mm");

    // Color weekend rows or holidays
    if (isWeekend || isHoliday(dateObj)) {
      sheet.getRange(currentRow, 1, 2, 11).setBackground(WEEKEND_BG);
    }

    // Add data validation for Status (column E)
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPTIONS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 5, 2, 1).setDataValidation(statusRule);

    // Add data validation for Activity Type (column F)
    const activityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ACTIVITY_OPTIONS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 6, 2, 1).setDataValidation(activityRule);

    currentRow += 2;
  });

  const lastDataRow = currentRow - 1;

  // Add "Monthly hours spend" row
  const monthlyRow = currentRow + 1;
  sheet.getRange(monthlyRow, 2, 1, 8).merge();
  sheet
    .getRange(monthlyRow, 2)
    .setValue("Monthly hours spend")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  // Monthly total formula
  const monthlyFormula =
    "=TEXT(INT(SUM(J6:J" +
    lastDataRow +
    "))*24+HOUR(SUM(J6:J" +
    lastDataRow +
    ')),"00")&":"&' +
    "TEXT(MINUTE(SUM(J6:J" +
    lastDataRow +
    ')),"00")';
  sheet
    .getRange(monthlyRow, 10)
    .setFormula(monthlyFormula)
    .setFontWeight("bold");

  // Add borders to all cells
  const allDataRange = sheet.getRange(4, 1, monthlyRow - 3, 11);
  allDataRange.setBorder(true, true, true, true, true, true);

  // Set font size 11 for all data rows (from row 6 to last data row)
  sheet
    .getRange(6, 1, sheet.getMaxRows() - 5, sheet.getMaxColumns())
    .setFontSize(11);
  // Center-align text in column I (Task column)
  sheet
    .getRange(6, 7, sheet.getMaxRows() - 5, 3)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  Logger.log("✓ " + empName + " completed (" + dates.length + " dates)");
}

/**
 * Create a custom menu when opening a spreadsheet
 * (Optional - only works if this script is bound to a spreadsheet)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Timesheet Tools")
    .addItem("Create Monthly Timesheet", "createMonthlyTimesheet")
    .addItem("Create Monthly Timesheet (Test)", "createMonthlyTimesheetTest")
    .addToUi();
}

// Wrapper to force test mode from menu/UI
function createMonthlyTimesheetTest() {
  createMonthlyTimesheet(null, null, true);
}

/**
 * Test function - creates timesheet for current month
 */
function testCreateCurrentMonth() {
  createMonthlyTimesheet();
}

/**
 * Test function - creates timesheet for next month
 */
function testCreateNextMonth() {
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const monthName = Utilities.formatDate(
    nextMonth,
    Session.getScriptTimeZone(),
    "MMMM"
  );
  const year = nextMonth.getFullYear();

  createMonthlyTimesheet(monthName, year);
}

/**
 * Send notification to Google Chat when a new timesheet is created
 */
function sendChatNotification(monthName, year, spreadsheetUrl, testMode = false) {
  if (!WEBHOOK_URL) {
    Logger.log(
      "WEBHOOK_URL not configured in Script Properties. Skipping notification."
    );
    return;
  }

  if (testMode) {
    Logger.log("TEST MODE: Would send Google Chat notification for " + monthName + " " + year + ". URL: " + spreadsheetUrl);
    return;
  }

  try {
    // Create the message payload for Google Chat with proper @all mention
    const message = {
      text:
        "<users/all> Timesheet for the month of " +
        monthName +
        " " +
        year +
        " Generated.\n" +
        spreadsheetUrl,
    };

    // Send POST request to webhook
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(message),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      Logger.log("✓ Google Chat notification sent successfully");
    } else {
      Logger.log(
        "⚠ Google Chat notification failed with status: " + responseCode
      );
      Logger.log("Response: " + response.getContentText());
    }
  } catch (err) {
    Logger.log("ERROR sending Google Chat notification: " + err);
  }
}
