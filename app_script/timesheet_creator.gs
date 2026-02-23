// Language: javascript
// File: `app_script/multi_team_timesheet_creator.gs`
// Multi-Team Timesheet Generator - Reads configuration from Master Sheet

// ============================================================================
// CONFIGURATION
// ============================================================================

// Master Configuration Sheet ID - Set this in Script Properties as "MASTER_SHEET_ID"
// or hardcode it here
const MASTER_SHEET_ID = PropertiesService.getScriptProperties().getProperty("MASTER_SHEET_ID") || "";

// Master Sheet Tab Name
const MASTER_SHEET_TAB = "Teams Config";

// Default options (can be overridden per team in master sheet)
const DEFAULT_STATUS_OPTIONS = ["In progress", "Completed", "On hold"];

const DEFAULT_ACTIVITY_OPTIONS = [
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

const DEFAULT_HOLIDAY_DATES = [
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

// Test mode env var
const TEST_ENV = (function () {
  const props = PropertiesService.getScriptProperties();
  const raw = (
    props.getProperty("TEST") || props.getProperty("TEST_MODE") || ""
  ).toString();
  const normalized = raw.trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
})();

// ============================================================================
// MASTER SHEET READING
// ============================================================================

/**
 * Read team configurations from the master sheet
 * Expected columns:
 * - Project Name
 * - Team Members (comma-separated)
 * - Drive Folder ID
 * - Webhook URL
 * - Employee Emails (comma-separated)
 * - Status Options (optional, comma-separated)
 * - Activity Options (optional, comma-separated)
 * - Holiday Dates (optional, comma-separated)
 *
 * @returns {Array} Array of team configuration objects
 */
function readTeamConfigurations() {
  if (!MASTER_SHEET_ID) {
    throw new Error("MASTER_SHEET_ID not configured. Please set it in Script Properties or hardcode it.");
  }

  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(MASTER_SHEET_TAB);

    if (!sheet) {
      throw new Error("Sheet '" + MASTER_SHEET_TAB + "' not found in master spreadsheet");
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      throw new Error("Master sheet must have header row and at least one data row");
    }

    // Read header row to map columns
    const headers = data[0].map(function(h) { return h.toString().trim(); });

    const colMap = {
      projectName: findColumnIndex(headers, ["Project Name", "ProjectName", "Project"]),
      teamMembers: findColumnIndex(headers, ["Team Members", "TeamMembers", "Members", "Employees"]),
      driveFolderId: findColumnIndex(headers, ["Drive Folder ID", "DriveFolderID", "Folder ID", "FolderID"]),
      webhookUrl: findColumnIndex(headers, ["Webhook URL", "WebhookURL", "Webhook", "Chat Webhook"]),
      employeeEmails: findColumnIndex(headers, ["Employee Emails", "EmployeeEmails", "Emails"]),
      statusOptions: findColumnIndex(headers, ["Status Options", "StatusOptions"]),
      activityOptions: findColumnIndex(headers, ["Activity Options", "ActivityOptions"]),
      holidayDates: findColumnIndex(headers, ["Holiday Dates", "HolidayDates", "Holidays"]),
      enabled: findColumnIndex(headers, ["Enabled (Create)", "Enabled", "Active"])
    };

    // Validate required columns
    if (colMap.projectName === -1) {
      throw new Error("Required column 'Project Name' not found in master sheet");
    }
    if (colMap.teamMembers === -1) {
      throw new Error("Required column 'Team Members' not found in master sheet");
    }

    // Parse data rows
    const teams = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[colMap.projectName]) {
        continue;
      }

      // Check if enabled (if column exists)
      if (colMap.enabled !== -1) {
        const enabledValue = row[colMap.enabled].toString().trim().toLowerCase();
        if (enabledValue === "no" || enabledValue === "false" || enabledValue === "0") {
          Logger.log("Skipping disabled team: " + row[colMap.projectName]);
          continue;
        }
      }

      const team = {
        projectName: row[colMap.projectName].toString().trim(),
        employees: parseListField(row[colMap.teamMembers]),
        driveFolderId: colMap.driveFolderId !== -1 ? row[colMap.driveFolderId].toString().trim() : "",
        webhookUrl: colMap.webhookUrl !== -1 ? row[colMap.webhookUrl].toString().trim() : "",
        employeeEmails: colMap.employeeEmails !== -1 ? parseListField(row[colMap.employeeEmails]) : [],
        statusOptions: colMap.statusOptions !== -1 && row[colMap.statusOptions]
          ? parseListField(row[colMap.statusOptions])
          : DEFAULT_STATUS_OPTIONS,
        activityOptions: colMap.activityOptions !== -1 && row[colMap.activityOptions]
          ? parseListField(row[colMap.activityOptions])
          : DEFAULT_ACTIVITY_OPTIONS,
        holidayDates: colMap.holidayDates !== -1 && row[colMap.holidayDates]
          ? parseListField(row[colMap.holidayDates])
          : DEFAULT_HOLIDAY_DATES
      };

      // Validate team has employees
      if (team.employees.length === 0) {
        Logger.log("WARNING: Team '" + team.projectName + "' has no team members. Skipping.");
        continue;
      }

      teams.push(team);
    }

    Logger.log("Loaded " + teams.length + " team configurations from master sheet");
    return teams;

  } catch (err) {
    Logger.log("ERROR reading master sheet: " + err);
    throw err;
  }
}

/**
 * Find column index by matching header names (case-insensitive)
 */
function findColumnIndex(headers, possibleNames) {
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toLowerCase();
    for (let j = 0; j < possibleNames.length; j++) {
      if (header === possibleNames[j].toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Parse comma-separated or newline-separated list field
 */
function parseListField(value) {
  if (!value) return [];

  const str = value.toString().trim();
  if (!str) return [];

  // Split by comma or newline
  const items = str.split(/[,\n]+/).map(function(item) {
    return item.trim();
  }).filter(function(item) {
    return item.length > 0;
  });

  return items;
}

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Create timesheets for ALL teams from master configuration
 */
function createAllTeamTimesheets(monthName = null, year = null, testMode = false) {
  // Determine test mode
  if (arguments.length < 3) {
    testMode = TEST_ENV;
  } else {
    testMode = !!testMode;
  }

  // Use current month/year if not specified
  const now = new Date();
  if (!monthName) {
    monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  }
  if (!year) {
    year = now.getFullYear();
  }

  Logger.log("=".repeat(60));
  Logger.log("Creating timesheets for ALL teams - " + monthName + " " + year + (testMode ? " (TEST MODE)" : ""));
  Logger.log("=".repeat(60));

  try {
    const teams = readTeamConfigurations();

    if (teams.length === 0) {
      Logger.log("No teams found in master configuration. Nothing to create.");
      return;
    }

    const results = [];

    teams.forEach(function(team, index) {
      Logger.log("\n" + "-".repeat(60));
      Logger.log("Processing team " + (index + 1) + "/" + teams.length + ": " + team.projectName);
      Logger.log("-".repeat(60));

      try {
        const url = createTeamTimesheet(team, monthName, year, testMode);
        results.push({
          projectName: team.projectName,
          success: true,
          url: url
        });
      } catch (err) {
        Logger.log("ERROR creating timesheet for " + team.projectName + ": " + err);
        results.push({
          projectName: team.projectName,
          success: false,
          error: err.toString()
        });
      }
    });

    // Summary
    Logger.log("\n" + "=".repeat(60));
    Logger.log("SUMMARY");
    Logger.log("=".repeat(60));
    const successful = results.filter(function(r) { return r.success; }).length;
    Logger.log("Total teams: " + results.length);
    Logger.log("Successful: " + successful);
    Logger.log("Failed: " + (results.length - successful));

    results.forEach(function(result) {
      if (result.success) {
        Logger.log("✓ " + result.projectName + " - " + result.url);
      } else {
        Logger.log("✗ " + result.projectName + " - " + result.error);
      }
    });
    Logger.log("=".repeat(60));

    return results;

  } catch (err) {
    Logger.log("FATAL ERROR: " + err);
    throw err;
  }
}

/**
 * Create timesheet for a specific team by project name
 */
function createTimesheetByProjectName(projectName, monthName = null, year = null, testMode = false) {
  // Determine test mode
  if (arguments.length < 4) {
    testMode = TEST_ENV;
  } else {
    testMode = !!testMode;
  }

  const teams = readTeamConfigurations();
  const team = teams.filter(function(t) {
    return t.projectName.toLowerCase() === projectName.toLowerCase();
  })[0];

  if (!team) {
    throw new Error("Project '" + projectName + "' not found in master configuration");
  }

  // Use current month/year if not specified
  const now = new Date();
  if (!monthName) {
    monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  }
  if (!year) {
    year = now.getFullYear();
  }

  return createTeamTimesheet(team, monthName, year, testMode);
}

/**
 * Create timesheet for a single team
 */
function createTeamTimesheet(teamConfig, monthName, year, testMode) {
  Logger.log("Creating timesheet for project: " + teamConfig.projectName);
  Logger.log("Team members: " + teamConfig.employees.join(", "));

  // Create the workbook
  const workbookName = teamConfig.projectName + "_work_report_" + monthName;

  // Get dates for the month
  const dates = getDatesForMonth(monthName, year);
  Logger.log("Generated " + dates.length + " dates");

  if (testMode) {
    // Test mode - no actual file creation
    const fakeId = "TEST_SPREADSHEET_ID_" + Utilities.getUuid();
    const url = "https://docs.google.com/spreadsheets/d/" + fakeId;

    Logger.log("TEST MODE: Would create workbook: " + workbookName + " (id=" + fakeId + ")");
    if (teamConfig.driveFolderId) {
      Logger.log("TEST MODE: Would move file to folder ID: " + teamConfig.driveFolderId);
    } else {
      Logger.log("TEST MODE: Would leave file in root Drive");
    }

    teamConfig.employees.forEach(function (empName) {
      Logger.log("TEST MODE: Would create sheet for employee: " + empName);
    });

    Logger.log("TEST MODE: Would delete default sheet and finalize workbook");
    Logger.log("TEST MODE: Workbook: " + workbookName);
    Logger.log("TEST MODE: Sheets: " + teamConfig.employees.length + " employee tabs");
    Logger.log("TEST MODE: URL: " + url);

    // Update config JSON
    updateConfigJson(fakeId, monthName, year, teamConfig.projectName, teamConfig.driveFolderId, true);

    // Send notification
    sendChatNotification(teamConfig, monthName, year, url, true);

    return url;
  }

  // Create the actual spreadsheet
  const ss = SpreadsheetApp.create(workbookName);
  Logger.log("Created spreadsheet id: " + ss.getId());

  // Move to destination folder if specified
  const file = DriveApp.getFileById(ss.getId());
  if (teamConfig.driveFolderId) {
    try {
      const folder = DriveApp.getFolderById(teamConfig.driveFolderId);
      file.moveTo(folder);
      Logger.log("Moved to folder: " + folder.getName());
    } catch (err) {
      Logger.log("WARNING: Failed to move file to folder: " + err);
    }
  }

  // Share the file with employee emails
  if (teamConfig.employeeEmails && teamConfig.employeeEmails.length > 0) {
    teamConfig.employeeEmails.forEach(function (email) {
      try {
        file.addEditor(email);
        Logger.log("Shared with: " + email);
      } catch (err) {
        Logger.log("Failed to share with " + email + ": " + err);
      }
    });
  }

  // Delete default "Sheet1"
  const defaultSheet = ss.getSheets()[0];

  // Create employee sheets
  teamConfig.employees.forEach(function (empName) {
    Logger.log("Creating sheet: " + empName);
    setupEmployeeSheet(
      ss,
      empName,
      dates,
      teamConfig.projectName,
      teamConfig.statusOptions,
      teamConfig.activityOptions,
      teamConfig.holidayDates
    );
  });

  // Delete the default sheet
  ss.deleteSheet(defaultSheet);

  // Get URL
  const url = ss.getUrl();
  Logger.log("====================================");
  Logger.log("✓ Timesheet created successfully!");
  Logger.log("====================================");
  Logger.log("Project: " + teamConfig.projectName);
  Logger.log("Workbook: " + workbookName);
  Logger.log("Sheets: " + teamConfig.employees.length + " employee tabs");
  Logger.log("URL: " + url);
  Logger.log("====================================");

  // Update the config JSON file
  updateConfigJson(ss.getId(), monthName, year, teamConfig.projectName, teamConfig.driveFolderId);

  // Send notification
  sendChatNotification(teamConfig, monthName, year, url);

  return url;
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Generate list of dates for the given month
 */
function getDatesForMonth(monthName, year) {
  const monthMap = {
    January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
    July: 6, August: 7, September: 8, October: 9, November: 10, December: 11,
    Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
    Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11,
  };

  const monthNum = monthMap[monthName];
  if (monthNum === undefined) {
    throw new Error("Invalid month name: " + monthName);
  }

  const lastDay = new Date(year, monthNum + 1, 0).getDate();
  const monthAbbr = Utilities.formatDate(
    new Date(year, monthNum, 1),
    Session.getScriptTimeZone(),
    "MMM"
  );

  const dates = [];
  for (let day = 1; day <= lastDay; day++) {
    const date = new Date(year, monthNum, day);
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const dateStr = day + "-" + monthAbbr;

    dates.push({
      date: dateStr,
      isWeekend: isWeekend,
      dateObj: date,
    });
  }

  return dates;
}

/**
 * Check if date is a holiday
 */
function isHoliday(dateObj, holidayDates) {
  try {
    if (!Array.isArray(holidayDates) || holidayDates.length === 0) {
      return false;
    }

    const tz = Session.getScriptTimeZone();
    const candidates = [
      Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(dateObj, tz, 'd-MMM'),
      Utilities.formatDate(dateObj, tz, 'dd-MMM'),
      Utilities.formatDate(dateObj, tz, 'd-MMMM'),
      Utilities.formatDate(dateObj, tz, 'dd-MMMM'),
    ].map(function(s){ return (s||'').toString().trim().toLowerCase(); });

    for (let i = 0; i < holidayDates.length; i++) {
      const raw = (holidayDates[i] || '').toString().trim().toLowerCase();
      if (!raw) continue;
      if (candidates.indexOf(raw) !== -1) return true;

      const numeric1 = Utilities.formatDate(dateObj, tz, 'M-d');
      const numeric2 = Utilities.formatDate(dateObj, tz, 'd-M');
      if (raw === numeric1.toLowerCase() || raw === numeric2.toLowerCase()) return true;
    }

    return false;
  } catch (e) {
    Logger.log('isHoliday error: ' + e);
    return false;
  }
}

/**
 * Create and format an individual employee timesheet
 */
function setupEmployeeSheet(ss, empName, dates, projectName, statusOptions, activityOptions, holidayDates) {
  const sheet = ss.insertSheet(empName);

  // Set column widths
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 400);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 70);
  sheet.setColumnWidth(8, 70);
  sheet.setColumnWidth(9, 70);
  sheet.setColumnWidth(10, 70);
  sheet.setColumnWidth(11, 420);

  // Set font
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily("Calibri");

  // Title
  sheet.getRange("B2:J2").merge();
  sheet
    .getRange("B2")
    .setValue(projectName + " - " + empName)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(12);

  // Headers
  const mainHeaders = [
    "", "Date", "Module/Area", "Task details/Ticket number",
    "Status", "Activity Type", "", "", "", "", "Remarks"
  ];
  sheet
    .getRange(4, 1, 1, mainHeaders.length)
    .setValues([mainHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  sheet.getRange(4, 7, 1, 2).merge();
  sheet.getRange(4, 7).setValue("Time");
  sheet.getRange(4, 9, 1, 2).merge();
  sheet.getRange(4, 9).setValue("Duration");

  const subHeaders = ["", "", "", "", "", "", "Start", "End", "Task", "Total", ""];
  sheet
    .getRange(5, 1, 1, subHeaders.length)
    .setValues([subHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  sheet.getRange(4, 2, 2, 1).merge();
  sheet.getRange(4, 3, 2, 1).merge();
  sheet.getRange(4, 4, 2, 1).merge();
  sheet.getRange(4, 5, 2, 1).merge();
  sheet.getRange(4, 6, 2, 1).merge();
  sheet.getRange(4, 11, 2, 1).merge();

  // Data rows
  let currentRow = 6;

  dates.forEach(function (dateInfo) {
    const dateStr = dateInfo.date;
    const isWeekend = dateInfo.isWeekend;
    const dateObj = dateInfo.dateObj;

    const dateRange = sheet.getRange(currentRow, 2, 2, 1);
    dateRange.merge();
    dateRange
      .setValue(dateStr)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    for (let offset = 0; offset < 2; offset++) {
      const row = currentRow + offset;
      const taskFormula =
        "=IF(AND(G" + row + "<TIME(12,30,0),H" + row + ">TIME(13,0,0)),H" +
        row + "-G" + row + "-TIME(0,30,0),H" + row + "-G" + row + ")";
      sheet.getRange(row, 9).setFormula(taskFormula);
    }

    const totalFormula = "=SUM(I" + currentRow + ":I" + (currentRow + 1) + ")";
    const totalRange = sheet.getRange(currentRow, 10, 2, 1);
    totalRange.merge();
    totalRange
      .setFormula(totalFormula)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setFontSize(11);

    sheet.getRange(currentRow, 7, 2, 2).setNumberFormat("hh:mm");
    sheet.getRange(currentRow, 9, 2, 2).setNumberFormat("hh:mm");

    if (isWeekend || isHoliday(dateObj, holidayDates)) {
      sheet.getRange(currentRow, 1, 2, 11).setBackground(WEEKEND_BG);
    }

    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusOptions, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 5, 2, 1).setDataValidation(statusRule);

    const activityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(activityOptions, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 6, 2, 1).setDataValidation(activityRule);

    currentRow += 2;
  });

  const lastDataRow = currentRow - 1;

  const monthlyRow = currentRow + 1;
  sheet.getRange(monthlyRow, 2, 1, 8).merge();
  sheet
    .getRange(monthlyRow, 2)
    .setValue("Monthly hours spend")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  const monthlyFormula =
    "=TEXT(INT(SUM(J6:J" + lastDataRow + "))*24+HOUR(SUM(J6:J" + lastDataRow +
    ')),"00")&":"&TEXT(MINUTE(SUM(J6:J' + lastDataRow + ')),"00")';
  sheet
    .getRange(monthlyRow, 10)
    .setFormula(monthlyFormula)
    .setFontWeight("bold");

  const allDataRange = sheet.getRange(4, 1, monthlyRow - 3, 11);
  allDataRange.setBorder(true, true, true, true, true, true);

  sheet
    .getRange(6, 1, sheet.getMaxRows() - 5, sheet.getMaxColumns())
    .setFontSize(11);
  sheet
    .getRange(6, 7, sheet.getMaxRows() - 5, 3)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  Logger.log("✓ " + empName + " completed (" + dates.length + " dates)");
}

// ============================================================================
// CONFIG AND NOTIFICATION FUNCTIONS
// ============================================================================

/**
 * Updates or creates a JSON config file with spreadsheet IDs
 * Now includes project name in the key: "ProjectName_2026-January"
 */
function updateConfigJson(spreadsheetId, monthName, year, projectName, driveFolderId, testMode = false) {
  const configFileName = "timesheet_config.json";
  const configKey = projectName + "_" + year + "-" + monthName;

  try {
    const newEntry = {};
    newEntry[configKey] = spreadsheetId;

    const configContent = JSON.stringify(newEntry, null, 2);

    if (testMode) {
      Logger.log("TEST MODE: Would update config file '" + configFileName + "' in folder: " + (driveFolderId || 'root'));
      Logger.log("TEST MODE: Config content:\n" + configContent);
      return;
    }

    const folder = driveFolderId
      ? DriveApp.getFolderById(driveFolderId)
      : DriveApp.getRootFolder();

    const existingFiles = folder.getFilesByName(configFileName);
    let configFile = null;

    if (existingFiles.hasNext()) {
      configFile = existingFiles.next();
      try {
        const existingContent = configFile.getBlob().getDataAsString();
        const parsed = JSON.parse(existingContent || '{}');

        const merged = {};
        for (const k in parsed) {
          if (Object.prototype.hasOwnProperty.call(parsed, k)) {
            if (typeof parsed[k] === 'string') {
              merged[k] = parsed[k];
            }
          }
        }

        merged[configKey] = spreadsheetId;
        const blob = Utilities.newBlob(JSON.stringify(merged, null, 2), 'text/plain', configFileName);
        configFile.setContent(blob.getDataAsString());
        Logger.log("✓ Updated existing config file: " + configFile.getId());
        Logger.log("Config file MIME type: " + configFile.getMimeType());
      } catch (e) {
        Logger.log("Warning: Could not parse existing config, overwriting");
        const blob = Utilities.newBlob(JSON.stringify(newEntry, null, 2), 'text/plain', configFileName);
        configFile.setContent(blob.getDataAsString());
      }
    } else {
      // Create as blob to ensure plain text MIME type
      const blob = Utilities.newBlob(JSON.stringify(newEntry, null, 2), 'text/plain', configFileName);
      configFile = folder.createFile(blob);
      Logger.log("✓ Created new config file: " + configFile.getId());
    }

    Logger.log("Config file URL: " + configFile.getUrl());
    Logger.log("Config file MIME type: " + configFile.getMimeType());
  } catch (err) {
    Logger.log("ERROR updating config file: " + err);
  }
}

/**
 * Send notification to Google Chat
 */
function sendChatNotification(teamConfig, monthName, year, spreadsheetUrl, testMode = false) {
  if (!teamConfig.webhookUrl) {
    Logger.log("No webhook URL configured for " + teamConfig.projectName + ". Skipping notification.");
    return;
  }

  if (testMode) {
    Logger.log("TEST MODE: Would send Google Chat notification for " + teamConfig.projectName +
               " - " + monthName + " " + year + ". URL: " + spreadsheetUrl);
    return;
  }

  try {
    const message = {
      text:
        "<users/all> Timesheet for " + teamConfig.projectName +
        " - " + monthName + " " + year + " Generated.\n" + spreadsheetUrl,
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(message),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(teamConfig.webhookUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      Logger.log("✓ Google Chat notification sent successfully for " + teamConfig.projectName);
    } else {
      Logger.log("⚠ Google Chat notification failed for " + teamConfig.projectName +
                 " with status: " + responseCode);
    }
  } catch (err) {
    Logger.log("ERROR sending Google Chat notification for " + teamConfig.projectName + ": " + err);
  }
}

// ============================================================================
// UI FUNCTIONS
// ============================================================================

/**
 * Create custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Multi-Team Timesheet")
    .addItem("Create All Team Timesheets", "menuCreateAllTeams")
    .addItem("Create All Team Timesheets (Test)", "menuCreateAllTeamsTest")
    .addSeparator()
    .addItem("Create for Specific Project...", "menuCreateSpecificProject")
    .addSeparator()
    .addItem("Reload Team Configurations", "menuReloadConfig")
    .addToUi();
}

function menuCreateAllTeams() {
  createAllTeamTimesheets();
}

function menuCreateAllTeamsTest() {
  createAllTeamTimesheets(null, null, true);
}

function menuCreateSpecificProject() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Create Timesheet for Specific Project',
    'Enter the project name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const projectName = result.getResponseText().trim();
    if (projectName) {
      try {
        createTimesheetByProjectName(projectName);
        ui.alert('Success', 'Timesheet created for ' + projectName, ui.ButtonSet.OK);
      } catch (err) {
        ui.alert('Error', err.toString(), ui.ButtonSet.OK);
      }
    }
  }
}

function menuReloadConfig() {
  try {
    const teams = readTeamConfigurations();
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Configuration Loaded',
      'Found ' + teams.length + ' team(s):\n' +
      teams.map(function(t) { return '• ' + t.projectName; }).join('\n'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error', err.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

function testCreateAllTeams() {
  createAllTeamTimesheets();
}

function testCreateAllTeamsTestMode() {
  createAllTeamTimesheets(null, null, true);
}

function testCreateNextMonthAllTeams() {
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const monthName = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), "MMMM");
  const year = nextMonth.getFullYear();
  createAllTeamTimesheets(monthName, year);
}

function testReadConfig() {
  const teams = readTeamConfigurations();
  Logger.log("Found " + teams.length + " teams:");
  teams.forEach(function(team) {
    Logger.log("\nProject: " + team.projectName);
    Logger.log("Employees: " + team.employees.join(", "));
    Logger.log("Drive Folder ID: " + team.driveFolderId);
    Logger.log("Webhook: " + (team.webhookUrl ? "Configured" : "Not configured"));
    Logger.log("Emails: " + team.employeeEmails.join(", "));
  });
}