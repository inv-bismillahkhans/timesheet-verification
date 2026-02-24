// Language: javascript
// File: `app_script/create_test_master_sheet.gs`
// Creates a test Master Sheet with multiple projects for use with timesheet_creator.gs

// ============================================================================
// CONFIGURATION
// ============================================================================

// Set this to an existing folder ID to place the master sheet there,
// or leave empty to create in root Drive.
const TEST_MASTER_FOLDER_ID = "";

// Name of the spreadsheet to create
const TEST_MASTER_SHEET_NAME = "Timesheet Master Config (Test)";

// Tab name must match MASTER_SHEET_TAB in timesheet_creator.gs
const TEST_MASTER_TAB_NAME = "Teams Config";

// ============================================================================
// SAMPLE DATA
// ============================================================================

const SAMPLE_TEAMS = [
  {
    projectName:            "Backend Team",
    teamMembers:            "Alice Johnson, Bob Smith, Charlie Davis",
    driveFolderId:          "",
    webhookUrl:             "",
    employeeEmails:         "alice@example.com, bob@example.com, charlie@example.com",
    alertWebhookUrl:        "",
    employeeChatIds:        "Alice Johnson:112233445566778899001, Bob Smith:998877665544332211009",
    configFileId:           "",
    enabledCreate:          "Yes",
    enabledVerify:          "Yes",
    statusOptions:          "",
    activityOptions:        "",
    holidayDates:           "",
    backlogUrl:             "",
    backlogApiKey:          "",
    backlogProjectIds:      "",
    backlogEngineerMapping: ""
  },
  {
    projectName:            "Frontend Team",
    teamMembers:            "Diana Prince, Ethan Hunt, Fiona Green",
    driveFolderId:          "",
    webhookUrl:             "",
    employeeEmails:         "diana@example.com, ethan@example.com, fiona@example.com",
    alertWebhookUrl:        "",
    employeeChatIds:        "Diana Prince:111222333444555666777, Ethan Hunt:777666555444333222111",
    configFileId:           "",
    enabledCreate:          "Yes",
    enabledVerify:          "Yes",
    statusOptions:          "",
    activityOptions:        "",
    holidayDates:           "",
    backlogUrl:             "",
    backlogApiKey:          "",
    backlogProjectIds:      "",
    backlogEngineerMapping: ""
  },
  {
    projectName:            "QA Team",
    teamMembers:            "George Miller, Hannah White",
    driveFolderId:          "",
    webhookUrl:             "",
    employeeEmails:         "george@example.com, hannah@example.com",
    alertWebhookUrl:        "",
    employeeChatIds:        "George Miller:444555666777888999000",
    configFileId:           "",
    enabledCreate:          "Yes",
    enabledVerify:          "Yes",
    statusOptions:          "In progress, Completed, On hold, Blocked",
    activityOptions:        "Manual Testing, Automation Testing, Bug Reporting, Bug Verification, Meeting, Documentation",
    holidayDates:           "",
    backlogUrl:             "",
    backlogApiKey:          "",
    backlogProjectIds:      "",
    backlogEngineerMapping: ""
  },
  {
    projectName:            "DevOps Team",
    teamMembers:            "Ivan Drago, Julia Roberts",
    driveFolderId:          "",
    webhookUrl:             "",
    employeeEmails:         "",
    alertWebhookUrl:        "",
    employeeChatIds:        "",
    configFileId:           "",
    enabledCreate:          "Yes",
    enabledVerify:          "Yes",
    statusOptions:          "",
    activityOptions:        "Infrastructure, Deployment, Monitoring, Bug fix(Customer), Meeting, Documentation",
    holidayDates:           "2026-01-01, 2026-01-26, 2026-04-03, 2026-08-15, 2026-10-02, 2026-12-25",
    backlogUrl:             "https://mycompany.backlog.com",
    backlogApiKey:          "YOUR_BACKLOG_API_KEY",
    backlogProjectIds:      "PROJECT_ID_1,PROJECT_ID_2",
    backlogEngineerMapping: "Ivan Drago:Ivan, Julia Roberts:JuliaR"
  },
  {
    projectName:            "Archived Project",
    teamMembers:            "Kevin Bacon",
    driveFolderId:          "",
    webhookUrl:             "",
    employeeEmails:         "",
    alertWebhookUrl:        "",
    employeeChatIds:        "",
    configFileId:           "",
    enabledCreate:          "No",  // Skipped by timesheet creator
    enabledVerify:          "No",  // Skipped by verifier
    statusOptions:          "",
    activityOptions:        "",
    holidayDates:           "",
    backlogUrl:             "",
    backlogApiKey:          "",
    backlogProjectIds:      "",
    backlogEngineerMapping: ""
  }
];

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Creates the test master sheet and logs its ID and URL.
 * Run this function once from the Apps Script editor.
 */
function createTestMasterSheet() {
  Logger.log("Creating test master sheet: " + TEST_MASTER_SHEET_NAME);

  // Create spreadsheet
  const ss = SpreadsheetApp.create(TEST_MASTER_SHEET_NAME);
  Logger.log("Created spreadsheet ID: " + ss.getId());

  // Move to folder if specified
  const file = DriveApp.getFileById(ss.getId());
  if (TEST_MASTER_FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(TEST_MASTER_FOLDER_ID);
      file.moveTo(folder);
      Logger.log("Moved to folder: " + folder.getName());
    } catch (e) {
      Logger.log("WARNING: Could not move to folder: " + e);
    }
  }

  // Rename default sheet to match MASTER_SHEET_TAB
  const sheet = ss.getSheets()[0];
  sheet.setName(TEST_MASTER_TAB_NAME);

  // Write headers
  const headers = [
    "Project Name",
    "Team Members",
    "Drive Folder ID",
    "Webhook URL",
    "Employee Emails",
    "Alert Webhook URL",
    "Employee Chat IDs",
    "Config File ID",
    "Enabled (Create)",
    "Enabled (Verify)",
    "Status Options",
    "Activity Options",
    "Holiday Dates",
    "Backlog URL",
    "Backlog API Key",
    "Backlog Project IDs",
    "Backlog Engineer Mapping"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#4A86E8")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(11);

  // Write data rows
  const rows = SAMPLE_TEAMS.map(function(team) {
    return [
      team.projectName,
      team.teamMembers,
      team.driveFolderId,
      team.webhookUrl,
      team.employeeEmails,
      team.alertWebhookUrl,
      team.employeeChatIds,
      team.configFileId,
      team.enabledCreate,
      team.enabledVerify,
      team.statusOptions,
      team.activityOptions,
      team.holidayDates,
      team.backlogUrl,
      team.backlogApiKey,
      team.backlogProjectIds,
      team.backlogEngineerMapping
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Style data rows (alternating)
  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const bg = (i % 2 === 0) ? "#FFFFFF" : "#F3F3F3";
    sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bg);

    // Highlight fully-disabled rows (both Create and Verify = No)
    const disabledCreate = rows[i][8].toString().trim().toLowerCase() === "no";
    const disabledVerify = rows[i][9].toString().trim().toLowerCase() === "no";
    if (disabledCreate && disabledVerify) {
      sheet.getRange(rowNum, 1, 1, headers.length)
        .setBackground("#FFE0E0")
        .setFontColor("#999999");
    }
  }

  // Set column widths
  sheet.setColumnWidth(1,  180);  // Project Name
  sheet.setColumnWidth(2,  250);  // Team Members
  sheet.setColumnWidth(3,  200);  // Drive Folder ID
  sheet.setColumnWidth(4,  220);  // Webhook URL (Manager)
  sheet.setColumnWidth(5,  250);  // Employee Emails
  sheet.setColumnWidth(6,  220);  // Alert Webhook URL
  sheet.setColumnWidth(7,  300);  // Employee Chat IDs
  sheet.setColumnWidth(8,  200);  // Config File ID
  sheet.setColumnWidth(9,  120);  // Enabled (Create)
  sheet.setColumnWidth(10, 120);  // Enabled (Verify)
  sheet.setColumnWidth(11, 250);  // Status Options
  sheet.setColumnWidth(12, 350);  // Activity Options
  sheet.setColumnWidth(13, 250);  // Holiday Dates
  sheet.setColumnWidth(14, 200);  // Backlog URL
  sheet.setColumnWidth(15, 180);  // Backlog API Key
  sheet.setColumnWidth(16, 200);  // Backlog Project IDs
  sheet.setColumnWidth(17, 280);  // Backlog Engineer Mapping

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add border to all data
  sheet
    .getRange(1, 1, rows.length + 1, headers.length)
    .setBorder(true, true, true, true, true, true);

  // Wrap text for multi-value columns
  sheet.getRange(1, 1, rows.length + 1, headers.length)
    .setWrap(true)
    .setVerticalAlignment("middle");

  // Auto-resize rows
  for (let r = 1; r <= rows.length + 1; r++) {
    sheet.setRowHeight(r, 40);
  }

  // Log output
  Logger.log("=".repeat(60));
  Logger.log("✓ Test Master Sheet created successfully!");
  Logger.log("=".repeat(60));
  Logger.log("Spreadsheet ID : " + ss.getId());
  Logger.log("URL            : " + ss.getUrl());
  Logger.log("Tab Name       : " + TEST_MASTER_TAB_NAME);
  Logger.log("Teams added    : " + SAMPLE_TEAMS.length + " (" + (SAMPLE_TEAMS.length - 1) + " enabled, 1 disabled)");
  Logger.log("Columns        : " + headers.length + " (includes Alert Webhook, Chat IDs, Config File ID, Verify toggle, Backlog)");
  Logger.log("=".repeat(60));
  Logger.log("");
  Logger.log("NEXT STEPS:");
  Logger.log("1. Copy the Spreadsheet ID above and set it as MASTER_SHEET_ID");
  Logger.log("   Extensions > Apps Script > Project Settings > Script Properties");
  Logger.log("   Key: MASTER_SHEET_ID   Value: " + ss.getId());
  Logger.log("2. Fill in real Webhook URLs, Chat IDs, Drive Folder IDs per team.");
  Logger.log("3. After creating timesheets, paste the timesheet_config.json file");
  Logger.log("   ID into the Config File ID column for each team.");
  Logger.log("4. Run runDailyVerificationAllTeams() to verify all teams at once.");
  Logger.log("=".repeat(60));

  return ss.getId();
}

// ============================================================================
// UTILITY: Print current Script Property value for MASTER_SHEET_ID
// ============================================================================

function printMasterSheetId() {
  const id = PropertiesService.getScriptProperties().getProperty("MASTER_SHEET_ID") || "(not set)";
  Logger.log("MASTER_SHEET_ID = " + id);
}
