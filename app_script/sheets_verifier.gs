/**
 * Google Apps Script version of sheets_verifier.py
 * Timesheet Verification System
 *
 * This script verifies employee timesheet entries and sends daily reports
 *
 * ============================================================================
 * QUICK START - Simplified Configuration
 * ============================================================================
 *
 * 1. Set basic configuration:
 *    - Run setupConfiguration() and update the values
 *
 * 2. Configure engineers (EASY WAY):
 *
 *    Option A - Quick setup (recommended):
 *    quickSetupEngineers([
 *      {name: 'John Doe', chatId: '123456789012345678901'},
 *      {name: 'Jane Smith', chatId: '987654321098765432109'},
 *      {name: 'Bob Johnson'}  // No chat ID needed
 *    ]);
 *
 *    Option B - Add one at a time:
 *    addEngineer('John Doe', '123456789012345678901');
 *    addEngineer('Jane Smith', '987654321098765432109');
 *
 *    Option C - Just set names (chat IDs optional):
 *    setEngineerNames(['John Doe', 'Jane Smith', 'Bob Johnson']);
 *    setEngineerChatId('John Doe', '123456789012345678901');  // Add chat ID later
 *
 * 3. View your configuration:
 *    viewConfiguration();
 *
 * 4. Remove an engineer:
 *    removeEngineer('John Doe');
 *
 * ============================================================================
 * ADVANCED CONFIGURATION - Script Properties
 * ============================================================================
 * Required Script Properties:
 * - CONFIG_FILE_ID: (Optional) ID of timesheet_config.json file in Google Drive
 * - DESTINATION_FOLDER_ID: (Optional) Folder ID where timesheet_config.json is stored
 * - MAIN_SPREADSHEET_ID: (Fallback) The main spreadsheet ID - used if config file not found
 * - ENGINEER_NAMES: Comma-separated list of engineer names (sheet names)
 * - GOOGLE_CHAT_WEBHOOK_URL: Webhook URL for manager reports
 * - EMPLOYEE_ALERT_WEBHOOK_URL: Webhook URL for employee reminders
 * - TEST_MODE: "true" or "false" (default: false)
 * - EMPLOYEE_CHAT_IDS: JSON object mapping employee names to chat IDs
 * - HOLIDAYS: JSON array of holiday dates in YYYY-MM-DD format
 * - ROWS_TO_CHECK_AFTER_DATE: Number of rows to check after finding a date (default: 5)
 * - RSS_ARTICLE_COUNT: Number of RSS articles to fetch and send (default: 5)
 * - BACKLOG_URL: Backlog service URL (default: "https://ilabs.backlog.com")
 * - BACKLOG_API_KEY: Backlog API key for authentication
 * - BACKLOG_PROJECT_IDS: Comma-separated list of project IDs to fetch issues from
 * - BACKLOG_ENGINEER_MAPPING: JSON object mapping engineer names to Backlog assignee names
 *
 * Note: The script automatically fetches the spreadsheet ID from timesheet_config.json
 * based on the current year-month (format: "2025-December"). If the config file is not
 * found or the current month's entry doesn't exist, it falls back to MAIN_SPREADSHEET_ID.
 * ============================================================================
 */

/**
 * Main class for Sheets Verification
 */
class SheetsVerifier {
  constructor() {
    this.props = PropertiesService.getScriptProperties();
    this.config = this._loadConfig();
  }

  /**
   * Get current month name (helper for config lookup)
   */
  _getCurrentMonthName() {
    const now = new Date();
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[now.getMonth()];
  }

  /**
   * Get month name from a Date object
   */
  _getMonthNameFromDate(date) {
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[date.getMonth()];
  }
_getCurrentMonthSpreadsheetId() {
  try {
    // Delegate to the existing function which accepts an optional date
    return this._getSpreadsheetIdForDate(new Date());
  } catch (e) {
    Logger.log("Error in _getCurrentMonthSpreadsheetId: " + e);
    return null;
  }
}
  /**
   * Fetch spreadsheet ID from timesheet_config.json based on a specific date's year-month
   * Returns the sheet ID for the specified date's month or null if not found
   * @param {Date} targetDate - Optional date to get sheet ID for (defaults to current date)
   */
  _getSpreadsheetIdForDate(targetDate = null) {
    try {
      const props = PropertiesService.getScriptProperties();
      const configFileId = props.getProperty("CONFIG_FILE_ID");
      const destinationFolderId = props.getProperty("DESTINATION_FOLDER_ID");

      if (!configFileId && !destinationFolderId) {
        Logger.log("CONFIG_FILE_ID or DESTINATION_FOLDER_ID not configured");
        return null;
      }

      // Try to get config file by ID first
      let configFile = null;
      if (configFileId) {
        try {
          configFile = DriveApp.getFileById(configFileId);
        } catch (e) {
          Logger.log("Could not find config file by ID, trying by name");
        }
      }

      // If not found by ID, try to find by name in folder
      if (!configFile && destinationFolderId) {
        try {
          const folder = DriveApp.getFolderById(destinationFolderId);
          const files = folder.getFilesByName("timesheet_config.json");
          if (files.hasNext()) {
            configFile = files.next();
          }
        } catch (e) {
          Logger.log("Could not find config file in folder");
        }
      }

      // If still not found, try root folder
      if (!configFile) {
        try {
          const files = DriveApp.getRootFolder().getFilesByName(
            "timesheet_config.json"
          );
          if (files.hasNext()) {
            configFile = files.next();
          }
        } catch (e) {
          Logger.log("Could not find config file in root folder");
        }
      }

      if (!configFile) {
        Logger.log("timesheet_config.json not found");
        return null;
      }

      // Read and parse config file
      const configContent = configFile.getBlob().getDataAsString();
      const config = JSON.parse(configContent);

      // Use targetDate if provided, otherwise use current date
      const dateToUse = targetDate || new Date();
      const targetYear = dateToUse.getFullYear();
      const targetMonth = this._getMonthNameFromDate(dateToUse);
      const configKey = targetYear + "-" + targetMonth;

      // Get spreadsheet ID for target month
      const spreadsheetId = config[configKey];

      if (spreadsheetId) {
        Logger.log(
          "✓ Found spreadsheet ID for " + configKey + ": " + spreadsheetId
        );
        return spreadsheetId;
      } else {
        Logger.log("No spreadsheet ID found for " + configKey);
        const availableKeys = Object.keys(config).filter(
          (k) => k !== "updated_at" && k !== "latest"
        );
        if (availableKeys.length > 0) {
          Logger.log("Available keys: " + availableKeys.join(", "));
        }
        return null;
      }
    } catch (e) {
      Logger.log("Error fetching spreadsheet ID from config: " + e);
      return null;
    }
  }

  /**
   * Load configuration from Script Properties
   */
  _loadConfig() {
    const props = PropertiesService.getScriptProperties();

    // Parse engineer names
    const engineerNamesStr =
      props.getProperty("ENGINEER_NAMES") ||
      "Jinu T J,Bismillakhan S,Midhun,Aravind,Akhil Mohan,Akash T K,Shinoj";
    const engineerNames = engineerNamesStr.split(",").map((n) => n.trim());

    // Parse employee chat IDs
    let employeeChatIds = {};
    try {
      const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
      if (chatIdsStr) {
        employeeChatIds = JSON.parse(chatIdsStr);
      }
    } catch (e) {
      Logger.log("Error parsing EMPLOYEE_CHAT_IDS: " + e);
    }

    // Parse holidays
    let holidays = [];
    try {
      const holidaysStr = props.getProperty("HOLIDAYS");
      if (holidaysStr) {
        holidays = JSON.parse(holidaysStr);
      } else {
        // Default holidays for 2025
        holidays = [
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
      }
    } catch (e) {
      Logger.log("Error parsing HOLIDAYS: " + e);
    }

    // Parse Backlog engineer name mapping
    let backlogEngineerMapping = {};
    try {
      const mappingStr = props.getProperty("BACKLOG_ENGINEER_MAPPING");
      if (mappingStr) {
        backlogEngineerMapping = JSON.parse(mappingStr);
      }
    } catch (e) {
      Logger.log("Error parsing BACKLOG_ENGINEER_MAPPING: " + e);
    }

    // Get spreadsheet ID for current month from config file
    // Fallback to MAIN_SPREADSHEET_ID if config file not available
    let mainSpreadsheetId = this._getCurrentMonthSpreadsheetId();
    if (!mainSpreadsheetId) {
      mainSpreadsheetId = props.getProperty("MAIN_SPREADSHEET_ID");
      if (mainSpreadsheetId) {
        Logger.log(
          "Using MAIN_SPREADSHEET_ID from Script Properties as fallback"
        );
      }
    }

    return {
      mainSpreadsheetId: mainSpreadsheetId,
      engineerNames: engineerNames,
      googleChatWebhookUrl: props.getProperty("GOOGLE_CHAT_WEBHOOK_URL"),
      employeeAlertWebhookUrl: props.getProperty("EMPLOYEE_ALERT_WEBHOOK_URL"),
      testMode: props.getProperty("TEST_MODE") === "true",
      employeeChatIds: employeeChatIds,
      holidays: holidays,
      rowsToCheckAfterDate: parseInt(
        props.getProperty("ROWS_TO_CHECK_AFTER_DATE") || "5",
        10
      ),
      sheetDataRange: props.getProperty("SHEET_DATA_RANGE") || "A1:K150",
      rssArticleCount: parseInt(
        props.getProperty("RSS_ARTICLE_COUNT") || "5",
        10
      ),
      backlogUrl:
        props.getProperty("BACKLOG_URL") || "https://ilabs.backlog.com",
      backlogApiKey: props.getProperty("BACKLOG_API_KEY"),
      backlogProjectIds: props.getProperty("BACKLOG_PROJECT_IDS") || "",
      backlogEngineerMapping: backlogEngineerMapping,
    };
  }

  /**
   * Get current month sheet name (e.g., 'July', 'August')
   */
  getCurrentMonthSheetName() {
    const now = new Date();
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[now.getMonth()];
  }

  /**
   * Check if a given date is a holiday
   */
  isHoliday(date) {
    const dateStr = Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    return this.config.holidays.indexOf(dateStr) !== -1;
  }

  /**
   * Get the last working day (excluding weekends and holidays)
   * Returns an object with both the formatted date string and the Date object
   */
  getLastWorkingDay() {
    const today = new Date();
    const currentDay = today.getDay(); // 0=Sunday, 6=Saturday

    // Skip if today is weekend
    if (currentDay === 0 || currentDay === 6) {
      return null;
    }

    // Find the last working day
    let checkDate = new Date(today);
    checkDate.setDate(checkDate.getDate() - 1);

    // Keep going back until we find a non-weekend, non-holiday day
    while (
      checkDate.getDay() === 0 ||
      checkDate.getDay() === 6 ||
      this.isHoliday(checkDate)
    ) {
      checkDate.setDate(checkDate.getDate() - 1);
    }

    return {
      dateString: Utilities.formatDate(
        checkDate,
        Session.getScriptTimeZone(),
        "d-MMM"
      ),
      dateObject: checkDate,
    };
  }

  /**
   * Get current day
   */
  getCurrentDay() {
    const today = new Date();
    const currentDay = today.getDay();

    // Skip if today is weekend or holiday
    if (currentDay === 0 || currentDay === 6 || this.isHoliday(today)) {
      return null;
    }

    return Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM");
  }

  /**
   * Get sheet data from Google Sheets
   */
  getSheetData(spreadsheetId, rangeName) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      // Remove quotes if present (for compatibility), but we don't add them anymore
      const sheetName = rangeName.split("!")[0].replace(/^'|'$/g, "");
      const range = rangeName.split("!")[1];

      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log("Sheet not found: " + sheetName);
        return [];
      }

      const dataRange = sheet.getRange(range);
      const values = dataRange.getValues();

      return values;
    } catch (e) {
      Logger.log("Error fetching sheet data: " + e);
      return [];
    }
  }

  /**
   * Get engineer sheet data
   * @param {string} engineerName - Name of the engineer (sheet name)
   * @param {string} monthName - Month name (e.g., "December") - for logging purposes
   * @param {string} spreadsheetId - Optional spreadsheet ID (defaults to config.mainSpreadsheetId)
   */
  getEngineerSheetData(engineerName, monthName, spreadsheetId = null) {
    const sheetRange = engineerName + "!" + this.config.sheetDataRange;
    const targetSpreadsheetId = spreadsheetId || this.config.mainSpreadsheetId;

    try {
      return this.getSheetData(targetSpreadsheetId, sheetRange);
    } catch (e) {
      Logger.log("Error fetching data for engineer " + engineerName + ": " + e);
      return [];
    }
  }

  /**
   * Helper function to pad numbers with leading zeros
   */
  _padZero(num, length) {
    const str = String(num);
    if (str.length >= length) {
      return str;
    }
    // Use Array.join for compatibility
    const zeros = Array(length - str.length + 1).join("0");
    return zeros + str;
  }

  /**
   * Convert time value (Date object or string) to time string format
   * Handles: Date objects (like "Sat Dec 30 1899 12:30:00"), time strings ("12:30"), etc.
   * Returns: "HH:mm" format (e.g., "12:30") or "HH:mm:ss" if seconds present, or empty string if cannot parse
   */
  _normalizeTime(timeValue) {
    if (!timeValue) {
      return "";
    }

    // If it's a Date object (Google Sheets time values come as Date objects)
    if (timeValue instanceof Date) {
      const hours = timeValue.getHours();
      const minutes = timeValue.getMinutes();
      const seconds = timeValue.getSeconds();

      // Format as HH:mm (or HH:mm:ss if seconds are present)
      if (seconds > 0) {
        return (
          this._padZero(hours, 2) +
          ":" +
          this._padZero(minutes, 2) +
          ":" +
          this._padZero(seconds, 2)
        );
      } else {
        return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
      }
    }

    const timeStr = String(timeValue).trim();
    if (!timeStr) {
      return "";
    }

    // If it's already in time format (HH:mm or HH:mm:ss)
    const timeRegex = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;
    const match = timeStr.match(timeRegex);
    if (match) {
      const hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      const seconds = match[3] ? parseInt(match[3], 10) : 0;

      if (seconds > 0) {
        return (
          this._padZero(hours, 2) +
          ":" +
          this._padZero(minutes, 2) +
          ":" +
          this._padZero(seconds, 2)
        );
      } else {
        return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
      }
    }

    // Try to parse as Date string (e.g., "Sat Dec 30 1899 12:30:00 GMT+0521")
    try {
      const parsedDate = new Date(timeStr);
      if (!isNaN(parsedDate.getTime())) {
        // Check if it looks like a time-only date (Dec 30, 1899 is the epoch for time values)
        // Also check for Jan 1, 1900 which is another common epoch
        if (
          parsedDate.getFullYear() === 1899 ||
          parsedDate.getFullYear() === 1900 ||
          timeStr.match(/Dec 30 1899/i) ||
          timeStr.match(/Jan 01 1900/i) ||
          timeStr.match(/Jan 1 1900/i)
        ) {
          const hours = parsedDate.getHours();
          const minutes = parsedDate.getMinutes();
          const seconds = parsedDate.getSeconds();

          if (seconds > 0) {
            return (
              this._padZero(hours, 2) +
              ":" +
              this._padZero(minutes, 2) +
              ":" +
              this._padZero(seconds, 2)
            );
          } else {
            return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
          }
        }
      }
    } catch (e) {
      // Not a parseable date, continue
    }

    // If we can't parse it, return the original string (might already be formatted)
    return timeStr;
  }

  /**
   * Normalize date from various formats to a standard format for comparison
   * Handles: Date objects, "d-MMM" format, "dd-MMM" format, full date strings, etc.
   * Returns: "d-MMM" format (e.g., "15-Dec") or null if cannot parse
   */
  _normalizeDate(dateValue) {
    if (!dateValue) {
      return null;
    }

    // If it's already a Date object
    if (dateValue instanceof Date) {
      return Utilities.formatDate(
        dateValue,
        Session.getScriptTimeZone(),
        "d-MMM"
      );
    }

    const dateStr = String(dateValue).trim();
    if (!dateStr) {
      return null;
    }

    // Try to parse as Date if it looks like a full date string
    // e.g., "Mon Dec 15 2025 00:00:00 GMT+0530"
    try {
      const parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        // Check if it's a valid date (not just a string that can't be parsed)
        // If the original string contains day names and full month names, it's likely a Date string
        if (
          dateStr.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)/i) ||
          (dateStr.match(/\d{4}/) && dateStr.length > 10)
        ) {
          return Utilities.formatDate(
            parsedDate,
            Session.getScriptTimeZone(),
            "d-MMM"
          );
        }
      }
    } catch (e) {
      // Not a Date object, continue with string parsing
    }

    // Check if it's already in "d-MMM" or "dd-MMM" format (e.g., "15-Dec" or "5-Dec")
    const dMMMRegex = /^(\d{1,2})-([A-Za-z]{3})$/i;
    const match = dateStr.match(dMMMRegex);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = match[2];
      // Return in consistent format (remove leading zero if present)
      return (
        day + "-" + month.charAt(0).toUpperCase() + month.slice(1).toLowerCase()
      );
    }

    // Try other common formats
    // Format: "DD/MM/YYYY" or "MM/DD/YYYY"
    const slashRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const slashMatch = dateStr.match(slashRegex);
    if (slashMatch) {
      try {
        // Try DD/MM/YYYY first (more common internationally)
        const day = parseInt(slashMatch[1], 10);
        const month = parseInt(slashMatch[2], 10) - 1; // Month is 0-indexed
        const year = parseInt(slashMatch[3], 10);
        const date = new Date(year, month, day);
        if (date.getDate() === day && date.getMonth() === month) {
          return Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            "d-MMM"
          );
        }
      } catch (e) {
        // Try MM/DD/YYYY
        try {
          const month = parseInt(slashMatch[1], 10) - 1;
          const day = parseInt(slashMatch[2], 10);
          const year = parseInt(slashMatch[3], 10);
          const date = new Date(year, month, day);
          if (date.getDate() === day && date.getMonth() === month) {
            return Utilities.formatDate(
              date,
              Session.getScriptTimeZone(),
              "d-MMM"
            );
          }
        } catch (e2) {
          // Ignore
        }
      }
    }

    // Format: "YYYY-MM-DD"
    const isoRegex = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    const isoMatch = dateStr.match(isoRegex);
    if (isoMatch) {
      try {
        const year = parseInt(isoMatch[1], 10);
        const month = parseInt(isoMatch[2], 10) - 1;
        const day = parseInt(isoMatch[3], 10);
        const date = new Date(year, month, day);
        if (date.getDate() === day && date.getMonth() === month) {
          return Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            "d-MMM"
          );
        }
      } catch (e) {
        // Ignore
      }
    }

    // If we can't parse it, return null
    Logger.log("Warning: Could not parse date format: " + dateStr);
    return null;
  }

  /**
   * Parse last working day entries from sheet data
   */
  parseLastWorkingDayEntries(sheetData, targetDate) {
    if (!sheetData || sheetData.length < 2) {
      return [];
    }

    // Normalize target date to "d-MMM" format
    const normalizedTargetDate = this._normalizeDate(targetDate);
    if (!normalizedTargetDate) {
      Logger.log("Error: Could not normalize target date: " + targetDate);
      return [];
    }

    Logger.log("Looking for normalized target date: " + normalizedTargetDate);

    // Find header row (look for 'Date' in column B, index 1)
    let headerRow = null;
    for (let i = 0; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row.length > 1 && String(row[1]).indexOf("Date") !== -1) {
        headerRow = i;
        break;
      }
    }

    if (headerRow === null) {
      Logger.log("Could not find header row with 'Date' in column B");
      return [];
    }

    // Extract entries for the target date and next configurable rows
    const targetEntries = [];
    let i = headerRow + 1;

    while (i < sheetData.length) {
      const row = sheetData[i];
      if (row.length === 0) {
        i++;
        continue;
      }

      // Check if this row has the target date in column B (index 1)
      const dateCellValue = row.length > 1 ? row[1] : null;
      const normalizedDateCell = this._normalizeDate(dateCellValue);

      // Compare normalized dates
      if (normalizedDateCell && normalizedDateCell === normalizedTargetDate) {
        Logger.log(
          "Found target date '" +
            normalizedTargetDate +
            "' (original: " +
            dateCellValue +
            ") at row " +
            (i + 1) +
            " (column B)"
        );

        // Parse this row and the next configurable rows for all tasks on this date
        const rowsToCheck = this.config.rowsToCheckAfterDate + 1; // +1 to include the current row

        // First, check if the first row has total_duration (merged cell case)
        const firstRow = sheetData[i];
        let mergedTotalDuration = "";
        if (firstRow.length > 9) {
          mergedTotalDuration = this._normalizeTime(firstRow[9]);
        }

        for (let j = i; j < Math.min(i + rowsToCheck, sheetData.length); j++) {
          const taskRow = sheetData[j];
          if (taskRow.length === 0) {
            continue;
          }

          // Skip if this is another date row (unless it's the first one)
          if (j > i && taskRow.length > 1) {
            const nextDateCellValue = taskRow[1];
            const normalizedNextDate = this._normalizeDate(nextDateCellValue);
            // If we can normalize it and it's different from target date, it's a new date
            if (
              normalizedNextDate &&
              normalizedNextDate !== normalizedTargetDate
            ) {
              Logger.log(
                "  Found new date at row " +
                  (j + 1) +
                  " (" +
                  normalizedNextDate +
                  "), stopping task collection"
              );
              break;
            }
            // Also check if it looks like a date string (has numbers)
            const nextDateStr = String(nextDateCellValue).trim();
            if (
              nextDateStr.length > 0 &&
              /[0-9]/.test(nextDateStr) &&
              !normalizedNextDate
            ) {
              // This might be a new date we couldn't parse, but has numbers - be cautious
              Logger.log(
                "  Found potential new date at row " +
                  (j + 1) +
                  ", stopping task collection"
              );
              break;
            }
          }

          // Parse the task data according to correct column structure
          const entry = {
            row_number: j + 1,
            date: normalizedTargetDate,
            module_area: taskRow.length > 2 ? String(taskRow[2]) : "",
            task_details: taskRow.length > 3 ? String(taskRow[3]) : "",
            status: taskRow.length > 4 ? String(taskRow[4]) : "",
            activity_type: taskRow.length > 5 ? String(taskRow[5]) : "",
            start_time:
              taskRow.length > 6 ? this._normalizeTime(taskRow[6]) : "",
            end_time: taskRow.length > 7 ? this._normalizeTime(taskRow[7]) : "",
            total_duration:
              taskRow.length > 9 ? this._normalizeTime(taskRow[9]) : "",
            remarks: taskRow.length > 10 ? String(taskRow[10]) : "",
          };

          // If this entry doesn't have total_duration but we found it in merged cell, use it
          if (!entry.total_duration.trim() && mergedTotalDuration) {
            entry.total_duration = mergedTotalDuration;
          }

          // Only add entries that have meaningful task details
          if (entry.task_details.trim() || entry.module_area.trim()) {
            targetEntries.push(entry);
            Logger.log(
              "  Added task from row " +
                (j + 1) +
                ": " +
                entry.task_details.substring(0, 50) +
                "..."
            );
          } else if (
            entry.start_time.trim() &&
            entry.end_time.trim() &&
            targetEntries.length > 0
          ) {
            // This indicates additional time for the previous task
            const previousEntry = targetEntries[targetEntries.length - 1];
            const additionalEntry = {
              row_number: entry.row_number,
              date: entry.date,
              module_area: previousEntry.module_area,
              task_details: previousEntry.task_details,
              status: previousEntry.status,
              activity_type: previousEntry.activity_type,
              start_time: entry.start_time,
              end_time: entry.end_time,
              total_duration: entry.total_duration,
              remarks: entry.remarks,
            };
            targetEntries.push(additionalEntry);
            Logger.log(
              "  Added additional time entry from row " +
                (j + 1) +
                " for previous task: " +
                previousEntry.task_details.substring(0, 50) +
                "..."
            );
          }
        }

        // Return immediately once entries are found and processed
        Logger.log(
          "Total entries found for " +
            normalizedTargetDate +
            ": " +
            targetEntries.length
        );
        return targetEntries;
      } else {
        i++;
      }
    }

    Logger.log(
      "Total entries found for " +
        normalizedTargetDate +
        ": " +
        targetEntries.length
    );
    return targetEntries;
  }

  /**
   * Generate formatted analysis message from timesheet entries
   */
  generateAnalysisMessage(employeeName, lastDayEntries) {
    if (!lastDayEntries || lastDayEntries.length === 0) {
      return "👤 " + employeeName + " (Total: 0h 0m)\n└ No entries found";
    }

    // Get total duration from the first entry (since it's a merged cell)
    const totalDuration = lastDayEntries[0].total_duration || "0:00";

    // Convert duration to hours format
    const durationToHours = function (durationStr) {
      if (!durationStr || durationStr.trim() === "") {
        return 0.0;
      }
      try {
        if (durationStr.indexOf(":") !== -1) {
          const parts = durationStr.split(":");
          return parseFloat(parts[0]) + parseFloat(parts[1]) / 60;
        } else {
          return parseFloat(durationStr);
        }
      } catch (e) {
        return 0.0;
      }
    };

    // Calculate task duration
    const calculateTaskDuration = function (start, end) {
      if (!start || !end) {
        return 0.0;
      }
      try {
        const startParts = start.split(":");
        const endParts = end.split(":");
        const startMinutes =
          parseInt(startParts[0]) * 60 + parseInt(startParts[1]);
        const endMinutes = parseInt(endParts[0]) * 60 + parseInt(endParts[1]);
        const durationMinutes = endMinutes - startMinutes;
        return durationMinutes / 60.0;
      } catch (e) {
        return 0.0;
      }
    };

    // Format duration
    const formatDuration = function (hoursDecimal) {
      if (hoursDecimal === 0) {
        return "0h 0m";
      }
      const hours = Math.floor(hoursDecimal);
      const minutes = Math.floor((hoursDecimal - hours) * 60);
      if (hours === 0) {
        return minutes + "m";
      } else if (minutes === 0) {
        return hours + "h";
      } else {
        return hours + "h " + minutes + "m";
      }
    };

    const totalHours = durationToHours(totalDuration);

    // Group entries by activity type
    const activityGroups = {};

    for (let i = 0; i < lastDayEntries.length; i++) {
      const entry = lastDayEntries[i];
      const activityType = entry.activity_type || "Task";

      if (!activityGroups[activityType]) {
        activityGroups[activityType] = [];
      }

      // Calculate individual task duration
      const startTime = entry.start_time || "";
      const endTime = entry.end_time || "";
      const taskDuration = calculateTaskDuration(startTime, endTime);

      // Get task details
      const taskDetails = entry.task_details || "No details";
      const moduleArea = entry.module_area || "";

      // Prefer module_area over task_details for description
      const description =
        moduleArea && moduleArea.trim() ? moduleArea : taskDetails;

      activityGroups[activityType].push({
        description: description,
        duration: taskDuration,
      });
    }

    // Format the header
    let message =
      "\n👤 " + employeeName + " (Total: " + formatDuration(totalHours) + ")\n";

    // Process grouped activities
    const activityTypes = Object.keys(activityGroups);

    for (let i = 0; i < activityTypes.length; i++) {
      const activityType = activityTypes[i];
      const tasks = activityGroups[activityType];

      // Combine descriptions and calculate total duration for this activity type
      const descriptions = [];
      let totalActivityDuration = 0.0;

      for (let j = 0; j < tasks.length; j++) {
        let desc = tasks[j].description;
        const duration = tasks[j].duration;

        // Truncate individual descriptions if too long
        if (desc.length > 40) {
          desc = desc.substring(0, 37) + "...";
        }

        descriptions.push(desc + " (" + formatDuration(duration) + ")");
        totalActivityDuration += duration;
      }

      // Combine all descriptions for this activity type
      const combinedDesc = descriptions.join(", ");

      // Use appropriate tree symbol
      const symbol = i === activityTypes.length - 1 ? "└" : "├";

      message += symbol + " " + activityType + " – " + combinedDesc + "\n";
    }

    // Add remarks summary if available
    const remarksList = [];
    for (let i = 0; i < lastDayEntries.length; i++) {
      const remarks = (lastDayEntries[i].remarks || "").trim();
      if (remarks) {
        remarksList.push(remarks);
      }
    }

    if (remarksList.length > 0) {
      // Combine all remarks
      let allRemarks = remarksList.join("; ");
      if (allRemarks.length > 100) {
        allRemarks = allRemarks.substring(0, 97) + "...";
      }
      message += "📝 Additional: " + allRemarks + "\n";
    }

    return message.trim();
  }

  /**
   * Verify all employee sheets and return analysis results
   */
  verifyAllEmployees() {
    const results = [];

    // Get last working day (returns object with dateString and dateObject)
    const lastWorkingDayInfo = this.getLastWorkingDay();
    if (!lastWorkingDayInfo) {
      Logger.log("No last working day found");
      return results;
    }

    const lastWorkingDay = lastWorkingDayInfo.dateString;
    const lastWorkingDayDate = lastWorkingDayInfo.dateObject;

    // Get the month of the last working day (not current month!)
    const lastWorkingDayMonth = this._getMonthNameFromDate(lastWorkingDayDate);
    const lastWorkingDayYear = lastWorkingDayDate.getFullYear();

    // Fetch spreadsheet ID for the last working day's month
    const spreadsheetId = this._getSpreadsheetIdForDate(lastWorkingDayDate);
    if (!spreadsheetId) {
      Logger.log(
        "Warning: Could not find spreadsheet ID for " +
          lastWorkingDayYear +
          "-" +
          lastWorkingDayMonth +
          ", using fallback"
      );
      // Use fallback from config
    }

    const effectiveSpreadsheetId =
      spreadsheetId || this.config.mainSpreadsheetId;

    Logger.log("Checking entries for working day: " + lastWorkingDay);
    Logger.log(
      "Last working day is in month: " +
        lastWorkingDayMonth +
        " " +
        lastWorkingDayYear
    );
    Logger.log("Using spreadsheet ID: " + effectiveSpreadsheetId);

    for (let i = 0; i < this.config.engineerNames.length; i++) {
      const engineerName = this.config.engineerNames[i];
      Logger.log("Checking " + engineerName + "'s sheet...");

      // Get sheet data from the engineer's sheet within the spreadsheet for last working day's month
      const sheetData = this.getEngineerSheetData(
        engineerName,
        lastWorkingDayMonth,
        effectiveSpreadsheetId
      );

      // Parse last working day entries
      const lastDayEntries = this.parseLastWorkingDayEntries(
        sheetData,
        lastWorkingDay
      );

      Logger.log(
        "Found " +
          lastDayEntries.length +
          " entries for " +
          engineerName +
          " on " +
          lastWorkingDay
      );

      if (lastDayEntries.length > 0) {
        // Analyze the data
        const analysis = this.generateAnalysisMessage(
          engineerName,
          lastDayEntries
        );
        results.push(analysis);
      } else {
        results.push([engineerName, false]);
      }
    }

    return results;
  }

  /**
   * Send message to Google Chat space
   */
  sendGoogleChatMessage(message, webhookUrl) {
    if (this.config.testMode) {
      Logger.log("TEST MODE: Would send Google Chat message: " + message);
      return true;
    }

    const url = webhookUrl || this.config.googleChatWebhookUrl;
    if (!url) {
      Logger.log("Google Chat webhook URL not configured");
      return false;
    }

    const payload = { text: message };

    try {
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        Logger.log("Message sent to Google Chat successfully");
        return true;
      } else {
        Logger.log("Error sending message to Google Chat: " + responseCode);
        return false;
      }
    } catch (e) {
      Logger.log("Error sending message to Google Chat: " + e);
      return false;
    }
  }

  /**
   * Send reminder message to employees who haven't submitted timesheets
   */
  sendEmployeeReminder(employeesToRemind) {
    if (!this.config.employeeAlertWebhookUrl) {
      return false;
    }

    // Create mentions for employees with chat IDs
    const employeeMentions = [];

    for (let i = 0; i < employeesToRemind.length; i++) {
      const employee = employeesToRemind[i];
      const chatId = this.config.employeeChatIds[employee];
      if (chatId) {
        employeeMentions.push("<users/" + chatId + ">");
      } else {
        // Fallback to name if no chat ID configured
        employeeMentions.push(employee);
        Logger.log(
          "Warning: No chat ID configured for " +
            employee +
            ", using name instead"
        );
      }
    }

    // Build message with proper mentions
    if (employeeMentions.length > 0) {
      let message = "Dear " + employeeMentions.join(", ") + ",\n\n";
      message += "\nPlease update your timesheet for the last working day. ✅";

      // Send message with mentions if available
      return this.sendGoogleChatMessage(
        message,
        this.config.employeeAlertWebhookUrl
      );
    }

    return false;
  }

  /**
   * Fetch open issues from Backlog API
   * Returns array of issues or empty array on error
   */
  _fetchBacklogIssues() {
    if (!this.config.backlogApiKey || !this.config.backlogUrl) {
      Logger.log("Backlog API key or URL not configured");
      return [];
    }

    try {
      // Parse project IDs (comma-separated)
      const projectIds = this.config.backlogProjectIds
        ? this.config.backlogProjectIds.split(",").map((id) => id.trim())
        : [];

      if (projectIds.length === 0) {
        Logger.log("No Backlog project IDs configured");
        return [];
      }

      // Build URL with query parameters
      let url =
        this.config.backlogUrl +
        "/api/v2/issues?apiKey=" +
        this.config.backlogApiKey;

      // Add project IDs
      for (let i = 0; i < projectIds.length; i++) {
        url += "&projectId[]=" + encodeURIComponent(projectIds[i]);
      }

      // Add status IDs (1 = Open, 2 = In Progress)
      url += "&statusId[]=1&statusId[]=2";

      Logger.log(
        "Fetching Backlog issues from: " +
          url.replace(this.config.backlogApiKey, "***")
      );

      const options = {
        method: "get",
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const issues = JSON.parse(response.getContentText());
        Logger.log("✓ Fetched " + issues.length + " open issues from Backlog");
        return issues;
      } else {
        Logger.log(
          "Error fetching Backlog issues: " +
            responseCode +
            " - " +
            response.getContentText()
        );
        return [];
      }
    } catch (e) {
      Logger.log("Error fetching Backlog issues: " + e);
      return [];
    }
  }

  /**
   * Count open Backlog tickets per engineer
   * Returns object mapping engineer names to array of issue objects {issueKey, dueDate, summary}
   */
  _countBacklogTicketsPerEngineer() {
    const issues = this._fetchBacklogIssues();
    const ticketsByEngineer = {};

    // Initialize all engineers with empty arrays
    for (let i = 0; i < this.config.engineerNames.length; i++) {
      ticketsByEngineer[this.config.engineerNames[i]] = [];
    }

    // Collect tickets per assignee
    for (let i = 0; i < issues.length; i++) {
      const issue = issues[i];
      if (issue.assignee && issue.assignee.name) {
        const assigneeName = issue.assignee.name;
        const issueData = {
          issueKey: issue.issueKey,
          dueDate: issue.dueDate || null,
          summary: issue.summary || "",
        };

        // Check if this assignee matches any engineer (direct match or via mapping)
        for (let j = 0; j < this.config.engineerNames.length; j++) {
          const engineerName = this.config.engineerNames[j];

          // Direct name match
          if (assigneeName === engineerName) {
            ticketsByEngineer[engineerName].push(issueData);
            break;
          }

          // Check mapping (engineer name -> Backlog assignee name)
          const mappedName = this.config.backlogEngineerMapping[engineerName];
          if (mappedName && assigneeName === mappedName) {
            ticketsByEngineer[engineerName].push(issueData);
            break;
          }

          // Partial match (case-insensitive, contains)
          if (
            assigneeName.toLowerCase().indexOf(engineerName.toLowerCase()) !==
              -1 ||
            engineerName.toLowerCase().indexOf(assigneeName.toLowerCase()) !==
              -1
          ) {
            ticketsByEngineer[engineerName].push(issueData);
            break;
          }
        }
      }
    }

    return ticketsByEngineer;
  }

  /**
   * Get due date icon based on urgency
   * Returns object with icon and isUrgent flag
   * 🔥 = overdue, 🚨 = due today, ⏰ = due tomorrow
   */
  _getDueDateIcon(dueDate) {
    if (!dueDate) {
      return { icon: "", isUrgent: false };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const dueDateObj = new Date(dueDate);
    dueDateObj.setHours(0, 0, 0, 0);

    if (dueDateObj < today) {
      return { icon: "🔥", isUrgent: true }; // Overdue
    } else if (dueDateObj.getTime() === today.getTime()) {
      return { icon: "🚨", isUrgent: true }; // Due today
    } else if (dueDateObj.getTime() === tomorrow.getTime()) {
      return { icon: "⏰", isUrgent: true }; // Due tomorrow
    }

    return { icon: "", isUrgent: false };
  }

  /**
   * Truncate summary to first N words
   */
  _truncateSummary(summary, wordCount) {
    if (!summary) {
      return "No summary";
    }
    const words = summary.trim().split(/\s+/);
    if (words.length <= wordCount) {
      return summary;
    }
    return words.slice(0, wordCount).join(" ") + "...";
  }

  /**
   * Format Backlog ticket counts for display
   * Shows total count but only lists urgent tickets (overdue, due today, due tomorrow)
   * @param {Object} ticketsByEngineer - Object mapping engineer names to arrays of issue objects
   */
  _formatBacklogTicketCounts(ticketsByEngineer) {
    let message = "\n📋 *Backlog Tickets*\n\n";
    const backlogUrl = this.config.backlogUrl.replace(/\/$/, ""); // Remove trailing slash

    // Build clean, professional list format
    let hasTickets = false;
    for (let i = 0; i < this.config.engineerNames.length; i++) {
      const engineerName = this.config.engineerNames[i];
      const issues = ticketsByEngineer[engineerName] || [];
      const totalCount = issues.length;

      // Filter only urgent tickets (overdue, today, tomorrow)
      const urgentIssues = [];
      for (let j = 0; j < issues.length; j++) {
        const issueData = issues[j];
        const dueDateInfo = this._getDueDateIcon(issueData.dueDate);
        if (dueDateInfo.isUrgent) {
          urgentIssues.push({
            issueKey: issueData.issueKey,
            summary: issueData.summary,
            icon: dueDateInfo.icon,
          });
        }
      }

      if (totalCount > 0) {
        hasTickets = true;

        // Build issue links with due date icons and truncated summary (only urgent)
        const issueLinks = [];
        for (let j = 0; j < urgentIssues.length; j++) {
          const issue = urgentIssues[j];
          const issueUrl = backlogUrl + "/view/" + issue.issueKey;
          const truncatedSummary = this._truncateSummary(issue.summary, 3);
          issueLinks.push(
            "<" + issueUrl + "|" + truncatedSummary + ">" + issue.icon
          );
        }

        // Show total count, but only list urgent tickets
        if (urgentIssues.length > 0) {
          message +=
            "• *" +
            engineerName +
            "*  `" +
            totalCount +
            "` (" +
            issueLinks.join(", ") +
            ")\n";
        } else {
          message += "• *" + engineerName + "*  `" + totalCount + "`\n";
        }
      }
    }

    if (!hasTickets) {
      message += "✓ No open tickets\n";
    }

    return message;
  }

  /**
   * Main method to run the daily verification process
   */
  runDailyVerification() {
    Logger.log("Starting daily task verification...");
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Check if we should run verification today
    const lastWorkingDayInfo = this.getLastWorkingDay();
    if (!lastWorkingDayInfo) {
      Logger.log(
        "No verification needed - today is weekend or no valid working day to check"
      );
      return [];
    }
    const lastWorkingDay = lastWorkingDayInfo.dateString;

    // Verify all employee sheets
    const results = this.verifyAllEmployees();

    // Find employees who need to be reminded (no data or poor performance)
    const employeesToRemind = [];
    let summaryMessage =
      "📊 Daily Task Report Summary - " + lastWorkingDay + " \n";

    for (let i = 0; i < results.length; i++) {
      const data = results[i];
      if (Array.isArray(data) && !data[1]) {
        employeesToRemind.push(data[0]);
        summaryMessage += data[0] + "❌ Not Added \n";
      } else {
        summaryMessage += data + "\n";
      }
    }

    // Get RSS feed articles
    const rssAggregator = new CategorizedRSSFeedAggregator();
    const category = rssAggregator.getCurrentDayCategory();
    const articleCount = this.config.rssArticleCount;
    const pickedArticles = rssAggregator.getRandomArticles(
      articleCount,
      category
    );

    if (pickedArticles && pickedArticles.length > 0) {
      Logger.log("\n✅ Found " + pickedArticles.length + " articles:");
      for (let i = 0; i < pickedArticles.length; i++) {
        const article = pickedArticles[i];
        Logger.log(i + 1 + ". " + article.title.substring(0, 80) + "...");
        Logger.log("   🔗 " + article.link);
      }
    } else {
      Logger.log("❌ No new articles found (all may have been sent before)");
    }

    const linkedinPost = rssAggregator.generateLinkedInPost(
      articleCount,
      category
    );

    // Get Backlog ticket counts
    let backlogMessage = "";
    if (this.config.backlogApiKey && this.config.backlogUrl) {
      const ticketCounts = this._countBacklogTicketsPerEngineer();
      backlogMessage = this._formatBacklogTicketCounts(ticketCounts);
      Logger.log(backlogMessage);
    }

    // Send employee reminders if needed
    if (employeesToRemind.length > 0) {
      Logger.log("Sending reminders to: " + employeesToRemind.join(", "));
      this.sendEmployeeReminder(employeesToRemind);
    } else {
      Logger.log(
        "No employee reminders needed - all timesheets are properly submitted"
      );
    }

    // Send LinkedIn post if available (include Backlog info)
    if (linkedinPost) {
      let postWithBacklog = linkedinPost;
      if (backlogMessage) {
        postWithBacklog = backlogMessage + "\n" + linkedinPost;
      }
      this.sendGoogleChatMessage(
        postWithBacklog,
        this.config.employeeAlertWebhookUrl
      );
    } else if (backlogMessage) {
      // Send Backlog info even if no articles
      this.sendGoogleChatMessage(
        backlogMessage,
        this.config.employeeAlertWebhookUrl
      );
    }

    Logger.log("\n" + "=".repeat(50));
    Logger.log("VERIFICATION SUMMARY");
    Logger.log("=".repeat(50));
    Logger.log(summaryMessage);
    if (backlogMessage) {
      Logger.log(backlogMessage);
    }
    Logger.log("=".repeat(50));

    // Include Backlog info in summary message
    let finalSummaryMessage = summaryMessage;
    if (backlogMessage) {
      finalSummaryMessage += backlogMessage;
    }
    this.sendGoogleChatMessage(finalSummaryMessage);

    return results;
  }
}

/**
 * Main function to run the verification
 * This can be called manually or set up as a time-driven trigger
 */
function runDailyVerification() {
  const verifier = new SheetsVerifier();
  const results = verifier.runDailyVerification();
  return results;
}

/**
 * RSS Feed Aggregator Class
 * Simplified version for Apps Script (basic RSS parsing)
 */
class CategorizedRSSFeedAggregator {
  constructor() {
    this.categorizedFeeds = {
      monday: {
        name: "Infrastructure & DevOps",
        feeds: [
          "https://devops.com/feed",
          "https://atlassian.com/blog/devops/feed",
          "https://www.docker.com/blog/feed/",
          "https://kubernetes.io/feed.xml",
        ],
      },
      tuesday: {
        name: "Python & Backend Technologies",
        feeds: [
          "https://realpython.com/atom.xml",
          "https://planetpython.org/rss20.xml",
          "https://fastapi.tiangolo.com/rss.xml",
        ],
      },
      wednesday: {
        name: "React & Frontend Technologies",
        feeds: [
          "https://reactjs.org/feed.xml",
          "https://blog.logrocket.com/feed/",
          "https://css-tricks.com/feed/",
        ],
      },
      thursday: {
        name: "Testing & QA Tools",
        feeds: [
          "https://blog.qasource.com/rss.xml",
          "https://browserstack.com/blog/rss",
        ],
      },
      friday: {
        name: "Fun & Interesting Tech",
        feeds: [
          "https://xkcd.com/rss.xml",
          "https://github.blog/feed/",
          "https://news.ycombinator.com/rss",
        ],
      },
    };

    this.categoryEmojis = {
      monday: ["🚀", "🛠️", "⚙️", "🔧"],
      tuesday: ["🐍", "⚡", "🔥", "💻"],
      wednesday: ["⚛️", "🎨", "💡", "🌟"],
      thursday: ["🧪", "🔍", "🎯", "✅"],
      friday: ["😄", "🎉", "🤖", "🎮"],
    };

    this.categoryIntros = {
      monday: ["🚀 Monday Motivation: Infrastructure & DevOps Edition!"],
      tuesday: [
        "🐍 Tuesday's Python Power! Backend technologies that'll boost your skills:",
      ],
      wednesday: ["⚛️ React Wednesday: Frontend frameworks and UI magic:"],
      thursday: [
        "🧪 Testing Thursday: QA tools and quality assurance insights:",
      ],
      friday: [
        "🎉 Fun Friday: Tech humor, interesting projects, and weekend inspiration:",
      ],
    };

    this.categoryHashtags = {
      monday: "#DevOps #Infrastructure #Docker #Kubernetes",
      tuesday: "#Python #Backend #API #Django",
      wednesday: "#React #Frontend #JavaScript #CSS",
      thursday: "#Testing #QA #Automation #Selenium",
      friday: "#TechFun #Programming #GitHub #OpenSource",
    };

    // Load sent links from PropertiesService
    this.sentLinks = this._loadSentLinks();
  }

  _loadSentLinks() {
    try {
      const linksStr =
        PropertiesService.getScriptProperties().getProperty("SENT_LINKS");
      if (linksStr) {
        return JSON.parse(linksStr);
      }
    } catch (e) {
      Logger.log("Error loading sent links: " + e);
    }
    return [];
  }

  _saveSentLinks() {
    PropertiesService.getScriptProperties().setProperty(
      "SENT_LINKS",
      JSON.stringify(this.sentLinks)
    );
  }

  getCurrentDayCategory() {
    const dayMapping = {
      0: "monday", // Sunday defaults to Monday
      1: "monday",
      2: "tuesday",
      3: "wednesday",
      4: "thursday",
      5: "friday",
      6: "monday", // Saturday defaults to Monday
    };
    const currentDay = new Date().getDay();
    return dayMapping[currentDay];
  }

  fetchArticlesByCategory(category) {
    if (!this.categorizedFeeds[category]) {
      return [];
    }

    const articles = [];
    const feeds = this.categorizedFeeds[category].feeds;

    for (let i = 0; i < feeds.length; i++) {
      try {
        const feedUrl = feeds[i];
        const response = UrlFetchApp.fetch(feedUrl, {
          muteHttpExceptions: true,
        });

        if (response.getResponseCode() === 200) {
          const xml = response.getContentText();
          const items = this._parseRSS(xml);

          for (let j = 0; j < items.length; j++) {
            articles.push({
              title: items[j].title,
              link: items[j].link,
              category: category,
            });
          }
        }
      } catch (e) {
        Logger.log("Error fetching from feed: " + e);
      }
    }

    return articles;
  }

  _parseRSS(xml) {
    // Simple RSS parser - extracts title and link from RSS/Atom feeds
    const items = [];

    // Try to parse as RSS
    const itemMatches = xml.match(/<item[^>]*>([\s\S]*?)<\/item>/gi);
    if (itemMatches) {
      for (let i = 0; i < itemMatches.length; i++) {
        const item = itemMatches[i];
        const titleMatch = item.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
        const linkMatch = item.match(/<link[^>]*>([\s\S]*?)<\/link>/i);

        if (titleMatch && linkMatch) {
          items.push({
            title: this._cleanText(titleMatch[1]),
            link: this._cleanText(linkMatch[1]),
          });
        }
      }
    } else {
      // Try to parse as Atom
      const entryMatches = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/gi);
      if (entryMatches) {
        for (let i = 0; i < entryMatches.length; i++) {
          const entry = entryMatches[i];
          const titleMatch = entry.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
          const linkMatch = entry.match(/<link[^>]*href=["']([^"']+)["']/i);

          if (titleMatch && linkMatch) {
            items.push({
              title: this._cleanText(titleMatch[1]),
              link: linkMatch[1],
            });
          }
        }
      }
    }

    return items;
  }

  _cleanText(text) {
    return text
      .replace(/<[^>]+>/g, "")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .trim();
  }

  getRandomArticles(count, category) {
    if (!category) {
      category = this.getCurrentDayCategory();
    }

    const allArticles = this.fetchArticlesByCategory(category);
    const filteredArticles = [];

    for (let i = 0; i < allArticles.length; i++) {
      if (this.sentLinks.indexOf(allArticles[i].link) === -1) {
        filteredArticles.push(allArticles[i]);
      }
    }

    // Shuffle and return up to count
    for (let i = filteredArticles.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [filteredArticles[i], filteredArticles[j]] = [
        filteredArticles[j],
        filteredArticles[i],
      ];
    }

    return filteredArticles.slice(0, count);
  }

  generateLinkedInPost(articleCount, category) {
    if (!category) {
      category = this.getCurrentDayCategory();
    }

    const articles = this.getRandomArticles(articleCount, category);
    if (articles.length === 0) {
      return null;
    }

    const intro = this.categoryIntros[category][0];
    const emojis = this.categoryEmojis[category];
    const hashtags = this.categoryHashtags[category];

    let post = intro + "\n\n";

    for (let i = 0; i < articles.length; i++) {
      const emoji = emojis[Math.floor(Math.random() * emojis.length)];
      post +=
        i +
        1 +
        ". " +
        emoji +
        " " +
        articles[i].title +
        "  \n" +
        articles[i].link +
        "\n\n";
    }

    post += "\n" + hashtags + "\n";

    // Track sent links
    for (let i = 0; i < articles.length; i++) {
      this.sentLinks.push(articles[i].link);
    }
    this._saveSentLinks();

    return post;
  }
}

/**
 * ============================================================================
 * SIMPLIFIED CONFIGURATION FUNCTIONS
 * ============================================================================
 * Use these functions to easily configure your timesheet verifier
 */

/**
 * Set engineer names (simplified - just pass an array)
 * Example: setEngineerNames(['John Doe', 'Jane Smith', 'Bob Johnson'])
 */
function setEngineerNames(engineerNames) {
  if (!Array.isArray(engineerNames)) {
    Logger.log("ERROR: engineerNames must be an array");
    Logger.log('Example: setEngineerNames(["John Doe", "Jane Smith"])');
    return false;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty("ENGINEER_NAMES", engineerNames.join(","));
  Logger.log("✓ Engineer names set: " + engineerNames.join(", "));
  return true;
}

/**
 * Add or update a single engineer with their chat ID
 * Example: addEngineer('John Doe', '123456789012345678901')
 *
 * @param {string} engineerName - The engineer's name (must match sheet name exactly)
 * @param {string} chatId - Google Chat user ID (optional, can be null)
 */
function addEngineer(engineerName, chatId) {
  if (!engineerName) {
    Logger.log("ERROR: engineerName is required");
    return false;
  }

  const props = PropertiesService.getScriptProperties();

  // Get current engineer names
  const currentNamesStr = props.getProperty("ENGINEER_NAMES") || "";
  const currentNames = currentNamesStr
    ? currentNamesStr.split(",").map((n) => n.trim())
    : [];

  // Add engineer name if not already present
  if (currentNames.indexOf(engineerName) === -1) {
    currentNames.push(engineerName);
    props.setProperty("ENGINEER_NAMES", currentNames.join(","));
    Logger.log("✓ Added engineer: " + engineerName);
  } else {
    Logger.log("ℹ Engineer already exists: " + engineerName);
  }

  // Update chat ID if provided
  if (chatId) {
    let chatIds = {};
    try {
      const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
      if (chatIdsStr) {
        chatIds = JSON.parse(chatIdsStr);
      }
    } catch (e) {
      Logger.log("Warning: Could not parse existing chat IDs");
    }

    chatIds[engineerName] = chatId;
    props.setProperty("EMPLOYEE_CHAT_IDS", JSON.stringify(chatIds));
    Logger.log("✓ Set chat ID for " + engineerName + ": " + chatId);
  }

  return true;
}

/**
 * Remove an engineer from the list
 * Example: removeEngineer('John Doe')
 */
function removeEngineer(engineerName) {
  if (!engineerName) {
    Logger.log("ERROR: engineerName is required");
    return false;
  }

  const props = PropertiesService.getScriptProperties();

  // Remove from engineer names
  const currentNamesStr = props.getProperty("ENGINEER_NAMES") || "";
  const currentNames = currentNamesStr
    ? currentNamesStr.split(",").map((n) => n.trim())
    : [];
  const index = currentNames.indexOf(engineerName);

  if (index !== -1) {
    currentNames.splice(index, 1);
    props.setProperty("ENGINEER_NAMES", currentNames.join(","));
    Logger.log("✓ Removed engineer: " + engineerName);
  } else {
    Logger.log("ℹ Engineer not found: " + engineerName);
  }

  // Remove from chat IDs
  try {
    const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
    if (chatIdsStr) {
      const chatIds = JSON.parse(chatIdsStr);
      if (chatIds[engineerName]) {
        delete chatIds[engineerName];
        props.setProperty("EMPLOYEE_CHAT_IDS", JSON.stringify(chatIds));
        Logger.log("✓ Removed chat ID for " + engineerName);
      }
    }
  } catch (e) {
    Logger.log("Warning: Could not update chat IDs");
  }

  return true;
}

/**
 * Set chat ID for an existing engineer
 * Example: setEngineerChatId('John Doe', '123456789012345678901')
 */
function setEngineerChatId(engineerName, chatId) {
  if (!engineerName || !chatId) {
    Logger.log("ERROR: Both engineerName and chatId are required");
    return false;
  }

  const props = PropertiesService.getScriptProperties();

  let chatIds = {};
  try {
    const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
    if (chatIdsStr) {
      chatIds = JSON.parse(chatIdsStr);
    }
  } catch (e) {
    Logger.log("Warning: Could not parse existing chat IDs");
  }

  chatIds[engineerName] = chatId;
  props.setProperty("EMPLOYEE_CHAT_IDS", JSON.stringify(chatIds));
  Logger.log("✓ Set chat ID for " + engineerName + ": " + chatId);
  return true;
}

/**
 * View current configuration
 * Shows all engineers and their chat IDs
 */
function viewConfiguration() {
  const props = PropertiesService.getScriptProperties();

  Logger.log("=".repeat(60));
  Logger.log("CURRENT CONFIGURATION");
  Logger.log("=".repeat(60));

  // Main settings
  Logger.log("\nMain Settings:");
  Logger.log(
    "  CONFIG_FILE_ID: " + (props.getProperty("CONFIG_FILE_ID") || "NOT SET")
  );
  Logger.log(
    "  DESTINATION_FOLDER_ID: " +
      (props.getProperty("DESTINATION_FOLDER_ID") || "NOT SET")
  );
  Logger.log(
    "  MAIN_SPREADSHEET_ID (fallback): " +
      (props.getProperty("MAIN_SPREADSHEET_ID") || "NOT SET")
  );

  // Show which spreadsheet ID is actually being used
  try {
    const verifier = new SheetsVerifier();
    const currentSpreadsheetId = verifier.config.mainSpreadsheetId;
    if (currentSpreadsheetId) {
      Logger.log(
        "  ✓ Current Spreadsheet ID (from config file): " + currentSpreadsheetId
      );
    } else {
      Logger.log(
        "  ⚠ No spreadsheet ID found (check config file or MAIN_SPREADSHEET_ID)"
      );
    }
  } catch (e) {
    Logger.log("  ⚠ Could not determine current spreadsheet ID: " + e);
  }

  Logger.log(
    "  GOOGLE_CHAT_WEBHOOK_URL: " +
      (props.getProperty("GOOGLE_CHAT_WEBHOOK_URL") ? "SET" : "NOT SET")
  );
  Logger.log(
    "  EMPLOYEE_ALERT_WEBHOOK_URL: " +
      (props.getProperty("EMPLOYEE_ALERT_WEBHOOK_URL") ? "SET" : "NOT SET")
  );
  Logger.log("  TEST_MODE: " + (props.getProperty("TEST_MODE") || "false"));
  Logger.log(
    "  ROWS_TO_CHECK_AFTER_DATE: " +
      (props.getProperty("ROWS_TO_CHECK_AFTER_DATE") || "5")
  );
  Logger.log(
    "  SHEET_DATA_RANGE: " +
      (props.getProperty("SHEET_DATA_RANGE") || "A1:K150")
  );
  Logger.log(
    "  RSS_ARTICLE_COUNT: " + (props.getProperty("RSS_ARTICLE_COUNT") || "5")
  );

  // Backlog settings
  Logger.log("\nBacklog Settings:");
  Logger.log(
    "  BACKLOG_URL: " + (props.getProperty("BACKLOG_URL") || "NOT SET")
  );
  Logger.log(
    "  BACKLOG_API_KEY: " +
      (props.getProperty("BACKLOG_API_KEY") ? "SET" : "NOT SET")
  );
  Logger.log(
    "  BACKLOG_PROJECT_IDS: " +
      (props.getProperty("BACKLOG_PROJECT_IDS") || "NOT SET")
  );
  try {
    const mappingStr = props.getProperty("BACKLOG_ENGINEER_MAPPING");
    if (mappingStr) {
      const mapping = JSON.parse(mappingStr);
      Logger.log(
        "  BACKLOG_ENGINEER_MAPPING: " +
          Object.keys(mapping).length +
          " mappings"
      );
    } else {
      Logger.log("  BACKLOG_ENGINEER_MAPPING: NOT SET");
    }
  } catch (e) {
    Logger.log("  BACKLOG_ENGINEER_MAPPING: Error reading");
  }

  // Engineers
  Logger.log("\nEngineers:");
  const engineerNamesStr = props.getProperty("ENGINEER_NAMES") || "";
  const engineerNames = engineerNamesStr
    ? engineerNamesStr.split(",").map((n) => n.trim())
    : [];

  if (engineerNames.length === 0) {
    Logger.log("  No engineers configured");
  } else {
    // Get chat IDs
    let chatIds = {};
    try {
      const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
      if (chatIdsStr) {
        chatIds = JSON.parse(chatIdsStr);
      }
    } catch (e) {
      // Ignore
    }

    for (let i = 0; i < engineerNames.length; i++) {
      const name = engineerNames[i];
      const chatId = chatIds[name] || "NOT SET";
      Logger.log("  " + (i + 1) + ". " + name + " (Chat ID: " + chatId + ")");
    }
  }

  // Holidays
  Logger.log("\nHolidays:");
  try {
    const holidaysStr = props.getProperty("HOLIDAYS");
    if (holidaysStr) {
      const holidays = JSON.parse(holidaysStr);
      Logger.log("  " + holidays.length + " holidays configured");
      for (let i = 0; i < holidays.length; i++) {
        Logger.log("    - " + holidays[i]);
      }
    } else {
      Logger.log("  No holidays configured");
    }
  } catch (e) {
    Logger.log("  Error reading holidays");
  }

  Logger.log("\n" + "=".repeat(60));
}

/**
 * Ultra-simple setup - configure engineers from a single string
 * Format: "name1|chat_id1,name2|chat_id2,name3" (chat ID optional)
 * Example: setupEngineersFromString("John Doe|123456789012345678901,Jane Smith|987654321098765432109,Bob Johnson")
 *
 * This is the simplest format - just one string, easy to copy/paste!
 */
function setupEngineersFromString(engineersString) {
  if (!engineersString || typeof engineersString !== "string") {
    Logger.log("ERROR: engineersString must be a string");
    Logger.log(
      'Example: setupEngineersFromString("John Doe|123456789,Jane Smith|987654321")'
    );
    return false;
  }

  const engineers = [];
  const parts = engineersString.split(",");

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    if (!part) continue;

    const pipeIndex = part.indexOf("|");
    if (pipeIndex !== -1) {
      // Has chat ID
      const name = part.substring(0, pipeIndex).trim();
      const chatId = part.substring(pipeIndex + 1).trim();
      engineers.push({ name: name, chatId: chatId });
    } else {
      // No chat ID
      engineers.push({ name: part });
    }
  }

  if (engineers.length === 0) {
    Logger.log("ERROR: No engineers found in string");
    return false;
  }

  return quickSetupEngineers(engineers);
}

/**
 * Quick setup - configure multiple engineers at once
 * Example: quickSetupEngineers([
 *   {name: 'John Doe', chatId: '123456789012345678901'},
 *   {name: 'Jane Smith', chatId: '987654321098765432109'},
 *   {name: 'Bob Johnson'}  // No chat ID
 * ])
 */
function quickSetupEngineers(engineers) {
  if (!Array.isArray(engineers)) {
    Logger.log("ERROR: engineers must be an array");
    Logger.log(
      'Example: quickSetupEngineers([{name: "John Doe", chatId: "123..."}])'
    );
    return false;
  }

  const names = [];
  const chatIds = {};

  for (let i = 0; i < engineers.length; i++) {
    const eng = engineers[i];
    if (!eng.name) {
      Logger.log('ERROR: Each engineer must have a "name" property');
      return false;
    }

    names.push(eng.name);
    if (eng.chatId) {
      chatIds[eng.name] = eng.chatId;
    }
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty("ENGINEER_NAMES", names.join(","));
  props.setProperty("EMPLOYEE_CHAT_IDS", JSON.stringify(chatIds));

  Logger.log("✓ Configured " + names.length + " engineers");
  Logger.log("  Names: " + names.join(", "));
  Logger.log(
    "  Chat IDs set for: " + Object.keys(chatIds).length + " engineers"
  );

  return true;
}

/**
 * Setup function to configure Script Properties
 * Run this once to set up your configuration
 *
 * For easier setup, use the helper functions above instead:
 * - setEngineerNames(['Name1', 'Name2'])
 * - addEngineer('Name', 'ChatID')
 * - quickSetupEngineers([{name: 'Name', chatId: 'ID'}])
 */
function setupConfiguration() {
  const props = PropertiesService.getScriptProperties();

  // Set your configuration values here
  // Note: The script will automatically fetch spreadsheet ID from timesheet_config.json
  // based on current year-month. Set these only if you want to use a fallback:
  props.setProperty("MAIN_SPREADSHEET_ID", "YOUR_SPREADSHEET_ID_HERE"); // Fallback only
  props.setProperty("CONFIG_FILE_ID", ""); // Optional: ID of timesheet_config.json file
  props.setProperty("DESTINATION_FOLDER_ID", ""); // Optional: Folder ID where config file is stored
  props.setProperty("GOOGLE_CHAT_WEBHOOK_URL", "YOUR_WEBHOOK_URL");
  props.setProperty("EMPLOYEE_ALERT_WEBHOOK_URL", "YOUR_ALERT_WEBHOOK_URL");
  props.setProperty("TEST_MODE", "false");
  props.setProperty("ROWS_TO_CHECK_AFTER_DATE", "5");
  props.setProperty("SHEET_DATA_RANGE", "A1:K150");
  props.setProperty("RSS_ARTICLE_COUNT", "5");

  // Backlog configuration (optional)
  props.setProperty("BACKLOG_URL", "https://ilabs.backlog.com");
  props.setProperty("BACKLOG_API_KEY", "YOUR_BACKLOG_API_KEY");
  props.setProperty("BACKLOG_PROJECT_IDS", "PROJECT_ID_1,PROJECT_ID_2"); // Comma-separated

  // Backlog engineer name mapping (optional - maps engineer names to Backlog assignee names)
  // Example: If engineer is "Bismillakhan S" but Backlog shows "Bismillakhan", add mapping
  const backlogMapping = {
    // "Bismillakhan S": "Bismillakhan",
    // "Jinu T J": "Jinu",
  };
  props.setProperty("BACKLOG_ENGINEER_MAPPING", JSON.stringify(backlogMapping));

  // Use simplified functions for engineers
  Logger.log("Setting up engineers...");
  Logger.log(
    'You can also use: quickSetupEngineers([{name: "Name", chatId: "ID"}])'
  );

  // Example: Set engineers using the simplified function
  setEngineerNames([
    "Jinu T J",
    "Bismillakhan S",
    "Midhun",
    "Aravind",
    "Akhil Mohan",
    "Akash T K",
    "Shinoj",
  ]);

  // Example: Set chat IDs using the simplified function
  addEngineer("Bismillakhan S", "115773514265520053097");
  addEngineer("Jinu T J", "115142352767738659510");
  addEngineer("Midhun", "108660074852905944786");
  addEngineer("Aravind", "107574045682955308493");
  addEngineer("Akhil Mohan", "104052864034579334603");
  addEngineer("Akash T K", "112652200937440721836");
  addEngineer("Shinoj", "117004017244973974527");

  // Set holidays as JSON array
  const holidays = [
    "2025-04-10",
    "2025-05-01",
    "2025-08-03",
    "2025-08-15",
    "2025-09-14",
    "2025-10-01",
    "2025-10-02",
    "2025-10-20",
    "2025-12-25",
  ];
  props.setProperty("HOLIDAYS", JSON.stringify(holidays));

  Logger.log("\n✓ Configuration set up successfully!");
  Logger.log("Run viewConfiguration() to see your current settings.");
}
