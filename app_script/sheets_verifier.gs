/**
 * Google Apps Script - Timesheet Verification System
 *
 * This script verifies employee timesheet entries and sends daily reports.
 * All team configuration lives in the master "Teams Config" sheet (same sheet
 * used by the timesheet creator). The verifier reads it on every run.
 *
 * ============================================================================
 * QUICK START
 * ============================================================================
 *
 * 1. Set the master sheet ID:
 *    setupMultiTeam('YOUR_MASTER_SHEET_ID');
 *
 * 2. Verify your config:
 *    viewMultiTeamConfiguration();
 *
 * 3. Set your time-driven trigger to: runDailyVerification
 *
 * ============================================================================
 * SCRIPT PROPERTIES
 * ============================================================================
 *
 * Required:
 * - MASTER_SHEET_ID: Spreadsheet ID of the master "Teams Config" sheet
 *
 * Optional (global, not team-specific):
 * - TEST_MODE: "true" or "false" (default: false)
 * - ROWS_TO_CHECK_AFTER_DATE: Number of rows to check after date (default: 5)
 * - SHEET_DATA_RANGE: Range to read per engineer sheet (default: "A1:K150")
 * - RSS_ARTICLE_COUNT: Number of RSS articles to fetch (default: 5)
 *
 * All other config (engineers, webhooks, holidays, Backlog settings) is read
 * per-team from the master sheet columns. See create_test_master_sheet.gs for
 * the full column layout.
 * ============================================================================
 */

/**
 * Main class for Sheets Verification
 * @param {Object} teamConfig - Team configuration object from master sheet (via readVerifierTeamConfigs())
 */
class SheetsVerifier {
  constructor(teamConfig) {
    this.props = PropertiesService.getScriptProperties();
    this._teamConfig = teamConfig;
    this.config = this._loadConfigFromTeam(teamConfig);
  }

  /**
   * Build this.config from a team configuration object (from master sheet).
   * Global settings (rowsToCheckAfterDate, sheetDataRange, rssArticleCount, testMode)
   * are still read from Script Properties since they are not team-specific.
   */
  _loadConfigFromTeam(team) {
    const props = this.props;

    // Resolve spreadsheet ID for this team
    let mainSpreadsheetId = this._getSpreadsheetIdForDate(new Date(), team.projectName, team.configFileId, team.driveFolderId);
    if (!mainSpreadsheetId) {
      mainSpreadsheetId = null; // no fallback in multi-team mode
    }

    return {
      projectName: team.projectName,
      mainSpreadsheetId: mainSpreadsheetId,
      engineerNames: team.engineerNames || [],
      googleChatWebhookUrl: team.webhookUrl || '',
      employeeAlertWebhookUrl: team.alertWebhookUrl || '',
      testMode: props.getProperty('TEST_MODE') === 'true',
      employeeChatIds: team.employeeChatIds || {},
      holidays: team.holidays || [],
      rowsToCheckAfterDate: parseInt(props.getProperty('ROWS_TO_CHECK_AFTER_DATE') || '5', 10),
      sheetDataRange: props.getProperty('SHEET_DATA_RANGE') || 'A1:K150',
      rssArticleCount: parseInt(props.getProperty('RSS_ARTICLE_COUNT') || '5', 10),
      backlogUrl: team.backlogUrl || '',
      backlogApiKey: team.backlogApiKey || '',
      backlogProjectIds: team.backlogProjectIds || '',
      backlogEngineerMapping: team.backlogEngineerMapping || {},
    };
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
      return this._getSpreadsheetIdForDate(
        new Date(),
        this._teamConfig.projectName,
        this._teamConfig.configFileId,
        this._teamConfig.driveFolderId
      );
    } catch (e) {
      Logger.log("Error in _getCurrentMonthSpreadsheetId: " + e);
      return null;
    }
  }
  /**
   * Fetch spreadsheet ID from timesheet_config.json based on a specific date's year-month
   * Returns the sheet ID for the specified date's month or null if not found
   * @param {Date} targetDate - Optional date to get sheet ID for (defaults to current date)
   * @param {string} [projectName] - Optional project name for multi-team key lookup (e.g. "Backend Team")
   * @param {string} [overrideConfigFileId] - Optional config file ID (from master sheet)
   * @param {string} [overrideFolderId] - Optional folder ID (from master sheet)
   */
  _getSpreadsheetIdForDate(targetDate, projectName, overrideConfigFileId, overrideFolderId) {
    try {
      const props = PropertiesService.getScriptProperties();
      const configFileId = overrideConfigFileId || props.getProperty("CONFIG_FILE_ID");
      const destinationFolderId = overrideFolderId || props.getProperty("DESTINATION_FOLDER_ID");

      if (!configFileId && !destinationFolderId) {
        Logger.log("CONFIG_FILE_ID or DESTINATION_FOLDER_ID not configured");
        return null;
      }

      // Try to get config file by ID first
      let configFile = null;
      if (configFileId) {
        try {
          const candidate = DriveApp.getFileById(configFileId);
          // Validate it's actually a JSON/text file, not a PDF or Google Doc
          const mimeType = candidate.getMimeType();
          if (mimeType === MimeType.PLAIN_TEXT || mimeType === 'application/json' || mimeType === 'text/plain') {
            configFile = candidate;
          } else {
            Logger.log("Config File ID points to a " + mimeType + " file, not JSON. Searching by name instead.");
          }
        } catch (e) {
          Logger.log("Could not find config file by ID, trying by name");
        }
      }

      // If not found by ID, try to find by name in folder
      if (!configFile && destinationFolderId) {
        try {
          const folder = DriveApp.getFolderById(destinationFolderId);
          const files = folder.getFilesByName("timesheet_config.json");
          while (files.hasNext()) {
            const candidate = files.next();
            const mimeType = candidate.getMimeType();
            if (mimeType === MimeType.PLAIN_TEXT || mimeType === 'application/json' || mimeType === 'text/plain') {
              configFile = candidate;
              break;
            } else {
              Logger.log("Skipping non-JSON file in folder: " + candidate.getName() + " (type: " + mimeType + ")");
            }
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
          while (files.hasNext()) {
            const candidate = files.next();
            const mimeType = candidate.getMimeType();
            if (mimeType === MimeType.PLAIN_TEXT || mimeType === 'application/json' || mimeType === 'text/plain') {
              configFile = candidate;
              break;
            } else {
              Logger.log("Skipping non-JSON file in root: " + candidate.getName() + " (type: " + mimeType + ")");
            }
          }
        } catch (e) {
          Logger.log("Could not find config file in root folder");
        }
      }

      if (!configFile) {
        Logger.log("timesheet_config.json not found");
        return null;
      }

      Logger.log("✓ Found config file: " + configFile.getName() + " (ID: " + configFile.getId() + ", MIME: " + configFile.getMimeType() + ")");

      // Read and parse config file
      const configContent = configFile.getBlob().getDataAsString();
      const config = JSON.parse(configContent);

      // Use targetDate if provided, otherwise use current date
      const dateToUse = targetDate || new Date();
      const targetYear = dateToUse.getFullYear();
      const targetMonth = this._getMonthNameFromDate(dateToUse);

      // Try project-prefixed key first (matches creator's format: "ProjectName_Year-Month")
      // Then fall back to non-prefixed key for backward compatibility
      let spreadsheetId = null;
      if (projectName) {
        const prefixedKey = projectName + "_" + targetYear + "-" + targetMonth;
        spreadsheetId = config[prefixedKey];
        if (spreadsheetId) {
          Logger.log("✓ Found spreadsheet ID for " + prefixedKey + ": " + spreadsheetId);
          return spreadsheetId;
        }
        Logger.log("No spreadsheet ID found for prefixed key: " + prefixedKey + ", trying non-prefixed...");
      }

      // Fall back to non-prefixed key (legacy format: "Year-Month")
      const configKey = targetYear + "-" + targetMonth;
      spreadsheetId = config[configKey];

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

  /**\n   * Get current month sheet name (e.g., 'July', 'August')\n   */
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
    const spreadsheetId = this._getSpreadsheetIdForDate(
      lastWorkingDayDate,
      this._teamConfig.projectName,
      this._teamConfig.configFileId,
      this._teamConfig.driveFolderId
    );
    if (!spreadsheetId) {
      Logger.log(
        "Warning: Could not find spreadsheet ID for " +
          this._teamConfig.projectName + '_' +
          lastWorkingDayYear +
          "-" +
          lastWorkingDayMonth +
          ", using fallback"
      );
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
   * Also fetches resolved issues for Bug type tickets
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

      const allIssues = [];

      // Fetch open and in-progress issues (status 1 = Open, 2 = In Progress)
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
        "Fetching open/in-progress Backlog issues from: " +
          url.replace(this.config.backlogApiKey, "***")
      );

      const options = {
        method: "get",
        muteHttpExceptions: true,
      };

      let response = UrlFetchApp.fetch(url, options);
      let responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const issues = JSON.parse(response.getContentText());
        Logger.log(
          "✓ Fetched " + issues.length + " open/in-progress issues from Backlog"
        );
        for (let k = 0; k < issues.length; k++) {
          allIssues.push(issues[k]);
        }
      } else {
        Logger.log(
          "Error fetching open Backlog issues: " +
            responseCode +
            " - " +
            response.getContentText()
        );
      }

      // Also fetch resolved issues (status 3 = Resolved)
      // These will be filtered to only include Bugs in the counting function
      let resolvedUrl =
        this.config.backlogUrl +
        "/api/v2/issues?apiKey=" +
        this.config.backlogApiKey;

      // Add project IDs
      for (let i = 0; i < projectIds.length; i++) {
        resolvedUrl += "&projectId[]=" + encodeURIComponent(projectIds[i]);
      }

      // Add resolved status ID (3 = Resolved)
      resolvedUrl += "&statusId[]=3";

      Logger.log(
        "Fetching resolved Backlog issues from: " +
          resolvedUrl.replace(this.config.backlogApiKey, "***")
      );

      response = UrlFetchApp.fetch(resolvedUrl, options);
      responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const resolvedIssues = JSON.parse(response.getContentText());
        Logger.log(
          "✓ Fetched " + resolvedIssues.length + " resolved issues from Backlog"
        );
        for (let k = 0; k < resolvedIssues.length; k++) {
          allIssues.push(resolvedIssues[k]);
        }
      } else {
        Logger.log(
          "Error fetching resolved Backlog issues: " +
            responseCode +
            " - " +
            response.getContentText()
        );
      }

      Logger.log("✓ Total issues fetched: " + allIssues.length);
      return allIssues;
    } catch (e) {
      Logger.log("Error fetching Backlog issues: " + e);
      return [];
    }
  }

  /**
   * Count open Backlog tickets per engineer
   * Includes all open/in-progress tickets and resolved Bug tickets
   * Returns object mapping engineer names to array of issue objects {issueKey, dueDate, summary, issueType, status}
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

        // Get issue type and status
        const issueType = issue.issueType ? issue.issueType.name : "";
        const statusId = issue.status ? issue.status.id : null;
        const statusName = issue.status ? issue.status.name : "";

        // For resolved tickets (status ID 3), only include if it's a Bug
        // For open/in-progress tickets (status ID 1 or 2), include all types
        const isResolved = statusId === 3;
        const isBug = issueType && issueType.toLowerCase() === "bug";

        if (isResolved && !isBug) {
          // Skip resolved tickets that are not bugs
          continue;
        }

        const issueData = {
          issueKey: issue.issueKey,
          dueDate: issue.dueDate || null,
          summary: issue.summary || "",
          issueType: issueType,
          status: statusName,
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
   * Run verification for this team.
   * Accepts a pre-fetched RSS post to avoid redundant fetches across teams.
   *
   * @param {string|null} linkedinPost - Pre-fetched RSS/LinkedIn post (shared across teams)
   */
  runTeamVerification(linkedinPost) {
    const teamLabel = this.config.projectName || 'Unknown Team';
    Logger.log('Starting verification for team: ' + teamLabel);

    // Get last working day
    const lastWorkingDayInfo = this.getLastWorkingDay();
    if (!lastWorkingDayInfo) {
      Logger.log('No verification needed for ' + teamLabel + ' - no valid working day to check');
      return;
    }
    const lastWorkingDay = lastWorkingDayInfo.dateString;

    // Verify all employee sheets
    const results = this.verifyAllEmployees();

    // Build summary
    const employeesToRemind = [];
    let summaryMessage = '\uD83D\uDCCA *' + teamLabel + '* - Daily Task Report - ' + lastWorkingDay + '\n';

    for (let i = 0; i < results.length; i++) {
      const data = results[i];
      if (Array.isArray(data) && !data[1]) {
        employeesToRemind.push(data[0]);
        summaryMessage += data[0] + '\u274C Not Added \n';
      } else {
        summaryMessage += data + '\n';
      }
    }

    // Get Backlog ticket counts (per-team, since Backlog config is team-specific)
    let backlogMessage = '';
    if (this.config.backlogApiKey && this.config.backlogUrl) {
      const ticketCounts = this._countBacklogTicketsPerEngineer();
      backlogMessage = this._formatBacklogTicketCounts(ticketCounts);
      Logger.log(backlogMessage);
    }

    // Send employee reminders
    if (employeesToRemind.length > 0) {
      Logger.log('[' + teamLabel + '] Sending reminders to: ' + employeesToRemind.join(', '));
      this.sendEmployeeReminder(employeesToRemind);
    } else {
      Logger.log('[' + teamLabel + '] All timesheets submitted');
    }

    // Send RSS + Backlog to employee alert webhook (RSS is shared, Backlog is team-specific)
    if (this.config.employeeAlertWebhookUrl) {
      let employeePost = '';
      if (backlogMessage) {
        employeePost += backlogMessage + '\n';
      }
      if (linkedinPost) {
        employeePost += linkedinPost;
      }
      if (employeePost) {
        this.sendGoogleChatMessage(employeePost, this.config.employeeAlertWebhookUrl);
      }
    }

    // Send summary to manager webhook
    let finalSummaryMessage = summaryMessage;
    if (backlogMessage) {
      finalSummaryMessage += backlogMessage;
    }
    this.sendGoogleChatMessage(finalSummaryMessage);

    Logger.log('[' + teamLabel + '] Verification complete');
  }
}

// ============================================================================
// ENTRY POINTS & TEAM CONFIG
// ============================================================================

/**
 * Read team configurations from the master "Teams Config" sheet.
 * Reads all columns including verifier-specific ones (Alert Webhook, Chat IDs,
 * Config File ID, Backlog settings, etc.).
 * Only returns teams where "Enabled (Verify)" is "Yes".
 *
 * Requires Script Property: MASTER_SHEET_ID
 *
 * @returns {Array} Array of team config objects ready for SheetsVerifier constructor
 */
function readVerifierTeamConfigs() {
  const props = PropertiesService.getScriptProperties();
  const masterSheetId = props.getProperty('MASTER_SHEET_ID');

  if (!masterSheetId) {
    throw new Error('MASTER_SHEET_ID not configured. Set it in Script Properties.');
  }

  const ss = SpreadsheetApp.openById(masterSheetId);
  const sheet = ss.getSheetByName('Teams Config');

  if (!sheet) {
    throw new Error("Sheet 'Teams Config' not found in master spreadsheet");
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('Master sheet must have header row and at least one data row');
  }

  // Fuzzy column header matching (same pattern as timesheet_creator.gs)
  const headers = data[0].map(function(h) { return h.toString().trim(); });

  function findCol(possibleNames) {
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

  function parseList(value) {
    if (!value) return [];
    const str = value.toString().trim();
    if (!str) return [];
    return str.split(/[,\n]+/).map(function(item) { return item.trim(); }).filter(function(item) { return item.length > 0; });
  }

  function parseColonMap(value) {
    // Parse "Name:Value, Name2:Value2" into {Name: Value, Name2: Value2}
    if (!value) return {};
    const str = value.toString().trim();
    if (!str) return {};
    const result = {};
    var pairs = str.split(/[,\n]+/);
    for (let i = 0; i < pairs.length; i++) {
      const pair = pairs[i].trim();
      const colonIdx = pair.indexOf(':');
      if (colonIdx > 0) {
        const key = pair.substring(0, colonIdx).trim();
        const val = pair.substring(colonIdx + 1).trim();
        if (key) result[key] = val;
      }
    }
    return result;
  }

  const colMap = {
    projectName:            findCol(['Project Name', 'ProjectName', 'Project']),
    teamMembers:            findCol(['Team Members', 'TeamMembers', 'Members', 'Employees']),
    driveFolderId:          findCol(['Drive Folder ID', 'DriveFolderID', 'Folder ID', 'FolderID']),
    webhookUrl:             findCol(['Webhook URL', 'WebhookURL', 'Webhook', 'Chat Webhook']),
    employeeEmails:         findCol(['Employee Emails', 'EmployeeEmails', 'Emails']),
    alertWebhookUrl:        findCol(['Alert Webhook URL', 'AlertWebhookURL', 'Alert Webhook', 'Employee Alert Webhook']),
    employeeChatIds:        findCol(['Employee Chat IDs', 'EmployeeChatIDs', 'Chat IDs']),
    configFileId:           findCol(['Config File ID', 'ConfigFileID', 'Config File']),
    enabledCreate:          findCol(['Enabled (Create)', 'Enabled Create']),
    enabledVerify:          findCol(['Enabled (Verify)', 'Enabled Verify']),
    statusOptions:          findCol(['Status Options', 'StatusOptions']),
    activityOptions:        findCol(['Activity Options', 'ActivityOptions']),
    holidayDates:           findCol(['Holiday Dates', 'HolidayDates', 'Holidays']),
    backlogUrl:             findCol(['Backlog URL', 'BacklogURL']),
    backlogApiKey:          findCol(['Backlog API Key', 'BacklogAPIKey', 'Backlog Key']),
    backlogProjectIds:      findCol(['Backlog Project IDs', 'BacklogProjectIDs']),
    backlogEngineerMapping: findCol(['Backlog Engineer Mapping', 'BacklogEngineerMapping'])
  };

  // Validate required columns
  if (colMap.projectName === -1) {
    throw new Error("Required column 'Project Name' not found in master sheet");
  }
  if (colMap.teamMembers === -1) {
    throw new Error("Required column 'Team Members' not found in master sheet");
  }

  const teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[colMap.projectName]) continue;

    // Check Enabled (Verify) column — skip unless "Yes"
    if (colMap.enabledVerify !== -1) {
      const val = row[colMap.enabledVerify].toString().trim().toLowerCase();
      if (val !== 'yes' && val !== 'true' && val !== '1') {
        Logger.log('Skipping team (Verify disabled): ' + row[colMap.projectName]);
        continue;
      }
    }

    const team = {
      projectName:           row[colMap.projectName].toString().trim(),
      engineerNames:         parseList(row[colMap.teamMembers]),
      driveFolderId:         colMap.driveFolderId !== -1 ? row[colMap.driveFolderId].toString().trim() : '',
      webhookUrl:            colMap.webhookUrl !== -1 ? row[colMap.webhookUrl].toString().trim() : '',
      alertWebhookUrl:       colMap.alertWebhookUrl !== -1 ? row[colMap.alertWebhookUrl].toString().trim() : '',
      employeeChatIds:       colMap.employeeChatIds !== -1 ? parseColonMap(row[colMap.employeeChatIds]) : {},
      configFileId:          colMap.configFileId !== -1 ? row[colMap.configFileId].toString().trim() : '',
      holidays:              colMap.holidayDates !== -1 ? parseList(row[colMap.holidayDates]) : [],
      backlogUrl:            colMap.backlogUrl !== -1 ? row[colMap.backlogUrl].toString().trim() : '',
      backlogApiKey:         colMap.backlogApiKey !== -1 ? row[colMap.backlogApiKey].toString().trim() : '',
      backlogProjectIds:     colMap.backlogProjectIds !== -1 ? row[colMap.backlogProjectIds].toString().trim() : '',
      backlogEngineerMapping: colMap.backlogEngineerMapping !== -1 ? parseColonMap(row[colMap.backlogEngineerMapping]) : {},
    };

    // Validate team has engineers
    if (team.engineerNames.length === 0) {
      Logger.log('WARNING: Team "' + team.projectName + '" has no team members. Skipping.');
      continue;
    }

    teams.push(team);
  }

  Logger.log('Loaded ' + teams.length + ' team configurations for verification from master sheet');
  return teams;
}

/**
 * Daily verification orchestrator.
 * Reads all enabled teams from the master sheet and runs verification for each.
 *
 * Requires Script Property: MASTER_SHEET_ID
 * Optional Script Properties: TEST_MODE, RSS_ARTICLE_COUNT, ROWS_TO_CHECK_AFTER_DATE, SHEET_DATA_RANGE
 *
 * Set this as your time-driven trigger.
 */
function runDailyVerification() {
  const startTime = new Date().getTime();
  const MAX_EXECUTION_MS = 5 * 60 * 1000; // 5 minutes (leave 1 min buffer from Apps Script 6-min limit)

  Logger.log('='.repeat(60));
  Logger.log('MULTI-TEAM VERIFICATION - Starting at ' + new Date().toISOString());
  Logger.log('='.repeat(60));

  // Check if today is a weekend first (global check, done once)
  const today = new Date();
  const currentDay = today.getDay();
  if (currentDay === 0 || currentDay === 6) {
    Logger.log('No verification needed - today is a weekend');
    return;
  }

  // Read team configurations from master sheet
  let teams;
  try {
    teams = readVerifierTeamConfigs();
  } catch (e) {
    Logger.log('FATAL: Could not read master sheet: ' + e);
    return;
  }

  if (teams.length === 0) {
    Logger.log('No teams enabled for verification.');
    return;
  }

  // Fetch RSS feed once (shared across all teams, not team-specific)
  let linkedinPost = null;
  try {
    const rssAggregator = new CategorizedRSSFeedAggregator();
    const category = rssAggregator.getCurrentDayCategory();
    const props = PropertiesService.getScriptProperties();
    const articleCount = parseInt(props.getProperty('RSS_ARTICLE_COUNT') || '5', 10);
    linkedinPost = rssAggregator.generateLinkedInPost(articleCount, category);
  } catch (e) {
    Logger.log('Warning: RSS feed fetch failed: ' + e);
  }

  // Run verification for each team
  const teamResults = [];
  for (let i = 0; i < teams.length; i++) {
    // Check execution time guard
    const elapsed = new Date().getTime() - startTime;
    if (elapsed > MAX_EXECUTION_MS) {
      Logger.log('⚠ Approaching execution time limit (' + Math.round(elapsed / 1000) + 's elapsed). Stopping.');
      Logger.log('Skipped teams: ' + teams.slice(i).map(function(t) { return t.projectName; }).join(', '));
      break;
    }

    const team = teams[i];
    Logger.log('\n' + '-'.repeat(50));
    Logger.log('Verifying team: ' + team.projectName + ' (' + team.engineerNames.length + ' engineers)');
    Logger.log('-'.repeat(50));

    try {
      const verifier = new SheetsVerifier(team);

      // Check if today is a holiday for THIS team specifically
      if (verifier.isHoliday(today)) {
        Logger.log('Skipping ' + team.projectName + ' - today is a team holiday');
        teamResults.push({ team: team.projectName, status: 'skipped (holiday)' });
        continue;
      }

      // Run daily verification for this team, passing pre-fetched RSS post
      verifier.runTeamVerification(linkedinPost);
      teamResults.push({ team: team.projectName, status: 'success' });

    } catch (e) {
      Logger.log('ERROR verifying team ' + team.projectName + ': ' + e);
      teamResults.push({ team: team.projectName, status: 'error: ' + e.message });
    }
  }

  // Summary
  const totalElapsed = Math.round((new Date().getTime() - startTime) / 1000);
  Logger.log('\n' + '='.repeat(60));
  Logger.log('MULTI-TEAM VERIFICATION SUMMARY (' + totalElapsed + 's)');
  Logger.log('='.repeat(60));
  for (let i = 0; i < teamResults.length; i++) {
    const r = teamResults[i];
    const icon = r.status === 'success' ? '✅' : (r.status.indexOf('error') !== -1 ? '❌' : '⏭');
    Logger.log(icon + ' ' + r.team + ': ' + r.status);
  }
  Logger.log('='.repeat(60));
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

// ============================================================================
// CONFIGURATION HELPERS
// ============================================================================

/**
 * Set up the verifier.
 * Only one script property needed: MASTER_SHEET_ID.
 * All team-specific config comes from the master "Teams Config" sheet.
 *
 * @param {string} masterSheetId - The spreadsheet ID of the master config sheet
 */
function setupMultiTeam(masterSheetId) {
  if (!masterSheetId) {
    Logger.log('ERROR: masterSheetId is required');
    Logger.log('Example: setupMultiTeam("1ABCdef.....")');
    return false;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('MASTER_SHEET_ID', masterSheetId);

  Logger.log('✓ MASTER_SHEET_ID set to: ' + masterSheetId);
  Logger.log('');
  Logger.log('Next steps:');
  Logger.log('  1. Run viewMultiTeamConfiguration() to verify your team configs');
  Logger.log('  2. Set your time-driven trigger to: runDailyVerification');
  Logger.log('  3. Optional: Set TEST_MODE to "true" for a dry run first');

  return true;
}

/**
 * View the multi-team configuration as read from the master sheet.
 * Shows all teams, their enabled/disabled status, and their configuration.
 * Run this to verify the master sheet is set up correctly before switching triggers.
 */
function viewMultiTeamConfiguration() {
  const props = PropertiesService.getScriptProperties();
  const masterSheetId = props.getProperty('MASTER_SHEET_ID');

  Logger.log('='.repeat(60));
  Logger.log('MULTI-TEAM CONFIGURATION');
  Logger.log('='.repeat(60));

  if (!masterSheetId) {
    Logger.log('MASTER_SHEET_ID: NOT SET');
    Logger.log('Run setupMultiTeam("your_spreadsheet_id") first.');
    return;
  }

  Logger.log('MASTER_SHEET_ID: ' + masterSheetId);
  Logger.log('TEST_MODE: ' + (props.getProperty('TEST_MODE') || 'false'));
  Logger.log('ROWS_TO_CHECK_AFTER_DATE: ' + (props.getProperty('ROWS_TO_CHECK_AFTER_DATE') || '5'));
  Logger.log('SHEET_DATA_RANGE: ' + (props.getProperty('SHEET_DATA_RANGE') || 'A1:K150'));
  Logger.log('RSS_ARTICLE_COUNT: ' + (props.getProperty('RSS_ARTICLE_COUNT') || '5'));

  try {
    const teams = readVerifierTeamConfigs();

    Logger.log('\n' + teams.length + ' team(s) enabled for verification:\n');

    for (let i = 0; i < teams.length; i++) {
      const team = teams[i];
      Logger.log((i + 1) + '. ' + team.projectName);
      Logger.log('   Engineers: ' + team.engineerNames.join(', '));
      Logger.log('   Drive Folder ID: ' + (team.driveFolderId || 'NOT SET'));
      Logger.log('   Config File ID: ' + (team.configFileId || 'NOT SET'));
      Logger.log('   Manager Webhook: ' + (team.webhookUrl ? 'SET' : 'NOT SET'));
      Logger.log('   Employee Alert Webhook: ' + (team.alertWebhookUrl ? 'SET' : 'NOT SET'));
      Logger.log('   Chat IDs: ' + Object.keys(team.employeeChatIds).length + ' configured');
      Logger.log('   Holidays: ' + team.holidays.length + ' dates');
      Logger.log('   Backlog URL: ' + (team.backlogUrl || 'NOT SET'));
      Logger.log('   Backlog API Key: ' + (team.backlogApiKey ? 'SET' : 'NOT SET'));
      Logger.log('   Backlog Project IDs: ' + (team.backlogProjectIds || 'NOT SET'));
      Logger.log('   Backlog Engineer Mapping: ' + Object.keys(team.backlogEngineerMapping).length + ' mappings');
      Logger.log('');
    }
  } catch (e) {
    Logger.log('ERROR reading master sheet: ' + e);
  }

  Logger.log('='.repeat(60));
}
