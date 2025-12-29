# Google Apps Script Timesheet Verifier

This is the Google Apps Script version of the Python `sheets_verifier.py` script. It verifies employee timesheet entries and sends daily reports to Google Chat.

## 🚀 Quick Start

### 1. Create Your Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click **New project**
3. Delete the default `function myFunction() {}` code
4. Copy ALL the code from `sheets_verifier.gs` and paste it
5. Click the save icon (💾) or Ctrl+S
6. Name your project: "Timesheet Verifier"

### 2. Configure Basic Settings

Run the `setupConfiguration()` function once to initialize basic configuration:

1. In the Apps Script editor, select `setupConfiguration` from the function dropdown
2. Click **Run** (▶️ button)
3. Grant permissions when prompted
4. Go to **Project Settings** (⚙️ icon) → **Script Properties**
5. Update these required values:
   - **MAIN_SPREADSHEET_ID**: Your Google Sheets ID
   - **GOOGLE_CHAT_WEBHOOK_URL**: Webhook URL for manager reports
   - **EMPLOYEE_ALERT_WEBHOOK_URL**: Webhook URL for employee reminders

### 3. Configure Engineers (Simplified! ✨)

**No more JSON editing!** Choose the format that works best for you:

#### Option A: Ultra-Simple String Format (Easiest! 🎯)

Just one string - perfect for copy/paste from a spreadsheet or text file:

```javascript
setupEngineersFromString(
  "John Doe|123456789012345678901,Jane Smith|987654321098765432109,Bob Johnson"
);
```

**Format**: `name1|chat_id1,name2|chat_id2,name3` (chat ID is optional - omit the `|chat_id` part if not needed)

**Why this is great:**

- ✅ Single string - easiest to copy/paste
- ✅ Can export from Excel/Sheets directly
- ✅ No quotes, brackets, or JSON syntax
- ✅ Chat ID optional (just omit `|chat_id`)

#### Option B: Quick Setup with Array (Most Flexible)

Configure all engineers at once with their chat IDs:

```javascript
quickSetupEngineers([
  { name: "John Doe", chatId: "123456789012345678901" },
  { name: "Jane Smith", chatId: "987654321098765432109" },
  { name: "Bob Johnson" }, // Chat ID optional
]);
```

**Why this is great:**

- ✅ Very readable
- ✅ Easy to add/remove engineers
- ✅ Type-safe (JavaScript validates it)

#### Option C: Add Engineers One by One

```javascript
addEngineer("John Doe", "123456789012345678901");
addEngineer("Jane Smith", "987654321098765432109");
addEngineer("Bob Johnson"); // Without chat ID
```

**Why this is great:**

- ✅ Good for adding engineers incrementally
- ✅ Can add chat IDs later

#### Option D: Set Names First, Add Chat IDs Later

```javascript
setEngineerNames(["John Doe", "Jane Smith", "Bob Johnson"]);
// Add chat IDs later if needed:
setEngineerChatId("John Doe", "123456789012345678901");
```

#### Getting Google Chat User IDs

To enable @mentions in Google Chat:

1. In Google Chat, right-click on a user's name
2. Copy their user ID (it's a long number like `115773514265520053097`)
3. Use it in the functions above

#### View Your Configuration

Check what you've configured:

```javascript
viewConfiguration(); // Shows all engineers and settings
```

#### Remove an Engineer

```javascript
removeEngineer("John Doe");
```

### 4. Additional Configuration (Optional)

These can be set in Script Properties or left as defaults:

- **TEST_MODE**: Set to "true" to prevent sending messages (default: "false")
- **HOLIDAYS**: JSON array of holiday dates in YYYY-MM-DD format
- **ROWS_TO_CHECK_AFTER_DATE**: Number of rows to check after finding a date (default: 5)
- **SHEET_DATA_RANGE**: Range to read from sheets (default: "A1:K150")

### 5. Run the Verification

1. Select `runDailyVerification` from the function dropdown
2. Click **Run** (▶️ button)
3. Check the **Execution log** (View → Logs) to see the results

## 🤖 Automation

### Set Up Daily Trigger

To run the verification automatically every day:

1. In Apps Script editor, click the clock icon (⏰ **Triggers**)
2. Click **Add Trigger**
3. Configure:
   - **Function**: `runDailyVerification`
   - **Event source**: **Time-driven**
   - **Type**: **Day timer**
   - **Time of day**: Choose when to run (e.g., 10:00 AM)
4. Click **Save**

The script will automatically:

- Check the last working day's entries
- Generate summaries for each employee
- Send reminders to employees who haven't submitted timesheets
- Post a daily summary to Google Chat
- Generate and post LinkedIn content from RSS feeds

## ⚙️ Simplified Configuration Functions

The script includes easy-to-use functions for managing engineers. No more manual JSON editing!

### Available Functions

| Function                     | Purpose                   | Example                                                |
| ---------------------------- | ------------------------- | ------------------------------------------------------ | -------- | ------ |
| `setupEngineersFromString()` | **Simplest!** One string  | `setupEngineersFromString("John                        | 123,Jane | 456")` |
| `quickSetupEngineers()`      | Set all engineers at once | `quickSetupEngineers([{name: 'John', chatId: '123'}])` |
| `addEngineer()`              | Add a single engineer     | `addEngineer('John Doe', '123456789')`                 |
| `setEngineerNames()`         | Set engineer names only   | `setEngineerNames(['John', 'Jane'])`                   |
| `setEngineerChatId()`        | Add/update chat ID        | `setEngineerChatId('John', '123456789')`               |
| `removeEngineer()`           | Remove an engineer        | `removeEngineer('John Doe')`                           |
| `viewConfiguration()`        | View current settings     | `viewConfiguration()`                                  |

### Example Workflows

#### Ultra-Simple (Recommended for Quick Setup)

```javascript
// Just one line - copy from spreadsheet or text file!
setupEngineersFromString(
  "Jinu T J|115142352767738659510,Bismillakhan S|115773514265520053097,Midhun|108660074852905944786"
);

// View what you configured
viewConfiguration();
```

#### Flexible Setup (Recommended for Ongoing Management)

```javascript
// 1. Quick setup all engineers
quickSetupEngineers([
  { name: "Jinu T J", chatId: "115142352767738659510" },
  { name: "Bismillakhan S", chatId: "115773514265520053097" },
  { name: "Midhun", chatId: "108660074852905944786" },
]);

// 2. View what you configured
viewConfiguration();

// 3. Add a new engineer later
addEngineer("New Employee", "123456789012345678901");

// 4. Update chat ID for existing engineer
setEngineerChatId("Jinu T J", "NEW_CHAT_ID_HERE");

// 5. Remove an engineer
removeEngineer("Old Employee");
```

All functions provide helpful log messages so you can see what changed.

## 📋 Features

### Timesheet Verification

- Automatically detects the last working day (excluding weekends and holidays)
- Parses timesheet entries for each engineer
- Groups tasks by activity type
- Calculates total hours worked
- Formats output with tree-like structure

### Employee Reminders

- Identifies employees who haven't submitted timesheets
- Sends personalized reminders with @mentions
- Uses Google Chat webhooks for notifications

### RSS Feed Integration

- Fetches articles from categorized RSS feeds based on day of week
- Monday: Infrastructure & DevOps
- Tuesday: Python & Backend
- Wednesday: React & Frontend
- Thursday: Testing & QA
- Friday: Fun & Interesting Tech
- Generates LinkedIn-style posts
- Tracks sent articles to avoid duplicates

### Automated Analysis

- Generates formatted summaries from timesheet data
- Groups tasks by activity type
- Calculates total hours worked
- Creates tree-structured output for easy reading

## 🔧 Configuration Details

### Holiday Configuration

Set holidays in the `HOLIDAYS` property as a JSON array:

```json
["2025-04-10", "2025-05-01", "2025-08-15", "2025-12-25"]
```

Or update via Script Properties → **HOLIDAYS**

### Engineer Names

**Important**: Engineer names must match the sheet names in your main spreadsheet exactly (case-sensitive).

**Tip**: Use `viewConfiguration()` to verify your engineer names match your sheet names.

### Sheet Structure

The script expects sheets with the following structure:

- Column B: Date
- Column C: Module/Area
- Column D: Task details
- Column E: Status
- Column F: Activity Type
- Column G: Start time
- Column H: End time
- Column I: Task duration (calculated)
- Column J: Total duration (merged cell)
- Column K: Remarks

## 🛠 Troubleshooting

### "Spreadsheet not found"

- Verify `MAIN_SPREADSHEET_ID` is correct
- Ensure the script has access to the spreadsheet
- Check that the spreadsheet is not in a restricted folder

### "Sheet not found for engineer"

- Verify engineer names match sheet names exactly
- Check for typos or extra spaces
- Ensure sheets exist in the main spreadsheet

### "No entries found"

- Check if the last working day calculation is correct
- Verify the date format in sheets matches expected format (e.g., "15-Jul")
- Check that `ROWS_TO_CHECK_AFTER_DATE` is sufficient

### "Error sending Google Chat message"

- Verify webhook URLs are correct
- Check that webhooks are still active
- Ensure `TEST_MODE` is not set to "true"

### "Engineer not found"

- Run `viewConfiguration()` to see current engineers
- Verify engineer names match sheet names exactly (case-sensitive)
- Use `addEngineer()` to add missing engineers

## 📝 Differences from Python Version

1. **Authentication**: Uses Apps Script's built-in OAuth instead of service account
2. **File I/O**: Uses PropertiesService instead of JSON files
3. **HTTP Requests**: Uses UrlFetchApp instead of requests library
4. **RSS Parsing**: Simplified RSS parser (basic XML parsing)
5. **Date Handling**: Uses JavaScript Date objects and Utilities.formatDate
6. **No AI Analysis**: Removed OpenRouter/AI integration - uses rule-based analysis instead
7. **Simplified Configuration**: Easy-to-use functions instead of manual JSON editing

## 🔐 Security Notes

- Never commit Script Properties with sensitive data
- Store API keys and webhook URLs in Script Properties, not in code
- Use `TEST_MODE` during development to prevent accidental notifications
- Review permissions before deploying

## 📚 Additional Resources

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Google Sheets API](https://developers.google.com/sheets/api)
- [Google Chat API](https://developers.google.com/chat)

## 🆘 Support

If you encounter issues:

1. Check the Execution log for error messages
2. Verify all Script Properties are set correctly
3. Test with `TEST_MODE` enabled first
4. Check that all required permissions are granted
