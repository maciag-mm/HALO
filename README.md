# HALO
Internal group communicator for QGIS 3.34 or higher.

========================================
           HALO – USER GUIDE
========================================

OVERVIEW
Halo is a QGIS plugin that displays messages based on an online
spreadsheet (e.g. Google Forms / Google Sheets). Messages are
automatically updated, and users can add new entries via a form.


DATA SOURCE CONFIGURATION
1. Click the settings icon (left side of the Halo panel)
2. Select a data source:
   - Local file (CSV/XLSX)
   - Online spreadsheet link (e.g. Google Forms response sheet)

For Google Forms:
Provide the link to the RESPONSE SPREADSHEET (not the form).

The plugin automatically interprets:
   A -> timestamp (date)
   B -> message content
   C -> signature
   Row 1 -> headers (ignored)


DISPLAYING MESSAGES
Messages appear in the Halo panel.

Each message includes:
   - number (automatic)
   - date (localized, no time)
   - content
   - signature (italic, preceded by ~)

Example:
   [12] 05.03.2026
   Message content ~John Smith


AUTOMATIC REFRESH
- Data refreshes every 60 seconds
- No manual reload required


ADDING MESSAGES

Form setup:
   1. Right-click the rocket button
   2. Paste the Google Forms respondent link
   3. Link is saved automatically

Adding a message:
   - Left-click the rocket button
   - The form opens in your web browser
   - Fill and submit

The message will appear after the next refresh.


MESSAGE NUMBERING
- Message number = row number - 1
- (Row 1 = headers → first message = 1)


SPREADSHEET REQUIREMENTS
- Must be accessible via link (public/shared)
- Required structure:
     Column A -> date
     Column B -> message
     Column C -> signature


NOTES
- Requires internet connection
- No local data storage
- Messages added only via Google Forms


SUMMARY
Halo provides:
- Centralized message source
- Automatic synchronization
- Simple message submission via form
- No server setup required

========================================
