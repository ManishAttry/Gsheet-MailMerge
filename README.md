# Gsheet-MailMerge (Google Sheets + Gmail Draft Template)
A Google Apps Script mail merge tool for Google Sheets and Gmail. Create personalized email drafts or send directly using a Gmail draft template with support for merge fields, inline images, attachments, and scheduling. Includes menu options for processing all rows or selected rows, and records status in your sheet.


A robust Mail Merge script for Google Sheets using Google Apps Script.  
Features:
- Create Gmail drafts from a single draft template (supports inline images & attachments)
- Send emails now or create drafts
- Process all rows or only rows you select
- Schedule selected rows for Today or Tomorrow (single one-time trigger)
- Records DRAFT/SENT status in the sheet

---

## Demo screenshot
![ui-screenshot](https://github.com/user-attachments/assets/aa26be78-cc92-41a0-9bd5-e2bc98a594d5)



---

## Installation

1. Open the Google Sheet you want to use for the mail merge.
2. Click **Extensions → Apps Script**.
3. Replace the contents of the default code with the script in `mailmerge.gs` (or paste the script above).
4. In the script editor set the `RECIPIENT_COL`, `EMAIL_SENT_COL`, `SUBJECT_COL` constants if your column names differ.
5. Save the project (`File → Save`).
6. Back in the spreadsheet, refresh the page. You should see a **Mail Merge** menu added to the top bar.
7. The first time you run functions the script will ask for permission scopes (Gmail & Sheets). Accept to allow the script to create drafts and send email.

---

## Sheet layout (required)
The sheet **must** have a header row (Row 1). Minimum required columns (case sensitive by default in script):

- `EMAIL` — recipient email address. Can contain multiple comma-separated recipients.
- `GENDER` — Recommend to use `Sir` or `Ma'am` only.
- `COMPANY` — You can use any Name.
- `SUBJECT` — optional: per-row subject. If empty, the draft's subject is used.
- `DRAFT DATE` — status column written by the script (shows `Draft: <date>` or `Sent: <date>` or errors).

Any additional columns can be used as merge placeholders in your Gmail draft body using `{{COLUMN_NAME}}`.

**Example header row**:
![GoogleSheet-screenshot](https://github.com/user-attachments/assets/497694d1-0e30-4fe1-bd5f-d3813b9492e3)

