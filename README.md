# Budget Gmail Automation

This Apps Script project processes financial emails and stores the extracted
information into monthly sheets. The automation searches Gmail using keywords
in both the subject and body and filters by trusted senders. Attachments are
converted to Google Docs/Sheets when necessary so that their content can be
analysed to extract transaction details.

## Main features

- Fetch emails by configurable date range and query.
- Supported attachment parsing for PDF, images and spreadsheets using the
  Drive advanced service.
- Simple analysis of text to obtain amount, payment method and authorization.
- Results are stored in sheets named `YYYY-MM` using Material colors.
- Menu entries allow updating the date range manually.
- Daily trigger at 6:00AM can be created via `createDailyTrigger()`.

From the **📊 Sistema Financiero** menu you can set the start and end dates
used to search Gmail or reset them at any time.

Run `processEmails` from the spreadsheet menu **📊 Sistema Financiero** or via
the trigger to keep your budget sheets up to date.

