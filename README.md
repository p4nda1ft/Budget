# Budget Gmail Automation

This Apps Script project processes financial emails and stores the extracted
information into monthly sheets. The automation searches Gmail using keywords
and trusted senders, converts attachments to text and analyses the content to
find transaction details.

## Main features

- Fetch emails by configurable date range and query.
- Supported attachment parsing for PDF, images and spreadsheets.
- Simple analysis of text to obtain amount, payment method and authorization.
- Results are stored in sheets named `YYYY-MM` using Material colors.
- Menu entries allow updating the date range manually.
- Daily trigger at 6:00AM can be created via `createDailyTrigger()`.

Run `processEmails` from the spreadsheet menu **📊 Sistema Financiero** or via
the trigger to keep your budget sheets up to date.
