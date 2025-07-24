# Budget Gmail Automation

This Apps Script project processes financial emails and stores the extracted
information into monthly sheets. The automation searches Gmail using keywords
found in the subject or body and filters by trusted senders. Attachments are
converted to Google Docs/Sheets when necessary so that their content can be
analysed to extract transaction details.

## Main features

- Fetch emails by configurable date range and query using keywords such as
  "pago", "crédito", "tarjeta", "transacción", "valor", "factura", "recibo",
  "compra" and "cargo" from trusted senders.
- Supported attachment parsing for PDF, images, spreadsheets and plain text
  using the Drive advanced service. Compressed archives are logged so the
  user can review them manually.
- Simple analysis of text to obtain amount, payment method and authorization.
- Results are stored in sheets named `YYYY-MM` using Material colors.
- Menu entries allow updating the date range manually.
- Daily trigger at 6:00AM can be created via `createDailyTrigger()`.

From the **📊 Sistema Financiero** menu you can set the start and end dates
used to search Gmail or reset them at any time.

Run `processEmails` from the spreadsheet menu **📊 Sistema Financiero** or via
the trigger to keep your budget sheets up to date.

