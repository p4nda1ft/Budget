const GmailAutomation = {
  keywords: [
    "pago",
    "crédito",
    "tarjeta",
    "transacción",
    "valor",
    "factura",
    "recibo",
    "compra",
    "cargo",
  ],
  trustedSenders: [
    "banco@ejemplo.com",
    "pagos@ejemplo.com",
    "notificaciones@paypal.com",
  ],
  monthColors: {
    "01": "#bbdefb",
    "02": "#f8bbd0",
    "03": "#c8e6c9",
    "04": "#d1c4e9",
    "05": "#ffe0b2",
    "06": "#ffccbc",
    "07": "#cfd8dc",
    "08": "#b2dfdb",
    "09": "#dcedc8",
    "10": "#ffcdd2",
    "11": "#e1bee7",
    "12": "#b3e5fc",
  },

  processEmails() {
    const query = this.buildQuery();
    const threads = GmailApp.search(query, 0, 50);
    threads.forEach((thread) => {
      thread.getMessages().forEach((msg) => this.handleMessage(msg));
    });
  },

  buildQuery() {
    const props = PropertiesService.getScriptProperties();
    const start = props.getProperty("START_DATE");
    const end = props.getProperty("END_DATE");
    const parts = [];
    if (this.keywords.length) {
      const kw = this.keywords.map((k) => `"${k}"`).join(" OR ");
      parts.push(`(${kw})`);
    }
    if (this.trustedSenders.length) {
      const from = this.trustedSenders.join(" OR ");
      parts.push(`from:(${from})`);
    }
    if (start) {
      const startSec = Math.floor(new Date(start).getTime() / 1000);
      parts.push(`after:${startSec}`);
    }
    if (end) {
      const d = new Date(end);
      d.setHours(23, 59, 59, 999);
      const endSec = Math.floor(d.getTime() / 1000);
      parts.push(`before:${endSec}`);
    }
    return parts.join(" ");
  },

  handleMessage(msg) {
    const sheet = this.getSheetForDate(msg.getDate());
    const idCol = 10; // ID column index (1-based)
    const id = msg.getId();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol - 1] === id) return;
    }

    const body = msg.getPlainBody();
    const attachText = this.extractAttachmentsText(msg);
    const ai = this.analyzeText(body + "\n" + attachText);

    const row = [
      msg.getDate(),
      msg.getSubject(),
      msg.getFrom(),
      body,
      attachText,
      ai.amount || "",
      ai.paymentMethod || "",
      ai.authorizationNumber || "",
      "",
      id,
    ];
    sheet.appendRow(row);
    this.colorLastRow(sheet, msg.getDate());
  },

  getSheetForDate(date) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const year = date.getFullYear();
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const name = `${year}-${month}`;
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      this.setupSheet(sheet, month);
    }
    return sheet;
  },

  setupSheet(sheet, month) {
    const headers = [
      "Fecha",
      "Asunto",
      "Remitente",
      "Cuerpo",
      "Texto Adjuntos",
      "Monto",
      "Metodo Pago",
      "Numero Autorizacion",
      "Consola",
      "ID",
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    const color = this.monthColors[month];
    if (color) sheet.getRange("A1:J1").setBackground(color);
    sheet.setColumnWidths(1, 10, 150);
    sheet.setRowHeights(1, 1, 21);
    sheet.getRange("A1:J1").setFontFamily("Roboto").setFontSize(10);
  },

  colorLastRow(sheet, date) {
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const color = this.monthColors[month];
    const lr = sheet.getLastRow();
    sheet.getRange(lr, 1, 1, sheet.getLastColumn()).setBackground(color);
  },

  extractAttachmentsText(msg) {
    const attachments = msg.getAttachments();
    let texts = [];
    attachments.forEach((att) => {
      try {
        const mime = att.getContentType();
        if (mime === MimeType.PDF) {
          texts.push(this.readPdf(att));
        } else if (mime.indexOf("spreadsheet") !== -1 || mime.indexOf("excel") !== -1) {
          texts.push(this.readSpreadsheet(att));
        } else if (mime.startsWith("image/")) {
          texts.push(this.readImage(att));
        } else {
          texts.push(att.getName());
        }
      } catch (e) {
        texts.push("[Adjunto no procesado]");
        Logger.log(e);
      }
    });
    return texts.join("\n");
  },

  readPdf(blob) {
    const file = DriveApp.createFile(blob);
    const converted = Drive.Files.copy({}, file.getId(), {
      mimeType: MimeType.GOOGLE_DOCS,
    });
    file.setTrashed(true);
    const doc = DocumentApp.openById(converted.id);
    const text = doc.getBody().getText();
    DriveApp.getFileById(converted.id).setTrashed(true);
    return text;
  },

  readImage(blob) {
    const res = Drive.Files.insert(
      { mimeType: MimeType.GOOGLE_DOCS, title: blob.getName() },
      blob,
      { ocr: true }
    );
    const doc = DocumentApp.openById(res.id);
    const text = doc.getBody().getText();
    DriveApp.getFileById(res.id).setTrashed(true);
    return text;
  },

  readSpreadsheet(blob) {
    const file = DriveApp.createFile(blob);
    const converted = Drive.Files.copy({}, file.getId(), {
      mimeType: MimeType.GOOGLE_SHEETS,
    });
    file.setTrashed(true);
    const ss = SpreadsheetApp.openById(converted.id);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getDisplayValues();
    DriveApp.getFileById(converted.id).setTrashed(true);
    return data.map((r) => r.join("\t")).join("\n");
  },

  analyzeText(text) {
    const amount = text.match(/\$\s?([0-9.,]+)/i);
    const method = text.match(/tarjeta|pse|paypal|crédito|debito/i);
    const auth = text.match(/(?:CUS|ID|Autorizaci[oó]n)\s*[:#]?\s*([A-Z0-9-]+)/i);
    return {
      amount: amount ? amount[0] : "",
      paymentMethod: method ? method[0] : "",
      authorizationNumber: auth ? auth[1] : "",
    };
  },
};

function processEmails() {
  GmailAutomation.processEmails();
}

function setStartDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Fecha inicio (AAAA-MM-DD)");
  if (resp.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty(
      "START_DATE",
      resp.getResponseText()
    );
    ui.alert("Fecha inicio guardada");
  }
}

function setEndDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Fecha fin (AAAA-MM-DD)");
  if (resp.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty(
      "END_DATE",
      resp.getResponseText()
    );
    ui.alert("Fecha fin guardada");
  }
}

function resetDates() {
  PropertiesService.getScriptProperties().deleteProperty("START_DATE");
  PropertiesService.getScriptProperties().deleteProperty("END_DATE");
  SpreadsheetApp.getUi().alert("Fechas reiniciadas");
}

function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some((t) => t.getHandlerFunction() === "processEmails");
  if (!exists) {
    ScriptApp.newTrigger("processEmails").timeBased().atHour(6).everyDays(1).create();
  }
}
