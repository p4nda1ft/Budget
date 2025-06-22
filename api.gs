const PREFS = () => ({
  locale: "en-US",
  currency: "USD",
});

// const SEND_EMAIL_ID = "techlever45@gmail.com";

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setFaviconUrl("https://heartstchr.github.io/img/borl.png")
    .setTitle("Budgeting")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function includes(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getLoggedInUser() {
  return Session.getActiveUser().getEmail();
}

function getStringifiedTables(tableNames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var o = {};
  tableNames.forEach((name) => {
    var rng = ss.getRangeByName(name);
    var vals = trimRangeRows(rng).getValues();
    o[name] = vals;
  });
  return o;
}

function addEntry(formDataObject, formName) {
  var Agent;
  initialize();
  if (formDataObject.ID) {
    formDataObject.ID = Number(formDataObject.ID);
  }
  try {
    Agent = Table.define({ sheetName: formName, idColumn: "id" });
    formDataObject.created_date = JSON.stringify(new Date());
    var srfEntry = new Agent(formDataObject);
    return Agent.createOrUpdate(srfEntry);
  } catch (error) {
    console.log(error);
    return false;
  }
}
function batchCreate(row, x, formName) {
  var Agent;
  initialize();
  Agent = Table.define({ sheetName: formName, idColumn: "id" });
  const rows = Array(x).fill({ model: row });
  return Agent.batchCreate(rows);
}

function deleteEntryByID(id, tableName) {
  var Agent;
  Tamotsu.initialize();

  try {
    Agent = Tamotsu.Table.define({ sheetName: tableName, idColumn: "id" });
    Agent.find(id).destroy();
    return true;
  } catch (error) {
    console.log(error);
    return false;
  }
}

function GetRecordById(id, tableName) {
  Tamotsu.initialize();
  var Agent;
  Agent = Tamotsu.Table.define({ sheetName: tableName, idColumn: "id" });
  return Agent.find(id);
}
function UpdateAttributes(sheetName, id, updateObject) {
  var Agent;
  Tamotsu.initialize();
  Agent = Tamotsu.Table.define({ sheetName: sheetName, idColumn: "id" });
  var row = Agent.find(id);
  return row.updateAttributes(updateObject);
}
function sendMail(htmlMessage, subject) {
  MailApp.sendEmail({
    to: SEND_EMAIL_ID,
    subject: subject,
    htmlBody: htmlMessage,
  });
}

function saveFile(fileObjects) {
  fileObjects.forEach((o) => {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(o.data),
      o.mimeType,
      o.fileName
    );
    DriveApp.createFile(blob).getId();
  });
}

function saveFile2(obj) {
  const FOLDER_ID = "1cHyWooe0RLMPR9zhQIuh0K3vgT6r22oz";
  var blob = Utilities.newBlob(
    Utilities.base64Decode(obj.data),
    obj.mimeType,
    obj.fileName
  );
  var folder = DriveApp.getFolderById(FOLDER_ID);
  return folder.createFile(blob).getUrl();
}
function commitFilesToDB(sheetName, id, formDataObject) {
  var Agent;
  Tamotsu.initialize();
  Agent = Tamotsu.Table.define({ sheetName: sheetName, idColumn: "id" });
  var row = Agent.find(id);
  return row.updateAttributes(formDataObject);
}
