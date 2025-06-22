function SetUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("template");

  var schema = ["SavingCategory", "Savings"];
  //  _DeleteSheets(schema)

  //  Check If All the sheets in array exist
  schema.forEach((name) => {
    if (ss.getSheetByName(name) == null) {
      ss.insertSheet(name, { template: templateSheet });
    } else {
      return ss.getSheetByName(name);
    }
  });

  // Set Named Range
  schema.forEach((name) => {
    if (ss.getRangeByName(name) == null) {
      ss.setNamedRange(
        "tbl_" + name,
        ss.getSheetByName(name).getRange("A1:D1000")
      );
    }
  });
}

function createNamedRange(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("template");
  const shtName = name.replace("tbl_", "");
  if (ss.getSheetByName(shtName) == null) {
    ss.insertSheet(shtName, { template: templateSheet });
  } else {
    return ss.getSheetByName(shtName);
  }
  if (ss.getRangeByName(name) == null) {
    ss.setNamedRange(name, ss.getSheetByName(shtName).getRange("A1:D1000"));
  }
}

function deleteAllNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getNamedRanges().forEach((range) => {
    range.remove();
  });
}

function _GetSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("template");
  if (ss.getSheetByName(name) == null) {
    ss.insertSheet(name, { template: templateSheet });
  } else {
    return ss.getSheetByName(name);
  }
}

function _GetNamedRange(name) {
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
  if (range != null) {
    Logger.log(range.getNumColumns());
  }
}

function _DeleteSheets(arrayOfSheetNames) {
  var ss = SpreadsheetApp.getActive();
  arrayOfSheetNames.forEach((name) => {
    var sheet = ss.getSheetByName(name);
    if (sheet != null) {
      ss.deleteSheet(sheet);
    }
  });
}
