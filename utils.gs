var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
function getDateBefore(days) {
  return new Date(now.getTime() - days * MILLIS_PER_DAY);
}

function dateDiff(dt_f, dt_i) {
  var diff = new Date(dt_f).getTime() - new Date(dt_i).getTime();
  return diff / MILLIS_PER_DAY;
}
function humanize(str) {
  var i,
    frags = str.split("_");
  for (i = 0; i < frags.length; i++) {
    frags[i] = frags[i].charAt(0).toUpperCase() + frags[i].slice(1);
  }
  return frags.join(" ");
}

function JSONToArray(jsonArray) {
  var header = [];
  for (const [k, v] of Object.entries(jsonArray[0])) {
    header.push(k);
  }
  return jsonArray.map((r) => {
    const a = [];
    for (let i = 0; i < header.length; ++i) {
      a.push(r[header[i]]);
    }
    return a;
  });
}

function ArrayToJSON(array2D) {
  const header = array2D[0];
  const body = array2D.slice(1);
  const arr2 = body.map((el) => {
    let obj = {};
    for (let i = 0; i < el.length; ++i) {
      obj[header[i]] = el[i];
    }
    return obj;
  });
  return arr2;
}
/**
 * Function to trim the rows of a range. The range should contain a header in the first row.
 * @param {Range} range: a range object from Google spreadsheet. First row of range must be the headers.
 * @returns {Range}
 */
function trimRangeRows(range) {
  var values = range.getValues();
  for (var rowIndex = values.length - 1; rowIndex >= 0; rowIndex--) {
    if (values[rowIndex].join("") !== "") {
      break;
    }
  }
  return range.offset(
    (rowOffset = 0),
    (columnOffset = 0),
    (numRows = rowIndex + 1)
  );
}
/**
 * Function to get JSON from named range.
 * @param {Array} optionSources: Array of range name
 * @returns {Object}
 */
function getDropdowns(optionSources) {
  let obj = {};
  try {
    optionSources.forEach((source) => {
      obj[source] = flattenModelInJSON(NamedRangeToJSON(source));
    });
    return obj;
  } catch (error) {}
}

/**
 * Function to get Indexed dropdowns  from named range.
 * @param {Array} optionSources: Array of range name
 * @returns {Object}
 */
function getIndexedDropdowns(optionSources) {
  let obj = {};
  try {
    optionSources.forEach((source) => {
      let o = {};
      const jsons = flattenModelInJSON(NamedRangeToJSON(source));
      jsons.map((row) => {
        o[row["id"]] = row["name"];
      });
      obj[source] = o;
    });
    return obj;
  } catch (error) {}
}

/**
 * Function to get JSON from named range.
 * @param {Array} JSONArray: Array of JSON object with model field
 * @returns {Array}
 */
function flattenModelInJSON(JSONArray) {
  try {
    let arr = JSONArray.map((el) => {
      // Logger.log(el)
      let o = JSON.parse(el.model);
      o.id = el.id;
      Logger.log(o);
      return o;
    });
    return arr;
  } catch (error) {}
}

/**
 * Function to get JSON from named range.
 * @param {Range} rangeName: Range Name
 * @returns {Object}
 */

function NamedRangeToJSON(rangeName) {
  var rng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
  var vals = trimRangeRows(rng).getValues();
  return ArrayToJSON(vals);
}

function indexedTable(rangeName, index_col_name, value_col_name) {
  var rng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
  var vals = trimRangeRows(rng).getValues();
  const jsons = ArrayToJSON(vals);
  let o = {};
  jsons.map((row) => {
    o[row[index_col_name]] = row[value_col_name];
  });
  return o;
}

