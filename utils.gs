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



/**
 * Trae correos.
 * @param {Range} rangeName: Range Name
 * @returns {Object}
 */
function processEmails() {
  try {
    const startTime = new Date();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("E-mail") || ss.insertSheet("E-mail");
    
    // Crear encabezados si no existen
    const headers = ["Fecha", "Asunto", "Remitente", "Email", "Contenido", "Adjuntos", "Monto", "Método Pago", "Autorización", "Categoría", "Log IA", "Estado", "ID Hilo"];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    
    // Palabras clave para filtrar
    const keywords = ["transacción", "compra", "factura", "pse", "Alertas y Notificaciones"];
    const query = keywords.map(k => `subject:${k}`).join(' OR ');
    
    // Buscar correos
    const threads = GmailApp.search(query, 0, 200); // Procesa hasta 200 hilos
    let newEmailsCount = 0;
    
    threads.forEach(thread => {
      const threadId = thread.getId();
      const messages = thread.getMessages();
      
      messages.forEach(message => {
        const existing = findRowByThreadId(sheet, threadId);
        if (existing) return; // Saltar si ya existe
        
        // Extraer información del correo
        const date = message.getDate();
        const subject = message.getSubject();
        const from = message.getFrom();
        const email = extractEmail(from);
        const body = message.getPlainBody().substring(0, 1000); // Limitar a 1000 caracteres
        const attachments = message.getAttachments().map(a => a.getName()).join(', ');
        
        // Crear nueva fila
        const newRow = [
          date,
          subject,
          from.replace(/<[^>]+>/g, '').trim(), // Nombre sin email
          email,
          body,
          attachments,
          '', // Monto (vacío)
          '', // Método Pago (vacío)
          '', // Autorización (vacío)
          '', // Categoría (vacío)
          '', // Log IA (vacío)
          'Pendiente', // Estado
          threadId
        ];
        
        sheet.appendRow(newRow);
        newEmailsCount++;
      });
    });
    
    // Ordenar por fecha (más recientes primero)
    sortSheetByDate(sheet);
    
    const processTime = ((new Date() - startTime) / 1000).toFixed(1);
    const msg = `✅ Proceso completado!\n\n📬 Correos nuevos: ${newEmailsCount}\n⏱ Tiempo: ${processTime} segundos`;
    
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.error(e);
    SpreadsheetApp.getUi().alert(`❌ Error: ${e.message}`);
  }
}

// Función auxiliar para extraer email
function extractEmail(str) {
  const emailRegex = /<([^>]+)>/;
  const match = emailRegex.exec(str);
  return match ? match[1] : str;
}

// Buscar si el hilo ya existe en la hoja
function findRowByThreadId(sheet, threadId) {
  const data = sheet.getDataRange().getValues();
  const threadCol = 12; // Columna del ID Hilo (columna M)
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][threadCol] === threadId) return i + 1;
  }
  return null;
}

// Ordenar hoja por fecha (columna A)
function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({column: 1, ascending: false});
}

// Obtener fechas configuradas
    const props = PropertiesService.getScriptProperties();
    const startDateProp = props.getProperty('START_DATE');
    const endDateProp = props.getProperty('END_DATE');
    
    const startDate = startDateProp ? new Date(startDateProp) : null;
    const endDate = endDateProp ? new Date(endDateProp) : null;
    
    // Ajustar fecha final para incluir todo el día
    if (endDate) endDate.setHours(23, 59, 59, 999);
    
    // Construir query con filtro de fecha
    let dateQuery = '';
    if (startDate) dateQuery += ` after:${Math.floor(startDate.getTime()/1000)}`;
    if (endDate) dateQuery += ` before:${Math.floor(endDate.getTime()/1000)}`;
    
    const keywords = ["transacción", "compra", "factura", "pse", "Alertas y Notificaciones"];
    const query = `${keywords.map(k => `subject:${k}`).join(' OR ')}${dateQuery}`;
