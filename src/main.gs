function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Gestión de Ingresos y Egresos_v1.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getFormOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('listas');
  const data = sheet.getDataRange().getValues();
  const options = { tiposGasto: [], formasPago: [], tags: [], categorias: [], seDivide: [], abrevs: [] };

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (row[1])  options.tiposGasto.push(row[1]);
    if (row[3])  options.formasPago.push(row[3]);
    if (row[9])  options.tags.push(row[9]);
    if (row[12]) options.categorias.push(row[12]);
    if (row[14]) options.seDivide.push(row[14]);
    if (row[16]) options.abrevs.push(row[16]);
  }
  return options;
}


function registrarGasto(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('unificado_v4');
  const lastRow = sheet.getLastRow();
  const lastId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
  const nextId = (parseInt(lastId) || 0) + 1;

  const fechaCargo = new Date(form.fecha_cargo + "T00:00:00");
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  
  const nuevaFila = [
    nextId, form.concepto, parseFloat(form.monto), form.tipo_gasto, form.forma_pago,
    meses[fechaCargo.getMonth()], fechaCargo.getFullYear(), form.fecha_cargo,
    form.fecha_pago || form.fecha_cargo, form.categoria, parseInt(form.no_mens) || 0,
    parseInt(form.total_meses) || 0, form.tag, form.se_divide, form.gasto_x_mes
  ];

  sheet.appendRow(nuevaFila);
  return "Registro #" + nextId + " guardado correctamente.";
}

function getUnificadoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('unificado_v4');
  if (!sheet) return { headers: [], data: [] };
  
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) return { headers: [], data: [] };
  
  const headers = values[0];
  const timezone = Session.getScriptTimeZone();
  
  const data = values.slice(1).map(row => {
    return row.map(val => {
      if (val instanceof Date) {
        // Return YYYY-MM-DD string to avoid timezone parsing issues on the frontend
        return Utilities.formatDate(val, timezone, "yyyy-MM-dd");
      }
      return val;
    });
  });
  
  return { headers, data };
}

function editarGasto(id, form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('unificado_v4');
  if (!sheet) throw new Error("La hoja 'unificado_v4' no existe.");
  
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  
  // Find row matching ID in column 1 (0-indexed 0)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      targetRowIndex = i + 1; // Google Sheets row numbers are 1-based
      break;
    }
  }
  
  if (targetRowIndex === -1) {
    throw new Error("Registro con ID #" + id + " no encontrado.");
  }
  
  const fechaCargo = new Date(form.fecha_cargo + "T00:00:00");
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  
  const filaActualizada = [
    parseInt(id), form.concepto, parseFloat(form.monto), form.tipo_gasto, form.forma_pago,
    meses[fechaCargo.getMonth()], fechaCargo.getFullYear(), form.fecha_cargo,
    form.fecha_pago || form.fecha_cargo, form.categoria, parseInt(form.no_mens) || 0,
    parseInt(form.total_meses) || 0, form.tag, form.se_divide, form.gasto_x_mes
  ];
  
  // Update the entire row in the sheet
  sheet.getRange(targetRowIndex, 1, 1, filaActualizada.length).setValues([filaActualizada]);
  return "Registro #" + id + " guardado correctamente.";
}

function getCalculosData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('cálculos');
  if (!sheet) return null;
  
  // Read current filter cell values
  const t1AnioVal = sheet.getRange("C2").getValue();
  const t1Anio = (typeof t1AnioVal === 'number') ? Math.round(t1AnioVal) : t1AnioVal;
  const t1Mes = sheet.getRange("C3").getValue();
  
  const t2Tarjeta = sheet.getRange("C7").getValue();
  const t2AnioVal = sheet.getRange("E7").getValue();
  const t2Anio = (typeof t2AnioVal === 'number') ? Math.round(t2AnioVal) : t2AnioVal;
  const t2Mes = sheet.getRange("C8").getValue();
  
  const t3AnioVal = sheet.getRange("C12").getValue();
  const t3Anio = (typeof t3AnioVal === 'number') ? Math.round(t3AnioVal) : t3AnioVal;
  const t3Mes = sheet.getRange("C13").getValue();

  // Read headers and data rows
  const t1Headers = sheet.getRange("B4:D4").getValues()[0];
  const t1Values = sheet.getRange("B5:D5").getValues()[0];
  
  const t2Headers = sheet.getRange("B9:I9").getValues()[0];
  const t2Values = sheet.getRange("B10:I10").getValues()[0];
  
  const t3Headers = sheet.getRange("B14:F14").getValues()[0];
  const t3Values = sheet.getRange("B15:F15").getValues()[0];

  return {
    t1: { anio: t1Anio, mes: t1Mes, headers: t1Headers, values: t1Values },
    t2: { tarjeta: t2Tarjeta, anio: t2Anio, mes: t2Mes, headers: t2Headers, values: t2Values },
    t3: { anio: t3Anio, mes: t3Mes, headers: t3Headers, values: t3Values }
  };
}

function updateCalculosTable(tableNum, filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('cálculos');
  if (!sheet) throw new Error("La hoja 'cálculos' no existe.");

  if (tableNum === 1) {
    if (filters.anio !== undefined) sheet.getRange("C2").setValue(Number(filters.anio) || filters.anio);
    if (filters.mes !== undefined) sheet.getRange("C3").setValue(filters.mes);
  } else if (tableNum === 2) {
    if (filters.tarjeta !== undefined) sheet.getRange("C7").setValue(filters.tarjeta);
    if (filters.anio !== undefined) sheet.getRange("E7").setValue(Number(filters.anio) || filters.anio);
    if (filters.mes !== undefined) sheet.getRange("C8").setValue(filters.mes);
  } else if (tableNum === 3) {
    if (filters.anio !== undefined) sheet.getRange("C12").setValue(Number(filters.anio) || filters.anio);
    if (filters.mes !== undefined) sheet.getRange("C13").setValue(filters.mes);
  }

  SpreadsheetApp.flush(); // Recalculate sheet formulas
  
  return getCalculosData();
}