function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gestión de Gastos - Paco')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Obtiene todos los valores de una hoja tal cual se muestran en la UI.
 */
function getSheetTable(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;

    // Usamos getDisplayValues para mantener el formato de moneda y fechas de tu Excel
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * Obtiene las opciones de configuración desde la hoja 'listas'
 * para dinamizar el formulario.
 */
function getFormOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('listas');
  const data = sheet.getDataRange().getValues();
  
  // Objeto para clasificar las listas
  const options = {
    tiposGasto: [],
    formasPago: [],
    tags: [],
    categorias: [],
    seDivide: [],
    abrevs: []
  };

  /**
   * AJUSTE TÉCNICO:
   * i = 0 -> Fila vacía
   * i = 1 -> Encabezados ("tipo", "forma", etc.)
   * i = 2 -> Primeros datos reales ("Fijo", "BBVA Oro", "Ingreso"...)
   */
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (row[1])  options.tiposGasto.push(row[1]);  // Col B: Fijo, MCI, MSI...
    if (row[3])  options.formasPago.push(row[3]);  // Col D: BBVA Oro, Klar...
    if (row[9])  options.tags.push(row[9]);        // Col I: Debo a Lulú, MD...
    if (row[12]) options.categorias.push(row[12]); // Col L: Ingreso, Egreso
    if (row[14]) options.seDivide.push(row[14]);   // Col O: Sí, No
    if (row[16]) options.abrevs.push(row[16]);     // Col Q: ENE, FEB, MAR...
  }

  return options;
}

function registrarGasto(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('unificado_v4');
  
  // 1. Cálculo de ID y Fechas
  const lastRow = sheet.getLastRow();
  const lastId = (lastRow > 1) ? sheet.getRange(lastRow, 1).getValue() : 0;
  const nextId = parseInt(lastId) + 1;

  const fechaCargo = new Date(form.fecha_cargo + "T00:00:00");
  const mesesNombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];  
  const mesFull = mesesNombres[fechaCargo.getMonth()];
  const añoNum = fechaCargo.getFullYear();

  // 2. Mapeo exacto de las 16 columnas
  // [0]id, [1]concepto, [2]monto, [3]tipo_gasto, [4]forma_pago, [5]mes, [6]año, 
  // [7]fecha_cargo, [8]fecha_pago, [9]categoría, [10]a_pagos, [11]no_mens, 
  // [12]total_meses, [13]tag, [14]se_divide, [15]gasto_x_mes
  
  const nuevaFila = [
    nextId,                                // id
    form.concepto,                         // concepto
    parseFloat(form.monto),                // monto
    form.tipo_gasto,                       // tipo_gasto
    form.forma_pago,                       // forma_pago
    mesFull,                               // mes
    añoNum,                                // año
    form.fecha_cargo,                      // fecha_cargo
    form.fecha_pago || form.fecha_cargo,   // fecha_pago (fallback a cargo)
    form.categoria,                        // categoría (E/I)
    parseInt(form.no_mens) || 0,           // no_mens
    parseInt(form.total_meses) || 0,       // total_meses
    form.tag,                              // tag (MD/D/NA)
    form.se_divide,                        // se_divide (Sí/No)
    form.gasto_x_mes                       // gasto_x_mes (JUL, AGO...)
  ];

  sheet.appendRow(nuevaFila);
  return "Registro #" + nextId + " guardado correctamente en unificado_v4.";
}

/**
 * Obtiene los datos de una hoja para mostrarlos en la UI.
 */
function getSheetTable(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName) || ss.getSheetByName(sheetName.charAt(0).toUpperCase() + sheetName.slice(1));
    if (!sheet) return "No se encontró la hoja: " + sheetName;
    
    // Obtenemos los valores visibles de la tabla
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return e.toString();
  }
}