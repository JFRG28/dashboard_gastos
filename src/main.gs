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