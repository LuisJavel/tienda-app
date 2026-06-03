// ==========================================
// STOCKFLOW - GOOGLE APPS SCRIPT
// ==========================================
// Instrucciones:
// 1. Abre tu Google Sheets
// 2. Ve a Extensiones > Apps Script
// 3. Borra todo el código y pega este
// 4. Guarda (Ctrl+S)
// 5. Implementa > Implementar como aplicación web
// 6. Configura: Ejecutar como: Yo, Quien tiene acceso: Cualquiera
// 7. Copia la URL del script y ponla en tu index.html variable URL_API
// ==========================================

function doPost(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { op, codigo, cantidad } = JSON.parse(e.postData.contents);
  const ahora = new Date();
  const fecha = ahora.toLocaleDateString("es-ES");
  const hora = ahora.toLocaleTimeString("es-ES", {hour: "2-digit", minute: "2-digit"});

  if (op === "agregarproducto") {
    const datos = buscarProducto(codigo);
    if (datos.fila) {
      return jsonResponse(false, "Producto ya existe");
    }
    const nuevaFila = [codigo, e.parameter.producto, parseFloat(e.parameter.precio), parseInt(e.parameter.stock), fecha, hora, "ACTIVO"];
    hoja.appendRow(nuevaFila);
    return jsonResponse(true, "Producto agregado correctamente");
  }

  if (op === "consultaprecio") {
    const datos = buscarProducto(codigo);
    if (!datos.fila) return jsonResponse(false, "Producto no encontrado");
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      producto: datos.data[1],
      precio: datos.data[2],
      stock: datos.data[3]
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (op === "registraventa") {
    const datos = buscarProducto(codigo);
    if (!datos.fila) return jsonResponse(false, "Producto no encontrado");
    const stockActual = datos.data[3];
    const cantidadVenta = cantidad ? parseInt(cantidad) : 1;
    if (stockActual < cantidadVenta) {
      return jsonResponse(false, "Stock insuficiente");
    }
    const nuevoStock = stockActual - cantidadVenta;
    hoja.getRange(datos.fila, 4).setValue(nuevoStock);
    hoja.getRange(datos.fila, 6).setValue(hora);
    registrarMovimiento(codigo, datos.data[1], cantidadVenta, "VENTA", fecha, hora);
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: `Venta realizada (x${cantidadVenta})`,
      stock: nuevoStock
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (op === "recarga") {
    const datos = buscarProducto(codigo);
    if (!datos.fila) return jsonResponse(false, "Producto no encontrado");
    const stockActual = datos.data[3];
    const nuevoStock = stockActual + 1;
    hoja.getRange(datos.fila, 4).setValue(nuevoStock);
    hoja.getRange(datos.fila, 6).setValue(hora);
    registrarMovimiento(codigo, datos.data[1], 1, "RECARGA", fecha, hora);
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Stock recargado +1",
      stock: nuevoStock
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return jsonResponse(false, "Operación no válida");
}

function doGet() {
  return jsonResponse(true, "StockFlow API activo");
}

function buscarProducto(codigo) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == codigo) {
      return { fila: i + 1, data: datos[i] };
    }
  }
  return { fila: null };
}

function registrarMovimiento(codigo, producto, cantidad, tipo, fecha, hora) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaMov;
  try {
    hojaMov = ss.getSheetByName("Movimientos");
    if (!hojaMov) {
      hojaMov = ss.insertSheet("Movimientos");
      hojaMov.appendRow(["Fecha", "Hora", "Código", "Producto", "Cantidad", "Tipo"]);
    }
  } catch(e) {
    hojaMov = ss.insertSheet("Movimientos");
    hojaMov.appendRow(["Fecha", "Hora", "Código", "Producto", "Cantidad", "Tipo"]);
  }
  hojaMov.appendRow([fecha, hora, codigo, producto, cantidad, tipo]);
}

function jsonResponse(success, message) {
  return ContentService.createTextOutput(JSON.stringify({ success, message }))
    .setMimeType(ContentService.MimeType.JSON);
}