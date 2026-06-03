// ==========================================
// STOCKFLOW - GOOGLE APPS SCRIPT v2.0
// ==========================================
// INSTRUCCIONES:
// 1. Abre tu Google Sheets
// 2. Ve a Extensiones > Apps Script
// 3. Borra todo el código y pega este
// 4. Guarda (Ctrl+S / Cmd+S)
// 5. Implementa > Implementar como aplicación web
// 6. Ejecutar como: Yo
// 7. Quien tiene acceso: Cualquiera
// 8. Copia la URL del script y ponla en index.html variable URL_API
// ==========================================

function doPost(e) {
  var parametros = JSON.parse(e.postData.contents);
  var op = parametros.op;
  var codigo = parametros.codigo;
  var cantidad = parametros.cantidad || 1;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet();
  var hojaMov = getHojaMovimientos(ss);
  
  var ahora = new Date();
  var fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm");
  
  // CONSULTAR PRECIO
  if (op === "consultaprecio") {
    var datos = buscarProducto(hoja, codigo);
    if (!datos.fila) {
      return jsonResponse(false, "Producto no encontrado");
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      producto: datos.data[1],
      precio: datos.data[2],
      stock: datos.data[3]
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // AGREGAR PRODUCTO
  if (op === "agregarproducto") {
    var datos = buscarProducto(hoja, codigo);
    if (datos.fila) {
      return jsonResponse(false, "Producto ya existe");
    }
    var producto = parametros.producto;
    var stock = parseInt(parametros.stock) || 0;
    var precio = parseFloat(parametros.precio) || 0;
    
    hoja.appendRow([codigo, producto, precio, stock, fecha, hora, "ACTIVO"]);
    registrarMovimiento(hojaMov, fecha, hora, codigo, producto, stock, "NUEVO PRODUCTO");
    
    return jsonResponse(true, producto + " agregado correctamente");
  }
  
  // REGISTRAR VENTA
  if (op === "registraventa") {
    var datos = buscarProducto(hoja, codigo);
    if (!datos.fila) {
      return jsonResponse(false, "Producto no encontrado");
    }
    
    var stockActual = parseInt(datos.data[3]) || 0;
    var cantidadVenta = parseInt(cantidad) || 1;
    
    if (stockActual < cantidadVenta) {
      return jsonResponse(false, "Stock insuficiente. Stock actual: " + stockActual);
    }
    
    var nuevoStock = stockActual - cantidadVenta;
    hoja.getRange(datos.fila, 4).setValue(nuevoStock);
    hoja.getRange(datos.fila, 6).setValue(hora);
    
    var producto = datos.data[1];
    var precioUnit = datos.data[2];
    var totalVenta = precioUnit * cantidadVenta;
    
    registrarMovimiento(hojaMov, fecha, hora, codigo, producto, cantidadVenta, "VENTA");
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Vendido x" + cantidadVenta + " - $" + totalVenta.toFixed(2),
      stock: nuevoStock
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // RECARGAR (1 producto)
  if (op === "recarga") {
    var datos = buscarProducto(hoja, codigo);
    if (!datos.fila) {
      return jsonResponse(false, "Producto no encontrado");
    }
    
    var stockActual = parseInt(datos.data[3]) || 0;
    var nuevoStock = stockActual + 1;
    hoja.getRange(datos.fila, 4).setValue(nuevoStock);
    hoja.getRange(datos.fila, 6).setValue(hora);
    
    registrarMovimiento(hojaMov, fecha, hora, codigo, datos.data[1], 1, "RECARGA +1");
    
    return jsonResponse(true, "Stock recargado +1. Nuevo: " + nuevoStock);
  }
  
  // RECARGAR VARIOS PRODUCTOS
  if (op === "recargavarios") {
    var datos = buscarProducto(hoja, codigo);
    if (!datos.fila) {
      return jsonResponse(false, "Producto no encontrado");
    }
    
    var stockActual = parseInt(datos.data[3]) || 0;
    var cant = parseInt(cantidad) || 1;
    var nuevoStock = stockActual + cant;
    hoja.getRange(datos.fila, 4).setValue(nuevoStock);
    hoja.getRange(datos.fila, 6).setValue(hora);
    
    registrarMovimiento(hojaMov, fecha, hora, codigo, datos.data[1], cant, "RECARGA +" + cant);
    
    return jsonResponse(true, "Recargado +" + cant + ". Nuevo stock: " + nuevoStock);
  }
  
  return jsonResponse(false, "Operación no válida");
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "StockFlow API activo",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function buscarProducto(hoja, codigo) {
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(codigo)) {
      return { fila: i + 1, data: datos[i] };
    }
  }
  return { fila: null };
}

function getHojaMovimientos(ss) {
  var hojaMov;
  try {
    hojaMov = ss.getSheetByName("Movimientos");
    if (!hojaMov) {
      hojaMov = ss.insertSheet("Movimientos");
      hojaMov.getRange(1, 1, 1, 6).setValues([["Fecha", "Hora", "Código", "Producto", "Cantidad", "Tipo"]]);
      hojaMov.getRange(1, 1, 1, 6).setFontWeight("bold");
    }
  } catch (e) {
    hojaMov = ss.insertSheet("Movimientos");
    hojaMov.getRange(1, 1, 1, 6).setValues([["Fecha", "Hora", "Código", "Producto", "Cantidad", "Tipo"]]);
    hojaMov.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return hojaMov;
}

function registrarMovimiento(hoja, fecha, hora, codigo, producto, cantidad, tipo) {
  try {
    hoja.appendRow([fecha, hora, codigo, producto, cantidad, tipo]);
  } catch (e) {
    console.log("Error al registrar movimiento: " + e);
  }
}

function jsonResponse(success, message) {
  return ContentService.createTextOutput(JSON.stringify({
    success: success,
    message: message
  })).setMimeType(ContentService.MimeType.JSON);
}