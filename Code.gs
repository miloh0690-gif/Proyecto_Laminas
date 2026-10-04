/**
 * ============================================================
 *  BACKEND - Control de Inventario y Ventas de Láminas Acrílicas (v3.1)
 * ============================================================
 *  PEGADO EN UN SOLO ARCHIVO: borra TODO el contenido del editor y pega esto.
 *
 *  Antes de desplegar, CORRE  diagnostico()  desde la lista de funciones de
 *  arriba. Debe decir "TODO OK". También puedes verificarlo por URL:
 *    <URL_DEL_WEB_APP>?action=diagnostico
 *  (sin desplegar todavía sirve, porque doGet lee el código guardado)
 *
 *  Novedades v3:
 *   - ANULAR VENTAS: la venta NO se borra, se marca como "Anulada" con
 *     fecha y motivo. El stock vuelve al inventario y el reporte del mes
 *     descuenta el monto (total neto = bruto - anulado).
 *   - REACTIVAR VENTA: devuelve una venta anulada a "Activa" (descuenta stock).
 *   - 3 columnas nuevas en la hoja "Ventas": Estado, FechaAnulacion,
 *     MotivoAnulacion. Se crean solas, sin borrar datos existentes.
 *   - editarVenta y eliminarVenta ya no rompen el stock en ventas anuladas.
 *   - getReporte ahora devuelve el neto (bruto, anulado y neto).
 *
 *  Novedades v3.2:
 *   - Botones Editar / Anular / ELIMINAR: no aparecían porque la hoja "Ventas"
 *     no tenía la columna ID, y la app solo muestra esas acciones en las ventas
 *     que traen ID. Ahora la columna ID se crea sola (al final) y se le asigna un
 *     ID a cada venta que esté sin él, sin borrar ni tocar las ventas viejas.
 *   - buscarVenta_ ya no asume que el ID está en la columna 1: lo busca por
 *     cabecera, así que da igual en qué posición quedó.
 *   - getVentas también crea las columnas que falten, para que al abrir la app
 *     las ventas ya lleguen con ID y los botones estén visibles.
 *
 *  Zona horaria: en Apps Script > Configuración del proyecto, verifica que
 *  sea "America/La_Paz" para que los filtros por fecha coincidan con tu día.
 *
 *  Hoja "Inventario":
 *   ID | Producto | Espesor | Medida | PrecioConFactura | PrecioSinFactura | Stock | StockMinimo
 *  Hoja "Ventas":
 *   ID | FechaHora | ProductoID | Producto | Cantidad | TipoVenta | PrecioUnitario | Total | Notas
 *       | Estado | FechaAnulacion | MotivoAnulacion
 * ============================================================
 */

const SHEET_INVENTARIO = 'Inventario';
const SHEET_VENTAS = 'Ventas';

const ESTADO_ACTIVA = 'Activa';
const ESTADO_ANULADA = 'Anulada';
const COLS_ANULACION = ['Estado', 'FechaAnulacion', 'MotivoAnulacion'];

/** Crea las hojas SOLO si no existen o están vacías (no borra datos reales). */
function configurarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inv = ss.getSheetByName(SHEET_INVENTARIO) || ss.insertSheet(SHEET_INVENTARIO);
  if (inv.getLastRow() === 0) {
    inv.appendRow(['ID', 'Producto', 'Espesor', 'Medida', 'PrecioConFactura', 'PrecioSinFactura', 'Stock', 'StockMinimo']);
    inv.appendRow(['P1', 'Lámina Acrílica 2mm', '2 mm', '1.22 x 2.44', 580, 520, 20, 5]);
    inv.appendRow(['P2', 'Lámina Acrílica 2mm', '2 mm', '1.22 x 1.22', 305, 275, 20, 5]);
    inv.appendRow(['P3', 'Lámina Acrílica 3mm', '3 mm', '1.22 x 2.44', 810, 730, 20, 5]);
    inv.appendRow(['P4', 'Lámina Acrílica 3mm', '3 mm', '1.22 x 1.22', 420, 380, 20, 5]);
    inv.setFrozenRows(1);
  }

  let ven = ss.getSheetByName(SHEET_VENTAS) || ss.insertSheet(SHEET_VENTAS);
  if (ven.getLastRow() === 0) {
    ven.appendRow(['ID', 'FechaHora', 'ProductoID', 'Producto', 'Cantidad', 'TipoVenta', 'PrecioUnitario', 'Total', 'Notas']);
    ven.setFrozenRows(1);
  }
  asegurarColumnasVentas_(ven);

  SpreadsheetApp.flush();
  Logger.log('Hojas listas y columnas de anulación verificadas (no se borró ningún dato).');
}

/* ---------------- ENTRADAS ---------------- */

function doGet(e) {
  const p = e.parameter || {};
  try {
    let data;
    switch (p.action) {
      case 'getInventario': data = getInventario(); break;
      case 'getVentas':
        data = getVentas(p.desde, p.hasta, p.limite === undefined || p.limite === '' ? 50 : Number(p.limite));
        break;
      case 'getReporte': data = getReporte(p.desde, p.hasta); break;
      case 'diagnostico': data = diagnostico(); break;
      default: throw new Error('Acción GET no reconocida: ' + p.action);
    }
    return jsonOut({ ok: true, data: data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let result;
    switch (body.action) {
      case 'registrarVenta': result = registrarVenta(body); break;
      case 'editarVenta': result = editarVenta(body); break;
      case 'anularVenta': result = anularVenta(body); break;
      case 'restaurarVenta': result = restaurarVenta(body); break;
      case 'eliminarVenta': result = eliminarVenta(body); break;
      case 'ajustarStock': result = ajustarStock(body.productoId, Number(body.cantidad), body.tipo || 'ingreso'); break;
      case 'actualizarProducto': result = actualizarProducto(body); break;
      case 'crearProducto': result = crearProducto(body); break;
      default: throw new Error('Acción POST no reconocida: ' + body.action);
    }
    return jsonOut({ ok: true, data: result });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function hoja_(nombre) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!s) throw new Error('No existe la hoja "' + nombre + '"');
  return s;
}

/* ---------------- COLUMNAS / UTILIDADES ---------------- */

/** Número de columna (1-based) de una cabecera, o -1 si no existe. */
function colIdx_(sheet, nombre) {
  const n = sheet.getLastColumn();
  if (n < 1) return -1;
  const cab = sheet.getRange(1, 1, 1, n).getValues()[0];
  for (let i = 0; i < cab.length; i++) {
    if (String(cab[i]).trim().toLowerCase() === nombre.toLowerCase()) return i + 1;
  }
  return -1;
}

/** Agrega ID / Estado / FechaAnulacion / MotivoAnulacion al final si faltan. No toca datos. */
function asegurarColumnasVentas_(sheet) {
  asegurarColumnaID_(sheet);
  const faltan = COLS_ANULACION.filter(c => colIdx_(sheet, c) === -1);
  if (!faltan.length) return;

  const ult = sheet.getLastColumn();
  const filasDatos = Math.max(sheet.getLastRow() - 1, 0);
  faltan.forEach((col, i) => {
    const c = ult + 1 + i;
    sheet.getRange(1, c, 1, 1).setValue(col);
    if (filasDatos > 0) sheet.getRange(2, c, filasDatos, 1).setValue(col === 'Estado' ? ESTADO_ACTIVA : '');
  });
}

/**
 * La columna ID es la que permite EDITAR, ANULAR y ELIMINAR una venta: sin ID
 * la app no muestra esos botones. Si la hoja se creó con una versión anterior
 * que no la tenía, se agrega al final y se le genera un ID a cada venta que
 * esté sin él. Las ventas viejas no se borran ni se modifican en nada más.
 */
function asegurarColumnaID_(sheet) {
  if (colIdx_(sheet, 'ID') !== -1) return;

  const c = sheet.getLastColumn() + 1;
  sheet.getRange(1, c, 1, 1).setValue('ID');

  const last = sheet.getLastRow();
  if (last < 2) return;

  const fFecha = colIdx_(sheet, 'FechaHora');
  const filas = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const usados = {};
  const ids = filas.map((fila, i) => {
    const f = fFecha > 0 ? fila[fFecha - 1] : null;
    const base = (f instanceof Date && !isNaN(f.getTime())) ? f.getTime() : (Date.now() + i);
    let id = 'V' + base;
    while (usados[id]) id += 'x';
    usados[id] = true;
    return [id];
  });
  sheet.getRange(2, c, ids.length, 1).setValues(ids);
}

/** Arma una fila respetando el orden real de las cabeceras de la hoja. */
function armarFila_(sheet, obj) {
  const n = sheet.getLastColumn();
  const cab = sheet.getRange(1, 1, 1, n).getValues()[0];
  const fila = [];
  for (let i = 0; i < n; i++) {
    const k = String(cab[i]).trim();
    fila.push(k && obj[k] !== undefined && obj[k] !== null ? obj[k] : '');
  }
  return fila;
}

function valorVenta_(sheet, v, nombre) {
  const c = colIdx_(sheet, nombre);
  return c > 0 ? v.valores[c - 1] : '';
}

function estadoVenta_(sheet, v) {
  return String(valorVenta_(sheet, v, 'Estado')).trim().toLowerCase() === 'anulada'
    ? ESTADO_ANULADA : ESTADO_ACTIVA;
}

/* ---------------- DIAGNÓSTICO ---------------- */

/**
 * Corre esto desde el editor (función diagnostico) o por URL:
 *   <URL_DEL_WEB_APP>?action=diagnostico
 * Dice si el código está pegado completo y si la hoja está bien montada.
 */
function diagnostico() {
  const salida = [];
  const requisite = [
    'doGet', 'doPost', 'jsonOut', 'hoja_', 'colIdx_', 'asegurarColumnasVentas_',
    'asegurarColumnaID_', 'armarFila_', 'valorVenta_', 'estadoVenta_', 'getInventario', 'findProductoRow',
    'crearProducto', 'actualizarProducto', 'ajustarStock', 'moverStock_',
    'registrarVenta', 'buscarVenta_', 'editarVenta', 'anularVenta',
    'restaurarVenta', 'eliminarVenta', 'getVentas', 'getReporte'
  ];
  const faltan = requisite.filter(n => eval('typeof ' + n) === 'undefined');
  salida.push('Funciones pegadas: ' + (requisite.length - faltan.length) + '/' + requisite.length +
    (faltan.length ? '   FALTAN: ' + faltan.join(', ') : '   (completo)'));

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    salida.push('Hoja de cálculo: ' + ss.getName());

    ['Inventario', 'Ventas'].forEach(n => {
      const s = ss.getSheetByName(n);
      if (!s) { salida.push('HOJA "' + n + '" NO EXISTE -> corre configurarHojas()'); return; }
      salida.push('Hoja ' + n + ': ' + s.getLastRow() + ' fila(s) x ' + s.getLastColumn() + ' columna(s)');
    });

    const ven = ss.getSheetByName('Ventas');
    if (ven) {
      salida.push('Cabecera Ventas: ' + ven.getRange(1, 1, 1, ven.getLastColumn()).getValues()[0].join(' | '));
      COLS_ANULACION.forEach(c => {
        if (colIdx_(ven, c) === -1) salida.push('FALTA la columna ' + c + ' -> se creará sola al primer uso');
      });
      if (colIdx_(ven, 'ID') === -1) {
        salida.push('FALTA la columna ID -> se creará sola y las ventas viejas recibirán un ID; sin ID la app NO muestra los botones Editar/Anular/Eliminar');
      }
    }
    salida.push('Zona horaria: ' + ss.getSpreadsheetTimeZone());
    salida.push('TODO OK: el sistema está conectado.');
  } catch (err) {
    salida.push('ERROR: ' + err.message);
  }
  Logger.log(salida.join('\n'));
  return salida;
}

/* ---------------- INVENTARIO ---------------- */

function getInventario() {
  const rows = hoja_(SHEET_INVENTARIO).getDataRange().getValues();
  const headers = rows.shift();
  return rows
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
}

function findProductoRow(sheet, productoId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(productoId)) return i + 2;
  }
  return -1;
}

function crearProducto(body) {
  const sheet = hoja_(SHEET_INVENTARIO);
  const data = sheet.getDataRange().getValues();
  const ids = data.slice(1).map(r => String(r[0]));
  let n = data.length;
  let newId;
  do { newId = 'P' + n; n++; } while (ids.indexOf(newId) !== -1);

  sheet.appendRow([
    newId, body.producto, body.espesor, body.medida,
    Number(body.precioConFactura) || 0, Number(body.precioSinFactura) || 0,
    Number(body.stock) || 0, Number(body.stockMinimo) || 5
  ]);
  return { id: newId };
}

function actualizarProducto(body) {
  const sheet = hoja_(SHEET_INVENTARIO);
  const row = findProductoRow(sheet, body.productoId);
  if (row === -1) throw new Error('Producto no encontrado: ' + body.productoId);

  if (body.producto !== undefined) sheet.getRange(row, 2).setValue(body.producto);
  if (body.espesor !== undefined) sheet.getRange(row, 3).setValue(body.espesor);
  if (body.medida !== undefined) sheet.getRange(row, 4).setValue(body.medida);
  if (body.precioConFactura !== undefined) sheet.getRange(row, 5).setValue(Number(body.precioConFactura));
  if (body.precioSinFactura !== undefined) sheet.getRange(row, 6).setValue(Number(body.precioSinFactura));
  if (body.stockMinimo !== undefined) sheet.getRange(row, 8).setValue(Number(body.stockMinimo));
  return { ok: true };
}

/** Reposición o corrección manual. tipo: 'ingreso' | 'correccion' */
function ajustarStock(productoId, cantidad, tipo) {
  if (isNaN(cantidad)) throw new Error('Cantidad inválida.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = hoja_(SHEET_INVENTARIO);
    const row = findProductoRow(sheet, productoId);
    if (row === -1) throw new Error('Producto no encontrado: ' + productoId);

    const stockCell = sheet.getRange(row, 7);
    const actual = Number(stockCell.getValue());
    const nuevo = tipo === 'correccion' ? cantidad : actual + cantidad;
    if (nuevo < 0) throw new Error('El stock no puede quedar negativo.');
    stockCell.setValue(nuevo);
    return { productoId: productoId, stockAnterior: actual, stockNuevo: nuevo };
  } finally {
    lock.releaseLock();
  }
}

/** Suma/resta stock validando que no quede negativo. Devuelve {nuevo, producto}. */
function moverStock_(invSheet, productoId, delta) {
  const row = findProductoRow(invSheet, productoId);
  if (row === -1) throw new Error('Producto no encontrado en inventario: ' + productoId);
  const fila = invSheet.getRange(row, 1, 1, 8).getValues()[0];
  const actual = Number(fila[6]);
  const nuevo = actual + delta;
  if (nuevo < 0) throw new Error('Stock insuficiente. Disponible: ' + actual);
  invSheet.getRange(row, 7).setValue(nuevo);
  return { nuevo: nuevo, producto: fila[1] };
}

/* ---------------- VENTAS ---------------- */

/**
 * body: { productoId, cantidad, tipoVenta, precioUnitario (opcional), notas (opcional) }
 * Descuenta stock dentro del mismo bloqueo (evita condiciones de carrera).
 * La venta nace siempre como "Activa".
 */
function registrarVenta(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invSheet = hoja_(SHEET_INVENTARIO);
    const venSheet = hoja_(SHEET_VENTAS);
    asegurarColumnasVentas_(venSheet);

    const row = findProductoRow(invSheet, body.productoId);
    if (row === -1) throw new Error('Producto no encontrado: ' + body.productoId);

    const f = invSheet.getRange(row, 1, 1, 8).getValues()[0];
    const producto = f[1];
    const precioConFactura = Number(f[4]);
    const precioSinFactura = Number(f[5]);
    const stockActual = Number(f[6]);

    const cantidad = Number(body.cantidad);
    if (!cantidad || cantidad <= 0) throw new Error('Cantidad inválida.');
    if (cantidad > stockActual) throw new Error('Stock insuficiente. Disponible: ' + stockActual);

    const tipoVenta = body.tipoVenta === 'Sin Factura' ? 'Sin Factura' : 'Con Factura';
    const precioUnitario = body.precioUnitario !== undefined && body.precioUnitario !== ''
      ? Number(body.precioUnitario)
      : (tipoVenta === 'Con Factura' ? precioConFactura : precioSinFactura);
    const total = precioUnitario * cantidad;

    const nuevoStock = stockActual - cantidad;
    invSheet.getRange(row, 7).setValue(nuevoStock);

    const ventaId = 'V' + new Date().getTime();
    venSheet.appendRow(armarFila_(venSheet, {
      ID: ventaId, FechaHora: new Date(), ProductoID: body.productoId, Producto: producto,
      Cantidad: cantidad, TipoVenta: tipoVenta, PrecioUnitario: precioUnitario,
      Total: total, Notas: body.notas || '', Estado: ESTADO_ACTIVA,
      FechaAnulacion: '', MotivoAnulacion: ''
    }));

    return {
      ventaId: ventaId, producto: producto, cantidad: cantidad, tipoVenta: tipoVenta,
      precioUnitario: precioUnitario, total: total, stockNuevo: nuevoStock
    };
  } finally {
    lock.releaseLock();
  }
}

/** Busca una venta por su ID. Devuelve { fila, valores: fila completa }. */
function buscarVenta_(venSheet, ventaId) {
  if (!ventaId) throw new Error('Falta el ID de la venta.');
  asegurarColumnaID_(venSheet);

  const cId = colIdx_(venSheet, 'ID');
  const last = venSheet.getLastRow();
  if (last < 2) throw new Error('No hay ventas registradas.');
  const ids = venSheet.getRange(2, cId, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(ventaId)) {
      const fila = i + 2;
      return { fila: fila, valores: venSheet.getRange(fila, 1, 1, venSheet.getLastColumn()).getValues()[0] };
    }
  }
  throw new Error('La venta ya no existe. Recarga la app.');
}

/**
 * body: { ventaId, productoId, cantidad, tipoVenta, precioUnitario, notas }
 * Ajusta el stock: devuelve lo de la venta original y descuenta lo nuevo.
 * Conserva el ID y la fecha originales. No permite editar ventas anuladas.
 */
function editarVenta(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invSheet = hoja_(SHEET_INVENTARIO);
    const venSheet = hoja_(SHEET_VENTAS);
    asegurarColumnasVentas_(venSheet);
    const v = buscarVenta_(venSheet, body.ventaId);
    if (estadoVenta_(venSheet, v) === ESTADO_ANULADA) {
      throw new Error('No se puede editar una venta anulada. Reactívala primero.');
    }

    const oldId = String(valorVenta_(venSheet, v, 'ProductoID'));
    const oldQ = Number(valorVenta_(venSheet, v, 'Cantidad'));
    const newId = String(body.productoId || oldId);
    const newQ = Number(body.cantidad);
    const precio = Number(body.precioUnitario);
    if (!newQ || newQ <= 0) throw new Error('Cantidad inválida.');
    if (isNaN(precio) || precio < 0) throw new Error('Precio inválido.');
    const tipoVenta = body.tipoVenta === 'Sin Factura' ? 'Sin Factura' : 'Con Factura';

    let r;
    if (newId === oldId) {
      r = moverStock_(invSheet, oldId, oldQ - newQ);
    } else {
      r = moverStock_(invSheet, newId, -newQ);      // primero valida el nuevo producto
      moverStock_(invSheet, oldId, oldQ);           // luego devuelve el stock del anterior
    }

    const cambios = {
      ProductoID: newId, Producto: r.producto, Cantidad: newQ, TipoVenta: tipoVenta,
      PrecioUnitario: precio, Total: precio * newQ, Notas: body.notas || ''
    };
    Object.keys(cambios).forEach(k => {
      const c = colIdx_(venSheet, k);
      if (c > 0) venSheet.getRange(v.fila, c, 1, 1).setValue(cambios[k]);
    });

    return { ventaId: body.ventaId, total: precio * newQ, stockNuevo: r.nuevo };
  } finally {
    lock.releaseLock();
  }
}

/**
 * body: { ventaId, motivo } — ANULA la venta: no la borra.
 * Queda con Estado=Anulada + fecha + motivo, las láminas vuelven al stock
 * y el reporte del mes resta ese monto del total del periodo.
 */
function anularVenta(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invSheet = hoja_(SHEET_INVENTARIO);
    const venSheet = hoja_(SHEET_VENTAS);
    asegurarColumnasVentas_(venSheet);
    const v = buscarVenta_(venSheet, body.ventaId);
    if (estadoVenta_(venSheet, v) === ESTADO_ANULADA) throw new Error('Esa venta ya está anulada.');

    const productoId = String(valorVenta_(venSheet, v, 'ProductoID'));
    const cantidad = Number(valorVenta_(venSheet, v, 'Cantidad')) || 0;
    const total = Number(valorVenta_(venSheet, v, 'Total')) || 0;
    const motivo = String(body.motivo || '').trim().slice(0, 250);

    const r = moverStock_(invSheet, productoId, cantidad);

    venSheet.getRange(v.fila, colIdx_(venSheet, 'Estado'), 1, 1).setValue(ESTADO_ANULADA);
    venSheet.getRange(v.fila, colIdx_(venSheet, 'FechaAnulacion'), 1, 1).setValue(new Date());
    venSheet.getRange(v.fila, colIdx_(venSheet, 'MotivoAnulacion'), 1, 1).setValue(motivo);

    return {
      ventaId: body.ventaId, estado: ESTADO_ANULADA, motivo: motivo,
      cantidadDevuelta: cantidad, montoAnulado: total, stockNuevo: r.nuevo
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * body: { ventaId } — Reactiva una venta anulada.
 * Vuelve a contar en el reporte y descuenta el stock otra vez.
 */
function restaurarVenta(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invSheet = hoja_(SHEET_INVENTARIO);
    const venSheet = hoja_(SHEET_VENTAS);
    asegurarColumnasVentas_(venSheet);
    const v = buscarVenta_(venSheet, body.ventaId);
    if (estadoVenta_(venSheet, v) !== ESTADO_ANULADA) throw new Error('Esa venta no está anulada.');

    const productoId = String(valorVenta_(venSheet, v, 'ProductoID'));
    const cantidad = Number(valorVenta_(venSheet, v, 'Cantidad')) || 0;

    const r = moverStock_(invSheet, productoId, -cantidad);

    venSheet.getRange(v.fila, colIdx_(venSheet, 'Estado'), 1, 1).setValue(ESTADO_ACTIVA);
    venSheet.getRange(v.fila, colIdx_(venSheet, 'FechaAnulacion'), 1, 1).setValue('');
    venSheet.getRange(v.fila, colIdx_(venSheet, 'MotivoAnulacion'), 1, 1).setValue('');

    return {
      ventaId: body.ventaId, estado: ESTADO_ACTIVA,
      cantidadDescontada: cantidad, stockNuevo: r.nuevo
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * body: { ventaId } — BORRADO DEFINITIVO de la venta (uso administrativo).
 * Si la venta estaba anulada ya devolvió el stock, así que no lo devuelve dos veces.
 * Desde la app se usa "anular"; esto queda como respaldo.
 */
function eliminarVenta(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invSheet = hoja_(SHEET_INVENTARIO);
    const venSheet = hoja_(SHEET_VENTAS);
    const v = buscarVenta_(venSheet, body.ventaId);
    const estabaAnulada = estadoVenta_(venSheet, v) === ESTADO_ANULADA;
    const productoId = String(valorVenta_(venSheet, v, 'ProductoID'));
    const cantidad = Number(valorVenta_(venSheet, v, 'Cantidad')) || 0;
    const r = estabaAnulada ? { nuevo: null } : moverStock_(invSheet, productoId, cantidad);
    venSheet.deleteRow(v.fila);
    return { ventaId: body.ventaId, stockNuevo: r.nuevo, borrada: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ventas más recientes primero. desde/hasta: 'YYYY-MM-DD' (opcionales).
 * limite: 0 = todas. Incluye las columnas Estado / MotivoAnulacion.
 */
function getVentas(desde, hasta, limite) {
  const sheet = hoja_(SHEET_VENTAS);
  asegurarColumnasVentas_(sheet);
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const d = desde ? new Date(desde + 'T00:00:00') : null;
  const h = hasta ? new Date(hasta + 'T23:59:59') : null;

  let ventas = [];
  rows.forEach(r => {
    if (r[0] === '') return;
    const f = new Date(r[1]);
    if (d && f < d) return;
    if (h && f > h) return;
    const obj = {};
    headers.forEach((k, i) => (obj[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]));
    ventas.push(obj);
  });
  ventas.sort((a, b) => new Date(b.FechaHora) - new Date(a.FechaHora));
  return limite > 0 ? ventas.slice(0, limite) : ventas;
}

/**
 * Reporte agregado: sólo cuenta las ventas ACTIVAS y expone el neto.
 * (La app calcula el detalle con getVentas, pero esto sirve para exportar.)
 */
function getReporte(desde, hasta) {
  const lista = getVentas(desde, hasta, 0);
  let totalConFactura = 0, totalSinFactura = 0, cantidadConFactura = 0, cantidadSinFactura = 0;
  let montoAnulado = 0, laminasAnuladas = 0, ventasAnuladas = 0;
  const porProducto = {};

  lista.forEach(v => {
    const cantidad = Number(v.Cantidad), total = Number(v.Total);
    const anulada = String(v.Estado || '').trim().toLowerCase() === 'anulada';

    if (anulada) {
      ventasAnuladas++;
      montoAnulado += total;
      laminasAnuladas += cantidad;
    } else if (v.TipoVenta === 'Con Factura') {
      totalConFactura += total; cantidadConFactura += cantidad;
    } else {
      totalSinFactura += total; cantidadSinFactura += cantidad;
    }

    const signo = anulada ? -1 : 1;
    const p = porProducto[v.Producto] || (porProducto[v.Producto] = { cantidad: 0, total: 0 });
    p.cantidad += signo * cantidad;
    p.total += signo * total;
  });

  const bruto = totalConFactura + totalSinFactura;
  return {
    totalConFactura: totalConFactura, totalSinFactura: totalSinFactura,
    totalBruto: bruto,
    montoAnulado: montoAnulado,
    totalNeto: bruto - montoAnulado,
    ventasAnuladas: ventasAnuladas,
    laminasAnuladas: laminasAnuladas,
    cantidadConFactura: cantidadConFactura, cantidadSinFactura: cantidadSinFactura,
    laminasNetas: cantidadConFactura + cantidadSinFactura - laminasAnuladas,
    porProducto: porProducto
  };
}