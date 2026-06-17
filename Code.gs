// ═══════════════════════════════════════════════════════════
//  SISTEMA DE DESPACHO · SUGAL GROUP
//  Apps Script — Backend Google Sheets
//  v5 — Registros + FCL + BBDD Rechazos + Email Gmail
// ═══════════════════════════════════════════════════════════

// ── NOMBRES DE HOJAS ──
const NOMBRE_HOJA         = 'Registros';
const NOMBRE_HOJA_FCL     = 'Status FCL';
const NOMBRE_HOJA_RECHAZOS = 'BBDD Rechazos';
const NOMBRE_HOJA_PEDIDOS = 'Pedidos Activos';

// ── COLUMNAS HOJA PEDIDOS ACTIVOS ──
// Cada pedido se guarda como un bloque JSON completo en una sola fila
const HEADERS_PEDIDOS = [
  'ID','Pedido','Cliente','Cargado En','Datos JSON','Últ. Actualización'
];

// ── COLUMNAS HOJA REGISTROS ──
const COLUMNAS = [
  'fecha','turno','operadores','fcl','pallets','accidentes',
  'trasvasije','brcsp','brsspA','brsspB','bolsaSuelta',
  'bolsaDesmetalizada','boquilla','tamborOxido','tamborPintura',
  'tamborRoto','tamborAbollado','otro','otroTexto','fechaMod'
];

const HEADERS_REGISTROS = [
  'Fecha','Turno','Operadores','FCL Preparados','Cambio Pallets',
  'Accidentes','Trasvasije','BRCSP','BRSSP-A (Leve)','BRSSP-B (Severo)',
  'Bolsa Suelta','Bolsa Desmetalizada','Boquilla No Conforme',
  'Tambor Oxido','Tambor Pintura','Tambor Roto','Tambor Abollado',
  'Otro','Descripción Otro','Últ. Modificación'
];

// ── COLUMNAS HOJA STATUS FCL ──
const HEADERS_FCL = [
  'Fecha','Turno','N° Pedido','Cantidad Preparada','Hora Registro'
];

// ── COLUMNAS HOJA BBDD RECHAZOS ──
const HEADERS_RECHAZOS = [
  'Fecha Rechazo','Cliente','Pedido SAP','N° Tambor/Tote',
  'Código Material','Fecha Elaboración','Peso Neto',
  'Motivo (COD)','Destino','Fecha Solicitud','Hora Registro'
];

// ── NOMBRES LEGIBLES DE RECHAZOS (para emails) ──
const NOMBRES_RECHAZO = {
  brcsp:              'BRCSP - Bolsa rota con salida de producto',
  brsspA:             'BRSSP-A - Bolsa rota sin salida de producto (Leve)',
  brsspB:             'BRSSP-B - Bolsa rota sin salida de producto (Severo)',
  bolsaSuelta:        'Bolsa suelta desprendida',
  bolsaDesmetalizada: 'Bolsa desmetalizada',
  boquilla:           'Boquilla no conforme',
  tamborOxido:        'Tambor con óxido severo',
  tamborPintura:      'Tambor con desprendimiento de pintura',
  tamborRoto:         'Tambor roto',
  tamborAbollado:     'Tambor abollado',
  otro:               'Otro'
};

const CAMPOS_RECHAZO = Object.keys(NOMBRES_RECHAZO);

// ── EMAIL ──
// Cambia esta dirección si quieres que el correo llegue a otro destino
const EMAIL_DESTINO = 'braulio.arroyo@sugal-group.com';

// ═══════════════════════════════════════════
//  ENTRY POINTS HTTP
// ═══════════════════════════════════════════

function doGet(e) {
  const params = e.parameter;
  const accion = params.accion || '';
  try {
    if (accion === 'obtener')    return respuesta(obtenerRegistro(params.fecha, params.turno));
    if (accion === 'historial')  return respuesta(obtenerHistorial());
    if (accion === 'obtenerFCL') return respuesta(obtenerFCL(params.fecha, params.turno));
    if (accion === 'obtenerPedidos') return respuesta(obtenerPedidos());
    return respuesta({ ok: false, error: 'Accion no reconocida' });
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.accion === 'guardar')        return respuesta(guardarRegistro(payload));
    if (payload.accion === 'guardarFCL')     return respuesta(guardarFCL(payload));
    if (payload.accion === 'guardarRechazo') return respuesta(guardarRechazo(payload));
    if (payload.accion === 'guardarPedido')  return respuesta(guardarPedido(payload));
    if (payload.accion === 'eliminarPedido') return respuesta(eliminarPedido(payload));
    return respuesta({ ok: false, error: 'Accion no reconocida' });
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  }
}

function respuesta(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════
//  NORMALIZAR FECHA
// ═══════════════════════════════════════════

function normalizarFecha(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(val).trim();
}

// ═══════════════════════════════════════════
//  GUARDAR REGISTRO DE TURNO (con upsert)
// ═══════════════════════════════════════════

function guardarRegistro(payload) {
  const hoja       = obtenerOCrearHoja(NOMBRE_HOJA, HEADERS_REGISTROS, formatearConfigRegistros);
  const fechaNorm  = String(payload.fecha).trim();
  const turnoNorm  = String(payload.turno).trim();
  const clave      = fechaNorm + '_' + turnoNorm;
  const filaExist  = buscarFila(hoja, clave);

  // Detectar rechazos nuevos para email
  let rechazosNuevos = [];
  if (filaExist > 0) {
    const anterior = {};
    const datosAnt = hoja.getRange(filaExist, 1, 1, COLUMNAS.length).getValues()[0];
    COLUMNAS.forEach((col, i) => { anterior[col] = datosAnt[i]; });
    CAMPOS_RECHAZO.forEach(campo => {
      const nuevo    = parseInt(payload[campo]) || 0;
      const antiguo  = parseInt(anterior[campo]) || 0;
      if (nuevo > antiguo) {
        rechazosNuevos.push({
          nombre: campo === 'otro' ? (payload.otroTexto || 'Otro') : NOMBRES_RECHAZO[campo],
          cantidad: nuevo
        });
      }
    });
  } else {
    CAMPOS_RECHAZO.forEach(campo => {
      const valor = parseInt(payload[campo]) || 0;
      if (valor > 0) {
        rechazosNuevos.push({
          nombre: campo === 'otro' ? (payload.otroTexto || 'Otro') : NOMBRES_RECHAZO[campo],
          cantidad: valor
        });
      }
    });
  }

  // Construir fila
  const fila = COLUMNAS.map(col => {
    if (col === 'fecha')   return fechaNorm;
    if (col === 'turno')   return turnoNorm;
    if (col === 'otroTexto' || col === 'fechaMod') return payload[col] || '';
    return parseInt(payload[col]) || 0;
  });

  if (filaExist > 0) {
    hoja.getRange(filaExist, 1, 1, fila.length).setValues([fila]);
  } else {
    hoja.appendRow(fila);
    colorearTurno(hoja, hoja.getLastRow(), turnoNorm, COLUMNAS.length);
  }

  // Enviar emails
  rechazosNuevos.forEach(r => {
    try { enviarEmailRechazo(r.nombre, fechaNorm, turnoNorm, r.cantidad); } catch(e) {}
  });

  return { ok: true, accion: filaExist > 0 ? 'actualizado' : 'creado', emails: rechazosNuevos.length };
}

// ═══════════════════════════════════════════
//  GUARDAR FCL
// ═══════════════════════════════════════════

function guardarFCL(payload) {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA_FCL, HEADERS_FCL, null);
  const ahora = new Date();
  const hora  = ahora.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' });

  hoja.appendRow([
    String(payload.fecha).trim(),
    String(payload.turno).trim(),
    String(payload.numeroPedido).trim(),
    parseInt(payload.cantidadPreparada) || 0,
    hora
  ]);

  // Alternar color de fila
  const uf = hoja.getLastRow();
  if (uf % 2 === 0) hoja.getRange(uf, 1, 1, HEADERS_FCL.length).setBackground('#fafafa');

  return { ok: true };
}

// ═══════════════════════════════════════════
//  OBTENER FCL POR FECHA + TURNO
// ═══════════════════════════════════════════

function obtenerFCL(fecha, turno) {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA_FCL, HEADERS_FCL, null);
  const uf   = hoja.getLastRow();
  if (uf < 2) return { ok: true, fcls: [] };

  const datos = hoja.getRange(2, 1, uf - 1, HEADERS_FCL.length).getValues();
  const fcls  = datos
    .filter(fila => normalizarFecha(fila[0]) === String(fecha).trim() && String(fila[1]).trim() === String(turno).trim())
    .map(fila => ({
      fecha:             normalizarFecha(fila[0]),
      turno:             fila[1],
      numeroPedido:      fila[2],
      cantidadPreparada: fila[3],
      hora:              fila[4]
    }));

  return { ok: true, fcls };
}

// ═══════════════════════════════════════════
//  GUARDAR RECHAZO EN BBDD RECHAZOS
//  Llamado desde Ventana 3 al generar solicitud
// ═══════════════════════════════════════════

function guardarRechazo(payload) {
  const hoja   = obtenerOCrearHoja(NOMBRE_HOJA_RECHAZOS, HEADERS_RECHAZOS, null);
  const ahora  = new Date();
  const hora   = ahora.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' });
  const hoy    = normalizarFecha(ahora);

  // payload.productos es un array de {tamb, mat, peso, fecha, mot, dest}
  const prods  = payload.productos || [];

  prods.forEach(p => {
    hoja.appendRow([
      payload.fecha    || hoy,        // Fecha Rechazo
      payload.cliente  || '',         // Cliente
      payload.pedidoSAP || '',        // Pedido SAP
      p.tamb           || '',         // N° Tambor/Tote
      p.mat            || '',         // Código Material
      p.fecha          || '',         // Fecha Elaboración
      p.peso           || '',         // Peso Neto
      p.mot            || '',         // Motivo (COD)
      p.dest           || '',         // Destino
      hoy,                            // Fecha Solicitud
      hora                            // Hora Registro
    ]);
    // Alternar color
    const uf = hoja.getLastRow();
    if (uf % 2 === 0) hoja.getRange(uf, 1, 1, HEADERS_RECHAZOS.length).setBackground('#fff0f2');
  });

  return { ok: true, filas: prods.length };
}

// ═══════════════════════════════════════════
//  OBTENER REGISTRO DE TURNO
// ═══════════════════════════════════════════

function obtenerRegistro(fecha, turno) {
  const hoja   = obtenerOCrearHoja(NOMBRE_HOJA, HEADERS_REGISTROS, formatearConfigRegistros);
  const clave  = String(fecha).trim() + '_' + String(turno).trim();
  const filaNum = buscarFila(hoja, clave);
  if (filaNum < 0) return { ok: true, registro: null };

  const datos    = hoja.getRange(filaNum, 1, 1, COLUMNAS.length).getValues()[0];
  const registro = {};
  COLUMNAS.forEach((col, i) => {
    registro[col] = col === 'fecha' ? normalizarFecha(datos[i]) : datos[i];
  });
  return { ok: true, registro };
}

// ═══════════════════════════════════════════
//  HISTORIAL DE TURNOS
// ═══════════════════════════════════════════

function obtenerHistorial() {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA, HEADERS_REGISTROS, formatearConfigRegistros);
  const uf   = hoja.getLastRow();
  if (uf < 2) return { ok: true, registros: [] };

  const datos     = hoja.getRange(2, 1, uf - 1, COLUMNAS.length).getValues();
  const registros = datos
    .filter(fila => fila[0])
    .map(fila => {
      const r = {};
      COLUMNAS.forEach((col, i) => {
        r[col] = col === 'fecha' ? normalizarFecha(fila[i]) : fila[i];
      });
      return r;
    })
    .sort((a, b) => {
      if (b.fecha > a.fecha) return 1;
      if (b.fecha < a.fecha) return -1;
      return String(a.turno).localeCompare(String(b.turno));
    });

  return { ok: true, registros };
}

// ═══════════════════════════════════════════
//  EMAIL AUTOMÁTICO AL REGISTRAR RECHAZO
// ═══════════════════════════════════════════

function enviarEmailRechazo(motivoRechazo, fecha, turno, cantidad) {
  const turnoLegible = turno === 'DIA' ? 'Turno Día' : 'Turno Noche';
  const asunto = 'RECHAZO PREPARACIÓN DE DESPACHO POR ' + motivoRechazo.toUpperCase() + ' - ' + fecha;
  const cuerpo =
    'Estimados,\n\n' +
    'Con fecha ' + fecha + ' (' + turnoLegible + ') se registró un rechazo por ' + motivoRechazo + '.\n\n' +
    'Cantidad registrada: ' + cantidad + '\n\n' +
    'Saludos!\n\n' +
    '---\n' +
    'Mensaje automático · Sistema Despacho Sugal\n' +
    'Sugal Group · Planta Quinta de Tilcoco';

  GmailApp.sendEmail(EMAIL_DESTINO, asunto, cuerpo);
}

// ═══════════════════════════════════════════
//  BUSCAR FILA POR CLAVE fecha_turno
// ═══════════════════════════════════════════

function buscarFila(hoja, clave) {
  const uf = hoja.getLastRow();
  if (uf < 2) return -1;

  const datos = hoja.getRange(2, 1, uf - 1, 2).getValues();
  for (let i = 0; i < datos.length; i++) {
    const fechaVal = normalizarFecha(datos[i][0]);
    const turnoVal = String(datos[i][1]).trim();
    if ((fechaVal + '_' + turnoVal) === clave) return i + 2;
  }
  return -1;
}

// ═══════════════════════════════════════════
//  CREAR O RETORNAR HOJA EXISTENTE
// ═══════════════════════════════════════════

function obtenerOCrearHoja(nombre, headers, configFn) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let hoja   = ss.getSheetByName(nombre);

  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Formato del header
    const rh = hoja.getRange(1, 1, 1, headers.length);
    rh.setBackground('#C8102E');
    rh.setFontColor('#ffffff');
    rh.setFontWeight('bold');
    rh.setFontSize(11);
    hoja.setFrozenRows(1);

    // Configuración específica por hoja
    if (configFn) configFn(hoja);
  }

  return hoja;
}

// Configuración específica de la hoja Registros
function formatearConfigRegistros(hoja) {
  hoja.setColumnWidth(1, 110);  // Fecha
  hoja.setColumnWidth(2, 80);   // Turno
  hoja.setColumnWidth(19, 200); // Descripción Otro
  hoja.setColumnWidth(20, 160); // Últ. Modificación
  hoja.getRange(1, 3, 1000, 17).setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════
//  GUARDAR/ACTUALIZAR PEDIDO (memoria compartida)
//  Cada pedido = una fila, datos completos en JSON
// ═══════════════════════════════════════════

function guardarPedido(payload) {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA_PEDIDOS, HEADERS_PEDIDOS, null);
  const id   = String(payload.id);
  const uf   = hoja.getLastRow();
  const ahora = new Date().toLocaleString('es-CL');

  let filaExist = -1;
  if (uf >= 2) {
    const ids = hoja.getRange(2, 1, uf - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) { filaExist = i + 2; break; }
    }
  }

  const fila = [
    id,
    payload.pedido || '',
    payload.cliente || '',
    payload.cargadoEn || '',
    JSON.stringify(payload.datos || {}),
    ahora
  ];

  if (filaExist > 0) {
    hoja.getRange(filaExist, 1, 1, fila.length).setValues([fila]);
  } else {
    hoja.appendRow(fila);
  }

  return { ok: true };
}

// ═══════════════════════════════════════════
//  OBTENER TODOS LOS PEDIDOS ACTIVOS
// ═══════════════════════════════════════════

function obtenerPedidos() {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA_PEDIDOS, HEADERS_PEDIDOS, null);
  const uf   = hoja.getLastRow();
  if (uf < 2) return { ok: true, pedidos: [] };

  const datos = hoja.getRange(2, 1, uf - 1, HEADERS_PEDIDOS.length).getValues();
  const pedidos = datos
    .filter(fila => fila[0])
    .map(fila => {
      let datosJSON = {};
      try { datosJSON = JSON.parse(fila[4]); } catch (e) {}
      return Object.assign({ id: fila[0], pedido: fila[1], cliente: fila[2], cargadoEn: fila[3] }, datosJSON);
    });

  return { ok: true, pedidos };
}

// ═══════════════════════════════════════════
//  ELIMINAR PEDIDO
// ═══════════════════════════════════════════

function eliminarPedido(payload) {
  const hoja = obtenerOCrearHoja(NOMBRE_HOJA_PEDIDOS, HEADERS_PEDIDOS, null);
  const id   = String(payload.id);
  const uf   = hoja.getLastRow();
  if (uf < 2) return { ok: true };

  const ids = hoja.getRange(2, 1, uf - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      hoja.deleteRow(i + 2);
      break;
    }
  }
  return { ok: true };
}
function colorearTurno(hoja, fila, turno, numCols) {
  if (fila % 2 === 0) {
    hoja.getRange(fila, 1, 1, numCols).setBackground('#fafafa');
  }
  if (turno === 'DIA') {
    hoja.getRange(fila, 2).setBackground('#fff8e1').setFontColor('#b36000');
  } else {
    hoja.getRange(fila, 2).setBackground('#e8eaf6').setFontColor('#3949ab');
  }
}
