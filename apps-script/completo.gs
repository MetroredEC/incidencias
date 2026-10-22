/* ═════════════════════════════════════════════════════════════════════════
 *  PUNTONET · CONTROL DE INCIDENCIAS — SCRIPT COMPLETO (un solo archivo)
 *
 *  1. Borra TODO lo que haya en el editor y pega este archivo completo.
 *  2. Cambia la clave en la línea marcada "CLAVE" (la misma de antes, o una nueva).
 *  3. Guarda, elige la función  cargarTodo  y pulsa Ejecutar.
 *     (Carga el histórico de 2026 si la hoja está vacía y el reporte de julio a octubre.
 *      Se puede ejecutar varias veces: no duplica nada.)
 *  4. Solo si cambiaste la clave: Implementar → Administrar implementaciones → lápiz →
 *     Versión "Nueva versión". La URL /exec no cambia.
 * ═════════════════════════════════════════════════════════════════════════ */

/* Puntonet · Control de incidencias Metrored — backend en Google Sheets
 *
 * Instalación:
 *  1. Crear una hoja de cálculo nueva → Extensiones → Apps Script → pegar este archivo.
 *  2. Cambiar CLAVE por una clave propia (la que usará el equipo de Puntonet en la página).
 *  3. Ejecutar una vez cargarHistorico() para subir las incidencias del 2026 ya registradas.
 *  4. Implementar → Nueva implementación → Aplicación web
 *     (Ejecutar como: yo · Quién tiene acceso: cualquier persona).
 *  5. Copiar la URL /exec en ENDPOINT dentro de assets/incidencias.js.
 */
var CLAVE = 'cambiar-esta-clave';   // <── CLAVE: cámbiala por la tuya (entre las comillas)
var HOJA = 'Incidencias';
var COLS = ['id', 'fecha', 'descripcion', 'sucursal', 'reporta', 'reportado', 'categoria', 'estado', 'seguimiento', 'creado', 'actualizado'];
var CAMPOS = ['fecha', 'descripcion', 'sucursal', 'reporta', 'reportado', 'categoria', 'estado', 'seguimiento'];

/* ───────── Web app ───────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.clave !== CLAVE) return json({ error: 'clave' });
  return json({ incidencias: leer() });
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.clave !== CLAVE) return json({ error: 'clave' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (p.accion === 'eliminar') eliminar(p.id);
    else guardar(p);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

/* ───────── Datos ───────── */

function hoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('@');                        // la fecha se guarda tal cual (AAAA-MM-DD)
  }
  return sh;
}

function leer() {
  var sh = hoja(), n = sh.getLastRow();
  if (n < 2) return [];
  var tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, n - 1, COLS.length).getValues().map(function (r) {
    var o = {};
    COLS.forEach(function (c, i) {
      var v = r[i];
      o[c] = v instanceof Date ? Utilities.formatDate(v, tz, c === 'fecha' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss") : String(v);
    });
    return o;
  });
}

function fila(id) {
  var sh = hoja(), n = sh.getLastRow();
  if (!id || n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function guardar(p) {
  if (!String(p.descripcion || '').trim()) return;
  var sh = hoja(), ahora = new Date(), r = fila(p.id);
  var valores = CAMPOS.map(function (c) { return limpio(p[c]); });
  if (r) {
    sh.getRange(r, 2, 1, CAMPOS.length).setValues([valores]);
    sh.getRange(r, COLS.indexOf('actualizado') + 1).setValue(ahora);
  } else {
    sh.appendRow([p.id || Utilities.getUuid()].concat(valores, [ahora, ahora]));
  }
}

function eliminar(id) {
  var r = fila(id);
  if (r) hoja().deleteRow(r);
}

/* ───────── Histórico 2026 (ejecutar una sola vez) ───────── */

function cargarHistorico() {
  var sh = hoja();
  if (sh.getLastRow() > 1) throw new Error('La hoja ya tiene incidencias; no se duplica el histórico.');
  var ahora = new Date();
  var filas = HISTORICO.map(function (h, i) {
    return ['h' + (i + 1), h[0], h[1], h[2], h[3], h[4], h[5], 'Abierta', '', ahora, ahora];
  });
  sh.getRange(2, 1, filas.length, COLS.length).setValues(filas);
}

/* fecha · descripción · sucursal · reporta · reportado · categoría
 * Los nombres de pacientes se reservan porque este archivo vive en un repositorio público. */
var HISTORICO = [
  ['2026-02-02', 'RADIOGRAFIA DE MALA CALIDAD, PIDEN ENVIAR REQUERIMIENTO POR CORREO PARA COORDINAR OTRA CITA', 'MANTA', 'MILENA ANDRADE', 'SI', 'Radiografía'],
  ['2026-02-23', 'NO SE CARGABAN LOS RESULTADOS DE UN COLABORADOR, FUERON CARGADOS DESPUES DE LA JORNADA DEL MEDICO DE GUAYAQUIL', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-03-06', 'RADIOGRAFIA INDICABA DESVIACION DE COLUMNA SIN PONER LOS ANGULOS, SE SOLICITO AYUDA DONDE INDICARON MANDAR CORREO A MI EXPERIENCIA METRORED PARA QUE REVISE GERENCIAL', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Radiografía'],
  ['2026-03-17', 'RADIOGRAFIAS DE MALA CALIDAD (BAJA CALIDAD Y FALTA DE NITIDEZ)', 'QUITO', 'MISHEL CARDENAS', 'SI', 'Radiografía'],
  ['2026-03-27', 'DEJAN PENDIENTE DE AGENDAMIENTO AUDIOMETRÍAS, NUNCA LAS AGENDAN SE RETRASA 1 SEMANA EL PROCESO DE SELECCIÓN', 'QUITO', 'EMILIA LAGOS / MISHEL CARDENAS', 'SI', 'Agendamiento'],
  ['2026-03-30', 'EXAMENES DE DROGAS INCOMPLETOS', 'EL COCA', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-03-31', 'EXAMENES NO ESTABAN COMPLETOS, ERA FICHA PERIODICA. ESPERAMOS 35 MINUTOS PARA LA SUBIDA COMPLETA DE RESULTADOS HECHOS HACE MAS DE 5 MESES', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-04-01', 'NO LES AGENDARON AUDIOMETRIA A DOS CANDIDATOS, SE AGENDARON AL DIA SIGUIENTE', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Agendamiento'],
  ['2026-04-03', 'SE SOLICITA RADIOGRAFIAS LUMBARES Y REALIZAN RADIOGRAFIAS CERVICALES, SE REAGENDA Y SE VUELVE A HACER LA RX DIAS DESPUES', 'QUITO', 'MISHEL CARDENAS', 'SI', 'Radiografía'],
  ['2026-04-09', 'CANDIDATA LLEGA A LA HORA AGENDADA Y EL CENTRO MEDICO NO SE ENCUENTRA ABIERTO, SE RETRASA SU OPTOMETRIA PORQUE LE HACEN TOMAR TURNO Y LUEGO LE INFORMAN QUE YA PASO SU TURNO PARA LA OPTOMETRIA', 'LA CAROLINA', 'ISABEL CAJO / MISHEL CARDENAS', 'SI', 'Atención en sucursal'],
  ['2026-04-17', 'ENVIAN SOLICITUD DE EXAMENES DIA VIERNES 17 Y NO SE TIENE RESPUESTA HASTA EL LUNES 20 HASTA LAS 12, SE DEBE SOLICITAR QUE POR FAVOR NOS AYUDEN PARA OBTENER UNA RESPUESTA', 'QUITO', 'HENRY ULLAURI / MISHEL CARDENAS', 'SI', 'Tiempo de respuesta'],
  ['2026-04-24', 'MALA CALIDAD EN REPORTES DE RX (REPORTAN UN DX SIN EMBARGO NO COLOCAN ANGULOS EN EL REPORTE)', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Radiografía'],
  ['2026-05-06', 'LES LLEGAN RESULTADOS DE LOS EXAMENES A LOS POSTULANTES', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Confidencialidad'],
  ['2026-05-21', 'NO SE ESTA PONIENDO ANGULOS EN LAS RADIOGRAFIAS DE QUITO', 'QUITO', 'MILENA ANDRADE', 'SI', 'Radiografía'],
  ['2026-05-26', 'EXAMEN AUDITIVO REALIZADO A PESAR DE INDICACIONES DE NO REALIZAR, DENTRO DE LA CABINA AUN SE ESCUCHABAN RUIDOS EXTERNOS', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Audiometría'],
  ['2026-05-26', 'RX CON ALTERACIÓN ANATÓMICA EVIDENTE, SIN EMBARGO NO SE REPORTA EN EL INFORME', 'LA CAROLINA', 'SILVIA MERA', 'NO', 'Radiografía'],
  ['2026-05-27', 'AUN SE DEBE SOLICITAR EXAMENES A PESAR DE HABER PASADO EL TIEMPO DE ENTREGA', '', 'MISHEL CARDENAS', '', 'Tiempo de respuesta'],
  ['2026-05-28', 'RX LATERAL DE COLUMNA LUMBAR ES TOTALMENTE ILEGIBLE', 'QUITO', 'SILVIA MERA', 'SI', 'Radiografía'],
  ['2026-06-01', 'SE SOLICITA EXAMENES EN HORARIO ACORDADO Y NO SE TIENE RESPUESTA', 'NACIONAL', 'MISHEL CARDENAS', 'SI', 'Tiempo de respuesta'],
  ['2026-06-02', 'NO SE TIENE RESPUESTA DE LA SOLICITUD DEL 01/02/2026 HASTA LA 1PM SE HACE UNA NUEVA SOLICITUD A LAS 3PM Y NO SE TIENE RESPUESTA SE INTENTA CONTACTAR CON JULY SIN RESPUESTA ALGUNA', 'NACIONAL', 'MISHEL CARDENAS', 'SI', 'Tiempo de respuesta'],
  ['2026-06-03', 'NO SE TIENE RESPUESTA DE LA SOLICITUD DEL 02/06/2026 HASTA LAS 10:30AM', 'NACIONAL', 'MISHEL CARDENAS', 'SI', 'Tiempo de respuesta'],
  ['2026-06-03', 'NO SE ATIENDEN LAS LLAMADAS DE PARTE DE LA PERSONA ENCARGADA', 'NACIONAL', 'MISHEL CARDENAS', '', 'Tiempo de respuesta'],
  ['2026-06-04', '8:37 AM NO SE SUBEN LOS RESULTADOS, NO HAY RESPUESTA TELEFONICA POR PARTE DE LA ENCARGADA, SUBEN RESULTADOS SIN OPTOMETRIAS NI AUDIOMETRIAS A PESAR DE HABER HABLADO CON ELLOS E INDICADO QUE SE NECESITABA ESOS RESULTADOS CON URGENCIA', 'NACIONAL', 'MISHEL CARDENAS', 'NO', 'Resultados'],
  ['2026-06-04', 'NO ESTUVIERON CARGADOS LOS RESULTADOS DE EXÁMENES AGENDADOS PREVIAMENTE', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-06-05', 'NO ESTUVIERON CARGADOS LOS RESULTADOS DE EXÁMENES AGENDADOS PREVIAMENTE', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-06-05', 'EN LOS RESULTADOS DE UNA COLABORADORA DE GUAYAQUIL [NOMBRE RESERVADO] ESTÁN LOS RESULTADOS DE OTRA PERSONA', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Confidencialidad'],
  ['2026-06-05', 'SOLICITAN INFORMACIÓN DEL PACIENTE [NOMBRE RESERVADO]: INDICA QUE SÍ SE HIZO SU AUDIOMETRÍA, SIN EMBARGO NO ESTÁ CARGADA', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Resultados'],
  ['2026-06-05', 'SE SOLICITA QUE ENVÍEN LA AUDIOMETRÍA DEL MISMO PACIENTE [NOMBRE RESERVADO] EL DÍA SÁBADO Y NO ENVÍAN; SE RECIBE EL DÍA DOMINGO', 'GUAYAQUIL', 'MISHEL CARDENAS', 'NO', 'Tiempo de respuesta'],
  ['2026-06-08', 'POR CAMBIO DE PLATAFORMAS NO CUENTAN CON EXAMENES SOLICITADOS, NO SE TIENE LOS ACCESOS PARA LOS MISMOS', 'QUITO', 'MISHEL CARDENAS', 'SI', 'Plataforma'],
  ['2026-06-09', 'NO HAY MIGRACIÓN DE EXAMENES ANTERIORES EN LA NUEVA PLATAFORMA', 'NACIONAL', 'MISHEL CARDENAS', 'SI', 'Plataforma'],
  ['2026-06-09', 'RADIOGRAFÍA DE COLABORADOR CON ÁNGULO DE COBB DE 13 GRADOS REPORTADO POR METRORED EN 2025 Y 2026. EN 2023 TIENE RADIOGRAFÍA CON ÁNGULO DE COBB DE 5 GRADOS Y EN 2024 DE 6 GRADOS. POSTERIOR A LA RADIOGRAFÍA DE METRORED ACUDIÓ A OTRO PROVEEDOR: ÁNGULO DE COBB DE 5 GRADOS', 'GUAYAQUIL', 'MILENA ANDRADE', 'SI', 'Radiografía']
];

/* ───────── Utilidades ───────── */

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function limpio(v) { v = String(v == null ? '' : v).slice(0, 2000); return /^[=+\-@]/.test(v) ? "'" + v : v; }   // evita fórmulas inyectadas

/* ───────── Reporte julio a 21 de octubre de 2026 ───────── */

function cargarReporteJulOct() {
  var sh = hoja(), n = sh.getLastRow();
  var existentes = {};
  if (n > 1) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { existentes[String(r[0])] = true; });
  var ahora = new Date(), filas = [];
  REPORTE_JUL_OCT.forEach(function (h, i) {
    var id = 'r' + (i + 1);
    if (!existentes[id]) filas.push([id, h[0], h[1], h[2], h[3], '', h[4], h[5], '', ahora, ahora]);
  });
  if (filas.length) sh.getRange(n + 1, 1, filas.length, COLS.length).setValues(filas);
  return filas.length + ' incidencias agregadas, ' + (REPORTE_JUL_OCT.length - filas.length) + ' ya estaban.';
}

/* fecha · descripción · sucursal · reporta · categoría · estado (reportado a Metrored: sin dato en este reporte) */
var REPORTE_JUL_OCT = [
  ["2026-07-07", "NO SE REALIZA LOS RX EN EL PRIMER AGENDAMIENTO EL CLIENTE SOLICITA AGENDAR AL SIGUIENTE DÍA TODO EL PERFIL COMPLETO", "QUITO", "KARYME MANTILLA", "Agendamiento", "Cerrada"],
  ["2026-07-09", "NO SE CARGA LA OPTOMETRÍA A TIEMPO", "GUAYAQUIL", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-07-14", "No se carga a tiempo la optometría", "QUITO", "KARYME MANTILLA", "Resultados", "Cerrada"],
  ["2026-07-15", "INGRESA SOLICITUD EL 13 DE JULIO, PERO POR DISPONIBILIDAD DE CENTRO MÉDICO CON ECOGRAFIAS SE ENVIARON 2 OPCIONES (14 Y 15), CLIENTE PUNTONET CONFIRMA ATENCIÓN PARA EL 15 DE JULIO.", "QUITO", "KARYME MANTILLA", "Agendamiento", "Cerrada"],
  ["2026-07-31", "reagendamiento/ OPTOMETRIA: no se realiza, no desea esperar a su turno", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-07-07", "Agendamiento de candidatos en GYE mandan a dirección UIO", "GUAYAQUIL", "KARYME MANTILLA", "Agendamiento", "Cerrada"],
  ["2026-07-07", "Candidatos a técnicos tienen que repetir su examen de RX", "QUITO", "JULY RUIZ", "Resultados", "Cerrada"],
  ["2026-07-09", "En la base no se encuentran los resultados de examene médicos candidato [NOMBRE RESERVADO], cuando suben se confunden y suben de otro candidato.", "QUITO", "KARYME MANTILLA", "Confidencialidad", "Cerrada"],
  ["2026-07-20", "No dan respuesta al pedido de examenes médicos", "QUITO", "KARYME MANTILLA", "Agendamiento", "Cerrada"],
  ["2026-07-21", "No dan respuesta al pedido de examenes médicos de los cortes, toca insistir con una llamada. Adicionalmente al momento de enviar la dirección se confunden de sucursal (candidatos de Manta envian con dirección a Quito)", "MANTA", "KARYME MANTILLA", "Agendamiento", "Cerrada"],
  ["2026-07-17", "Coordinador técnico informa que los colaboradores de Portoviejo no pudieron realizarse los RX a la hora programada por un daño de la máquina, se solventó con colaboradores y se reagendó a los 2 colaboradores faltantes", "PORTOVIEJO", "KARYME MANTILLA", "Atención en sucursal", "Cerrada"],
  ["2026-07-21", "Informan que ningún colaborador de la sucursal de Ambato asiste al examen agendado, se valida la información con Jefaturas técnicas de la sucursal y manifiesta que todos los técnicos asistieron, se pone en conocimiento dicha información, y efectivamente el laboratorio se confundió por lo que se valida que el examen si se realizó.", "AMBATO", "KARYME MANTILLA", "Resultados", "Cerrada"],
  ["2026-07-22", "Se informa que dos colaboradores de Ibarra no asistieron al examen, se valida la información con jefatura correspondiente e indican que todos asistieron,adicional, nos informan que les atendieron luego de una hora. Se comparte esta información y la respuesta del laboratorio es que no se acercaron los compañeros, se trata de comunicar con Luz sin obtener respuesta, se procede a llamar a Kevin, sin embargo, tampoco hay respuesta de su parte.", "IBARRA", "KARYME MANTILLA", "Resultados", "Cerrada"],
  ["2026-08-01", "NO SE ENTREGA RESULTADOS A TIEMPO", "GUAYAQUIL", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-08-06", "NO SE ENTREGA RESULTADOS A TIEMPO", "MANTA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-08-21", "CLIENTE INDICA QUE NO ESTÁN LOS RESULTADOS, SE ENVIA LA EVIDENCIA QUE SI ESTABAN CARGADOS", "CUENCA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-08-25", "CLIENTE INDICA QUE LA AUDIOMETRIA NO SE REALIZÓ CORRECTAMENTE. SOLICITAN VOLVER A REALIZAR EL PROCEDIMIENTO Y SE REALIZA NUEVAMENTE.", "CUENCA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-08-25", "CLIENTE INDICA QUE LA AUDIOMETRIA NO SE REALIZÓ CORRECTAMENTE. SOLICITAN VOLVER A REALIZAR EL PROCEDIMIENTO Y SE REALIZA NUEVAMENTE.", "CUENCA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-09-03", "SE REALIZA REAGENDAMIENTO DE AUDIOMETRIA, CLIENTE EXPRESA QUE EN EL 1ER AGENDAMIENTO NO SE HIZO UNA CORRECTA REVISIÓN AL MOMENTO DE REALIZAR EL PROCEDIMIENTO AUDIOMETRICO.", "CUENCA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-09-03", "SE REALIZA REAGENDAMIENTO DE AUDIOMETRIA, CLIENTE EXPRESA QUE EN EL 1ER AGENDAMIENTO NO SE HIZO UNA CORRECTA REVISIÓN AL MOMENTO DE REALIZAR EL PROCEDIMIENTO AUDIOMETRICO.", "CUENCA", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-09-03", "SE REALIZA REAGENDAMIENTO DE AUDIOMETRIA, CLIENTE EXPRESA QUE EN EL 1ER AGENDAMIENTO NO SE HIZO UNA CORRECTA REVISIÓN AL MOMENTO DE REALIZAR EL PROCEDIMIENTO AUDIOMETRICO.", "GUAYAQUIL", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-09-03", "SE REALIZA REAGENDAMIENTO DE AUDIOMETRIA, CLIENTE EXPRESA QUE EN EL 1ER AGENDAMIENTO NO SE HIZO UNA CORRECTA REVISIÓN AL MOMENTO DE REALIZAR EL PROCEDIMIENTO AUDIOMETRICO.", "GUAYAQUIL", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-09-02", "NO SE ENTREGA RESULTADOS A TIEMPO", "GUAYAQUIL", "MILENA ANDRADE", "Resultados", "Cerrada"],
  ["2026-08-31", "TENIA CITA EL 29AGO2026, PERO ÉL MISMO REAGENDA PARA EL 31AGO2026 EN CENTRO MÉDICO", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-09-29", "RAYOS X LUMBAR SIN REPORTE DE ANGULOS, SE ESCALA EL CASO AL AREA MÉDICA PARA DAR SOLUCIÓN LO ANTES POSIBLE", "QUITO", "SILVIA MERA", "Radiografía", "Cerrada"],
  ["2026-09-29", "RAYOS X LUMBAR SIN REPORTE DE ANGULOS, SE ESCALA EL CASO AL AREA MÉDICA PARA DAR SOLUCIÓN LO ANTES POSIBLE", "QUITO", "SILVIA MERA", "Radiografía", "Cerrada"],
  ["2026-10-02", "AUDIOMETRIA COORDINADA PARA EL LUNES, NO ASISTE AL RESTO DE EXAMENES", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-07", "Buenos días REPORTE DE NOVEDAD: 7 octubre 2026 [NOMBRE RESERVADO] , los nombres en sus examenes están al revés se reportó al personal de ayer. Hoy persiste el error en sus examenes de laboratorio. LOS DATOS ESTABAN MAL EN LA HOJA DE RUTA. HOY 8/10/2026 SE CORRIJE LOS NOMBRES EN LOS RESULTADOS DEL LABORATORIO", "QUITO", "SILVIA MERA", "Atención en sucursal", "Cerrada"],
  ["2026-10-07", "[NOMBRE RESERVADO] indica el paciente que también le colocaron los nombres al revés. Pero ya los han rectificado hoy. Optometría estaba agendado para las 7:00 y llegó 7:36 DATOS ERRONEOS EN EL NOMBRE EN LA HOJA DE RUTA, SE CORRIGIÓ DE MANERA INMEDIATA CON LA VALIDACION DE DATOS AL COLABORADOR. LOS RESULTADOS SALEN CON DATOS CORRECTOS. EL PACIENTE LLEGÓ A LAS 7:16 AM. SALIÓ A LAS 7:44 AM. LA HORA DE INICIO ERA A LAS 7:00 AM", "QUITO", "SILVIA MERA", "Atención en sucursal", "Cerrada"],
  ["2026-10-08", "En el agendamiento la paciente tenía tapon de cerúm, no se contaba con el profesional en ese momento para realizar el lavado de oidos, se comunicó al cliente y quién solicitó nos comuniquemos nosotros para el reagendamiento esta novedad se escaló y finalmente se estableció que el contacto con los colaboradores debe gestionar el cliente.", "GUAYAQUIL", "MILENA ANDRADE", "Agendamiento", "Cerrada"],
  ["2026-10-15", "LA AUDIOMETRIA SE CARGA A LAS 11:00 AM, PORQUE SE REALIZÓ EL DÍA ANTERIOR A LAS 11:30 AM", "QUITO", "SILVIA MERA", "Resultados", "Cerrada"],
  ["2026-10-16", "Cliente envia el requerimiento 15:40, fuera del tiempo acordado con Metrored, se realiza la coordinación.", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-16", "Cliente envia el requerimiento 15:40, fuera del tiempo acordado con Metrored, se realiza la coordinación.", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-16", "Cliente envia el requerimiento 15:40, fuera del tiempo acordado con Metrored, se realiza la coordinación.", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-16", "Cliente envia el requerimiento 15:40, fuera del tiempo acordado con Metrored, se realiza la coordinación.", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-16", "Cliente envia el requerimiento 15:40, fuera del tiempo acordado con Metrored, se realiza la coordinación.", "QUITO", "DIEGO VELA", "Agendamiento", "Cerrada"],
  ["2026-10-21", "CLIENTE ENVIA CI INCORRECTA PARA EL AGENDAMIENTO. EL CLIENTE REPORTA QUE NO ENCUENTRA LOS RESULTADOS, INFORMAMOS QUE LA CEDULA CON LA QUE BUSCABA NO ES LA CORRECTA. TAMBIEN REPORTA EL CLIENTE QUE NO SE VISUALIZA BIEN LAS IMAGENES, SE PROCEDIÓ A ENTREGAR UN LINK PARA QUE LO VISUALICE EN LINEA, LO CUAL ESTUVO CONFORME EL CLIENTE.", "LOJA", "MILENA ANDRADE", "Radiografía", "Cerrada"]
];


/* ───────── Carga inicial: ejecutar esta función una vez ───────── */

function cargarTodo() {
  var antes = '';
  if (hoja().getLastRow() < 2) { cargarHistorico(); antes = 'Histórico 2026 cargado. '; }
  return antes + cargarReporteJulOct();
}
