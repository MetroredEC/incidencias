/* Puntonet · Carga del reporte actualizado (julio a 21 de octubre de 2026)
 *
 * Cómo usarlo, sin tocar CLAVE ni el script principal:
 *  1. En el editor de Apps Script: + (junto a "Archivos") → Secuencia de comandos → nombre "Reporte".
 *  2. Pegar este archivo completo.
 *  3. Elegir la función cargarReporteJulOct y pulsar Ejecutar. Se puede ejecutar más de una vez:
 *     las filas ya cargadas (ids r1…rN) se omiten.
 *
 * Usa hoja() y COLS de incidencias.gs (comparten el mismo proyecto).
 * Los nombres de pacientes se reservan porque este archivo vive en un repositorio público.
 */
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
