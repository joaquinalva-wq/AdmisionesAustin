/*************************************************************************
 * AUSTIN EBS — ADMISIONES · Apps Script
 *
 * Copia de referencia del script que vive en script.google.com.
 * Si se edita allá, actualizar también este archivo (y al revés).
 *
 * Pegar TODO el contenido de este archivo reemplazando el código actual,
 * y después crear una NUEVA VERSIÓN de la implementación existente.
 * Guardar no alcanza: la URL /exec sigue sirviendo la versión publicada.
 *************************************************************************/

/************************
 * CONFIGURACIÓN SIMPLE *
 ************************/
const SENDER_NAME   = 'Laura Unamuno (Admisiones Austin EBS)';
const SENDER_EMAIL  = 'admisiones@austinebs-ah.edu.ar'; // Ideal: alias configurado en Gmail
const CALENDAR_LINK = 'https://calendar.app.google/WrfYGGSiRH3smkxL8'; // Link público para reservar cita

// Logo del colegio para los mails. Gmail bloquea imágenes embebidas en base64,
// por eso se referencia por URL desde el sitio de admisiones.
const MAIL_LOGO = 'https://admisionesaustin.com.ar/logo-email.png';

// Marca de versión: sirve para confirmar que la implementación se publicó.
// Al abrir la URL del script con ?action=ping tiene que aparecer este valor.
const API_VERSION = '2026-08-22-v4';

// ── Calendario ────────────────────────────────────────────────
const CAL_ADMISIONES        = 'admisiones@austinebs-ah.edu.ar';
const LUGAR_ENTREVISTA      = 'Planta alta de High School — Austin EBS';
const DURACION_ENTREVISTA_MIN = 60;

// ── Mails automáticos por trigger ─────────────────────────────
// El panel de admisiones ya manda su propio mail cuando se mueve la ficha a
// "Entrevistado". Si además lo mandara el trigger, la familia recibiría el
// mismo mail dos veces. Por eso el post-entrevista automático queda apagado:
// la fuente de verdad es el panel. El recordatorio de 24hs sí queda prendido
// porque el panel no lo manda.
const RECORDATORIO_AUTOMATICO   = true;
const POST_ENTREVISTA_AUTOMATICO = false;

// ── Hojas (el script las crea solo si no existen) ─────────────
const SHEET_TURNOS   = 'Turnos';
const SHEET_BLOQUEOS = 'Bloqueos';

const COLS_TURNOS   = ['ID','Fecha','Hora','AlumnoId','Alumno','Padre','Email','Tel','Curso','Estado','Creado'];
const COLS_BLOQUEOS = ['Fecha','Hora','Motivo','Creado'];

/****************************************
 * MAPEO DE COLUMNAS (Sheet del Form)   *
 * A=1 ... D=4 ... I=9 ... O=15 ... P=16
 ****************************************/
const COL = {
  TIMESTAMP: 1, // "Marca temporal" (A)
  TELEFONO: 2,  // "Teléfono" (B)
  POSTULANTE_NOMBRE: 4, // "Nombre del postulante-Hijo/a" (D)
  PADRE_NOMBRE: 9, // "Nombre y apellido de Madre-Padre-Tutor" (I)
  EMAIL2: 15, // "Dirección de correo electrónico" (O) -> destinatario
  ESTADO_P: 16 // "Observaciones de la entrevista" (P) -> USADA COMO ESTADO
};

/*******************************************************************
 * FECHAS Y HORAS — el problema más importante que resuelve esta   *
 * versión.                                                        *
 *                                                                 *
 * appendRow() deja que Sheets interprete el texto: "2026-08-30" se *
 * guarda como fecha y "14:30" como hora. Al leerlas vuelven como   *
 * objetos Date, y el panel las compara contra texto plano. Ninguna *
 * comparación daba verdadera: los horarios ocupados no se marcaban *
 * (se podía reservar dos veces el mismo turno), los bloqueos no se *
 * respetaban y los dos triggers automáticos no disparaban nunca.   *
 *                                                                 *
 * Se arregla por los dos lados: al escribir se fuerza formato      *
 * texto, y al leer se normaliza igual. Lo segundo también repara   *
 * las filas viejas que ya quedaron guardadas como fecha.           *
 *******************************************************************/

/** Zona horaria de la planilla. Con esta se leen bien las fechas guardadas. */
function tz_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() ||
           'America/Argentina/Buenos_Aires';
  } catch (_) { return 'America/Argentina/Buenos_Aires'; }
}

/** Devuelve siempre "yyyy-MM-dd", venga texto o venga un Date de la planilla. */
function aFecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : s;
}

/** Devuelve siempre "HH:mm". El "*" (día completo bloqueado) se respeta. */
function aHora_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  const s = String(v == null ? '' : v).trim();
  if (s === '*' || s === '') return s;
  const m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : s;
}

/**
 * Agrega una fila forzando a texto las columnas indicadas, para que Sheets no
 * vuelva a convertir las fechas y horas en valores numéricos.
 */
function appendFilaTexto_(sh, fila, colsTexto) {
  const r = sh.getLastRow() + 1;
  (colsTexto || []).forEach(function (c) {
    sh.getRange(r, c).setNumberFormat('@');
  });
  sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  return r;
}

/***********************
 * UTILIDADES BÁSICAS  *
 ***********************/

/** Escapa texto que viene de las familias antes de meterlo en un mail HTML. */
function esc_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Envía un mail.
 * El cuarto parámetro es la versión en texto plano. Antes se mandaba siempre
 * un espacio ' ', y eso es una señal clásica de spam: los filtros esperan que
 * un mail HTML traiga también su alternativa en texto. Si no se pasa, se genera
 * automáticamente a partir del HTML.
 */
function sendEmail_(to, subject, htmlBody, plainBody) {
  const texto = (plainBody && String(plainBody).trim())
    ? String(plainBody)
    : htmlAtexto_(htmlBody);
  const opts = { name: SENDER_NAME, htmlBody: htmlBody };
  // Si el alias está configurado en Gmail, usarlo:
  if (SENDER_EMAIL && SENDER_EMAIL.trim()) opts.from = SENDER_EMAIL;
  GmailApp.sendEmail(to, subject, texto || ' ', opts);
}

/** Convierte HTML a texto plano legible, para la alternativa del mail. */
function htmlAtexto_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&mdash;/gi, '—')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// obtiene una fila completa (array) y también la hoja activa
function getRow_(row, hoja) {
  const sh = hoja || SpreadsheetApp.getActiveSheet();
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  return { sh, values };
}

// ayuda: toma destinatario preferente (col O). Si vacío, intenta columna H (8).
function getRecipientEmail_(rowValues) {
  const emailO = (rowValues[COL.EMAIL2 - 1] || '').toString().trim();
  if (emailO) return emailO;
  // fallback opcional si quisieras: columna 8 ("Dirección de correo electrónico" anterior)
  const emailH = (rowValues[8 - 1] || '').toString().trim();
  return emailH || '';
}

/**********************************************
 * 1) MAIL AUTOMÁTICO AL LLEGAR UNA RESPUESTA *
 **********************************************/
function onFormSubmit_SendWelcome(e) {
  // --- Guardas para evitar que se ejecute "a mano" o en contexto incorrecto ---
  if (!e || !e.range || !e.namedValues) return; // solo corre en envíos reales del Form

  // Tomamos exactamente la fila recién agregada por el Form
  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

  const padre  = (values[COL.PADRE_NOMBRE - 1] || '').toString().trim();
  const alumno = (values[COL.POSTULANTE_NOMBRE - 1] || '').toString().trim();
  const to     = getRecipientEmail_(values);
  if (!to) return;

  const subject = '¡Gracias por postular a Austin EBS!';
  const contenido = `
    <p style="margin:0 0 12px">Hola <strong>${esc_(padre) || 'familia'}</strong>,</p>
    <p style="margin:0 0 12px">Soy Laura Unamuno, parte del staff de <strong>Austin Eco Bilingual School</strong> y encargada del proceso de admisiones.</p>
    <p style="margin:0 0 12px">¡Nos alegra muchísimo tu postulación para <strong>${esc_(alumno) || 'su hijo/a'}</strong>! Para conocer el proyecto en detalle y las tarifas, podés ver el brochure y los aranceles acá:</p>
    <p style="margin:0 0 18px"><a href="https://drive.google.com/drive/folders/1c0vFf1MWUcSC8-lqBjBCGDQ-hbXxGQDA" style="color:#3A607A">Brochure y aranceles</a></p>
    <p style="margin:0 0 12px">Para empezar con el proceso, reservá un horario de visita con entrevista y conocé el proyecto educativo en profundidad y las instalaciones del colegio:</p>
    <p style="margin:0 0 20px">${boton_(CALENDAR_LINK, 'Reservar horario de entrevista')}</p>
    <p style="margin:0">Quedo a disposición por cualquier consulta.</p>
  `;
  sendEmail_(to, subject, mailBase_('¡Gracias por postular!', contenido, firma_()));
}

/***************************************************************
 * 2) MAILS AUTOMÁTICOS CUANDO SE EDITA LA COLUMNA P (ESTADO)  *
 ***************************************************************/
function onEdit_SendByStatus(e) {
  // Sin evento no hay nada que hacer: evita que reviente si se ejecuta a mano
  // desde el editor.
  if (!e || !e.range) return;

  // Solo reaccionar si se editó la columna P (16)
  const range = e.range;
  const sh = range.getSheet();
  if (range.getColumn() !== COL.ESTADO_P || range.getRow() === 1) return;

  const newValue = (e.value || '').toString().trim();
  const oldValue = (e.oldValue || '').toString().trim();
  // Si no cambió el valor, no hacer nada
  if (newValue === oldValue) return;

  const row = range.getRow();
  // Se lee de la hoja editada, no de la hoja activa: si había otra pestaña
  // abierta, antes tomaba los datos de la fila equivocada.
  const { values } = getRow_(row, sh);

  const padre = (values[COL.PADRE_NOMBRE - 1] || '').toString().trim();
  const alumno = (values[COL.POSTULANTE_NOMBRE - 1] || '').toString().trim();
  const to = getRecipientEmail_(values);
  if (!to) return;

  const estado = newValue.toLowerCase();
  const P = esc_(padre) || 'familia';
  const A = esc_(alumno) || 'su hijo/a';

  let subject = '';
  let titulo  = '';
  let cuerpo  = '';

  switch (estado) {
    case 'postulado':
      subject = 'Siguiente paso en el proceso de admisión';
      titulo  = 'Siguiente paso';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Gracias por la postulación de <strong>${A}</strong>.</p>
        <p style="margin:0 0 12px">Estamos muy contentos de acompañarlos en los pasos a seguir para continuar con la reserva de vacante. Podrán encontrar las tarifas acá:</p>
        <p style="margin:0 0 18px"><a href="https://drive.google.com/drive/folders/16zIE1tlrVfoOFLEZumgFkKYSV3US9SNp?usp=sharing" style="color:#3A607A">Ver tarifas</a></p>
        <p style="margin:0">El proceso quedará confirmado una vez que recibamos el comprobante de pago de la reserva de vacante por este medio.</p>
      `;
      break;

    case 'citado entrevista':
      subject = '¡Tu entrevista quedó agendada!';
      titulo  = 'Entrevista agendada';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Gracias por agendar la entrevista para conocer Austin EBS. Será un gusto recibirlos en la fecha y hora seleccionada.</p>
        <p style="margin:0 0 12px">Si necesitás modificar el horario, podés hacerlo desde acá:</p>
        <p style="margin:0 0 20px">${boton_(CALENDAR_LINK, 'Ver o modificar el horario')}</p>
        <p style="margin:0">¡Nos vemos pronto!</p>
      `;
      break;

    case 'entrevistado':
      subject = '¡Gracias por la entrevista!';
      titulo  = 'Gracias por la entrevista';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Fue un gusto recibirlos y conocerlos. Estamos orgullosos de que nos contemplen para la elección escolar de <strong>${A}</strong>.</p>
        <p style="margin:0 0 12px">Para seguir acompañándolos en el proceso, les pedimos que completen este breve formulario post entrevista:</p>
        <p style="margin:0 0 20px">${boton_('https://forms.gle/rsQwi9ps1chN2q8x5', 'Completar formulario')}</p>
        <p style="margin:0">En caso de querer avanzar, esperamos su confirmación por este medio.</p>
      `;
      break;

    case 'desiste':
      subject = 'Proceso de admisión: baja solicitada';
      titulo  = 'Proceso de admisión';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Tomamos nota de que no continuarán con la postulación de ${A}. Gracias por habernos considerado.</p>
        <p style="margin:0">Si desean retomar más adelante, cuenten con nosotros.</p>
      `;
      break;

    case 'lista de espera':
      subject = 'Actualización: lista de espera';
      titulo  = 'Lista de espera';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Por el momento, la postulación de ${A} se encuentra en <strong>lista de espera</strong>. Les avisaremos ante cualquier novedad.</p>
        <p style="margin:0">¡Muchas gracias! Quedamos a disposición.</p>
      `;
      break;

    case 'matriculado':
      subject = '¡Bienvenidos a la familia Austin EBS!';
      titulo  = '¡Bienvenidos a Austin EBS!';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Recibimos el comprobante de pago y <strong>${A}</strong> ya se encuentra <strong>matriculado/a</strong>. ¡Bienvenidos a la comunidad Austin!</p>
        <p style="margin:0">Estamos para lo que necesiten.</p>
      `;
      break;

    case 'no queremos':
      subject = 'Gracias por haber considerado Austin EBS';
      titulo  = 'Proceso de admisión';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Muchas gracias por el interés y por haber considerado a <strong>Austin EBS</strong> para la educación de ${A}.</p>
        <p style="margin:0 0 12px">En esta oportunidad no contamos con vacantes disponibles para continuar con el proceso de admisión.</p>
        <p style="margin:0">Agradecemos sinceramente su tiempo y les deseamos lo mejor en su elección escolar.</p>
      `;
      break;

    case 'en proceso':
      subject = 'Seguimos acompañándolos en el proceso de admisión';
      titulo  = 'Seguimos en contacto';
      cuerpo = `
        <p style="margin:0 0 12px">Hola <strong>${P}</strong>,</p>
        <p style="margin:0 0 12px">Muchas gracias por seguir considerando a <strong>Austin EBS</strong> para la educación de ${A}.</p>
        <p style="margin:0 0 12px">Entendemos que se encuentran evaluando y avanzando en el proceso de postulación. Quedamos a disposición para acompañarlos y responder cualquier consulta.</p>
        <p style="margin:0">Será un gusto seguir en contacto.</p>
      `;
      break;

    default:
      // Si escriben otra cosa en P, no enviamos nada.
      return;
  }

  sendEmail_(to, subject, mailBase_(titulo, cuerpo, firma_()));
}

// ── PUNTO DE ENTRADA HTTP ──────────────────────────────────────
function doGet(e)  { return handleAPI_(e); }
function doPost(e) { return handleAPI_(e); }

function handleAPI_(e) {
  const params = (e && e.parameter) || {};

  // Leer datos: primero intenta POST body (JSON o texto plano), luego GET params
  let body = {};
  if (e && e.postData && e.postData.contents) {
    try { body = JSON.parse(e.postData.contents); } catch(_) {
      try { body = JSON.parse(e.postData.contents.trim()); } catch(_2) {}
    }
  }

  const action = body.action || params.action || '';
  const data   = body.data   || (params.data ? JSON.parse(params.data) : params);

  let result = { ok: false, error: 'Acción no reconocida: ' + action };

  try {
    switch (action) {
      case 'ping':          result = { ok: true, msg: 'Austin EBS API activa', version: API_VERSION }; break;
      case 'getTurnos':     result = getTurnos_();   break;
      case 'reservarTurno': result = reservarTurno_(data); break;
      case 'cancelarTurno': result = cancelarTurno_(data); break;
      case 'getBloqueos':   result = getBloqueos_(); break;
      case 'setBloqueo':    result = setBloqueo_(data); break;
      case 'sendMail':      result = sendMailDirect_(data); break;
      case 'crearEventoEntrevista': result = crearEventoEntrevista(data); break;
    }
  } catch(err) {
    result = { ok: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TURNOS ─────────────────────────────────────────────────────
function getTurnos_() {
  const sh = getOrCreate_(SHEET_TURNOS, COLS_TURNOS);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const turnos = rows.slice(1)
    .map(r => { const o = {}; headers.forEach((h,i) => o[h] = r[i]); return o; })
    .filter(t => t.ID)
    .map(t => {
      // El panel compara contra "yyyy-MM-dd" y "HH:mm". Se normaliza acá para
      // que también funcionen las filas viejas guardadas como fecha real.
      t.Fecha = aFecha_(t.Fecha);
      t.Hora  = aHora_(t.Hora);
      return t;
    });
  return { ok: true, turnos };
}

function reservarTurno_(data) {
  const sh = getOrCreate_(SHEET_TURNOS, COLS_TURNOS);

  const fecha = aFecha_(data.fecha);
  const hora  = aHora_(data.hora);
  if (!fecha || !hora) return { ok: false, error: 'Falta fecha u hora' };

  // Verificar que el slot no esté ya tomado
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const iEstado = headers.indexOf('Estado');
  const iFecha  = headers.indexOf('Fecha');
  const iHora   = headers.indexOf('Hora');
  for (let i = 1; i < rows.length; i++) {
    const estado = String(rows[i][iEstado] || '');
    const tomado = estado === 'reservado' || estado === 'recordatorio_enviado';
    if (aFecha_(rows[i][iFecha]) === fecha && aHora_(rows[i][iHora]) === hora && tomado) {
      return { ok: false, error: 'Este horario ya fue reservado. Por favor elegí otro.' };
    }
  }

  const id = 'T' + Date.now();
  appendFilaTexto_(sh, [
    id, fecha, hora, data.alumnoId || '',
    data.alumno, data.padre, data.email, data.tel || '',
    data.curso, 'reservado', new Date().toISOString()
  ], [2, 3]); // Fecha y Hora como texto

  // Mail de confirmación a la familia
  if (data.email) {
    const fechaLinda = formatFechaAPI_(fecha);
    sendEmail_(data.email,
      'Tu entrevista en Austin EBS está confirmada',
      mailBase_(
        'Entrevista confirmada',
        `<p style="margin:0 0 8px">Hola <strong>${esc_(data.padre)}</strong>,</p>
         <p style="margin:0 0 16px">Tu entrevista en <strong>Austin Eco Bilingual School</strong> está confirmada. Te esperamos con gusto.</p>
         ${datosEntrevista_(fechaLinda, hora)}
         <p style="margin:8px 0">Estudiante: <strong>${esc_(data.alumno)}</strong> &mdash; Curso: <strong>${esc_(data.curso)}</strong></p>
         <p style="margin:16px 0 0;color:#5A5A54;font-size:13px">Vas a recibir un recordatorio 24 horas antes. Si necesitás reprogramar, respondé este mail.</p>`,
        firma_()
      )
    );
  }

  // Notificar a admisiones
  GmailApp.sendEmail(
    SENDER_EMAIL,
    `Nueva entrevista agendada: ${data.alumno} — ${fecha} ${hora}`,
    `Estudiante: ${data.alumno}\nFamilia: ${data.padre}\nEmail: ${data.email}\nTel: ${data.tel||'—'}\nCurso: ${data.curso}\nFecha: ${fecha} ${hora}`,
    { name: 'Sistema Admisiones Austin EBS' }
  );

  // Evento en el calendario de Admisiones. Usa la misma función que el panel,
  // así una entrevista tiene un solo evento sin importar por dónde se cargó.
  crearEventoEntrevista({
    alumno: data.alumno, curso: data.curso, padre: data.padre,
    email: data.email, tel: data.tel, alumnoId: data.alumnoId,
    fecha: fecha, hora: hora
  });

  return { ok: true, turnoId: id };
}

function cancelarTurno_(data) {
  const sh = getOrCreate_(SHEET_TURNOS, COLS_TURNOS);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const iEstado = headers.indexOf('Estado');
  const iFecha  = headers.indexOf('Fecha');
  const iHora   = headers.indexOf('Hora');
  const iEmail  = headers.indexOf('Email');
  const iPadre  = headers.indexOf('Padre');
  const iAlumno = headers.indexOf('Alumno');

  const fecha = aFecha_(data.fecha);
  const hora  = aHora_(data.hora);

  for (let i = 1; i < rows.length; i++) {
    if (aFecha_(rows[i][iFecha]) === fecha && aHora_(rows[i][iHora]) === hora) {
      const nuevoEstado = data.nuevoEstado || 'cancelado';
      sh.getRange(i + 1, iEstado + 1).setValue(nuevoEstado);

      // El evento no se borra: se renombra, así queda el registro de que estaba
      // agendada y de que no se concretó.
      marcarEventoCancelado_(rows[i][iAlumno], nuevoEstado);

      // Si no asistió → mail de reprogramación
      if (nuevoEstado === 'reprogramar' && rows[i][iEmail]) {
        sendEmail_(rows[i][iEmail],
          '¡Qué lástima! Podemos reprogramar tu entrevista en Austin EBS',
          mailBase_(
            'Te esperábamos hoy',
            `<p style="margin:0 0 8px">Hola <strong>${esc_(rows[i][iPadre])}</strong>,</p>
             <p style="margin:0 0 16px">Lamentamos que no hayas podido llegar a la entrevista de hoy. Nos hubiera encantado conocerlos.</p>
             <p style="margin:0 0 16px">Si querés reprogramarla, respondé este mail y coordinamos un nuevo horario con gusto.</p>
             <p style="margin:0;color:#5A5A54;font-size:13px">Seguimos a tu disposición cuando quieras.</p>`,
            firma_()
          )
        );
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Turno no encontrado' };
}

// ── BLOQUEOS ───────────────────────────────────────────────────
function getBloqueos_() {
  const sh = getOrCreate_(SHEET_BLOQUEOS, COLS_BLOQUEOS);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const bloqueos = rows.slice(1)
    .map(r => { const o = {}; headers.forEach((h,i) => o[h] = r[i]); return o; })
    .filter(b => b.Fecha)
    .map(b => {
      b.Fecha = aFecha_(b.Fecha);
      b.Hora  = aHora_(b.Hora);
      return b;
    });
  return { ok: true, bloqueos };
}

function setBloqueo_(data) {
  const sh = getOrCreate_(SHEET_BLOQUEOS, COLS_BLOQUEOS);
  const fecha = aFecha_(data.fecha);
  const hora  = data.hora ? aHora_(data.hora) : '*';

  if (data.eliminar) {
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (aFecha_(rows[i][0]) === fecha &&
          (hora === '*' || aHora_(rows[i][1]) === hora)) {
        sh.deleteRow(i + 1);
      }
    }
    return { ok: true };
  }

  // No duplicar si ya estaba bloqueado
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (aFecha_(rows[i][0]) === fecha && aHora_(rows[i][1]) === hora) return { ok: true };
  }

  appendFilaTexto_(sh, [fecha, hora, data.motivo || '', new Date().toISOString()], [1, 2]);
  return { ok: true };
}

// ── CALENDARIO DE ADMISIONES ───────────────────────────────────

// Guarda el motivo real por el que no se pudo abrir el calendario. Sin esto, el
// error quedaba tapado por el try/catch y no se podia distinguir "falta el
// permiso de Calendar" de "el calendario no esta compartido".
let _calError = '';

/** Abre el calendario de Admisiones. Si no lo encuentra, usa el del script. */
function calAdmisiones_() {
  _calError = '';
  let cal = null;
  try { cal = CalendarApp.getCalendarById(CAL_ADMISIONES); }
  catch (e) { _calError = String(e); }
  if (!cal) {
    try { cal = CalendarApp.getDefaultCalendar(); }
    catch (e) { if (!_calError) _calError = String(e); }
  }
  if (!cal && !_calError) {
    _calError = 'El calendario ' + CAL_ADMISIONES +
                ' no existe o no esta compartido con la cuenta que ejecuta el script.';
  }
  return cal;
}

/**
 * Marca invisible que identifica al alumno dentro del evento.
 * Sirve para que, si se cambia la fecha de la entrevista, se mueva el evento
 * que ya existe en vez de crear uno nuevo al lado.
 */
function marcaAlumno_(nombre) {
  const s = String(nombre || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  return s ? 'AEBS[' + s + ']' : '';
}

/** Busca el evento de ese alumno en una ventana amplia alrededor de hoy. */
function buscarEventoAlumno_(cal, marca) {
  if (!cal || !marca) return null;
  const hoy = new Date();
  const desde = new Date(hoy.getTime() - 120 * 24 * 3600 * 1000);
  const hasta = new Date(hoy.getTime() + 500 * 24 * 3600 * 1000);
  let evs = [];
  try { evs = cal.getEvents(desde, hasta, { search: 'AEBS' }); } catch (_) { return null; }
  for (let i = 0; i < evs.length; i++) {
    try {
      if (String(evs[i].getDescription() || '').indexOf(marca) !== -1) return evs[i];
    } catch (_) {}
  }
  return null;
}

/**
 * Crea (o mueve, si ya existía) el evento de la entrevista en el calendario de
 * Admisiones. La llama el panel al agendar a mano y también reservarTurno_.
 *
 * Sin guion bajo al final a propósito: así aparece en el desplegable "Ejecutar"
 * del editor, que es como se autoriza el permiso de Calendar la primera vez.
 */
function crearEventoEntrevista(d) {
  try {
    if (!d || !d.fecha || !d.hora) return { ok: false, error: 'Falta fecha u hora' };

    const cal = calAdmisiones_();
    if (!cal) return { ok: false, error: 'No se pudo abrir el calendario de Admisiones: ' + _calError };

    const fecha = aFecha_(d.fecha);
    const hora  = aHora_(d.hora);
    const f  = fecha.split('-');
    const hm = hora.split(':');
    const inicio = new Date(Number(f[0]), Number(f[1]) - 1, Number(f[2]), Number(hm[0]), Number(hm[1]), 0);
    if (isNaN(inicio.getTime())) return { ok: false, error: 'Fecha u hora inválida: ' + fecha + ' ' + hora };
    const fin = new Date(inicio.getTime() + DURACION_ENTREVISTA_MIN * 60000);

    const marca  = marcaAlumno_(d.alumno);
    const titulo = 'ENTREVISTA: ' + (d.alumno || 'Sin nombre') + (d.curso ? ' (' + d.curso + ')' : '');
    const lugar  = d.lugar || LUGAR_ENTREVISTA;

    const detalle = [
      'Familia: '  + (d.padre || '—'),
      'Email: '    + (d.email || '—'),
      'Telefono: ' + (d.tel   || '—'),
      d.curso ? 'Curso: ' + d.curso : '',
      d.nivel ? 'Nivel: ' + d.nivel : '',
      d.anio  ? 'Ciclo lectivo: ' + d.anio : '',
      d.alumnoId ? 'Ficha N° ' + d.alumnoId : '',
      '',
      'Evento del Sistema de Admisiones. No borrar la linea de abajo:',
      marca
    ].filter(function (x) { return x !== ''; }).join('\n');

    // Si el alumno ya tenía evento, se mueve en vez de duplicarlo.
    const previo = buscarEventoAlumno_(cal, marca);
    if (previo) {
      previo.setTitle(titulo);
      previo.setTime(inicio, fin);
      previo.setDescription(detalle);
      previo.setLocation(lugar);
      return { ok: true, eventId: previo.getId(), actualizado: true };
    }

    const ev = cal.createEvent(titulo, inicio, fin, { description: detalle, location: lugar });
    try { ev.addPopupReminder(60); } catch (_) {}
    return { ok: true, eventId: ev.getId(), actualizado: false };

  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/** Renombra el evento cuando la entrevista se cancela o hay que reprogramarla. */
function marcarEventoCancelado_(alumno, estado) {
  try {
    const cal = calAdmisiones_();
    const ev = buscarEventoAlumno_(cal, marcaAlumno_(alumno));
    if (!ev) return;
    const prefijo = estado === 'reprogramar' ? 'NO ASISTIÓ — ' : 'CANCELADA — ';
    const t = ev.getTitle();
    if (t.indexOf(prefijo) !== 0) ev.setTitle(prefijo + t);
  } catch (_) {}
}

// ── MAIL DIRECTO DESDE EL PANEL DE ADMISIONES ──────────────────
function sendMailDirect_(data) {
  if (!data.to || !data.subject || !data.body) {
    return { ok: false, error: 'Faltan campos: to, subject, body' };
  }

  // El panel arma el mail ya diseñado y lo manda en 'html'. Si viene, se usa tal
  // cual y se aprovecha 'body' como alternativa en texto plano.
  if (data.html) {
    sendEmail_(data.to, data.subject, data.html, data.body);
    return { ok: true };
  }

  // Si no viene diseñado (por compatibilidad), se arma acá a partir del texto.
  const htmlBody = mailBase_(
    data.subject,
    data.body
      .replace(/\n\n/g, '</p><p style="margin:0 0 12px">')
      .replace(/\n/g, '<br>')
      .replace(/^/, '<p style="margin:0 0 12px">')
      .replace(/$/, '</p>'),
    firma_()
  );
  sendEmail_(data.to, data.subject, htmlBody, data.body);
  return { ok: true };
}

// ── TRIGGERS AUTOMÁTICOS ───────────────────────────────────────

/**
 * Recordatorio 24hs antes.
 * Ejecutar una vez: instalarTriggers_API()
 * Se corre todos los días a las 8am.
 */
function enviarRecordatorios_() {
  if (!RECORDATORIO_AUTOMATICO) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TURNOS);
  if (!sh) return;
  const manana = new Date();
  manana.setDate(manana.getDate() + 1);
  const fechaManana = Utilities.formatDate(manana, tz_(), 'yyyy-MM-dd');

  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idx = h => headers.indexOf(h);

  for (let i = 1; i < rows.length; i++) {
    const fecha  = aFecha_(rows[i][idx('Fecha')]);
    const estado = String(rows[i][idx('Estado')]);
    if (fecha === fechaManana && estado === 'reservado') {
      const email  = String(rows[i][idx('Email')]);
      const padre  = String(rows[i][idx('Padre')]);
      const hora   = aHora_(rows[i][idx('Hora')]);
      if (email) {
        sendEmail_(email,
          'Recordatorio: tu entrevista en Austin EBS es mañana',
          mailBase_(
            'Tu entrevista es mañana',
            `<p style="margin:0 0 8px">Hola <strong>${esc_(padre)}</strong>,</p>
             <p style="margin:0 0 16px">Te recordamos que mañana tenés tu entrevista en Austin EBS. ¡Te esperamos!</p>
             ${datosEntrevista_(formatFechaAPI_(fecha), hora)}
             <p style="margin:16px 0 0;color:#5A5A54;font-size:13px">Si necesitás reprogramar, respondé este mail antes de esta noche.</p>`,
            firma_()
          )
        );
        sh.getRange(i + 1, idx('Estado') + 1).setValue('recordatorio_enviado');
      }
    }
  }
}

/**
 * Mail post-entrevista automático.
 * Corre cada 30 min. Detecta entrevistas que ya pasaron.
 *
 * Apagado por defecto (POST_ENTREVISTA_AUTOMATICO): el panel manda este mismo
 * mail al mover la ficha a "Entrevistado", y con los dos prendidos la familia
 * lo recibiría dos veces. Igual marca el turno como pasado.
 */
function verificarPostEntrevista_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TURNOS);
  if (!sh) return;
  const ahora   = new Date();
  const hoyStr  = Utilities.formatDate(ahora, tz_(), 'yyyy-MM-dd');
  const horaMin = ahora.getHours() * 60 + ahora.getMinutes();

  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idx = h => headers.indexOf(h);

  for (let i = 1; i < rows.length; i++) {
    const fecha  = aFecha_(rows[i][idx('Fecha')]);
    const hora   = aHora_(rows[i][idx('Hora')]);
    const estado = String(rows[i][idx('Estado')]);
    if (estado !== 'reservado' && estado !== 'recordatorio_enviado') continue;

    const partes = hora.split(':');
    const hh = Number(partes[0]), mm = Number(partes[1]);
    if (isNaN(hh) || isNaN(mm)) continue;
    const yaPaso = fecha < hoyStr ||
      (fecha === hoyStr && horaMin >= (hh * 60 + mm + 60)); // 1h después

    if (yaPaso) {
      const email  = String(rows[i][idx('Email')]);
      const padre  = String(rows[i][idx('Padre')]);
      const alumno = String(rows[i][idx('Alumno')]);
      if (POST_ENTREVISTA_AUTOMATICO && email) {
        sendEmail_(email,
          'Gracias por la entrevista — Austin EBS',
          mailBase_(
            'Gracias por la entrevista',
            `<p style="margin:0 0 8px">Hola <strong>${esc_(padre)}</strong>,</p>
             <p style="margin:0 0 16px">Fue un gusto recibirlos y conocerlos. Estamos orgullosos de que nos contemplen para la elección escolar de <strong>${esc_(alumno)}</strong>.</p>
             <p style="margin:0 0 12px">Para continuar con el proceso, les pedimos que completen este breve formulario:</p>
             <p style="margin:0 0 20px">${boton_('https://forms.gle/rsQwi9ps1chN2q8x5', 'Completar formulario post-entrevista')}</p>
             <p style="margin:0;color:#5A5A54;font-size:13px">En caso de querer avanzar, esperamos su confirmación por este medio.</p>`,
            firma_()
          )
        );
      }
      // El turno se marca como pasado igual, así el horario deja de figurar
      // ocupado en el calendario de las familias.
      sh.getRange(i + 1, idx('Estado') + 1).setValue('post_entrevista_enviado');
    }
  }
}

/**
 * Instalar los triggers automáticos.
 * EJECUTAR UNA SOLA VEZ manualmente.
 */
function instalarTriggers_API() {
  // Solo agregar los nuevos — no toca los triggers del Form que ya existen
  const triggers = ScriptApp.getProjectTriggers();
  const nombres = triggers.map(t => t.getHandlerFunction());

  if (!nombres.includes('enviarRecordatorios_')) {
    ScriptApp.newTrigger('enviarRecordatorios_')
      .timeBased().everyDays(1).atHour(8).create();
  }
  if (!nombres.includes('verificarPostEntrevista_')) {
    ScriptApp.newTrigger('verificarPostEntrevista_')
      .timeBased().everyMinutes(30).create();
  }
  Logger.log('Triggers API instalados. Triggers existentes preservados.');
}

/**
 * Chequeo rápido para ejecutar desde el editor (menú Ejecutar).
 * Verifica el acceso al calendario y que las fechas se lean bien.
 * No manda mails ni crea eventos.
 */
function probarSistema() {
  const out = [];
  out.push('Zona horaria de la planilla: ' + tz_());

  const cal = calAdmisiones_();
  out.push('Calendario: ' + (cal ? cal.getName() + ' (' + cal.getId() + ')'
                                 : 'NO ACCESIBLE -> ' + _calError));
  out.push('Ejecuta como: ' + (function(){ try { return Session.getEffectiveUser().getEmail(); } catch (e) { return '?'; } })());

  const t = getTurnos_();
  out.push('Turnos leidos: ' + (t.turnos ? t.turnos.length : 0));
  (t.turnos || []).slice(-3).forEach(function (x) {
    out.push('  ' + x.Fecha + ' ' + x.Hora + '  [' + x.Estado + ']');
  });

  const b = getBloqueos_();
  out.push('Bloqueos leidos: ' + (b.bloqueos ? b.bloqueos.length : 0));
  (b.bloqueos || []).slice(-3).forEach(function (x) {
    out.push('  ' + x.Fecha + ' ' + x.Hora);
  });

  out.push('Las fechas tienen que verse como 2026-08-30 y las horas como 14:30.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

// ── HELPERS ────────────────────────────────────────────────────
function getOrCreate_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#2C4A5E')
      .setFontColor('#ffffff');
    // Fecha y Hora como texto desde el arranque.
    const cFecha = headers.indexOf('Fecha') + 1;
    const cHora  = headers.indexOf('Hora') + 1;
    if (cFecha) sh.getRange(2, cFecha, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    if (cHora)  sh.getRange(2, cHora,  sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sh;
}

/** Pie de firma común a todos los mails. */
function firma_() {
  return 'Laura Unamuno &mdash; Admisiones Austin EBS &mdash; ' +
         '<a href="mailto:admisiones@austinebs-ah.edu.ar" style="color:#3A607A">admisiones@austinebs-ah.edu.ar</a>';
}

/** Botón de acción. */
function boton_(url, texto) {
  return '<a href="' + url + '" style="background:#E8621A;color:#ffffff;padding:12px 24px;' +
         'border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;display:inline-block">' +
         texto + '</a>';
}

// ── TEMPLATE BASE DE MAIL ──────────────────────────────────────
function mailBase_(titulo, contenido, footer) {
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#EEE9DF;font-family:Arial,Helvetica,sans-serif;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEE9DF;padding:24px 12px">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #E3DED3">
  <!-- Encabezado -->
  <tr><td align="center" style="background:#F8F6F0;padding:26px 32px 20px;border-bottom:3px solid #E8621A">
    <img src="${MAIL_LOGO}" width="210" alt="Austin Eco Bilingual School" style="display:block;border:0;width:210px;max-width:74%;height:auto;margin:0 auto 10px">
    <div style="font:400 11.5px Arial,Helvetica,sans-serif;color:#95958E;letter-spacing:.09em;text-transform:uppercase">Admisiones</div>
  </td></tr>
  <!-- Título -->
  <tr><td style="padding:22px 32px 0">
    <div style="font:700 19px Georgia,serif;color:#2C4A5E">${titulo}</div>
  </td></tr>
  <!-- Contenido -->
  <tr><td style="padding:16px 32px 26px;color:#252520;font-size:15px;line-height:1.7;word-break:break-word">${contenido}</td></tr>
  <!-- Pie -->
  <tr><td style="background:#F8F6F0;padding:18px 32px;border-top:1px solid #E3DED3">
    <p style="margin:0;font-size:12px;line-height:1.7;color:#95958E">${footer || firma_()}</p>
  </td></tr>
</table>
</td></tr></table>
</body></html>`;
}

function datosEntrevista_(fecha, hora) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#EEE9DF;border-radius:8px;margin:20px 0">
    <tr><td style="padding:18px 20px">
      <table role="presentation" cellpadding="4" cellspacing="0">
        <tr><td style="color:#5A5A54;font-size:12px;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;padding-right:16px">Fecha</td><td style="color:#2C4A5E;font-size:14px;font-weight:bold">${fecha}</td></tr>
        <tr><td style="color:#5A5A54;font-size:12px;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;padding-right:16px">Hora</td><td style="color:#2C4A5E;font-size:14px;font-weight:bold">${hora}</td></tr>
        <tr><td style="color:#5A5A54;font-size:12px;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;padding-right:16px">Lugar</td><td style="color:#2C4A5E;font-size:14px;font-weight:bold">Planta alta de High School</td></tr>
      </table>
    </td></tr>
  </table>`;
}

function formatFechaAPI_(fechaStr) {
  try {
    const partes = String(fechaStr).split('-');
    const y = partes[0], m = partes[1], d = partes[2];
    const dias  = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
    const meses = ['enero','febrero','marzo','abril','mayo','junio','julio',
                   'agosto','septiembre','octubre','noviembre','diciembre'];
    const dt = new Date(Number(y), Number(m) - 1, Number(d));
    return `${dias[dt.getDay()]} ${d} de ${meses[Number(m)-1]} de ${y}`;
  } catch(_) { return fechaStr; }
}
