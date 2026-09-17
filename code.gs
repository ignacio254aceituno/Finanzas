/**
 * ZYNTO — Backend de Google Apps Script
 * ---------------------------------------------------------------
 * Maneja: lectura/escritura de Google Sheets + integración con Fintoc
 * (conexión bancaria PERSISTENTE: el widget solo se abre una vez por
 * banco nuevo; después, fintocSync reutiliza el link_token guardado
 * sin volver a pedir permiso).
 *
 * ── INSTALACIÓN ───────────────────────────────────────────────
 * 1) Extensiones > Apps Script en tu Google Sheet (script "atado" a la hoja).
 * 2) Pega este archivo completo reemplazando el Code.gs existente.
 * 3) Ejecuta el menú Project Settings (⚙️) > Script Properties y agrega:
 *      APP_SECRET          = Cacaflow123$   (debe ser IDÉNTICA a APP_SECRET en index.html)
 *      FINTOC_SECRET_KEY   = sk_live_xxxxxxxxxxxx   (Dashboard Fintoc > API Keys)
 *      FINTOC_PUBLIC_KEY   = pk_live_xxxxxxxxxxxx
 * 4) Implementar > Nueva implementación > Aplicación web:
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 *    Copia la URL /exec resultante y pégala en API_URL dentro de index.html.
 * 5) Cada vez que edites este script, vuelve a "Implementar > Gestionar
 *    implementaciones > ✏️ > Nueva versión > Implementar". Si no subes
 *    versión nueva, el /exec sigue sirviendo el código VIEJO (causa
 *    típica del error que se ve como "CORS").
 * -----------------------------------------------------------------
 */

// ── CONFIG ─────────────────────────────────────────────────────
function cfg_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}
const APP_SECRET_        = () => cfg_('APP_SECRET');
const FINTOC_SECRET_KEY_ = () => cfg_('FINTOC_SECRET_KEY');
const FINTOC_PUBLIC_KEY_ = () => cfg_('FINTOC_PUBLIC_KEY');
const FINTOC_API_        = 'https://api.fintoc.com/v1';

const SS_ = () => SpreadsheetApp.getActiveSpreadsheet();

// Hojas de control internas (no se muestran en la app, solo las usa el backend)
const SHEET_LINKS  = 'FintocLinks';   // link_token guardados por banco conectado
// Cada usuario (identificado por su correo, sanitizado en el frontend como "ns")
// tiene sus propias hojas: FintocLinks__ns, Cuentas__ns, Registros__ns, etc.
function nsName_(base, ns) { return ns ? (base + '__' + ns) : base; }
const SHEET_RESET  = 'ResetCodes';    // códigos de recuperación de contraseña
const SHEET_USERS  = 'Usuarios';      // cuentas: correo, nombre, usuario, hash+salt de contraseña (NUNCA texto plano)

// ── UTILIDADES DE RESPUESTA (siempre JSON válido, nunca una página de error) ──
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok_(extra)  { return json_(Object.assign({ status: 'ok' }, extra || {})); }
function err_(msg)   { return json_({ status: 'error', message: String(msg || 'Error desconocido') }); }

// doOptions existe por si algún día el frontend agrega headers custom
// (hoy no hace falta: los requests son "simple request" y no disparan preflight).
function doOptions() { return ContentService.createTextOutput(''); }

// ── ENTRADA GET: lectura de una hoja (?sheet=NombreHoja&secret=...) ──
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.secret !== APP_SECRET_()) return err_('secret inválido');
    if (!p.sheet) return err_('falta parámetro sheet');
    const rows = readSheet_(p.sheet);
    // El frontend espera el arreglo de filas DIRECTAMENTE (no envuelto en {status:...})
    return json_(rows);
  } catch (ex) {
    return err_(ex.message);
  }
}

// ── ENTRADA POST: todas las acciones de la app ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.secret !== APP_SECRET_()) return err_('secret inválido');

    switch (body.action) {
      case 'save':                 return actionSave_(body);
      case 'registerAccount':        return actionRegisterAccount_(body);
      case 'loginAccount':           return actionLoginAccount_(body);
      case 'changePasswordAccount':  return actionChangePasswordAccount_(body);
      case 'fintocCreateLinkIntent': return actionFintocCreateLinkIntent_();
      case 'fintocSaveLink':         return actionFintocSaveLink_(body);
      case 'fintocSync':             return actionFintocSync_(body);
      case 'sendResetCode':          return actionSendResetCode_(body);
      case 'verifyResetCode':        return actionVerifyResetCode_(body);
      default: return err_('acción no reconocida: ' + body.action);
    }
  } catch (ex) {
    // Nunca dejar que una excepción se escape sin JSON: eso es lo que
    // el navegador termina mostrando como error de "CORS".
    return err_(ex.message);
  }
}

// ── SHEETS genéricos ──────────────────────────────────────────
function ensureSheet_(name) {
  const ss = SS_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function readSheet_(name) {
  const sh = SS_().getSheetByName(name);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  return values;
}
function writeSheet_(name, rows) {
  const sh = ensureSheet_(name);
  sh.clearContents();
  if (rows && rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}
function actionSave_(body) {
  if (!body.sheet || !body.data) return err_('faltan sheet o data');
  writeSheet_(body.sheet, body.data);
  return ok_();
}

// ── FINTOC: helper genérico para llamar la API con la Secret Key ──
function extractFintocError_(data, code) {
  if (!data) return 'Fintoc respondió ' + code;
  if (typeof data.message === 'string') return data.message;
  if (typeof data.error === 'string') return data.error;
  if (data.error && typeof data.error.message === 'string') return data.error.message;
  if (Array.isArray(data.errors) && data.errors.length) {
    return data.errors.map(function(e){ return typeof e === 'string' ? e : (e.message || JSON.stringify(e)); }).join('; ');
  }
  try { return JSON.stringify(data); } catch (e2) { return 'Fintoc respondió ' + code; }
}
function fintocFetch_(path, opts) {
  opts = opts || {};
  const secret = FINTOC_SECRET_KEY_();
  if (!secret) throw new Error('Falta FINTOC_SECRET_KEY en Script Properties');
  const res = UrlFetchApp.fetch(FINTOC_API_ + path, {
    method: opts.method || 'get',
    contentType: 'application/json',
    headers: { Authorization: secret },
    payload: opts.payload ? JSON.stringify(opts.payload) : undefined,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  let data;
  try { data = JSON.parse(res.getContentText() || '{}'); } catch (e) { data = null; }
  if (code >= 400) {
    throw new Error(extractFintocError_(data, code));
  }
  return data;
}

// 1) Crear Link Intent → abre el widget en el frontend (SOLO la primera vez por banco)
function actionFintocCreateLinkIntent_() {
  const publicKey = FINTOC_PUBLIC_KEY_();
  if (!publicKey) return err_('Falta FINTOC_PUBLIC_KEY en Script Properties');
  const li = fintocFetch_('/link_intents', {
    method: 'post',
    payload: { product: 'movements', country: 'cl', holder_type: 'individual' }
  });
  return ok_({ publicKey: publicKey, widgetToken: li.widget_token, linkIntentId: li.id });
}

// 2) Canjear el exchange_token del widget por el Link definitivo y GUARDAR
//    su link_token de forma persistente → de aquí en adelante NUNCA más
//    se vuelve a pedir el widget para este banco, solo se usa fintocSync.
function actionFintocSaveLink_(body) {
  if (!body.exchangeToken) return err_('falta exchangeToken');
  const link = fintocFetch_('/links/' + encodeURIComponent(body.exchangeToken), { method: 'get' });
  const linkToken = link.link_token;
  const institucion = (link.institution && link.institution.name) || 'Banco';
  if (!linkToken) return err_('Fintoc no devolvió link_token');

  const sh = ensureSheet_(nsName_(SHEET_LINKS, body.ns));
  if (sh.getLastRow() === 0) sh.appendRow(['linkId', 'linkToken', 'institucion', 'creado']);
  // Evita duplicar si el usuario reconecta el mismo banco
  const rows = sh.getDataRange().getValues();
  const yaExiste = rows.some(r => r[0] === link.id);
  if (!yaExiste) sh.appendRow([link.id, linkToken, institucion, new Date().toISOString()]);

  return ok_({ banco: institucion });
}

// 3) Sincronizar: recorre TODOS los link_token guardados (sin abrir widget)
//    y trae cuentas + movimientos frescos a las hojas Registros y Cuentas.
function actionFintocSync_(body) {
  const ns = (body && body.ns) || '';
  const linkRows = readSheet_(nsName_(SHEET_LINKS, ns));
  if (linkRows.length <= 1) return err_('No hay bancos conectados todavía. Toca "Agregar banco o tarjeta".');

  const cuentasOut = [['cuenta', 'banco', 'numero', 'tipo', 'disponible', 'actual', 'cupo']];
  const registrosOut = [['id', 'cuenta', 'descripcion', 'monto', 'fecha', 'tipo', 'categoria']];

  for (let i = 1; i < linkRows.length; i++) {
    const linkToken = linkRows[i][1];
    const institucion = linkRows[i][2];
    if (!linkToken) continue;

    // Le pedimos a Fintoc que vaya a buscar datos FRESCOS al banco (si no,
    // devuelve lo último que tenía guardado en caché, que puede estar
    // atrasado). Fintoc solo permite un refresh cada 5 min por banco, así
    // que si ya hay uno reciente en curso, seguimos igual con lo disponible.
    try {
      fintocFetch_('/refresh_intents?link_token=' + encodeURIComponent(linkToken) + '&refresh_type=only_last', { method: 'post' });
    } catch (ex) {
      // Errores esperables aquí (ya hay un refresh en curso, o hay que
      // esperar 5 min): no son un problema, simplemente seguimos con los
      // datos que Fintoc tenga disponibles ahora mismo.
    }

    let accountsResp;
    try {
      accountsResp = fintocFetch_('/accounts?link_token=' + encodeURIComponent(linkToken));
    } catch (ex) {
      continue; // banco con error puntual: seguimos con los demás, no rompe todo el sync
    }
    const accounts = Array.isArray(accountsResp) ? accountsResp : (accountsResp.data || []);

    accounts.forEach(acc => {
      const last4 = String(acc.number || '').slice(-4);
      const cuentaNombre = (acc.name || 'Cuenta') + (last4 ? (' ' + last4) : '');
      const bal = acc.balance || {};
      cuentasOut.push([
        cuentaNombre, institucion, acc.number || '', acc.type || '',
        bal.available != null ? bal.available : '',
        bal.current != null ? bal.current : '',
        bal.limit != null ? bal.limit : ''
      ]);

      try {
        const movResp = fintocFetch_('/accounts/' + acc.id + '/movements?link_token=' + encodeURIComponent(linkToken) + '&per_page=100');
        const movs = Array.isArray(movResp) ? movResp : (movResp.data || []);
        movs.forEach(m => {
          registrosOut.push([
            m.id, cuentaNombre, m.description || '', m.amount,
            (m.transaction_date || m.post_date || '').toString().slice(0, 10),
            m.amount < 0 ? 'gasto' : 'ingreso',
            ''
          ]);
        });
      } catch (ex) { /* si falla un banco puntual, seguimos con el resto */ }
    });
  }

  writeSheet_(nsName_('Cuentas', ns), cuentasOut);
  writeSheet_(nsName_('Registros', ns), registrosOut);
  return ok_({ cuentas: cuentasOut.length - 1, movimientos: registrosOut.length - 1 });
}

// ── RECUPERACIÓN DE CONTRASEÑA POR CORREO ──────────────────────
// ── CUENTAS DE USUARIO (guardadas en el servidor, contraseña NUNCA en texto plano) ──
// La contraseña se guarda como SHA-256(password + salt); el salt es único por
// usuario y aleatorio, así nadie (ni tú mirando el Sheet) puede leer la
// contraseña real de nadie.
function hashPassword_(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '::' + salt);
  return digest.map(b => (b + 256).toString(16).slice(-2)).join('');
}
function findUserRow_(email) {
  const sh = ensureSheet_(SHEET_USERS);
  if (sh.getLastRow() === 0) sh.appendRow(['email', 'name', 'username', 'passwordHash', 'salt', 'creado']);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(email).toLowerCase()) return { row: i + 1, data: rows[i] };
  }
  return null;
}

function actionRegisterAccount_(body) {
  const { name, username, email, password } = body;
  if (!name || !username || !email || !password) return err_('Faltan datos para crear la cuenta');
  if (password.length < 4) return err_('La contraseña debe tener al menos 4 caracteres');
  if (findUserRow_(email)) return err_('Ya existe una cuenta con ese correo');

  const salt = Utilities.getUuid();
  const hash = hashPassword_(password, salt);
  const sh = ensureSheet_(SHEET_USERS);
  sh.appendRow([email.toLowerCase().trim(), name, username, hash, salt, new Date().toISOString()]);
  return ok_({ name, username, email });
}

function actionLoginAccount_(body) {
  const { email, password } = body;
  if (!email || !password) return err_('Faltan correo o contraseña');
  const found = findUserRow_(email);
  if (!found) return err_('No existe una cuenta con ese correo');
  const [, name, username, storedHash, salt] = found.data;
  const hash = hashPassword_(password, salt);
  if (hash !== storedHash) return err_('Contraseña incorrecta');
  return ok_({ name, username, email: found.data[0] });
}

function actionChangePasswordAccount_(body) {
  const { email, oldPassword, newPassword } = body;
  if (!email || !newPassword) return err_('Faltan datos');
  if (newPassword.length < 4) return err_('La nueva contraseña debe tener al menos 4 caracteres');
  const found = findUserRow_(email);
  if (!found) return err_('No existe una cuenta con ese correo');
  if (oldPassword) {
    const hash = hashPassword_(oldPassword, found.data[4]);
    if (hash !== found.data[3]) return err_('La contraseña actual no coincide');
  }
  const salt = Utilities.getUuid();
  const hash = hashPassword_(newPassword, salt);
  const sh = ensureSheet_(SHEET_USERS);
  sh.getRange(found.row, 4, 1, 2).setValues([[hash, salt]]);
  return ok_();
}

function actionSendResetCode_(body) {
  if (!body.email) return err_('falta email');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const sh = ensureSheet_(SHEET_RESET);
  if (sh.getLastRow() === 0) sh.appendRow(['email', 'code', 'expira', 'usado']);
  sh.appendRow([body.email, code, new Date(Date.now() + 15 * 60 * 1000).toISOString(), 'no']);
  MailApp.sendEmail({
    to: body.email,
    subject: 'Tu código de recuperación — Zynto',
    body: `Tu código de verificación es: ${code}\n\nExpira en 15 minutos. Si no lo solicitaste, ignora este correo.`
  });
  return ok_();
}
function actionVerifyResetCode_(body) {
  if (!body.email || !body.code) return err_('faltan email o code');
  const sh = ensureSheet_(SHEET_RESET);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const [email, code, expira, usado] = rows[i];
    if (email === body.email && String(code) === String(body.code)) {
      if (usado === 'si') return err_('Este código ya fue usado');
      if (new Date(expira).getTime() < Date.now()) return err_('El código expiró, solicita uno nuevo');
      sh.getRange(i + 1, 4).setValue('si');
      // Si además viene la nueva contraseña, se actualiza la cuenta real en el servidor
      if (body.newPassword) {
        const r = actionChangePasswordAccount_({ email: body.email, newPassword: body.newPassword });
        const parsed = JSON.parse(r.getContent());
        if (parsed.status !== 'ok') return r;
      }
      return ok_();
    }
  }
  return err_('Código inválido');
}
