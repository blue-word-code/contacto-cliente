/**
 * ============================================================
 *  Contacto-Cliente — Backend (Google Apps Script)
 *  Plataforma de distribución: clientes y notas de voz de ubicación.
 * ============================================================
 *  Módulos:
 *   1. Acceso: inicio de sesión, sesiones, bloqueo por intentos, roles.
 *   2. Usuarios: crear (contraseña escrita o generada), activar/desactivar,
 *      eliminar, restablecer contraseña, cambiar la propia.
 *   3. Recuperación: "¿Olvidaste tu contraseña?" con código al correo general.
 *   4. Clientes: listar, crear, editar, eliminar, importar desde Excel/CSV.
 *   5. Notas de voz: escuchar, subir, grabar, reemplazar, eliminar,
 *      sincronizar la carpeta de Drive y asociar audios sueltos.
 *   6. Actividad: registro de quién hizo qué y cuándo.
 *
 *  Datos:
 *   - Hoja de cálculo "Contacto-Cliente — Datos" con las hojas
 *     Usuarios, Clientes y Actividad.
 *   - Carpeta de Drive "Contacto-Cliente — Notas de voz" (privada).
 *
 *  Instalación: ver docs/instalacion.md
 *    1) Pega este archivo en un proyecto de Apps Script.
 *    2) Escribe la contraseña inicial en CLAVE_INICIAL_ADMIN y ejecuta  instalar  una vez.
 *       Después deja CLAVE_INICIAL_ADMIN vacía otra vez.
 *    3) Implementar > Nueva implementación > Aplicación web
 *       (Ejecutar como: Yo · Acceso: Cualquier usuario).
 * ============================================================
 */

/* ------------------------- CONFIGURACIÓN ------------------------- */
const CONFIG = {
  CORREO_GENERAL: 'distribucion.planetaazul@gmail.com', // recibe avisos y códigos
  NOMBRE_SISTEMA: 'Contacto-Cliente',
  INCLUIR_CLAVE_EN_CORREO: true, // true = el aviso incluye la nueva contraseña
  DURACION_SESION_MIN: 360,      // minutos de inactividad antes de cerrar sesión (máx. 360)
  INTENTOS_MAX: 5,               // intentos fallidos antes de bloquear el usuario
  BLOQUEO_MIN: 15,               // minutos de bloqueo
  CODIGO_MIN: 15,                // minutos de validez del código de recuperación
  LARGO_MIN_CLAVE: 8,
  ITERACIONES: 1000,             // NO cambiar después de instalar
  ZONA: 'America/Santo_Domingo',
  MAX_AUDIO_MB: 10,              // tamaño máximo de una nota de voz
  REGISTRAR_ESCUCHAS: true,      // true = la actividad registra quién escuchó cada audio
  MAX_IMPORTAR_POR_LOTE: 2000    // filas por envío al importar clientes
};

/*
 * CONTRASEÑA INICIAL DEL USUARIO "admin"
 * Escríbela entre las comillas SOLO aquí, en el editor de Apps Script
 * (nunca en GitHub). Ejecuta "instalar" y después vuelve a dejarla vacía: ''
 * La hoja guarda únicamente su versión cifrada.
 */
const CLAVE_INICIAL_ADMIN = '';

/* Usuario maestro. */
const ADMIN_INICIAL = {
  usuario: 'admin',
  nombre: 'Administrador'
};

const HOJA_USUARIOS = 'Usuarios';
const HOJA_CLIENTES = 'Clientes';
const HOJA_ACTIVIDAD = 'Actividad';

const COLUMNAS = ['usuario', 'nombre', 'rol', 'activo', 'salt', 'hash',
                  'creado', 'actualizado', 'ultimoAcceso', 'actualizadoPor'];
const COL_CLI = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono',
                 'audioId', 'audioNombre', 'audioFecha', 'audioPor', 'actualizado', 'actualizadoPor'];
const CAMPOS_CLI = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono'];
const COL_ACT = ['fecha', 'usuario', 'accion', 'detalle'];

const ROLES = ['usuario', 'editor', 'admin'];
const NIVEL = { usuario: 0, editor: 1, admin: 2 };
const EXT_AUDIO = ['mp3', 'm4a', 'mp4', 'aac', 'ogg', 'oga', 'opus', 'wav', 'webm', 'amr', '3gp', 'flac'];

/* ------------------------- INSTALACIÓN ------------------------- */

/**
 * Ejecutar desde el editor. Crea (o completa, si ya existen) la hoja de datos,
 * la carpeta de notas de voz y el usuario admin. Se puede ejecutar de nuevo sin perder datos.
 */
function instalar() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.NOMBRE_SISTEMA + ' — Datos');
    props.setProperty('SHEET_ID', ss.getId());
  } else {
    ss.rename(CONFIG.NOMBRE_SISTEMA + ' — Datos');
  }

  // Hoja Usuarios
  let sh = ss.getSheetByName(HOJA_USUARIOS);
  if (!sh) {
    const libre = ss.getSheets().filter(function (s) { return s.getLastRow() === 0; })[0];
    sh = libre || ss.insertSheet();
    sh.setName(HOJA_USUARIOS);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNAS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
    sh.hideColumns(COLUMNAS.indexOf('salt') + 1, 2); // oculta salt y hash
  }

  // Hoja Clientes
  let sc = ss.getSheetByName(HOJA_CLIENTES) || ss.insertSheet(HOJA_CLIENTES);
  if (sc.getLastRow() === 0) {
    sc.appendRow(COL_CLI);
    sc.setFrozenRows(1);
    sc.getRange(1, 1, 1, COL_CLI.length).setFontWeight('bold');
  }
  // Código y teléfono como texto (conserva ceros a la izquierda)
  sc.getRange(1, COL_CLI.indexOf('codigo') + 1, sc.getMaxRows(), 1).setNumberFormat('@');
  sc.getRange(1, COL_CLI.indexOf('telefono') + 1, sc.getMaxRows(), 1).setNumberFormat('@');

  // Hoja Actividad
  let sa = ss.getSheetByName(HOJA_ACTIVIDAD) || ss.insertSheet(HOJA_ACTIVIDAD);
  if (sa.getLastRow() === 0) {
    sa.appendRow(COL_ACT);
    sa.setFrozenRows(1);
    sa.getRange(1, 1, 1, COL_ACT.length).setFontWeight('bold');
  }

  // Usuario admin
  if (!buscarUsuario_(ADMIN_INICIAL.usuario)) {
    const clave = claveInicialAdmin_();
    const salt = nuevaSal_();
    const ahora = new Date();
    sh.appendRow([ADMIN_INICIAL.usuario, ADMIN_INICIAL.nombre, 'admin', true,
                  salt, hashClave_(clave, salt), ahora, ahora, '', 'instalación']);
  }

  // Carpeta de notas de voz
  const carpeta = carpeta_(true);

  invalidarClientes_();
  registrar_('sistema', 'Instalación', 'Hoja y carpeta verificadas');
  enviarCorreo_('Sistema instalado',
    '<p>La plataforma <b>' + esc_(CONFIG.NOMBRE_SISTEMA) + '</b> quedó instalada.</p>' +
    '<p>Hoja de datos (usuarios, clientes, actividad): <a href="' + ss.getUrl() + '">' + esc_(ss.getName()) + '</a></p>' +
    '<p>Carpeta de notas de voz: <a href="' + carpeta.getUrl() + '">' + esc_(carpeta.getName()) + '</a></p>' +
    '<p>A este correo llegarán los avisos de cambios de contraseña y los códigos de recuperación.</p>',
    true);
  Logger.log('Instalación completa.\nHoja de datos: ' + ss.getUrl() + '\nCarpeta de notas de voz: ' + carpeta.getUrl());
  return ss.getUrl();
}

/**
 * EMERGENCIA: pone al usuario "admin" la contraseña escrita en CLAVE_INICIAL_ADMIN y lo reactiva.
 * Ejecutar desde el editor solo si nadie puede entrar como administrador.
 */
function restablecerAdminInicial() {
  const clave = claveInicialAdmin_();
  const lock = bloquear_();
  try {
    const u = buscarUsuario_(ADMIN_INICIAL.usuario);
    if (!u) return instalar();
    const salt = nuevaSal_();
    guardarCampos_(u.fila, {
      salt: salt, hash: hashClave_(clave, salt), activo: true, rol: 'admin',
      actualizado: new Date(), actualizadoPor: 'editor de Apps Script'
    });
    CacheService.getScriptCache().remove('bloq_' + u.usuario);
    notificar_('Contraseña del admin devuelta a la inicial (emergencia)', u, '', 'editor de Apps Script');
    registrar_('sistema', 'Admin restablecido', 'Desde el editor de Apps Script');
    Logger.log('El usuario admin volvió a la contraseña de CLAVE_INICIAL_ADMIN.');
  } finally { lock.releaseLock(); }
}

/**
 * OPCIONAL: usar una carpeta de Drive que ya existe (por ejemplo, donde copiaste los audios).
 * Pega el ID de la carpeta (lo que va después de /folders/ en su dirección) y ejecuta.
 */
function usarCarpetaExistente() {
  const ID_CARPETA = ''; // ← pega aquí el ID entre las comillas
  if (!ID_CARPETA) throw new Error('Pega el ID de la carpeta en ID_CARPETA y vuelve a ejecutar.');
  const c = DriveApp.getFolderById(ID_CARPETA);
  PropertiesService.getScriptProperties().setProperty('CARPETA_ID', c.getId());
  Logger.log('Carpeta de notas de voz: ' + c.getName() + ' — ' + c.getUrl());
}

function claveInicialAdmin_() {
  const clave = String(CLAVE_INICIAL_ADMIN || '');
  if (clave.length < CONFIG.LARGO_MIN_CLAVE) {
    throw new Error('Escribe la contraseña inicial del admin en CLAVE_INICIAL_ADMIN (al inicio del código, ' +
                    'mínimo ' + CONFIG.LARGO_MIN_CLAVE + ' caracteres), guarda y vuelve a ejecutar.');
  }
  return clave;
}

/* ------------------------- PUNTOS DE ENTRADA ------------------------- */

/** Prueba rápida: abrir la URL de la aplicación web en el navegador. */
function doGet() {
  return json_({ ok: true, sistema: CONFIG.NOMBRE_SISTEMA, mensaje: 'API activa.' });
}

function doPost(e) {
  let res;
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const accion = ACCIONES[d.accion];
    if (!accion) throw new ErrorApp_('Acción no válida.');
    res = Object.assign({ ok: true }, accion(d) || {});
  } catch (err) {
    if (err instanceof ErrorApp_) {
      res = { ok: false, error: err.message, codigo: err.codigo || '' };
    } else {
      console.error(err && err.stack || err);
      res = { ok: false, error: 'Error interno del servidor. Intenta de nuevo.' };
    }
  }
  return json_(res);
}

const ACCIONES = {
  // Acceso y usuarios
  login: login_,
  sesion: sesion_,
  salir: salir_,
  cambiarClave: cambiarClave_,
  listarUsuarios: listarUsuarios_,
  crearUsuario: crearUsuario_,
  restablecerClave: restablecerClave_,
  cambiarEstado: cambiarEstado_,
  cambiarRol: cambiarRol_,
  eliminarUsuario: eliminarUsuario_,
  olvideClave: olvideClave_,
  restablecerConCodigo: restablecerConCodigo_,
  // Clientes
  listarClientes: listarClientes_,
  guardarCliente: guardarCliente_,
  eliminarCliente: eliminarCliente_,
  importarClientes: importarClientes_,
  // Notas de voz
  audio: audio_,
  subirAudio: subirAudio_,
  eliminarAudio: eliminarAudio_,
  sincronizarAudios: sincronizarAudios_,
  asignarAudio: asignarAudio_,
  // Actividad
  listarActividad: listarActividad_
};

/* =========================================================
 *  MÓDULO 1-3: ACCESO, USUARIOS Y RECUPERACIÓN
 * ========================================================= */

function login_(d) {
  const usuario = norm_(d.usuario);
  const clave = String(d.clave || '');
  if (!usuario || !clave) throw new ErrorApp_('Escribe tu usuario y contraseña.');
  const c = cache_();
  if (c.get('bloq_' + usuario)) {
    throw new ErrorApp_('Demasiados intentos fallidos. Espera ' + CONFIG.BLOQUEO_MIN + ' minutos e intenta de nuevo.');
  }
  const u = buscarUsuario_(usuario);
  if (!u || !verificarClave_(clave, u)) {
    const fallos = Number(c.get('fallos_' + usuario) || 0) + 1;
    if (fallos >= CONFIG.INTENTOS_MAX) {
      c.put('bloq_' + usuario, '1', CONFIG.BLOQUEO_MIN * 60);
      c.remove('fallos_' + usuario);
      registrar_(usuario, 'Usuario bloqueado', CONFIG.INTENTOS_MAX + ' intentos fallidos');
      throw new ErrorApp_('Demasiados intentos fallidos. El usuario quedó bloqueado ' + CONFIG.BLOQUEO_MIN + ' minutos.');
    }
    c.put('fallos_' + usuario, String(fallos), CONFIG.BLOQUEO_MIN * 60);
    throw new ErrorApp_('Usuario o contraseña incorrectos.');
  }
  if (!u.activo) throw new ErrorApp_('Este usuario está desactivado. Contacta al administrador.');
  c.remove('fallos_' + usuario);
  guardarCampos_(u.fila, { ultimoAcceso: new Date() });
  registrar_(u.usuario, 'Inicio de sesión', '');
  return { token: crearSesion_(u), usuario: perfil_(u) };
}

function sesion_(d) {
  return { usuario: perfil_(requerirSesion_(d)) };
}

function salir_(d) {
  if (d.token) cache_().remove('s_' + d.token);
  return {};
}

function cambiarClave_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d);
    const actual = String(d.actual || '');
    const nueva = String(d.nueva || '');
    if (!verificarClave_(actual, u)) throw new ErrorApp_('La contraseña actual no es correcta.');
    validarClave_(nueva);
    if (nueva === actual) throw new ErrorApp_('La nueva contraseña debe ser diferente a la actual.');
    asignarClave_(u, nueva, u.usuario);
    cache_().remove('s_' + d.token);
    notificar_('Cambio de contraseña hecho por el propio usuario', u, nueva, u.usuario);
    registrar_(u.usuario, 'Cambió su contraseña', '');
    return { token: crearSesion_(u), mensaje: 'Contraseña actualizada.' };
  } finally { lock.releaseLock(); }
}

function listarUsuarios_(d) {
  requerirSesion_(d, 'admin');
  const lista = leerUsuarios_().map(function (u) {
    return {
      usuario: u.usuario, nombre: u.nombre, rol: u.rol, activo: u.activo,
      creado: fecha_(u.creado), ultimoAcceso: fecha_(u.ultimoAcceso),
      principal: u.usuario === ADMIN_INICIAL.usuario
    };
  });
  return { usuarios: lista };
}

function crearUsuario_(d) {
  const lock = bloquear_();
  try {
    const admin = requerirSesion_(d, 'admin');
    const usuario = norm_(d.usuario);
    if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) {
      throw new ErrorApp_('El usuario debe tener de 3 a 30 caracteres: letras sin acento, números, punto, guion o guion bajo.');
    }
    if (buscarUsuario_(usuario)) throw new ErrorApp_('Ya existe un usuario "' + usuario + '".');
    const nombre = limpiarTexto_(d.nombre, 60);
    if (!nombre) throw new ErrorApp_('Escribe el nombre de la persona.');
    const rol = ROLES.indexOf(d.rol) >= 0 ? d.rol : 'usuario';
    const clave = String(d.clave || '');
    validarClave_(clave);
    const salt = nuevaSal_();
    const hash = hashClave_(clave, salt);
    const ahora = new Date();
    hoja_(HOJA_USUARIOS).appendRow([usuario, nombre, rol, true, salt, hash, ahora, ahora, '', admin.usuario]);
    const nuevo = { usuario: usuario, nombre: nombre, rol: rol };
    notificar_('Usuario creado (contraseña ' + (d.generada ? 'generada' : 'escrita por el administrador') + ')',
               nuevo, clave, admin.usuario);
    registrar_(admin.usuario, 'Creó usuario', usuario + ' (' + rol + ')');
    return { usuario: nuevo, mensaje: 'Usuario creado.' };
  } finally { lock.releaseLock(); }
}

function restablecerClave_(d) {
  const lock = bloquear_();
  try {
    const admin = requerirSesion_(d, 'admin');
    const u = buscarUsuario_(d.usuario);
    if (!u) throw new ErrorApp_('El usuario no existe.');
    if (u.usuario === ADMIN_INICIAL.usuario && admin.usuario !== ADMIN_INICIAL.usuario) {
      throw new ErrorApp_('Solo el propio usuario "' + ADMIN_INICIAL.usuario + '" puede cambiar su contraseña.');
    }
    const clave = String(d.clave || '');
    validarClave_(clave);
    asignarClave_(u, clave, admin.usuario);
    cache_().remove('bloq_' + u.usuario);
    notificar_('Contraseña restablecida por el administrador (' + (d.generada ? 'generada' : 'escrita') + ')',
               u, clave, admin.usuario);
    registrar_(admin.usuario, 'Restableció contraseña', u.usuario);
    const res = { mensaje: 'Contraseña actualizada.' };
    if (u.usuario === admin.usuario) { // el admin se cambió la suya: nueva sesión
      cache_().remove('s_' + d.token);
      res.token = crearSesion_(u);
    }
    return res;
  } finally { lock.releaseLock(); }
}

function cambiarEstado_(d) {
  const lock = bloquear_();
  try {
    const admin = requerirSesion_(d, 'admin');
    const u = objetivoProtegido_(d.usuario, admin);
    const activo = d.activo === true;
    guardarCampos_(u.fila, { activo: activo, actualizado: new Date(), actualizadoPor: admin.usuario });
    registrar_(admin.usuario, activo ? 'Activó usuario' : 'Desactivó usuario', u.usuario);
    return { mensaje: activo ? 'Usuario activado.' : 'Usuario desactivado.' };
  } finally { lock.releaseLock(); }
}

function cambiarRol_(d) {
  const lock = bloquear_();
  try {
    const admin = requerirSesion_(d, 'admin');
    const u = objetivoProtegido_(d.usuario, admin);
    if (ROLES.indexOf(d.rol) < 0) throw new ErrorApp_('Rol no válido.');
    guardarCampos_(u.fila, { rol: d.rol, actualizado: new Date(), actualizadoPor: admin.usuario });
    registrar_(admin.usuario, 'Cambió rol', u.usuario + ': ' + u.rol + ' → ' + d.rol);
    return { mensaje: 'Rol actualizado.' };
  } finally { lock.releaseLock(); }
}

function eliminarUsuario_(d) {
  const lock = bloquear_();
  try {
    const admin = requerirSesion_(d, 'admin');
    const u = objetivoProtegido_(d.usuario, admin);
    hoja_(HOJA_USUARIOS).deleteRow(u.fila);
    notificar_('Usuario eliminado', u, '', admin.usuario);
    registrar_(admin.usuario, 'Eliminó usuario', u.usuario);
    return { mensaje: 'Usuario eliminado.' };
  } finally { lock.releaseLock(); }
}

function olvideClave_(d) {
  const usuario = norm_(d.usuario);
  if (!usuario) throw new ErrorApp_('Escribe tu usuario.');
  const c = cache_();
  if (c.get('rcesp_' + usuario)) throw new ErrorApp_('Ya se envió un código. Espera un minuto antes de pedir otro.');
  const u = buscarUsuario_(usuario);
  if (u && u.activo) {
    const codigo = nuevoCodigo_();
    c.put('rc_' + usuario, JSON.stringify({ c: codigo, i: 0 }), CONFIG.CODIGO_MIN * 60);
    c.put('rcesp_' + usuario, '1', 60);
    enviarCorreo_('Código para restablecer la contraseña de ' + u.usuario,
      '<p>Se solicitó restablecer la contraseña del usuario <b>' + esc_(u.usuario) + '</b> (' + esc_(u.nombre) + ').</p>' +
      '<p style="font-size:28px;letter-spacing:6px;font-family:monospace;margin:16px 0"><b>' + codigo + '</b></p>' +
      '<p>El código vence en ' + CONFIG.CODIGO_MIN + ' minutos. Entrégaselo solo a esa persona.</p>' +
      '<p style="color:#777">Si nadie lo solicitó, ignora este correo; la contraseña no cambia.</p>',
      true);
    registrar_(u.usuario, 'Pidió código de recuperación', '');
  }
  // Misma respuesta exista o no el usuario.
  return { mensaje: 'Si el usuario existe, se envió un código de recuperación al correo general (' +
                    ocultarCorreo_(CONFIG.CORREO_GENERAL) + '). Pídeselo al encargado.' };
}

function restablecerConCodigo_(d) {
  const lock = bloquear_();
  try {
    const usuario = norm_(d.usuario);
    const c = cache_();
    const raw = c.get('rc_' + usuario);
    if (!raw) throw new ErrorApp_('El código venció o no es válido. Solicita uno nuevo.');
    const reg = JSON.parse(raw);
    if (reg.c !== String(d.codigo || '').trim()) {
      reg.i += 1;
      if (reg.i >= 5) {
        c.remove('rc_' + usuario);
        throw new ErrorApp_('Demasiados intentos con código incorrecto. Solicita un código nuevo.');
      }
      c.put('rc_' + usuario, JSON.stringify(reg), CONFIG.CODIGO_MIN * 60);
      throw new ErrorApp_('El código no es correcto.');
    }
    const nueva = String(d.nueva || '');
    validarClave_(nueva);
    const u = buscarUsuario_(usuario);
    if (!u || !u.activo) throw new ErrorApp_('El usuario no existe o está desactivado.');
    asignarClave_(u, nueva, u.usuario + ' (código de recuperación)');
    c.remove('rc_' + usuario);
    c.remove('bloq_' + usuario);
    c.remove('fallos_' + usuario);
    notificar_('Contraseña restablecida con "¿Olvidaste tu contraseña?"' + (d.generada ? ' (generada)' : ''),
               u, nueva, u.usuario);
    registrar_(u.usuario, 'Restableció su contraseña con código', '');
    return { mensaje: 'Contraseña restablecida. Ya puedes iniciar sesión.' };
  } finally { lock.releaseLock(); }
}

/* =========================================================
 *  MÓDULO 4: CLIENTES
 * ========================================================= */

/** Lista completa en formato compacto (una fila = un arreglo). La app busca y filtra en el teléfono. */
function listarClientes_(d) {
  const u = requerirSesion_(d);
  const res = { campos: COMPACTO, clientes: clientesCompactos_() };
  if (NIVEL[u.rol] >= NIVEL.editor) {
    const c = carpeta_(false);
    if (c) res.carpeta = { nombre: c.getName(), url: c.getUrl() };
  }
  return res;
}

function guardarCliente_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d, 'editor');
    const datos = limpiarCliente_(d.cliente || {});
    if (!datos.codigo) throw new ErrorApp_('Escribe el código del cliente.');
    if (!/^[A-Z0-9][A-Z0-9 ._\/-]{0,39}$/.test(datos.codigo)) {
      throw new ErrorApp_('El código solo puede tener letras, números, espacio, punto, guion o barra (máximo 40).');
    }
    if (!datos.nombre) throw new ErrorApp_('Escribe el nombre del cliente.');
    const t = tablaClientes_();
    const original = d.codigoOriginal ? normCodigo_(d.codigoOriginal) : '';
    const idxNuevo = t.indice[datos.codigo];
    let fila;
    if (!original) { // nuevo
      if (idxNuevo !== undefined) throw new ErrorApp_('Ya existe un cliente con el código "' + datos.codigo + '".');
      fila = COL_CLI.map(function () { return ''; });
      CAMPOS_CLI.forEach(function (k) { fila[ci_(k)] = datos[k]; });
      fila[ci_('actualizado')] = new Date();
      fila[ci_('actualizadoPor')] = u.usuario;
      escribirNuevos_(t, [fila]);
      registrar_(u.usuario, 'Creó cliente', datos.codigo + ' — ' + datos.nombre);
    } else {
      const idx = t.indice[original];
      if (idx === undefined) throw new ErrorApp_('El cliente "' + original + '" ya no existe.');
      if (datos.codigo !== original && idxNuevo !== undefined) {
        throw new ErrorApp_('Ya existe otro cliente con el código "' + datos.codigo + '".');
      }
      fila = t.valores[idx];
      CAMPOS_CLI.forEach(function (k) { fila[ci_(k)] = datos[k]; });
      fila[ci_('actualizado')] = new Date();
      fila[ci_('actualizadoPor')] = u.usuario;
      t.hoja.getRange(idx + 2, 1, 1, COL_CLI.length).setValues([fila]);
      registrar_(u.usuario, 'Editó cliente', datos.codigo + (datos.codigo !== original ? ' (antes ' + original + ')' : ''));
    }
    invalidarClientes_();
    return { cliente: compactar_(fila), mensaje: original ? 'Cliente actualizado.' : 'Cliente creado.' };
  } finally { lock.releaseLock(); }
}

function eliminarCliente_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d, 'admin');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined) throw new ErrorApp_('El cliente no existe.');
    const fila = t.valores[idx];
    if (fila[ci_('audioId')]) papelera_(fila[ci_('audioId')]);
    t.hoja.deleteRow(idx + 2);
    invalidarClientes_();
    registrar_(u.usuario, 'Eliminó cliente', codigo + ' — ' + fila[ci_('nombre')]);
    return { mensaje: 'Cliente eliminado.' };
  } finally { lock.releaseLock(); }
}

/**
 * Importa clientes por lotes. Si el código ya existe, actualiza solo las columnas que
 * vienen con valor (no borra datos con celdas vacías). Nunca toca la nota de voz.
 */
function importarClientes_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d, 'editor');
    const filas = Array.isArray(d.filas) ? d.filas : [];
    if (!filas.length) throw new ErrorApp_('No llegaron filas para importar.');
    if (filas.length > CONFIG.MAX_IMPORTAR_POR_LOTE) throw new ErrorApp_('Máximo ' + CONFIG.MAX_IMPORTAR_POR_LOTE + ' filas por envío.');
    const t = tablaClientes_();
    const nuevos = [];
    const indiceNuevos = {};
    let actualizados = 0, sinCambios = 0;
    const omitidos = [];
    const ahora = new Date();
    filas.forEach(function (bruto, i) {
      const c = limpiarCliente_(bruto || {});
      const n = Number(d.desde || 0) + i + 1;
      if (!c.codigo) { omitidos.push('Fila ' + n + ': sin código'); return; }
      const idx = t.indice[c.codigo];
      if (idx !== undefined) {
        const fila = t.valores[idx];
        let cambio = false;
        CAMPOS_CLI.forEach(function (k) {
          if (k !== 'codigo' && c[k] && String(fila[ci_(k)]) !== c[k]) { fila[ci_(k)] = c[k]; cambio = true; }
        });
        if (cambio) {
          fila[ci_('actualizado')] = ahora;
          fila[ci_('actualizadoPor')] = u.usuario + ' (importación)';
          t.cambiadas[idx] = true;
          actualizados++;
        } else sinCambios++;
      } else if (indiceNuevos[c.codigo] !== undefined) {
        const f = nuevos[indiceNuevos[c.codigo]];
        CAMPOS_CLI.forEach(function (k) { if (c[k]) f[ci_(k)] = c[k]; });
        omitidos.push('Fila ' + n + ': código ' + c.codigo + ' repetido en el archivo (se unieron los datos)');
      } else {
        if (!c.nombre) { omitidos.push('Fila ' + n + ': cliente nuevo ' + c.codigo + ' sin nombre'); return; }
        const f = COL_CLI.map(function () { return ''; });
        CAMPOS_CLI.forEach(function (k) { f[ci_(k)] = c[k]; });
        f[ci_('actualizado')] = ahora;
        f[ci_('actualizadoPor')] = u.usuario + ' (importación)';
        indiceNuevos[c.codigo] = nuevos.length;
        nuevos.push(f);
      }
    });
    // Escribe solo los bloques que cambiaron
    const idxs = Object.keys(t.cambiadas).map(Number).sort(function (a, b) { return a - b; });
    if (idxs.length) {
      const desde = idxs[0], hasta = idxs[idxs.length - 1];
      t.hoja.getRange(desde + 2, 1, hasta - desde + 1, COL_CLI.length).setValues(t.valores.slice(desde, hasta + 1));
    }
    if (nuevos.length) escribirNuevos_(t, nuevos);
    invalidarClientes_();
    registrar_(u.usuario, 'Importó clientes',
      'nuevos ' + nuevos.length + ', actualizados ' + actualizados + ', sin cambios ' + sinCambios + ', omitidos ' + omitidos.length);
    return { nuevos: nuevos.length, actualizados: actualizados, sinCambios: sinCambios,
             omitidos: omitidos.slice(0, 50), totalOmitidos: omitidos.length };
  } finally { lock.releaseLock(); }
}

/* =========================================================
 *  MÓDULO 5: NOTAS DE VOZ
 * ========================================================= */

/** Devuelve el audio en base64. Solo con sesión válida: los archivos nunca se comparten en Drive. */
function audio_(d) {
  const u = requerirSesion_(d);
  const codigo = normCodigo_(d.codigo);
  const t = tablaClientes_();
  const idx = t.indice[codigo];
  if (idx === undefined) throw new ErrorApp_('El cliente no existe.');
  const fila = t.valores[idx];
  const id = fila[ci_('audioId')];
  if (!id) throw new ErrorApp_('Este cliente no tiene nota de voz.');
  let archivo;
  try { archivo = DriveApp.getFileById(id); } catch (e) {
    throw new ErrorApp_('No se encontró el archivo de audio en Drive. Puede que lo hayan borrado.');
  }
  if (archivo.isTrashed()) throw new ErrorApp_('El archivo de audio está en la papelera de Drive.');
  const blob = archivo.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > 25 * 1024 * 1024) throw new ErrorApp_('El audio es demasiado grande para reproducirlo.');
  if (CONFIG.REGISTRAR_ESCUCHAS) registrar_(u.usuario, 'Escuchó nota de voz', codigo + ' — ' + fila[ci_('nombre')]);
  return { mime: tipoAudio_(archivo.getName(), blob.getContentType()), nombre: archivo.getName(),
           base64: Utilities.base64Encode(bytes) };
}

/** Sube (o graba) una nota de voz para un cliente. */
function subirAudio_(d) {
  const u = requerirSesion_(d, 'editor');
  const codigo = normCodigo_(d.codigo);
  if (!d.base64) throw new ErrorApp_('No llegó el audio.');
  const bytes = Utilities.base64Decode(String(d.base64));
  if (bytes.length > CONFIG.MAX_AUDIO_MB * 1024 * 1024) {
    throw new ErrorApp_('El audio pesa más de ' + CONFIG.MAX_AUDIO_MB + ' MB.');
  }
  const ext = extension_(d.nombreArchivo, d.mime);
  const lock = bloquear_();
  try {
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined) throw new ErrorApp_('No existe un cliente con el código "' + codigo + '".');
    const fila = t.valores[idx];
    const anterior = fila[ci_('audioId')];
    if (anterior && !d.reemplazar) throw new ErrorApp_('El cliente ' + codigo + ' ya tiene nota de voz.', 'YA_TIENE');
    const nombreArchivo = codigo.replace(/[\/\\]/g, '-') + '.' + ext;
    const archivo = carpeta_(true).createFile(Utilities.newBlob(bytes, tipoAudio_(nombreArchivo, d.mime), nombreArchivo));
    archivo.setDescription(String(fila[ci_('nombre')] || ''));
    if (anterior) papelera_(anterior);
    fila[ci_('audioId')] = archivo.getId();
    fila[ci_('audioNombre')] = nombreArchivo;
    fila[ci_('audioFecha')] = new Date();
    fila[ci_('audioPor')] = u.usuario + (d.grabado ? ' (grabado en la app)' : '');
    escribirAudio_(t, idx);
    invalidarClientes_();
    registrar_(u.usuario, anterior ? 'Reemplazó nota de voz' : (d.grabado ? 'Grabó nota de voz' : 'Subió nota de voz'),
               codigo + ' — ' + fila[ci_('nombre')]);
    return { cliente: compactar_(fila), mensaje: 'Nota de voz guardada.' };
  } finally { lock.releaseLock(); }
}

function eliminarAudio_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d, 'editor');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined) throw new ErrorApp_('El cliente no existe.');
    const fila = t.valores[idx];
    if (!fila[ci_('audioId')]) throw new ErrorApp_('Este cliente no tiene nota de voz.');
    papelera_(fila[ci_('audioId')]);
    ['audioId', 'audioNombre', 'audioFecha', 'audioPor'].forEach(function (k) { fila[ci_(k)] = ''; });
    escribirAudio_(t, idx);
    invalidarClientes_();
    registrar_(u.usuario, 'Eliminó nota de voz', codigo + ' — ' + fila[ci_('nombre')]);
    return { cliente: compactar_(fila), mensaje: 'Nota de voz eliminada (quedó en la papelera de Drive 30 días).' };
  } finally { lock.releaseLock(); }
}

/**
 * Recorre la carpeta de Drive y asocia cada audio al cliente cuyo código aparece en el nombre del archivo.
 * Trabaja por tandas de ~4,5 minutos; si quedan archivos devuelve "continuar" y la app vuelve a llamar.
 */
function sincronizarAudios_(d) {
  const u = requerirSesion_(d, 'editor');
  const inicio = Date.now();
  const carpeta = carpeta_(true);
  const it = d.continuar ? DriveApp.continueFileIterator(String(d.continuar)) : carpeta.getFiles();
  const t = tablaClientes_();
  const mapas = mapasCodigo_(t);
  const vinculados = {};
  t.valores.forEach(function (f) { if (f[ci_('audioId')]) vinculados[f[ci_('audioId')]] = true; });

  const cambios = {}; // codigo -> {id, nombre, fecha}
  const sinAsociar = [], conAudio = [];
  let procesados = 0, asociados = 0, yaAsociados = 0, ignorados = 0;
  while (it.hasNext() && Date.now() - inicio < 270000) {
    const f = it.next();
    procesados++;
    const id = f.getId();
    if (vinculados[id]) { yaAsociados++; continue; }
    const nombre = f.getName();
    if (!esAudio_(nombre, f.getMimeType())) { ignorados++; continue; }
    const idx = buscarCodigoEnNombre_(nombre, mapas);
    if (idx === null) { if (sinAsociar.length < 1000) sinAsociar.push({ id: id, nombre: nombre }); continue; }
    const codigo = String(t.valores[idx][ci_('codigo')]);
    const yaTiene = t.valores[idx][ci_('audioId')] || cambios[codigo];
    if (yaTiene && !d.reemplazar) {
      if (conAudio.length < 1000) conAudio.push({ id: id, nombre: nombre, codigo: codigo });
      continue;
    }
    cambios[codigo] = { id: id, nombre: nombre, fecha: f.getDateCreated(), anterior: t.valores[idx][ci_('audioId')] };
    vinculados[id] = true;
    asociados++;
  }
  const continuar = it.hasNext() ? it.getContinuationToken() : null;

  if (asociados) {
    const lock = bloquear_();
    try {
      const t2 = tablaClientes_(); // relee por si alguien cambió algo mientras tanto
      Object.keys(cambios).forEach(function (codigo) {
        const idx = t2.indice[codigo];
        if (idx === undefined) return;
        const fila = t2.valores[idx];
        const c = cambios[codigo];
        if (fila[ci_('audioId')] && fila[ci_('audioId')] !== c.id && d.reemplazar) papelera_(fila[ci_('audioId')]);
        fila[ci_('audioId')] = c.id;
        fila[ci_('audioNombre')] = c.nombre;
        fila[ci_('audioFecha')] = c.fecha;
        fila[ci_('audioPor')] = u.usuario + ' (sincronización)';
        t2.cambiadas[idx] = true;
      });
      escribirAudiosCambiados_(t2);
    } finally { lock.releaseLock(); }
    invalidarClientes_();
  }
  registrar_(u.usuario, 'Sincronizó carpeta de audios',
    'revisados ' + procesados + ', asociados ' + asociados + ', sin asociar ' + sinAsociar.length +
    (continuar ? ' (continúa)' : ''));
  return { procesados: procesados, asociados: asociados, yaAsociados: yaAsociados, ignorados: ignorados,
           sinAsociar: sinAsociar, conAudio: conAudio, continuar: continuar };
}

/** Asocia a mano un archivo de la carpeta (que no se pudo asociar solo) a un cliente. */
function asignarAudio_(d) {
  const lock = bloquear_();
  try {
    const u = requerirSesion_(d, 'editor');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined) throw new ErrorApp_('No existe un cliente con el código "' + codigo + '".');
    let archivo;
    try { archivo = DriveApp.getFileById(String(d.fileId)); } catch (e) { throw new ErrorApp_('No se encontró el archivo en Drive.'); }
    const fila = t.valores[idx];
    const anterior = fila[ci_('audioId')];
    if (anterior && anterior !== archivo.getId() && !d.reemplazar) {
      throw new ErrorApp_('El cliente ' + codigo + ' ya tiene nota de voz.', 'YA_TIENE');
    }
    if (anterior && anterior !== archivo.getId()) papelera_(anterior);
    fila[ci_('audioId')] = archivo.getId();
    fila[ci_('audioNombre')] = archivo.getName();
    fila[ci_('audioFecha')] = archivo.getDateCreated();
    fila[ci_('audioPor')] = u.usuario + ' (asignado a mano)';
    escribirAudio_(t, idx);
    invalidarClientes_();
    registrar_(u.usuario, 'Asignó nota de voz', codigo + ' ← ' + archivo.getName());
    return { cliente: compactar_(fila), mensaje: 'Audio asignado al cliente ' + codigo + '.' };
  } finally { lock.releaseLock(); }
}

/* =========================================================
 *  MÓDULO 6: ACTIVIDAD
 * ========================================================= */

function listarActividad_(d) {
  requerirSesion_(d, 'admin');
  const sh = hoja_(HOJA_ACTIVIDAD);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { actividad: [], total: 0 };
  const leer = Math.min(n, 5000);
  const filas = sh.getRange(n - leer + 2, 1, leer, COL_ACT.length).getValues().reverse();
  const q = String(d.filtro || '').trim().toLowerCase();
  const limite = Math.min(Number(d.limite) || 300, 1000);
  const out = [];
  for (let i = 0; i < filas.length && out.length < limite; i++) {
    const f = filas[i];
    const texto = (f[1] + ' ' + f[2] + ' ' + f[3]).toLowerCase();
    if (q && texto.indexOf(q) < 0) continue;
    out.push([fecha_(f[0]), String(f[1]), String(f[2]), String(f[3])]);
  }
  return { actividad: out, total: n };
}

function registrar_(usuario, accion, detalle) {
  try {
    hoja_(HOJA_ACTIVIDAD).appendRow([new Date(), String(usuario || ''), String(accion || ''),
                                     limpiarTexto_(detalle || '', 300)]);
  } catch (e) {
    console.error('No se pudo registrar actividad: ' + e);
  }
}

/* ------------------------- APOYO: CLIENTES ------------------------- */

const COMPACTO = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono', 'audio', 'audioFecha', 'audioPor'];

function ci_(campo) { return COL_CLI.indexOf(campo); }

function normCodigo_(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toUpperCase(); }

function limpiarCliente_(c) {
  const out = {};
  CAMPOS_CLI.forEach(function (k) {
    out[k] = k === 'codigo' ? normCodigo_(c[k]) : limpiarTexto_(c[k], k === 'direccion' ? 250 : 120);
  });
  out.telefono = String(c.telefono == null ? '' : c.telefono).replace(/[^\d+\-() \/]/g, '').trim().slice(0, 40);
  return out;
}

/** Agrega filas nuevas al final de Clientes (amplía la hoja si hace falta; código y teléfono como texto). */
function escribirNuevos_(t, filas) {
  const sh = t.hoja;
  const primera = t.valores.length + 2;
  const ultima = primera + filas.length - 1;
  if (ultima > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), ultima - sh.getMaxRows() + 100);
  sh.getRange(primera, ci_('codigo') + 1, filas.length, 1).setNumberFormat('@');
  sh.getRange(primera, ci_('telefono') + 1, filas.length, 1).setNumberFormat('@');
  sh.getRange(primera, 1, filas.length, COL_CLI.length).setValues(filas);
}

/** Lee la hoja Clientes completa e indexa por código. */
function tablaClientes_() {
  const sh = hoja_(HOJA_CLIENTES);
  const n = sh.getLastRow() - 1;
  const valores = n > 0 ? sh.getRange(2, 1, n, COL_CLI.length).getValues() : [];
  const indice = {};
  valores.forEach(function (f, i) {
    const c = normCodigo_(f[0]);
    f[0] = c;
    if (c && indice[c] === undefined) indice[c] = i;
  });
  return { hoja: sh, valores: valores, indice: indice, cambiadas: {} };
}

function compactar_(f) {
  return [String(f[ci_('codigo')]), String(f[ci_('nombre')]), String(f[ci_('direccion')]), String(f[ci_('sector')]),
          String(f[ci_('ruta')]), String(f[ci_('vendedor')]), String(f[ci_('telefono')]),
          f[ci_('audioId')] ? 1 : 0, fecha_(f[ci_('audioFecha')]), String(f[ci_('audioPor')] || '')];
}

/** Lista compacta con caché por trozos (CacheService admite 100 KB por clave). */
function clientesCompactos_() {
  const c = cache_();
  const n = Number(c.get('cli_n') || 0);
  if (n > 0) {
    const claves = [];
    for (let i = 0; i < n; i++) claves.push('cli_' + i);
    const trozos = c.getAll(claves);
    if (Object.keys(trozos).length === n) {
      try { return JSON.parse(claves.map(function (k) { return trozos[k]; }).join('')); } catch (e) { /* reconstruye */ }
    }
  }
  const lista = tablaClientes_().valores.filter(function (f) { return f[0]; }).map(compactar_);
  const texto = JSON.stringify(lista);
  const TAM = 30000; // caracteres por trozo (con margen para acentos en UTF-8)
  const pares = {};
  let k = 0;
  for (let i = 0; i < texto.length; i += TAM) pares['cli_' + (k++)] = texto.slice(i, i + TAM);
  try {
    if (k <= 900) { c.putAll(pares, 21600); c.put('cli_n', String(k), 21600); }
  } catch (e) { console.error('No se pudo guardar la lista en caché: ' + e); }
  return lista;
}

function invalidarClientes_() { cache_().remove('cli_n'); }

function escribirAudio_(t, idx) {
  const desde = ci_('audioId') + 1;
  t.hoja.getRange(idx + 2, desde, 1, 4).setValues([t.valores[idx].slice(desde - 1, desde + 3)]);
}

function escribirAudiosCambiados_(t) {
  const idxs = Object.keys(t.cambiadas).map(Number).sort(function (a, b) { return a - b; });
  if (!idxs.length) return;
  const desde = ci_('audioId');
  const ini = idxs[0], fin = idxs[idxs.length - 1];
  const bloque = t.valores.slice(ini, fin + 1).map(function (f) { return f.slice(desde, desde + 4); });
  t.hoja.getRange(ini + 2, desde + 1, bloque.length, 4).setValues(bloque);
}

/* ------------------------- APOYO: AUDIOS ------------------------- */

function carpeta_(crear) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CARPETA_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* se crea de nuevo */ }
  }
  if (!crear) return null;
  const nueva = DriveApp.createFolder(CONFIG.NOMBRE_SISTEMA + ' — Notas de voz');
  props.setProperty('CARPETA_ID', nueva.getId());
  return nueva;
}

function papelera_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { console.error('No se pudo mover a la papelera: ' + e); }
}

function extension_(nombre, mime) {
  const m = String(nombre || '').toLowerCase().match(/\.([a-z0-9]{2,5})$/);
  if (m && EXT_AUDIO.indexOf(m[1]) >= 0) return m[1];
  const t = String(mime || '').toLowerCase();
  if (t.indexOf('webm') >= 0) return 'webm';
  if (t.indexOf('ogg') >= 0 || t.indexOf('opus') >= 0) return 'ogg';
  if (t.indexOf('mp4') >= 0 || t.indexOf('m4a') >= 0 || t.indexOf('aac') >= 0) return 'm4a';
  if (t.indexOf('wav') >= 0) return 'wav';
  if (t.indexOf('amr') >= 0) return 'amr';
  return 'mp3';
}

function tipoAudio_(nombre, mime) {
  const ext = (String(nombre || '').toLowerCase().match(/\.([a-z0-9]{2,5})$/) || [])[1];
  const mapa = { mp3: 'audio/mpeg', m4a: 'audio/mp4', mp4: 'audio/mp4', aac: 'audio/aac', ogg: 'audio/ogg',
                 oga: 'audio/ogg', opus: 'audio/ogg', wav: 'audio/wav', webm: 'audio/webm', amr: 'audio/amr',
                 '3gp': 'audio/3gpp', flac: 'audio/flac' };
  if (ext && mapa[ext]) return mapa[ext];
  const m = String(mime || '');
  return m.indexOf('audio/') === 0 ? m.split(';')[0] : 'audio/mpeg';
}

function esAudio_(nombre, mime) {
  const ext = (String(nombre || '').toLowerCase().match(/\.([a-z0-9]{2,5})$/) || [])[1];
  return String(mime || '').indexOf('audio/') === 0 || (ext && EXT_AUDIO.indexOf(ext) >= 0);
}

/** Índices de códigos: exacto y sin ceros a la izquierda (solo si no hay ambigüedad). */
function mapasCodigo_(t) {
  const exacto = {}, sinCeros = {}, repetidoSinCeros = {};
  Object.keys(t.indice).forEach(function (c) {
    exacto[c] = t.indice[c];
    const s = c.replace(/^0+(?=.)/, '');
    if (sinCeros[s] !== undefined && sinCeros[s] !== t.indice[c]) repetidoSinCeros[s] = true;
    else sinCeros[s] = t.indice[c];
  });
  Object.keys(repetidoSinCeros).forEach(function (s) { delete sinCeros[s]; });
  return { exacto: exacto, sinCeros: sinCeros };
}

/**
 * Busca el código del cliente dentro del nombre del archivo.
 * "C00123.mp3", "00123 Juan Perez.ogg", "Ruta 5 - C00123 - colmado.m4a" → C00123
 */
function buscarCodigoEnNombre_(nombre, mapas) {
  const base = String(nombre).replace(/\.[^.]+$/, '').trim().toUpperCase();
  if (mapas.exacto[base] !== undefined) return mapas.exacto[base];
  const partes = base.split(/[^A-Z0-9]+/).filter(Boolean);
  for (let i = 0; i < partes.length; i++) {
    if (mapas.exacto[partes[i]] !== undefined) return mapas.exacto[partes[i]];
  }
  for (let i = 0; i < partes.length; i++) {
    const s = partes[i].replace(/^0+(?=.)/, '');
    if (/\d/.test(s) && mapas.sinCeros[s] !== undefined) return mapas.sinCeros[s];
  }
  return null;
}

/* ------------------------- SESIONES Y ROLES ------------------------- */

function segundosSesion_() {
  return Math.min(CONFIG.DURACION_SESION_MIN, 360) * 60; // CacheService permite máx. 6 h
}

function crearSesion_(u) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache_().put('s_' + token, JSON.stringify({ u: u.usuario, h: String(u.hash).slice(0, 16) }), segundosSesion_());
  return token;
}

/**
 * Valida el token y el rol mínimo ('usuario', 'editor' o 'admin').
 * Una sesión deja de servir si el usuario cambia de contraseña o es desactivado.
 */
function requerirSesion_(d, rolMinimo) {
  const c = cache_();
  const raw = d.token ? c.get('s_' + d.token) : null;
  if (!raw) throw new ErrorApp_('Tu sesión expiró. Inicia sesión de nuevo.', 'SESION');
  const s = JSON.parse(raw);
  const u = buscarUsuario_(s.u);
  if (!u || !u.activo || String(u.hash).slice(0, 16) !== s.h) {
    c.remove('s_' + d.token);
    throw new ErrorApp_('Tu sesión ya no es válida. Inicia sesión de nuevo.', 'SESION');
  }
  if (rolMinimo && NIVEL[u.rol] < NIVEL[rolMinimo]) {
    throw new ErrorApp_(rolMinimo === 'admin' ? 'Solo un administrador puede hacer esto.'
                                              : 'Tu usuario no tiene permiso para hacer esto.');
  }
  c.put('s_' + d.token, raw, segundosSesion_()); // renueva el tiempo
  return u;
}

/* ------------------------- CONTRASEÑAS ------------------------- */

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function sha256_(texto) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8));
}

function hashClave_(clave, salt) {
  let h = sha256_(salt + ':' + clave);
  for (let i = 0; i < CONFIG.ITERACIONES; i++) h = sha256_(h + salt);
  return h;
}

function verificarClave_(clave, u) {
  const a = hashClave_(clave, String(u.salt));
  const b = String(u.hash);
  if (a.length !== b.length) return false;
  let dif = 0;
  for (let i = 0; i < a.length; i++) dif |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return dif === 0;
}

function validarClave_(clave) {
  if (clave.length < CONFIG.LARGO_MIN_CLAVE) {
    throw new ErrorApp_('La contraseña debe tener al menos ' + CONFIG.LARGO_MIN_CLAVE + ' caracteres.');
  }
  if (clave.length > 100) throw new ErrorApp_('La contraseña es demasiado larga.');
}

function nuevaSal_() {
  return Utilities.getUuid().replace(/-/g, '');
}

function asignarClave_(u, clave, por) {
  const salt = nuevaSal_();
  const hash = hashClave_(clave, salt);
  guardarCampos_(u.fila, { salt: salt, hash: hash, actualizado: new Date(), actualizadoPor: por });
  u.salt = salt;
  u.hash = hash;
  return u;
}

function nuevoCodigo_() {
  const n = parseInt(sha256_(Utilities.getUuid() + Date.now()).slice(0, 12), 16) % 1000000;
  return ('00000' + n).slice(-6);
}

/* ------------------------- HOJA DE USUARIOS ------------------------- */

function libro_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new ErrorApp_('El sistema no está instalado. Ejecuta la función "instalar" en Apps Script.');
  return SpreadsheetApp.openById(id);
}

function hoja_(nombre) {
  nombre = nombre || HOJA_USUARIOS;
  const sh = libro_().getSheetByName(nombre);
  if (!sh) throw new ErrorApp_('No se encontró la hoja "' + nombre + '". Ejecuta "instalar" de nuevo.');
  return sh;
}

function leerUsuarios_() {
  const sh = hoja_(HOJA_USUARIOS);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, COLUMNAS.length).getValues().map(function (fila, i) {
    const u = { fila: i + 2 };
    COLUMNAS.forEach(function (col, j) { u[col] = fila[j]; });
    u.usuario = String(u.usuario || '').trim().toLowerCase();
    u.nombre = String(u.nombre || '');
    u.rol = ROLES.indexOf(u.rol) >= 0 ? u.rol : 'usuario';
    u.activo = u.activo === true || String(u.activo).toUpperCase() === 'TRUE';
    return u;
  }).filter(function (u) { return u.usuario; });
}

function buscarUsuario_(usuario) {
  const u = norm_(usuario);
  if (!u) return null;
  return leerUsuarios_().find(function (x) { return x.usuario === u; }) || null;
}

function guardarCampos_(fila, campos) {
  const sh = hoja_(HOJA_USUARIOS);
  Object.keys(campos).forEach(function (k) {
    sh.getRange(fila, COLUMNAS.indexOf(k) + 1).setValue(campos[k]);
  });
}

/** Usuario que no se puede desactivar, eliminar ni cambiar de rol: el admin principal y uno mismo. */
function objetivoProtegido_(usuario, admin) {
  const u = buscarUsuario_(usuario);
  if (!u) throw new ErrorApp_('El usuario no existe.');
  if (u.usuario === ADMIN_INICIAL.usuario) throw new ErrorApp_('El usuario principal "' + u.usuario + '" no se puede modificar así.');
  if (u.usuario === admin.usuario) throw new ErrorApp_('No puedes cambiar el estado, el rol ni eliminar tu propio usuario.');
  return u;
}

/* ------------------------- CORREOS ------------------------- */

function notificar_(accion, u, clave, por) {
  const filas = [
    ['Usuario', u.usuario],
    ['Nombre', u.nombre],
    ['Acción', accion],
    ['Realizado por', por],
    ['Fecha', fecha_(new Date())]
  ];
  if (CONFIG.INCLUIR_CLAVE_EN_CORREO && clave) filas.push(['Contraseña', clave]);
  const tabla = '<table cellpadding="8" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' +
    filas.map(function (f) {
      const valor = f[0] === 'Contraseña'
        ? '<code style="font-size:16px;background:#f1f5f9;padding:2px 6px">' + esc_(f[1]) + '</code>'
        : esc_(f[1]);
      return '<tr><td style="border:1px solid #ddd;background:#f7f7f7"><b>' + esc_(f[0]) + '</b></td>' +
             '<td style="border:1px solid #ddd">' + valor + '</td></tr>';
    }).join('') + '</table>';
  enviarCorreo_(accion + ' — ' + u.usuario, '<p>Aviso automático de cambios de acceso.</p>' + tabla, false);
}

function enviarCorreo_(asunto, html, obligatorio) {
  try {
    MailApp.sendEmail({
      to: CONFIG.CORREO_GENERAL,
      subject: '[' + CONFIG.NOMBRE_SISTEMA + '] ' + asunto,
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' + html + '</div>',
      name: CONFIG.NOMBRE_SISTEMA
    });
  } catch (e) {
    console.error('No se pudo enviar el correo: ' + e);
    if (obligatorio) throw new ErrorApp_('No se pudo enviar el correo. Intenta más tarde.');
  }
}

/* ------------------------- UTILIDADES ------------------------- */

function ErrorApp_(mensaje, codigo) {
  this.message = mensaje;
  this.codigo = codigo || '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cache_() { return CacheService.getScriptCache(); }

function bloquear_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  return lock;
}

function norm_(s) { return String(s || '').trim().toLowerCase(); }

function limpiarTexto_(s, max) {
  let t = String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t; // evita fórmulas en la hoja
  return t;
}

function perfil_(u) { return { usuario: u.usuario, nombre: u.nombre, rol: u.rol }; }

function fecha_(d) {
  return d instanceof Date ? Utilities.formatDate(d, CONFIG.ZONA, 'dd/MM/yyyy HH:mm') : '';
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
  });
}

function ocultarCorreo_(correo) {
  const p = correo.split('@');
  return p[0].slice(0, 4) + '•••@' + p[1];
}
