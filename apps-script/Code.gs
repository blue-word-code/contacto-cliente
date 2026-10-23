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
 *   7. Formularios: el administrador diseña formularios de preguntas; se llenan
 *      desde la ficha del cliente y cada respuesta queda en su historial.
 *
 *  Roles: usuario < editor < supervisor < admin  (ver PERMISOS más abajo).
 *
 *  Datos:
 *   - Hoja de cálculo "Contacto-Cliente — Datos" con las hojas
 *     Usuarios, Clientes, Actividad, Formularios y Respuestas.
 *   - Las hojas Formularios y Respuestas se crean solas la primera vez que se usan.
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
const HOJA_FORMULARIOS = 'Formularios';
const HOJA_RESPUESTAS = 'Respuestas';

const COLUMNAS = ['usuario', 'nombre', 'rol', 'activo', 'salt', 'hash',
                  'creado', 'actualizado', 'ultimoAcceso', 'actualizadoPor'];
// "oculto" va al final para no mover las columnas que ya existen.
const COL_CLI = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono',
                 'audioId', 'audioNombre', 'audioFecha', 'audioPor', 'actualizado', 'actualizadoPor', 'oculto'];
const CAMPOS_CLI = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono'];
const COL_ACT = ['fecha', 'usuario', 'accion', 'detalle'];
const COL_FORM = ['id', 'titulo', 'descripcion', 'activo', 'unaVez', 'preguntas',
                  'creado', 'creadoPor', 'actualizado', 'actualizadoPor'];
const COL_RES = ['id', 'fecha', 'formularioId', 'formulario', 'codigo', 'cliente', 'usuario', 'resumen', 'datos'];

const ROLES = ['usuario', 'editor', 'supervisor', 'admin'];
const NIVEL = { usuario: 0, editor: 1, supervisor: 2, admin: 3 };

/*
 * PERMISOS: rol mínimo para cada acción. Cambia el rol aquí si quieres
 * dar o quitar un permiso (usuario · editor · supervisor · admin).
 * La app lee esta misma tabla, así que los botones se ajustan solos.
 */
const PERMISOS = {
  responderFormularios: 'usuario',    // llenar formularios en la ficha del cliente
  editarClientes: 'editor',           // crear, editar e importar clientes
  agregarAudio: 'editor',             // grabar / subir nota a clientes que NO tienen nota
  verHistorialCompleto: 'supervisor', // ver en la ficha las respuestas de todos (los demás ven solo las suyas)
  verRespuestas: 'supervisor',        // ver y descargar en Excel las respuestas de cada formulario
  verActividad: 'supervisor',         // menú Actividad
  reemplazarAudio: 'admin',           // cambiar una nota de voz ya grabada por otra
  eliminarAudio: 'admin',
  ocultarClientes: 'admin',           // ocultar / volver a mostrar clientes y ver los ocultos
  disenarFormularios: 'admin',        // crear, editar, activar y eliminar formularios
  eliminarRespuestas: 'admin',
  eliminarClientes: 'admin',
  gestionarUsuarios: 'admin'
};

const TIPOS_PREGUNTA = ['texto', 'parrafo', 'numero', 'sino', 'opcion', 'varias', 'fecha', 'calificacion'];
const CACHE_CLI = 'cli2_'; // prefijo de la lista de clientes en caché
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
  } else {
    asegurarTitulosClientes_(sc); // agrega la columna "oculto" en instalaciones anteriores
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

  // Hojas de formularios y respuestas
  hojaFormularios_();
  hojaRespuestas_();

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
  ocultarCliente: ocultarCliente_,
  // Formularios e historial
  listarFormularios: listarFormularios_,
  guardarFormulario: guardarFormulario_,
  eliminarFormulario: eliminarFormulario_,
  responderFormulario: responderFormulario_,
  historialCliente: historialCliente_,
  respuestasFormulario: respuestasFormulario_,
  eliminarRespuesta: eliminarRespuesta_,
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
  let lista = clientesCompactos_();
  const iOculto = COMPACTO.indexOf('oculto');
  if (!permite_(u, 'ocultarClientes')) lista = lista.filter(function (f) { return !f[iOculto]; });
  const res = { campos: COMPACTO, clientes: lista };
  if (permite_(u, 'agregarAudio')) {
    const c = carpeta_(false);
    if (c) res.carpeta = { nombre: c.getName(), url: c.getUrl() };
  }
  return res;
}

function guardarCliente_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'editarClientes');
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
      if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('El cliente "' + original + '" ya no existe.');
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
    const u = requerirPermiso_(d, 'eliminarClientes');
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
    const u = requerirPermiso_(d, 'editarClientes');
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

/**
 * Oculta o vuelve a mostrar un cliente. Un cliente oculto no aparece en la lista,
 * la búsqueda ni el resumen de quienes no tienen el permiso "ocultarClientes".
 * No se borra nada: datos, nota de voz e historial se conservan.
 */
function ocultarCliente_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'ocultarClientes');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined) throw new ErrorApp_('El cliente no existe.');
    asegurarTitulosClientes_(t.hoja);
    const oculto = d.oculto === true;
    const fila = t.valores[idx];
    fila[ci_('oculto')] = oculto ? true : '';
    fila[ci_('actualizado')] = new Date();
    fila[ci_('actualizadoPor')] = u.usuario;
    t.hoja.getRange(idx + 2, 1, 1, COL_CLI.length).setValues([fila]);
    invalidarClientes_();
    registrar_(u.usuario, oculto ? 'Ocultó cliente' : 'Mostró cliente', codigo + ' — ' + fila[ci_('nombre')]);
    return { cliente: compactar_(fila), mensaje: oculto ? 'Cliente oculto.' : 'El cliente vuelve a estar visible.' };
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
  if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('El cliente no existe.');
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
  const u = requerirPermiso_(d, 'agregarAudio');
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
    if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('No existe un cliente con el código "' + codigo + '".');
    const fila = t.valores[idx];
    const anterior = fila[ci_('audioId')];
    if (anterior) exigirReemplazo_(u, codigo, d.reemplazar);
    const nombreArchivo = codigo.replace(/[\/\\]/g, '-') + '.' + ext;
    const archivo = carpeta_(true).createFile(Utilities.newBlob(bytes, tipoAudio_(nombreArchivo, d.mime), nombreArchivo));
    archivo.setDescription(String(fila[ci_('nombre')] || ''));
    if (anterior) archivar_(anterior);
    fila[ci_('audioId')] = archivo.getId();
    fila[ci_('audioNombre')] = nombreArchivo;
    fila[ci_('audioFecha')] = new Date();
    fila[ci_('audioPor')] = u.usuario + (d.grabado ? ' (grabado en la app)' : '');
    escribirAudio_(t, idx);
    invalidarClientes_();
    registrar_(u.usuario, anterior ? 'Reemplazó nota de voz' : (d.grabado ? 'Grabó nota de voz' : 'Subió nota de voz'),
               codigo + ' — ' + fila[ci_('nombre')] + (d.grabado && anterior ? ' (grabada en la app)' : ''));
    return { cliente: compactar_(fila),
             mensaje: anterior ? 'Nota de voz reemplazada. La anterior quedó en la carpeta "Notas reemplazadas".' : 'Nota de voz guardada.' };
  } finally { lock.releaseLock(); }
}

function eliminarAudio_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'eliminarAudio');
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
  const u = requerirPermiso_(d, 'agregarAudio');
  if (d.reemplazar && !permite_(u, 'reemplazarAudio')) {
    throw new ErrorApp_('Solo un ' + nombreRol_(PERMISOS.reemplazarAudio) + ' puede reemplazar notas de voz que ya existen.');
  }
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
        if (fila[ci_('audioId')] && fila[ci_('audioId')] !== c.id) {
          if (!d.reemplazar) return; // alguien le puso nota mientras tanto
          archivar_(fila[ci_('audioId')]);
        }
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
    const u = requerirPermiso_(d, 'agregarAudio');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('No existe un cliente con el código "' + codigo + '".');
    let archivo;
    try { archivo = DriveApp.getFileById(String(d.fileId)); } catch (e) { throw new ErrorApp_('No se encontró el archivo en Drive.'); }
    const fila = t.valores[idx];
    const anterior = fila[ci_('audioId')];
    if (anterior && anterior !== archivo.getId()) exigirReemplazo_(u, codigo, d.reemplazar);
    if (anterior && anterior !== archivo.getId()) archivar_(anterior);
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
  requerirPermiso_(d, 'verActividad');
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

/* =========================================================
 *  MÓDULO 7: FORMULARIOS E HISTORIAL DEL CLIENTE
 * ========================================================= */

/**
 * Formularios disponibles. Quien diseña formularios recibe todos (activos e inactivos)
 * y quien ve respuestas recibe además cuántas tiene cada uno; los demás solo los activos.
 */
function listarFormularios_(d) {
  const u = requerirSesion_(d);
  const disenar = permite_(u, 'disenarFormularios');
  let lista = leerFormularios_();
  if (!disenar) lista = lista.filter(function (f) { return f.activo; });
  if (permite_(u, 'verRespuestas')) {
    const conteo = {};
    const sh = hojaRespuestas_();
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      sh.getRange(2, COL_RES.indexOf('formularioId') + 1, n, 1).getValues().forEach(function (f) {
        conteo[f[0]] = (conteo[f[0]] || 0) + 1;
      });
    }
    lista.forEach(function (f) { f.respuestas = conteo[f.id] || 0; });
  }
  return { formularios: lista.map(function (f) { delete f.fila; return f; }) };
}

/** Crea o actualiza un formulario. Las respuestas ya guardadas no cambian (guardan su propia copia de las preguntas). */
function guardarFormulario_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'disenarFormularios');
    const f = d.formulario || {};
    const titulo = limpiarTexto_(f.titulo, 100);
    if (!titulo) throw new ErrorApp_('Escribe el título del formulario.');
    const descripcion = textoLibre_(f.descripcion, 500);
    const brutas = Array.isArray(f.preguntas) ? f.preguntas : [];
    if (!brutas.length) throw new ErrorApp_('Agrega al menos una pregunta.');
    if (brutas.length > 50) throw new ErrorApp_('Máximo 50 preguntas por formulario.');
    const usados = {};
    const preguntas = brutas.map(function (p, i) {
      const texto = textoLibre_(p && p.texto, 200).replace(/\s+/g, ' ');
      if (!texto) throw new ErrorApp_('La pregunta ' + (i + 1) + ' no tiene texto.');
      const tipo = TIPOS_PREGUNTA.indexOf(p.tipo) >= 0 ? p.tipo : 'texto';
      let id = /^[a-z0-9]{4,12}$/.test(String(p.id || '')) ? String(p.id) : nuevoId_('', 6).toLowerCase();
      while (usados[id]) id = nuevoId_('', 6).toLowerCase();
      usados[id] = true;
      const q = { id: id, texto: texto, tipo: tipo, requerida: p.requerida === true };
      if (tipo === 'opcion' || tipo === 'varias') {
        const vistas = {};
        const ops = (Array.isArray(p.opciones) ? p.opciones : String(p.opciones || '').split('\n'))
          .map(function (o) { return textoLibre_(o, 100).replace(/\s+/g, ' '); })
          .filter(function (o) { if (!o || vistas[o.toLowerCase()]) return false; vistas[o.toLowerCase()] = true; return true; });
        if (ops.length < 2) throw new ErrorApp_('La pregunta "' + texto + '" necesita al menos 2 opciones (una por línea).');
        if (ops.length > 30) throw new ErrorApp_('La pregunta "' + texto + '" tiene más de 30 opciones.');
        q.opciones = ops;
      }
      return q;
    });
    const json = JSON.stringify(preguntas);
    if (json.length > 45000) throw new ErrorApp_('El formulario es demasiado largo.');
    const sh = hojaFormularios_();
    const ahora = new Date();
    const existentes = leerFormularios_();
    const previo = f.id ? existentes.find(function (x) { return x.id === String(f.id); }) : null;
    if (f.id && !previo) throw new ErrorApp_('El formulario ya no existe.');
    const repetido = existentes.find(function (x) { return x.titulo.toLowerCase() === titulo.toLowerCase() && (!previo || x.id !== previo.id); });
    if (repetido) throw new ErrorApp_('Ya existe un formulario llamado "' + titulo + '".');
    const activo = f.activo !== false;
    const unaVez = f.unaVez === true;
    let id;
    if (previo) {
      id = previo.id;
      sh.getRange(previo.fila, 2, 1, 5).setValues([[titulo, descripcion, activo, unaVez, json]]);
      sh.getRange(previo.fila, COL_FORM.indexOf('actualizado') + 1, 1, 2).setValues([[ahora, u.usuario]]);
      registrar_(u.usuario, 'Editó formulario', titulo);
    } else {
      id = nuevoId_('F', 8);
      sh.appendRow([id, titulo, descripcion, activo, unaVez, json, ahora, u.usuario, ahora, u.usuario]);
      registrar_(u.usuario, 'Creó formulario', titulo + ' (' + preguntas.length + ' preguntas)');
    }
    return { formulario: { id: id, titulo: titulo, descripcion: descripcion, activo: activo, unaVez: unaVez,
                           preguntas: preguntas, actualizado: fecha_(ahora) },
             mensaje: previo ? 'Formulario actualizado.' : 'Formulario creado.' };
  } finally { lock.releaseLock(); }
}

/** Elimina el formulario. Sus respuestas se quedan en el historial de cada cliente. */
function eliminarFormulario_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'disenarFormularios');
    const f = leerFormularios_().find(function (x) { return x.id === String(d.id || ''); });
    if (!f) throw new ErrorApp_('El formulario no existe.');
    hojaFormularios_().deleteRow(f.fila);
    registrar_(u.usuario, 'Eliminó formulario', f.titulo);
    return { mensaje: 'Formulario eliminado. Sus respuestas siguen en el historial de los clientes.' };
  } finally { lock.releaseLock(); }
}

/** Guarda las respuestas de un formulario para un cliente (queda en su historial). */
function responderFormulario_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'responderFormularios');
    const codigo = normCodigo_(d.codigo);
    const t = tablaClientes_();
    const idx = t.indice[codigo];
    if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('El cliente no existe.');
    const nombreCliente = String(t.valores[idx][ci_('nombre')] || '');
    const form = leerFormularios_().find(function (x) { return x.id === String(d.formularioId || ''); });
    if (!form || !form.activo) throw new ErrorApp_('Este formulario ya no está disponible. Actualiza la página.');
    if (form.unaVez && respuestasDe_(codigo).some(function (r) { return r.formularioId === form.id; })) {
      throw new ErrorApp_('Este cliente ya respondió el formulario "' + form.titulo + '" (solo se responde una vez).');
    }
    const entrada = d.respuestas || {};
    const datos = form.preguntas.map(function (q) {
      const v = validarRespuesta_(q, entrada[q.id]);
      return { id: q.id, p: q.texto, t: q.tipo, v: v };
    });
    if (!datos.some(function (x) { return !vacio_(x.v); })) throw new ErrorApp_('Responde al menos una pregunta.');
    const json = JSON.stringify(datos);
    if (json.length > 45000) throw new ErrorApp_('Las respuestas son demasiado largas.');
    const resumen = datos.filter(function (x) { return !vacio_(x.v); })
      .map(function (x) { return x.p + ': ' + valorTexto_(x.t, x.v); }).join(' | ');
    const ahora = new Date();
    const id = nuevoId_('R', 10);
    const sh = hojaRespuestas_();
    const fila = sh.getLastRow() + 1;
    if (fila > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 500);
    sh.getRange(fila, COL_RES.indexOf('codigo') + 1).setNumberFormat('@');
    sh.getRange(fila, 1, 1, COL_RES.length).setValues([[id, ahora, form.id, form.titulo, codigo, nombreCliente,
      u.usuario, limpiarTexto_(resumen, 45000), json]]);
    registrar_(u.usuario, 'Llenó formulario', form.titulo + ' — ' + codigo + ' ' + nombreCliente);
    return { respuesta: { id: id, fecha: fecha_(ahora), formularioId: form.id, formulario: form.titulo,
                          usuario: u.usuario, respuestas: datos, propia: true },
             mensaje: 'Respuestas guardadas en el historial del cliente.' };
  } finally { lock.releaseLock(); }
}

/** Historial de formularios de un cliente (más reciente primero). */
function historialCliente_(d) {
  const u = requerirSesion_(d);
  const codigo = normCodigo_(d.codigo);
  const t = tablaClientes_();
  const idx = t.indice[codigo];
  if (idx === undefined || !clienteVisible_(u, t.valores[idx])) throw new ErrorApp_('El cliente no existe.');
  const todas = respuestasDe_(codigo);
  const completo = permite_(u, 'verHistorialCompleto');
  const respondidos = {};
  todas.forEach(function (r) { respondidos[r.formularioId] = r.fecha; });
  const visibles = todas.filter(function (r) { return completo || r.usuario === u.usuario; });
  visibles.forEach(function (r) { r.propia = r.usuario === u.usuario; });
  return { historial: visibles, respondidos: respondidos, completo: completo, ocultas: todas.length - visibles.length };
}

/** Todas las respuestas de un formulario (para ver en pantalla o descargar en Excel). */
function respuestasFormulario_(d) {
  const u = requerirPermiso_(d, 'verRespuestas');
  const id = String(d.id || '');
  if (!id) throw new ErrorApp_('Falta el formulario.');
  const sh = hojaRespuestas_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { respuestas: [] };
  const col = COL_RES.indexOf('formularioId') + 1;
  const encontrados = sh.getRange(2, col, n, 1).createTextFinder(id).matchEntireCell(true).findAll();
  const ocultos = {};
  if (!permite_(u, 'ocultarClientes')) {
    tablaClientes_().valores.forEach(function (f) { if (esVerdad_(f[ci_('oculto')])) ocultos[f[0]] = true; });
  }
  const lista = leerRespuestas_(sh, encontrados.map(function (r) { return r.getRow(); }))
    .filter(function (r) { return !ocultos[r.codigo]; })
    .sort(porRecientes_);
  lista.forEach(function (r) { delete r._t; delete r._f; });
  registrar_(u.usuario, 'Consultó respuestas de formulario', (lista[0] && lista[0].formulario) || id);
  return { respuestas: lista };
}

function eliminarRespuesta_(d) {
  const lock = bloquear_();
  try {
    const u = requerirPermiso_(d, 'eliminarRespuestas');
    const sh = hojaRespuestas_();
    const n = sh.getLastRow() - 1;
    const r = n > 0 ? sh.getRange(2, 1, n, 1).createTextFinder(String(d.id || '')).matchEntireCell(true).findNext() : null;
    if (!r || !d.id) throw new ErrorApp_('La respuesta ya no existe.');
    const datos = leerRespuesta_(sh, r.getRow());
    sh.deleteRow(r.getRow());
    registrar_(u.usuario, 'Eliminó respuesta de formulario', (datos ? datos.formulario + ' — ' + datos.codigo : d.id));
    return { mensaje: 'Respuesta eliminada del historial.' };
  } finally { lock.releaseLock(); }
}

/* ---- apoyo de formularios ---- */

function hojaFormularios_() { return asegurarHoja_(HOJA_FORMULARIOS, COL_FORM, []); }
function hojaRespuestas_() { return asegurarHoja_(HOJA_RESPUESTAS, COL_RES, ['codigo']); }

/** Devuelve la hoja; si no existe la crea con sus títulos (columnas de texto indicadas). */
function asegurarHoja_(nombre, columnas, deTexto) {
  const ss = libro_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(columnas);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, columnas.length).setFontWeight('bold');
    deTexto.forEach(function (c) { sh.getRange(1, columnas.indexOf(c) + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}

function leerFormularios_() {
  const sh = hojaFormularios_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COL_FORM.length).getValues().map(function (f, i) {
    let preguntas = [];
    try { preguntas = JSON.parse(f[5] || '[]'); } catch (e) { preguntas = []; }
    return { fila: i + 2, id: String(f[0]), titulo: String(f[1] || ''), descripcion: String(f[2] || ''),
             activo: esVerdad_(f[3]), unaVez: esVerdad_(f[4]), preguntas: preguntas, actualizado: fecha_(f[8]) };
  }).filter(function (f) { return f.id; });
}

/** Lee varias filas de Respuestas: pocas, una por una; muchas, en un solo bloque. */
function leerRespuestas_(sh, filas) {
  if (!filas.length) return [];
  if (filas.length <= 8) return filas.map(function (n) { return leerRespuesta_(sh, n); }).filter(Boolean);
  const min = Math.min.apply(null, filas), max = Math.max.apply(null, filas);
  const bloque = sh.getRange(min, 1, max - min + 1, COL_RES.length).getValues();
  return filas.map(function (n) { return convertirRespuesta_(bloque[n - min], n); }).filter(Boolean);
}

function leerRespuesta_(sh, fila) {
  return convertirRespuesta_(sh.getRange(fila, 1, 1, COL_RES.length).getValues()[0], fila);
}

/** Más reciente primero (a igual hora, la fila de más abajo es la más nueva). */
function porRecientes_(a, b) { return (b._t - a._t) || (b._f - a._f); }

function convertirRespuesta_(f, fila) {
  if (!f || !f[0]) return null;
  let datos = [];
  try { datos = JSON.parse(f[8] || '[]'); } catch (e) { datos = []; }
  return { id: String(f[0]), fecha: fecha_(f[1]), _t: f[1] instanceof Date ? f[1].getTime() : 0, _f: fila || 0,
           formularioId: String(f[2]), formulario: String(f[3]), codigo: normCodigo_(f[4]), cliente: String(f[5]),
           usuario: String(f[6]), respuestas: datos };
}

/** Respuestas de un cliente, más reciente primero. Usa el buscador de la hoja (rápido aunque haya miles). */
function respuestasDe_(codigo) {
  const sh = hojaRespuestas_();
  const n = sh.getLastRow() - 1;
  if (n < 1 || !codigo) return [];
  const col = COL_RES.indexOf('codigo') + 1;
  const encontrados = sh.getRange(2, col, n, 1).createTextFinder(codigo).matchEntireCell(true).findAll();
  const lista = leerRespuestas_(sh, encontrados.map(function (r) { return r.getRow(); }))
    .filter(function (r) { return r.codigo === codigo; })
    .sort(porRecientes_);
  lista.forEach(function (r) { delete r._t; delete r._f; });
  return lista;
}

function vacio_(v) { return v === '' || v === null || v === undefined || (Array.isArray(v) && !v.length); }

/** Valida y normaliza una respuesta según el tipo de pregunta. */
function validarRespuesta_(q, v) {
  let out = '';
  const falta = function () { return new ErrorApp_('Falta responder: "' + q.texto + '".'); };
  switch (q.tipo) {
    case 'parrafo': out = textoLibre_(v, 2000); break;
    case 'numero': {
      const s = String(v == null ? '' : v).trim().replace(/\s/g, '').replace(',', '.');
      if (s) {
        const num = Number(s);
        if (!isFinite(num)) throw new ErrorApp_('"' + q.texto + '" debe ser un número.');
        out = num;
      }
      break;
    }
    case 'sino': {
      const s = norm_(v);
      out = s === 'si' || s === 'sí' ? 'si' : (s === 'no' ? 'no' : '');
      break;
    }
    case 'opcion': {
      const s = String(v == null ? '' : v);
      if (s && q.opciones.indexOf(s) < 0) throw new ErrorApp_('Opción no válida en "' + q.texto + '".');
      out = s;
      break;
    }
    case 'varias': {
      const arr = Array.isArray(v) ? v.map(String) : [];
      arr.forEach(function (s) { if (q.opciones.indexOf(s) < 0) throw new ErrorApp_('Opción no válida en "' + q.texto + '".'); });
      out = q.opciones.filter(function (o) { return arr.indexOf(o) >= 0; });
      break;
    }
    case 'fecha': {
      const s = String(v == null ? '' : v).trim();
      if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new ErrorApp_('"' + q.texto + '" debe ser una fecha.');
      out = s;
      break;
    }
    case 'calificacion': {
      const s = String(v == null ? '' : v).trim();
      if (s) {
        const num = Number(s);
        if ([1, 2, 3, 4, 5].indexOf(num) < 0) throw new ErrorApp_('"' + q.texto + '" debe ser de 1 a 5.');
        out = num;
      }
      break;
    }
    default: out = textoLibre_(v, 300).replace(/\s+/g, ' ');
  }
  if (q.requerida && vacio_(out)) throw falta();
  return out;
}

function valorTexto_(tipo, v) {
  if (vacio_(v)) return '';
  if (tipo === 'sino') return v === 'si' ? 'Sí' : 'No';
  if (tipo === 'varias') return v.join(', ');
  if (tipo === 'calificacion') return v + '/5';
  if (tipo === 'fecha') { const p = String(v).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  return String(v);
}

/** Texto libre: conserva saltos de línea, quita caracteres de control. */
function textoLibre_(s, max) {
  return String(s == null ? '' : s).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000B-\u001F\u007F]/g, '')
    .replace(/\n{3,}/g, '\n\n').trim().slice(0, max);
}

function nuevoId_(prefijo, largo) {
  return prefijo + Utilities.getUuid().replace(/-/g, '').slice(0, largo).toUpperCase();
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

const COMPACTO = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono', 'audio', 'audioFecha', 'audioPor', 'oculto'];

function ci_(campo) { return COL_CLI.indexOf(campo); }

function esVerdad_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === 1 || v === '1'; }

/** Un cliente oculto solo lo ven quienes tienen el permiso "ocultarClientes". */
function clienteVisible_(u, fila) { return !esVerdad_(fila[ci_('oculto')]) || permite_(u, 'ocultarClientes'); }

/** Pone los títulos de la hoja Clientes (agrega "oculto" si la hoja es de una versión anterior). */
function asegurarTitulosClientes_(sh) {
  const actual = sh.getRange(1, 1, 1, COL_CLI.length).getValues()[0];
  if (actual[COL_CLI.length - 1] !== COL_CLI[COL_CLI.length - 1]) {
    sh.getRange(1, 1, 1, COL_CLI.length).setValues([COL_CLI]).setFontWeight('bold');
  }
}

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
          f[ci_('audioId')] ? 1 : 0, fecha_(f[ci_('audioFecha')]), String(f[ci_('audioPor')] || ''),
          esVerdad_(f[ci_('oculto')]) ? 1 : 0];
}

/** Lista compacta con caché por trozos (CacheService admite 100 KB por clave). */
function clientesCompactos_() {
  const c = cache_();
  const n = Number(c.get(CACHE_CLI + 'n') || 0);
  if (n > 0) {
    const claves = [];
    for (let i = 0; i < n; i++) claves.push(CACHE_CLI + i);
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
  for (let i = 0; i < texto.length; i += TAM) pares[CACHE_CLI + (k++)] = texto.slice(i, i + TAM);
  try {
    if (k <= 900) { c.putAll(pares, 21600); c.put(CACHE_CLI + 'n', String(k), 21600); }
  } catch (e) { console.error('No se pudo guardar la lista en caché: ' + e); }
  return lista;
}

function invalidarClientes_() { cache_().remove(CACHE_CLI + 'n'); }

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

/**
 * Nota reemplazada: en vez de borrarla, se mueve a la subcarpeta "Notas reemplazadas"
 * (dentro de la carpeta de notas de voz) con la fecha en el nombre, por si hay que recuperarla.
 */
function archivar_(id) {
  try {
    const f = DriveApp.getFileById(id);
    const sello = Utilities.formatDate(new Date(), CONFIG.ZONA, 'yyyy-MM-dd HHmm');
    const nombre = f.getName();
    const m = nombre.match(/^(.*?)(\.[a-z0-9]{2,5})?$/i);
    f.setName(m[1] + ' (reemplazada ' + sello + ')' + (m[2] || ''));
    f.moveTo(carpetaReemplazadas_());
  } catch (e) {
    console.error('No se pudo archivar la nota anterior, va a la papelera: ' + e);
    papelera_(id);
  }
}

function carpetaReemplazadas_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CARPETA_REEMPLAZADAS_ID');
  if (id) {
    try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* se busca o crea */ }
  }
  const madre = carpeta_(true);
  const it = madre.getFoldersByName('Notas reemplazadas');
  const sub = it.hasNext() ? it.next() : madre.createFolder('Notas reemplazadas');
  props.setProperty('CARPETA_REEMPLAZADAS_ID', sub.getId());
  return sub;
}

/** Si el cliente ya tiene nota: pide confirmación (YA_TIENE) y exige el permiso de reemplazo. */
function exigirReemplazo_(u, codigo, reemplazar) {
  if (!permite_(u, 'reemplazarAudio')) {
    throw new ErrorApp_('El cliente ' + codigo + ' ya tiene nota de voz. Solo un ' + nombreRol_(PERMISOS.reemplazarAudio) +
                        ' puede cambiarla por otra.', 'SIN_PERMISO');
  }
  if (!reemplazar) throw new ErrorApp_('El cliente ' + codigo + ' ya tiene nota de voz.', 'YA_TIENE');
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

/** ¿El usuario tiene este permiso? (ver tabla PERMISOS al inicio). */
function permite_(u, permiso) {
  const minimo = PERMISOS[permiso] || 'admin';
  return (NIVEL[u.rol] || 0) >= NIVEL[minimo];
}

/** Valida la sesión y el permiso indicado. */
function requerirPermiso_(d, permiso) {
  const u = requerirSesion_(d);
  if (!permite_(u, permiso)) {
    const minimo = PERMISOS[permiso] || 'admin';
    throw new ErrorApp_(minimo === 'admin' ? 'Solo un administrador puede hacer esto.'
                                           : 'Tu usuario no tiene permiso para hacer esto (se necesita rol ' + nombreRol_(minimo) + ' o superior).');
  }
  return u;
}

function nombreRol_(r) {
  return { usuario: 'usuario', editor: 'editor', supervisor: 'supervisor', admin: 'administrador' }[r] || r;
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

function perfil_(u) {
  const permisos = {};
  Object.keys(PERMISOS).forEach(function (p) { if (permite_(u, p)) permisos[p] = true; });
  return { usuario: u.usuario, nombre: u.nombre, rol: u.rol, permisos: permisos };
}

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
