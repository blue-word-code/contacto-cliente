/*
 * Contacto-Cliente — Acceso (login, usuarios, recuperación)
 * La URL del backend (Apps Script) se configura en assets/config.js.
 *
 * Para el contenido de la app (index.html):
 *   - Los elementos con  data-pa-nombre  muestran el nombre del usuario.
 *   - Los elementos con  data-pa-solo-admin  solo se ven si el usuario es administrador.
 *   - Los elementos con  data-pa-min-rol="editor"  solo se ven para ese rol o uno superior
 *     (usuario < editor < supervisor < admin).
 *   - Los elementos con  data-pa-permiso="verActividad"  solo se ven si el usuario tiene ese permiso
 *     (la tabla de permisos está en Code.gs → PERMISOS).
 *   - Evento  "pa:sesion"  en document al iniciar sesión (detail = {usuario, nombre, rol}).
 *   - window.PA.api(accion, datos)  llama al backend con la sesión actual.
 */
(function () {
  'use strict';

  var CFG = window.APP_CONFIG || {};
  var API_URL = String(CFG.apiUrl || '').trim();
  var SISTEMA = CFG.nombreSistema || 'Contacto-Cliente';
  var LOGO = CFG.logo || 'CC';
  var NIVEL = { usuario: 0, editor: 1, supervisor: 2, admin: 3 };
  var ROLES = ['usuario', 'editor', 'supervisor', 'admin'];
  var NOMBRE_ROL = { usuario: 'Usuario', editor: 'Editor', supervisor: 'Supervisor', admin: 'Administrador' };
  function nombreRol(r) { return NOMBRE_ROL[r] || 'Usuario'; }
  var CLAVE_TOKEN = 'pa_token';
  var LARGO_MIN = 8;
  var usuarioActual = null;

  /* ---------------- utilidades ---------------- */

  function h(tag, attrs, hijos) {
    var el = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.indexOf('on') === 0) el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    [].concat(hijos || []).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return el;
  }

  function guardarToken(t) { try { t ? sessionStorage.setItem(CLAVE_TOKEN, t) : sessionStorage.removeItem(CLAVE_TOKEN); } catch (e) {} }
  function leerToken() { try { return sessionStorage.getItem(CLAVE_TOKEN); } catch (e) { return null; } }

  function apiConfigurada() {
    return /^https:\/\/script\.google(usercontent)?\.com\//.test(API_URL);
  }

  function api(accion, datos) {
    datos = datos || {};
    datos.accion = accion;
    var token = leerToken();
    if (token && !datos.token) datos.token = token;
    // text/plain (por defecto) evita la verificación CORS previa de Apps Script.
    return fetch(API_URL, { method: 'POST', body: JSON.stringify(datos), redirect: 'follow' })
      .then(function (r) { return r.json(); })
      .catch(function () {
        return { ok: false, error: 'No hay conexión con el servidor. Revisa tu internet e intenta de nuevo.' };
      })
      .then(function (res) {
        if (!res.ok && res.codigo === 'SESION' && accion !== 'sesion') {
          cerrarLocal();
          mostrarLogin(res.error);
        }
        return res;
      });
  }

  function generarClave(largo) {
    largo = largo || 12;
    var may = 'ABCDEFGHJKLMNPQRSTUVWXYZ', min = 'abcdefghijkmnpqrstuvwxyz', num = '23456789', sim = '!@#$%*?';
    var todos = may + min + num + sim;
    function al(conj) {
      var lim = 256 - (256 % conj.length), b = new Uint8Array(1);
      do { crypto.getRandomValues(b); } while (b[0] >= lim);
      return conj[b[0] % conj.length];
    }
    var chars = [al(may), al(min), al(num), al(sim)];
    while (chars.length < largo) chars.push(al(todos));
    for (var i = chars.length - 1; i > 0; i--) { // mezcla
      var b = new Uint32Array(1); crypto.getRandomValues(b);
      var j = b[0] % (i + 1), t = chars[i]; chars[i] = chars[j]; chars[j] = t;
    }
    return chars.join('');
  }

  function copiar(texto, boton) {
    var listo = function () {
      if (!boton) return;
      var orig = boton.textContent;
      boton.textContent = '¡Copiada!';
      setTimeout(function () { boton.textContent = orig; }, 1500);
    };
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(texto).then(listo, function () { copiarViejo(texto); listo(); });
    } else { copiarViejo(texto); listo(); }
  }
  function copiarViejo(texto) {
    var ta = h('textarea', { style: 'position:fixed;opacity:0' }); ta.value = texto;
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } catch (e) {}
    ta.remove();
  }

  function aviso(msg, tipo) {
    var cont = document.getElementById('pa-avisos') || document.body.appendChild(h('div', { id: 'pa-avisos' }));
    var t = h('div', { class: 'pa-aviso pa-aviso-' + (tipo || 'ok'), role: 'status', text: msg });
    cont.appendChild(t);
    setTimeout(function () { t.classList.add('pa-salir'); setTimeout(function () { t.remove(); }, 300); }, 3500);
  }

  function mensaje(cont, texto, tipo) {
    cont.textContent = texto || '';
    cont.className = 'pa-msg' + (texto ? ' pa-msg-' + (tipo || 'error') : '');
  }

  function ocupado(boton, si, texto) {
    if (si) { boton.dataset.t = boton.textContent; boton.textContent = texto || 'Procesando…'; boton.disabled = true; }
    else { boton.textContent = boton.dataset.t || boton.textContent; boton.disabled = false; }
  }

  function campo(etiqueta, input, ayuda) {
    return h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: etiqueta }), input,
      ayuda ? h('small', { class: 'pa-ayuda', text: ayuda }) : null]);
  }

  function entrada(attrs) {
    attrs = attrs || {};
    attrs.class = 'pa-input' + (attrs.class ? ' ' + attrs.class : '');
    return h('input', attrs);
  }

  /* -------- campo de contraseña: escribirla o generarla -------- */

  function campoClave(opts) {
    opts = opts || {};
    var modo = 'manual';
    var nombreGrupo = 'pa-modo-' + Math.random().toString(36).slice(2);
    var input = entrada({ type: 'password', autocomplete: 'new-password', placeholder: 'Mínimo ' + LARGO_MIN + ' caracteres' });
    var confirmar = entrada({ type: 'password', autocomplete: 'new-password', placeholder: 'Repite la contraseña' });
    var ver = h('button', { type: 'button', class: 'pa-btn-icono', title: 'Mostrar u ocultar', text: 'Ver' });
    var otra = h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: '↻ Generar otra' });
    var copiarBtn = h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Copiar' });
    var campoConf = campo('Confirmar contraseña', confirmar);
    var accionesGen = h('div', { class: 'pa-fila pa-oculto' }, [otra, copiarBtn]);

    ver.addEventListener('click', function () {
      var tipo = input.type === 'password' ? 'text' : 'password';
      input.type = tipo; confirmar.type = tipo;
      ver.textContent = tipo === 'password' ? 'Ver' : 'Ocultar';
    });
    otra.addEventListener('click', function () { input.value = generarClave(); });
    copiarBtn.addEventListener('click', function () { copiar(input.value, copiarBtn); });

    function poner(m) {
      modo = m;
      var gen = m === 'gen';
      input.readOnly = gen;
      input.type = gen ? 'text' : 'password';
      ver.textContent = gen ? 'Ocultar' : 'Ver';
      ver.classList.toggle('pa-oculto', gen);
      input.value = gen ? generarClave() : '';
      confirmar.value = '';
      campoConf.classList.toggle('pa-oculto', gen);
      accionesGen.classList.toggle('pa-oculto', !gen);
    }

    var hijos = [];
    if (opts.generar !== false) {
      var r1 = h('input', { type: 'radio', name: nombreGrupo, value: 'manual', checked: true, onchange: function () { poner('manual'); } });
      var r2 = h('input', { type: 'radio', name: nombreGrupo, value: 'gen', onchange: function () { poner('gen'); } });
      hijos.push(h('div', { class: 'pa-modos', role: 'radiogroup' }, [
        h('label', { class: 'pa-modo' }, [r1, ' Escribirla']),
        h('label', { class: 'pa-modo' }, [r2, ' Generarla automáticamente'])
      ]));
    }
    hijos.push(h('label', { class: 'pa-campo' }, [
      h('span', { class: 'pa-etq', text: opts.etiqueta || 'Contraseña' }),
      h('div', { class: 'pa-con-icono' }, [input, ver])
    ]));
    hijos.push(accionesGen, campoConf);

    return {
      el: h('div', { class: 'pa-clave' }, hijos),
      valor: function () { return input.value; },
      generada: function () { return modo === 'gen'; },
      validar: function () {
        if (input.value.length < LARGO_MIN) return 'La contraseña debe tener al menos ' + LARGO_MIN + ' caracteres.';
        if (modo === 'manual' && input.value !== confirmar.value) return 'Las contraseñas no coinciden.';
        return '';
      },
      enfocar: function () { input.focus(); }
    };
  }

  /* Caja que muestra una contraseña recién asignada, con botón Copiar. */
  function cajaResultado(titulo, usuario, clave) {
    var btn = h('button', { type: 'button', class: 'pa-btn pa-btn-sm', text: 'Copiar' });
    btn.addEventListener('click', function () { copiar('Usuario: ' + usuario + '\nContraseña: ' + clave, btn); });
    return h('div', { class: 'pa-resultado' }, [
      h('strong', { text: titulo }),
      h('div', { class: 'pa-cred' }, [h('span', { text: 'Usuario: ' }), h('code', { text: usuario })]),
      h('div', { class: 'pa-cred' }, [h('span', { text: 'Contraseña: ' }), h('code', { text: clave })]),
      h('div', { class: 'pa-fila' }, [btn]),
      h('small', { class: 'pa-ayuda', text: 'También se envió un aviso al correo general.' })
    ]);
  }

  /* ---------------- pantalla de acceso ---------------- */

  function raiz() {
    var r = document.getElementById('pa-auth');
    if (!r) { r = h('div', { id: 'pa-auth' }); document.body.appendChild(r); }
    return r;
  }

  function tarjeta(titulo, subtitulo, contenido) {
    var r = raiz();
    r.innerHTML = '';
    r.appendChild(h('div', { class: 'pa-tarjeta', role: 'dialog', 'aria-modal': 'true', 'aria-label': titulo }, [
      h('div', { class: 'pa-logo', 'aria-hidden': 'true', text: LOGO }),
      h('p', { class: 'pa-sistema', text: SISTEMA }),
      h('h1', { class: 'pa-titulo', text: titulo }),
      subtitulo ? h('p', { class: 'pa-sub', text: subtitulo }) : null,
      contenido
    ]));
  }

  function mostrarLogin(msgInicial, tipoMsg) {
    document.documentElement.classList.add('pa-locked');
    quitarBarra();
    var usuario = entrada({ type: 'text', autocomplete: 'username', autocapitalize: 'none', spellcheck: 'false', required: true });
    var clave = entrada({ type: 'password', autocomplete: 'current-password', required: true });
    var ver = h('button', { type: 'button', class: 'pa-btn-icono', text: 'Ver', title: 'Mostrar u ocultar' });
    ver.addEventListener('click', function () { clave.type = clave.type === 'password' ? 'text' : 'password'; ver.textContent = clave.type === 'password' ? 'Ver' : 'Ocultar'; });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var entrar = h('button', { type: 'submit', class: 'pa-btn pa-btn-bloque', text: 'Entrar' });
    var olvide = h('button', { type: 'button', class: 'pa-enlace', text: '¿Olvidaste tu contraseña?' });
    olvide.addEventListener('click', function () { mostrarOlvide(usuario.value); });

    var form = h('form', { class: 'pa-form', novalidate: true }, [
      campo('Usuario', usuario),
      h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Contraseña' }), h('div', { class: 'pa-con-icono' }, [clave, ver])]),
      msg, entrar, h('div', { class: 'pa-centro' }, [olvide])
    ]);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!usuario.value.trim() || !clave.value) { mensaje(msg, 'Escribe tu usuario y contraseña.'); return; }
      ocupado(entrar, true, 'Entrando…'); mensaje(msg, '');
      api('login', { usuario: usuario.value, clave: clave.value }).then(function (res) {
        ocupado(entrar, false);
        if (!res.ok) { mensaje(msg, res.error); clave.select(); return; }
        guardarToken(res.token);
        desbloquear(res.usuario);
        aviso('Bienvenido, ' + res.usuario.nombre);
      });
    });

    tarjeta('Iniciar sesión', 'Acceso solo para personal autorizado.', form);
    if (!apiConfigurada()) {
      mensaje(msg, 'Falta configurar la URL del servidor: edita assets/config.js (ver el manual de instalación).');
      entrar.disabled = true;
    } else if (msgInicial) mensaje(msg, msgInicial, tipoMsg);
    setTimeout(function () { usuario.focus(); }, 50);
  }

  function mostrarOlvide(usuarioPrevio) {
    var usuario = entrada({ type: 'text', autocapitalize: 'none', spellcheck: 'false', value: usuarioPrevio || '' });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var enviar = h('button', { type: 'submit', class: 'pa-btn pa-btn-bloque', text: 'Enviar código' });
    var volver = h('button', { type: 'button', class: 'pa-enlace', text: '← Volver a iniciar sesión', onclick: function () { mostrarLogin(); } });
    var yaTengo = h('button', { type: 'button', class: 'pa-enlace', text: 'Ya tengo un código', onclick: function () { mostrarCodigo(usuario.value, ''); } });
    var form = h('form', { class: 'pa-form', novalidate: true }, [
      campo('Usuario', usuario, 'Se enviará un código al correo general de distribución.'),
      msg, enviar, h('div', { class: 'pa-centro pa-col' }, [yaTengo, volver])
    ]);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!usuario.value.trim()) { mensaje(msg, 'Escribe tu usuario.'); return; }
      ocupado(enviar, true, 'Enviando…');
      api('olvideClave', { usuario: usuario.value }).then(function (res) {
        ocupado(enviar, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        mostrarCodigo(usuario.value, res.mensaje);
      });
    });
    tarjeta('¿Olvidaste tu contraseña?', 'Te ayudamos a crear una nueva.', form);
    setTimeout(function () { usuario.focus(); }, 50);
  }

  function mostrarCodigo(usuarioPrevio, info) {
    var usuario = entrada({ type: 'text', autocapitalize: 'none', spellcheck: 'false', value: usuarioPrevio || '' });
    var codigo = entrada({ type: 'text', inputmode: 'numeric', autocomplete: 'one-time-code', maxlength: '6', placeholder: '000000', class: 'pa-codigo' });
    var clave = campoClave({ etiqueta: 'Nueva contraseña' });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn pa-btn-bloque', text: 'Cambiar contraseña' });
    var volver = h('button', { type: 'button', class: 'pa-enlace', text: '← Volver a iniciar sesión', onclick: function () { mostrarLogin(); } });
    var form = h('form', { class: 'pa-form', novalidate: true }, [
      info ? h('p', { class: 'pa-msg pa-msg-info', text: info }) : null,
      campo('Usuario', usuario),
      campo('Código de 6 dígitos', codigo),
      clave.el, msg, guardar, h('div', { class: 'pa-centro' }, [volver])
    ]);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var err = !usuario.value.trim() ? 'Escribe tu usuario.' : (!/^\d{6}$/.test(codigo.value.trim()) ? 'El código tiene 6 dígitos.' : clave.validar());
      if (err) { mensaje(msg, err); return; }
      ocupado(guardar, true, 'Guardando…');
      api('restablecerConCodigo', { usuario: usuario.value, codigo: codigo.value.trim(), nueva: clave.valor(), generada: clave.generada() })
        .then(function (res) {
          ocupado(guardar, false);
          if (!res.ok) { mensaje(msg, res.error); return; }
          if (clave.generada()) {
            var cont = h('div', { class: 'pa-form' }, [
              cajaResultado('Tu nueva contraseña', usuario.value.trim().toLowerCase(), clave.valor()),
              h('button', { type: 'button', class: 'pa-btn pa-btn-bloque', text: 'Ir a iniciar sesión', onclick: function () { mostrarLogin(res.mensaje, 'ok'); } })
            ]);
            tarjeta('Contraseña restablecida', 'Guárdala antes de continuar.', cont);
          } else mostrarLogin(res.mensaje, 'ok');
        });
    });
    tarjeta('Nueva contraseña', 'Escribe el código que llegó al correo general.', form);
    setTimeout(function () { (usuario.value ? codigo : usuario).focus(); }, 50);
  }

  function pantallaCargando() {
    tarjeta('Verificando sesión…', '', h('div', { class: 'pa-cargando', 'aria-hidden': 'true' }));
  }

  function desbloquear(u) {
    usuarioActual = u;
    var r = document.getElementById('pa-auth');
    if (r) r.remove();
    document.documentElement.classList.remove('pa-locked');
    ponerBarra();
    [].slice.call(document.querySelectorAll('[data-pa-nombre]')).forEach(function (el) { el.textContent = u.nombre || u.usuario; });
    [].slice.call(document.querySelectorAll('[data-pa-solo-admin]')).forEach(function (el) { el.hidden = u.rol !== 'admin'; });
    [].slice.call(document.querySelectorAll('[data-pa-min-rol]')).forEach(function (el) {
      el.hidden = (NIVEL[u.rol] || 0) < (NIVEL[el.getAttribute('data-pa-min-rol')] || 0);
    });
    [].slice.call(document.querySelectorAll('[data-pa-permiso]')).forEach(function (el) {
      el.hidden = !(u.permisos && u.permisos[el.getAttribute('data-pa-permiso')]);
    });
    document.dispatchEvent(new CustomEvent('pa:sesion', { detail: u }));
  }

  function cerrarLocal() {
    guardarToken(null);
    usuarioActual = null;
    cerrarModales();
  }

  /* ---------------- barra de usuario ---------------- */

  var cerrarMenu = function () {};
  document.addEventListener('click', function () { cerrarMenu(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') cerrarMenu(); });

  function quitarBarra() { var b = document.getElementById('pa-barra'); if (b) b.remove(); cerrarMenu = function () {}; }

  function ponerBarra() {
    quitarBarra();
    var u = usuarioActual;
    var iniciales = (u.nombre || u.usuario).split(' ').map(function (p) { return p[0]; }).join('').slice(0, 2).toUpperCase();
    var menu = h('div', { class: 'pa-menu pa-oculto', role: 'menu' });
    var boton = h('button', { type: 'button', class: 'pa-chip', 'aria-haspopup': 'true', 'aria-expanded': 'false' }, [
      h('span', { class: 'pa-avatar', text: iniciales }),
      h('span', { class: 'pa-chip-nombre', text: u.nombre || u.usuario })
    ]);
    function alternar(abrir) {
      var a = abrir === undefined ? menu.classList.contains('pa-oculto') : abrir;
      menu.classList.toggle('pa-oculto', !a);
      boton.setAttribute('aria-expanded', a ? 'true' : 'false');
    }
    boton.addEventListener('click', function (e) { e.stopPropagation(); alternar(); });
    cerrarMenu = function () { alternar(false); };

    function item(texto, fn) {
      return h('button', { type: 'button', class: 'pa-menu-item', role: 'menuitem', text: texto, onclick: function () { alternar(false); fn(); } });
    }
    menu.appendChild(h('div', { class: 'pa-menu-cab' }, [
      h('strong', { text: u.nombre }), h('small', { text: '@' + u.usuario + ' · ' + nombreRol(u.rol) })
    ]));
    if (u.rol === 'admin') menu.appendChild(item('Gestionar usuarios', panelUsuarios));
    menu.appendChild(item('Cambiar mi contraseña', modalMiClave));
    menu.appendChild(item('Cerrar sesión', function () {
      api('salir', {}); cerrarLocal(); mostrarLogin('Sesión cerrada.', 'ok');
    }));
    document.body.appendChild(h('div', { id: 'pa-barra' }, [boton, menu]));
  }

  /* ---------------- ventanas (modales) ---------------- */

  function cerrarModales() {
    [].slice.call(document.querySelectorAll('.pa-fondo')).forEach(function (m) { m.remove(); });
  }

  function modal(titulo, contenido, ancho) {
    var fondo = h('div', { class: 'pa-fondo' });
    var cerrarBtn = h('button', { type: 'button', class: 'pa-cerrar', 'aria-label': 'Cerrar', text: '×' });
    var caja = h('div', { class: 'pa-modal' + (ancho ? ' pa-modal-ancho' : ''), role: 'dialog', 'aria-modal': 'true', 'aria-label': titulo }, [
      h('div', { class: 'pa-modal-cab' }, [h('h2', { text: titulo }), cerrarBtn]),
      h('div', { class: 'pa-modal-cuerpo' }, [contenido])
    ]);
    fondo.appendChild(caja);
    function cerrar() { fondo.remove(); document.removeEventListener('keydown', esc); }
    function esc(e) { if (e.key === 'Escape' && fondo === [].slice.call(document.querySelectorAll('.pa-fondo')).pop()) cerrar(); }
    cerrarBtn.addEventListener('click', cerrar);
    fondo.addEventListener('mousedown', function (e) { if (e.target === fondo) cerrar(); });
    document.addEventListener('keydown', esc);
    document.body.appendChild(fondo);
    return { cerrar: cerrar, cuerpo: caja.querySelector('.pa-modal-cuerpo') };
  }

  function modalMiClave() {
    var actual = entrada({ type: 'password', autocomplete: 'current-password' });
    var nueva = campoClave({ etiqueta: 'Nueva contraseña' });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn', text: 'Guardar' });
    var form = h('form', { class: 'pa-form', novalidate: true }, [campo('Contraseña actual', actual), nueva.el, msg, h('div', { class: 'pa-fila pa-fin' }, [guardar])]);
    var m = modal('Cambiar mi contraseña', form);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var err = !actual.value ? 'Escribe tu contraseña actual.' : nueva.validar();
      if (err) { mensaje(msg, err); return; }
      ocupado(guardar, true, 'Guardando…');
      api('cambiarClave', { actual: actual.value, nueva: nueva.valor() }).then(function (res) {
        ocupado(guardar, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        guardarToken(res.token);
        if (nueva.generada()) {
          m.cuerpo.innerHTML = '';
          m.cuerpo.appendChild(cajaResultado('Tu nueva contraseña', usuarioActual.usuario, nueva.valor()));
        } else { m.cerrar(); }
        aviso(res.mensaje);
      });
    });
    setTimeout(function () { actual.focus(); }, 50);
  }

  /* ---------------- gestión de usuarios (admin) ---------------- */

  function panelUsuarios() {
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var buscar = entrada({ type: 'search', placeholder: 'Buscar por usuario o nombre…' });
    var nuevoBtn = h('button', { type: 'button', class: 'pa-btn', text: '+ Nuevo usuario' });
    var zonaForm = h('div');
    var tabla = h('div', { class: 'pa-tabla-cont' });
    var lista = [];

    var cont = h('div', {}, [
      h('div', { class: 'pa-fila pa-entre' }, [buscar, nuevoBtn]),
      zonaForm, msg, tabla
    ]);
    modal('Gestionar usuarios', cont, true);

    nuevoBtn.addEventListener('click', function () {
      var abierto = zonaForm.firstChild && zonaForm.firstChild.tagName === 'FORM';
      zonaForm.innerHTML = '';
      if (!abierto) zonaForm.appendChild(formNuevo());
    });
    buscar.addEventListener('input', pintar);

    function cargar() {
      tabla.innerHTML = '';
      tabla.appendChild(h('div', { class: 'pa-cargando', 'aria-hidden': 'true' }));
      api('listarUsuarios', {}).then(function (res) {
        if (!res.ok) { tabla.innerHTML = ''; mensaje(msg, res.error); return; }
        lista = res.usuarios; pintar();
      });
    }

    function pintar() {
      var q = buscar.value.trim().toLowerCase();
      var filas = lista.filter(function (u) { return !q || (u.usuario + ' ' + u.nombre).toLowerCase().indexOf(q) >= 0; });
      tabla.innerHTML = '';
      if (!filas.length) { tabla.appendChild(h('p', { class: 'pa-vacio', text: 'No hay usuarios que coincidan.' })); return; }
      var cuerpo = h('tbody');
      filas.forEach(function (u) {
        var yo = u.usuario === usuarioActual.usuario;
        var protegido = u.principal || yo;
        var acciones = h('div', { class: 'pa-acciones' }, [
          (!u.principal || usuarioActual.usuario === u.usuario)
            ? h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Restablecer contraseña', onclick: function () { modalRestablecer(u, cargar); } })
            : null,
          !protegido ? h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: u.activo ? 'Desactivar' : 'Activar', onclick: function (e) { cambiarEstado(u, e.currentTarget); } }) : null,
          !protegido ? h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar', onclick: function (e) { eliminar(u, e.currentTarget); } }) : null
        ]);
        cuerpo.appendChild(h('tr', { class: u.activo ? '' : 'pa-inactivo' }, [
          h('td', { 'data-etq': 'Usuario' }, [h('strong', { text: u.usuario }), u.principal ? h('span', { class: 'pa-tag', text: 'principal' }) : null, yo ? h('span', { class: 'pa-tag pa-tag-yo', text: 'tú' }) : null]),
          h('td', { 'data-etq': 'Nombre', text: u.nombre }),
          h('td', { 'data-etq': 'Rol' }, [protegido
            ? h('span', { class: 'pa-tag ' + ({ admin: 'pa-tag-admin', editor: 'pa-tag-editor', supervisor: 'pa-tag-supervisor' }[u.rol] || ''), text: nombreRol(u.rol) })
            : selectorRol(u)]),
          h('td', { 'data-etq': 'Estado' }, [h('span', { class: 'pa-estado ' + (u.activo ? 'pa-on' : 'pa-off'), text: u.activo ? 'Activo' : 'Inactivo' })]),
          h('td', { 'data-etq': 'Último acceso', text: u.ultimoAcceso || '—' }),
          h('td', { 'data-etq': 'Acciones' }, [acciones])
        ]));
      });
      tabla.appendChild(h('table', { class: 'pa-tabla' }, [
        h('thead', {}, [h('tr', {}, ['Usuario', 'Nombre', 'Rol', 'Estado', 'Último acceso', ''].map(function (t) { return h('th', { text: t }); }))]),
        cuerpo
      ]));
    }

    function formNuevo() {
      var usuario = entrada({ type: 'text', autocapitalize: 'none', spellcheck: 'false', placeholder: 'ej. juan.perez', maxlength: '30' });
      var nombre = entrada({ type: 'text', placeholder: 'Nombre y apellido', maxlength: '60' });
      var rol = h('select', { class: 'pa-input' }, [
        h('option', { value: 'usuario', text: 'Usuario — busca clientes y escucha notas de voz' }),
        h('option', { value: 'editor', text: 'Editor — además carga clientes y graba notas nuevas' }),
        h('option', { value: 'supervisor', text: 'Supervisor — además ve la actividad y todas las respuestas' }),
        h('option', { value: 'admin', text: 'Administrador — todo: usuarios, formularios, ocultar clientes, reemplazar notas' })]);
      var clave = campoClave();
      var m2 = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
      var crear = h('button', { type: 'submit', class: 'pa-btn', text: 'Crear usuario' });
      var cancelar = h('button', { type: 'button', class: 'pa-btn pa-btn-sec', text: 'Cancelar', onclick: function () { zonaForm.innerHTML = ''; } });
      var form = h('form', { class: 'pa-form pa-panel', novalidate: true }, [
        h('h3', { text: 'Nuevo usuario' }),
        h('div', { class: 'pa-rejilla' }, [
          campo('Usuario', usuario, 'Letras sin acento, números, punto, guion. Sin espacios.'),
          campo('Nombre completo', nombre),
          campo('Rol', rol)
        ]),
        clave.el, m2, h('div', { class: 'pa-fila pa-fin' }, [cancelar, crear])
      ]);
      form.addEventListener('submit', function (e) {
        e.preventDefault();
        var u = usuario.value.trim().toLowerCase();
        var err = !/^[a-z0-9._-]{3,30}$/.test(u) ? 'El usuario debe tener de 3 a 30 caracteres: letras sin acento, números, punto, guion o guion bajo.'
          : (!nombre.value.trim() ? 'Escribe el nombre de la persona.' : clave.validar());
        if (err) { mensaje(m2, err); return; }
        ocupado(crear, true, 'Creando…');
        api('crearUsuario', { usuario: u, nombre: nombre.value, rol: rol.value, clave: clave.valor(), generada: clave.generada() })
          .then(function (res) {
            ocupado(crear, false);
            if (!res.ok) { mensaje(m2, res.error); return; }
            zonaForm.innerHTML = '';
            zonaForm.appendChild(cajaResultado('Usuario creado: ' + res.usuario.nombre, res.usuario.usuario, clave.valor()));
            aviso('Usuario creado');
            cargar();
          });
      });
      setTimeout(function () { usuario.focus(); }, 50);
      return form;
    }

    function selectorRol(u) {
      var sel = h('select', { class: 'pa-input pa-input-sm', 'aria-label': 'Rol de ' + u.usuario },
        ROLES.map(function (r) { return h('option', { value: r, text: nombreRol(r), selected: r === u.rol }); }));
      sel.addEventListener('change', function () {
        sel.disabled = true;
        api('cambiarRol', { usuario: u.usuario, rol: sel.value }).then(function (res) {
          sel.disabled = false;
          if (!res.ok) { mensaje(msg, res.error); sel.value = u.rol; return; }
          aviso(res.mensaje); cargar();
        });
      });
      return sel;
    }

    function cambiarEstado(u, boton) {
      ocupado(boton, true, '…');
      api('cambiarEstado', { usuario: u.usuario, activo: !u.activo }).then(function (res) {
        ocupado(boton, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        aviso(res.mensaje); cargar();
      });
    }

    function eliminar(u, boton) {
      if (!confirm('¿Eliminar el usuario "' + u.usuario + '" (' + u.nombre + ')? Esta acción no se puede deshacer.')) return;
      ocupado(boton, true, '…');
      api('eliminarUsuario', { usuario: u.usuario }).then(function (res) {
        ocupado(boton, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        aviso(res.mensaje); cargar();
      });
    }

    cargar();
  }

  function modalRestablecer(u, alTerminar) {
    var clave = campoClave({ etiqueta: 'Nueva contraseña' });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn', text: 'Guardar contraseña' });
    var form = h('form', { class: 'pa-form', novalidate: true }, [
      h('p', { class: 'pa-sub' }, ['Usuario: ', h('strong', { text: u.usuario }), ' — ' + u.nombre]),
      clave.el, msg, h('div', { class: 'pa-fila pa-fin' }, [guardar])
    ]);
    var m = modal('Restablecer contraseña', form);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var err = clave.validar();
      if (err) { mensaje(msg, err); return; }
      ocupado(guardar, true, 'Guardando…');
      api('restablecerClave', { usuario: u.usuario, clave: clave.valor(), generada: clave.generada() }).then(function (res) {
        ocupado(guardar, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        if (res.token) guardarToken(res.token);
        m.cuerpo.innerHTML = '';
        m.cuerpo.appendChild(cajaResultado('Contraseña actualizada', u.usuario, clave.valor()));
        aviso(res.mensaje);
        if (alTerminar) alTerminar();
      });
    });
  }

  /* ---------------- inicio ---------------- */

  function iniciar() {
    var token = leerToken();
    if (!token || !apiConfigurada()) { mostrarLogin(); return; }
    pantallaCargando();
    api('sesion', { token: token }).then(function (res) {
      if (res.ok) desbloquear(res.usuario);
      else { guardarToken(null); mostrarLogin(res.codigo === 'SESION' ? '' : res.error); }
    });
  }

  window.PA = {
    api: api,
    usuario: function () { return usuarioActual; },
    gestionarUsuarios: function () { if (usuarioActual && usuarioActual.rol === 'admin') panelUsuarios(); },
    cambiarClave: function () { if (usuarioActual) modalMiClave(); }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
