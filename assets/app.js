/*
 * Contacto-Cliente — módulos de la app (después de iniciar sesión)
 *   Clientes · Ficha del cliente con nota de voz · Grabar/Subir audio
 *   Resumen de cobertura · Cargar (Excel, audios, sincronizar Drive) · Actividad
 *   Formularios de preguntas e historial del cliente · Clientes ocultos
 * Depende de assets/auth.js (window.PA y el evento "pa:sesion").
 */
(function () {
  'use strict';

  var CAMPOS = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono', 'audio', 'audioFecha', 'audioPor', 'oculto', 'tipo'];
  var ETIQUETAS = { codigo: 'Código', nombre: 'Nombre', tipo: 'Tipo de cliente', direccion: 'Dirección', sector: 'Sector / Zona', ruta: 'Ruta',
                    vendedor: 'Vendedor / Chofer', telefono: 'Teléfono' };
  var EDITABLES = ['codigo', 'nombre', 'tipo', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono'];
  var CON_LISTA = ['tipo', 'sector', 'ruta', 'vendedor']; // campos con sugerencias de valores ya usados
  var NIVEL = { usuario: 0, editor: 1, supervisor: 2, admin: 3 };
  var POR_PAGINA = 60;
  var MAX_MB = 10;
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  var estado = {
    usuario: null,
    clientes: [],       // objetos {codigo, nombre, ..., audio, _txt}
    porCodigo: {},
    carpeta: null,
    mostrados: POR_PAGINA,
    filtrados: [],
    grupo: 'ruta',
    cacheAudio: {},     // codigo -> {url, mime}
    formularios: [],    // formularios disponibles (activos; el admin recibe también los inactivos)
    conteoRoles: {},    // personas activas por rol (solo quien diseña formularios)
    maxPreguntas: 200
  };

  /* ---------------- utilidades ---------------- */

  function $(id) { return document.getElementById(id); }

  function h(tag, attrs, hijos) {
    var el = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'html') el.innerHTML = v; // solo con texto fijo del propio código
      else if (k.indexOf('on') === 0) el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    [].concat(hijos || []).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
    });
    return el;
  }

  function api(accion, datos) { return window.PA.api(accion, datos || {}); }

  function puede(rol) { return estado.usuario && (NIVEL[estado.usuario.rol] || 0) >= NIVEL[rol]; }

  /** Permisos que envía el servidor (tabla PERMISOS de Code.gs). */
  function tiene(permiso) { return !!(estado.usuario && estado.usuario.permisos && estado.usuario.permisos[permiso]); }

  function normalizar(s) {
    return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
  }

  /** ¿El formulario aplica a este tipo de cliente? (sin tipos elegidos = todos). Igual que en Code.gs. */
  function aplicaATipo(f, tipo) {
    var lista = f.tiposCliente || [];
    if (!lista.length) return true;
    var t = normalizar(String(tipo || '').replace(/^'/, ''));
    return !!t && lista.some(function (x) { return normalizar(String(x).replace(/^'/, '')) === t; });
  }

  function miles(n) { return Number(n || 0).toLocaleString('es-DO'); }

  function pct(a, b) { return b ? Math.round(a * 1000 / b) / 10 : 0; }

  function aviso(msg, tipo) {
    var cont = $('pa-avisos') || document.body.appendChild(h('div', { id: 'pa-avisos' }));
    var t = h('div', { class: 'pa-aviso pa-aviso-' + (tipo || 'ok'), role: 'status', text: msg });
    cont.appendChild(t);
    setTimeout(function () { t.classList.add('pa-salir'); setTimeout(function () { t.remove(); }, 300); }, 3500);
  }

  function mensaje(el, texto, tipo) {
    el.textContent = texto || '';
    el.className = 'pa-msg' + (texto ? ' pa-msg-' + (tipo || 'error') : '');
  }

  function ocupado(boton, si, texto) {
    if (si) { boton.dataset.t = boton.textContent; boton.textContent = texto || 'Procesando…'; boton.disabled = true; }
    else { boton.textContent = boton.dataset.t || boton.textContent; boton.disabled = false; }
  }

  function modal(titulo, contenido, opciones) {
    opciones = opciones || {};
    var fondo = h('div', { class: 'pa-fondo' });
    var cerrarBtn = h('button', { type: 'button', class: 'pa-cerrar', 'aria-label': 'Cerrar', text: '×' });
    var caja = h('div', { class: 'pa-modal cc-modal' + (opciones.ancho ? ' pa-modal-ancho' : ''), role: 'dialog', 'aria-modal': 'true', 'aria-label': titulo }, [
      h('div', { class: 'pa-modal-cab' }, [h('div', { class: 'cc-modal-titulo' }, [h('h2', { text: titulo }), opciones.sub ? h('small', { text: opciones.sub }) : null]), cerrarBtn]),
      h('div', { class: 'pa-modal-cuerpo' }, [contenido])
    ]);
    fondo.appendChild(caja);
    function cerrar() {
      fondo.remove();
      document.removeEventListener('keydown', esc);
      if (opciones.alCerrar) opciones.alCerrar();
    }
    function esc(e) { if (e.key === 'Escape' && fondo === [].slice.call(document.querySelectorAll('.pa-fondo')).pop()) cerrar(); }
    cerrarBtn.addEventListener('click', cerrar);
    fondo.addEventListener('mousedown', function (e) { if (e.target === fondo) cerrar(); });
    document.addEventListener('keydown', esc);
    document.body.appendChild(fondo);
    return { cerrar: cerrar, cuerpo: caja.querySelector('.pa-modal-cuerpo'), titulo: caja.querySelector('h2') };
  }

  function leerBase64(archivo) {
    return new Promise(function (ok, mal) {
      var r = new FileReader();
      r.onload = function () { ok(String(r.result).split(',')[1] || ''); };
      r.onerror = function () { mal(new Error('No se pudo leer el archivo.')); };
      r.readAsDataURL(archivo);
    });
  }

  function base64ABlob(b64, mime) {
    var bin = atob(b64), n = bin.length, bytes = new Uint8Array(n);
    for (var i = 0; i < n; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type: mime || 'audio/mpeg' });
  }

  function cargarScript(url) {
    return new Promise(function (ok, mal) {
      if (document.querySelector('script[src="' + url + '"]')) { ok(); return; }
      var s = document.createElement('script');
      s.src = url; s.onload = ok; s.onerror = function () { mal(new Error('No se pudo cargar ' + url)); };
      document.head.appendChild(s);
    });
  }

  var ICONOS = {
    play: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7Z"/></svg>',
    mic: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 14a3 3 0 0 0 3-3V5a3 3 0 0 0-6 0v6a3 3 0 0 0 3 3Zm5-3a5 5 0 0 1-10 0H5a7 7 0 0 0 6 6.9V21h2v-3.1A7 7 0 0 0 19 11Z"/></svg>',
    tel: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6.6 10.8a15 15 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.2 11.4 11.4 0 0 0 3.6.6 1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1 11.4 11.4 0 0 0 .6 3.6 1 1 0 0 1-.3 1Z"/></svg>',
    wa: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2Zm5.4 14.2c-.2.6-1.3 1.2-1.8 1.2-.5.1-1 .1-3.2-.7a11 11 0 0 1-4.3-3.8 4.8 4.8 0 0 1-1-2.6 2.8 2.8 0 0 1 .9-2.1 1 1 0 0 1 .7-.3h.5c.2 0 .4 0 .6.5l.8 2c.1.2.1.3 0 .5l-.3.5-.4.4c-.1.1-.3.3-.1.6a8 8 0 0 0 1.5 1.8 7.3 7.3 0 0 0 2.1 1.3c.3.1.4.1.6-.1l.8-1c.2-.3.4-.2.6-.1l1.9.9c.3.1.5.2.5.3a2.4 2.4 0 0 1-.2 1.4Z"/></svg>',
    mapa: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2a7 7 0 0 0-7 7c0 5.3 7 13 7 13s7-7.7 7-13a7 7 0 0 0-7-7Zm0 9.5A2.5 2.5 0 1 1 14.5 9 2.5 2.5 0 0 1 12 11.5Z"/></svg>',
    subir: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M11 16V7.8l-3.6 3.6L6 10l6-6 6 6-1.4 1.4L13 7.8V16Zm-7 4v-5h2v3h12v-3h2v5Z"/></svg>'
  };
  function icono(nombre) { return h('span', { class: 'cc-ico', html: ICONOS[nombre] }); }

  /* ---------------- navegación ---------------- */

  function irA(vista) {
    if (vista === 'usuarios') { window.PA.gestionarUsuarios(); return; }
    var btn = document.querySelector('.cc-nav [data-vista="' + vista + '"]');
    if (!btn || btn.hidden) vista = 'clientes';
    [].slice.call(document.querySelectorAll('.cc-nav [data-vista]')).forEach(function (b) {
      b.classList.toggle('activo', b.getAttribute('data-vista') === vista);
      if (b.getAttribute('data-vista') === vista) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
    });
    [].slice.call(document.querySelectorAll('.cc-vista')).forEach(function (s) {
      var activa = s.id === 'vista-' + vista;
      s.hidden = !activa;
      s.classList.toggle('activa', activa);
    });
    if (location.hash !== '#' + vista) history.replaceState(null, '', '#' + vista);
    if (vista === 'resumen') pintarResumen();
    if (vista === 'actividad') cargarActividad();
    if (vista === 'formularios') cargarFormularios().then(pintarListaFormularios);
    window.scrollTo(0, 0);
  }

  [].slice.call(document.querySelectorAll('.cc-nav [data-vista]')).forEach(function (b) {
    b.addEventListener('click', function () { irA(b.getAttribute('data-vista')); });
  });

  /* ---------------- datos de clientes ---------------- */

  function aObjeto(fila) {
    var c = {};
    CAMPOS.forEach(function (k, i) { c[k] = fila[i]; });
    c.audio = !!c.audio;
    c.oculto = !!c.oculto;
    c.tipo = String(c.tipo || '');
    c._txt = normalizar([c.codigo, c.nombre, c.tipo, c.direccion, c.sector, c.ruta, c.vendedor].join(' ')) + ' ' +
             String(c.telefono || '').replace(/\D/g, '');
    return c;
  }

  function cargarClientes() {
    $('cc-conteo').textContent = 'Cargando clientes…';
    $('cc-lista').innerHTML = '';
    for (var i = 0; i < 6; i++) $('cc-lista').appendChild(h('li', { class: 'cc-cliente cc-esqueleto', 'aria-hidden': 'true' }));
    return api('listarClientes').then(function (res) {
      if (!res.ok) { $('cc-conteo').textContent = res.error; $('cc-lista').innerHTML = ''; return; }
      estado.clientes = res.clientes.map(aObjeto);
      estado.porCodigo = {};
      estado.clientes.forEach(function (c) { estado.porCodigo[c.codigo] = c; });
      estado.carpeta = res.carpeta || null;
      llenarFiltros();
      filtrar();
      pintarCarpeta();
      if (!$('vista-resumen').hidden) pintarResumen();
    });
  }

  /** Reemplaza (o agrega) un cliente en memoria con la fila compacta que devuelve el servidor. */
  function actualizarLocal(fila, codigoAnterior) {
    var c = aObjeto(fila);
    var clave = codigoAnterior || c.codigo;
    var i = estado.clientes.findIndex(function (x) { return x.codigo === clave; });
    if (i >= 0) estado.clientes[i] = c; else estado.clientes.push(c);
    if (codigoAnterior && codigoAnterior !== c.codigo) delete estado.porCodigo[codigoAnterior];
    estado.porCodigo[c.codigo] = c;
    delete estado.cacheAudio[c.codigo];
    llenarFiltros();
    filtrar(true);
    return c;
  }

  function quitarLocal(codigo) {
    estado.clientes = estado.clientes.filter(function (x) { return x.codigo !== codigo; });
    delete estado.porCodigo[codigo];
    llenarFiltros();
    filtrar(true);
  }

  function valoresUnicos(campo) {
    var vistos = {};
    estado.clientes.forEach(function (c) { if (c[campo]) vistos[c[campo]] = true; });
    return Object.keys(vistos).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); });
  }

  function llenarFiltros() {
    [['cc-f-ruta', 'ruta', 'Todas las rutas'], ['cc-f-vendedor', 'vendedor', 'Todos los vendedores'],
     ['cc-f-sector', 'sector', 'Todos los sectores'], ['cc-f-tipo', 'tipo', 'Todos los tipos']].forEach(function (f) {
      var sel = $(f[0]), actual = sel.value;
      sel.innerHTML = '';
      sel.appendChild(h('option', { value: '', text: f[2] }));
      valoresUnicos(f[1]).forEach(function (v) { sel.appendChild(h('option', { value: v, text: v })); });
      sel.value = actual;
      sel.hidden = sel.options.length <= 1;
    });
  }

  /* ---------------- módulo: lista de clientes ---------------- */

  var temporizador = null;
  $('cc-buscar').addEventListener('input', function () {
    clearTimeout(temporizador);
    temporizador = setTimeout(function () { filtrar(); }, 120);
  });
  ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector', 'cc-f-tipo', 'cc-f-audio', 'cc-f-oculto'].forEach(function (id) {
    $(id).addEventListener('change', function () { filtrar(); });
  });
  $('cc-limpiar').addEventListener('click', function () {
    $('cc-buscar').value = '';
    ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector', 'cc-f-tipo', 'cc-f-audio'].forEach(function (id) { $(id).value = ''; });
    $('cc-f-oculto').value = 'vis';
    filtrar();
    $('cc-buscar').focus();
  });
  $('cc-mas').addEventListener('click', function () { estado.mostrados += POR_PAGINA; pintarLista(); });

  function filtrar(mantenerPagina) {
    var q = normalizar($('cc-buscar').value);
    var terminos = q ? q.split(' ') : [];
    var ruta = $('cc-f-ruta').value, vend = $('cc-f-vendedor').value, sector = $('cc-f-sector').value, aud = $('cc-f-audio').value;
    var tipo = $('cc-f-tipo').value;
    var ocu = tiene('ocultarClientes') ? $('cc-f-oculto').value : 'vis';
    var qCodigo = q.toUpperCase();
    estado.filtrados = estado.clientes.filter(function (c) {
      if (ocu === 'vis' && c.oculto) return false;
      if (ocu === 'ocu' && !c.oculto) return false;
      if (ruta && c.ruta !== ruta) return false;
      if (vend && c.vendedor !== vend) return false;
      if (sector && c.sector !== sector) return false;
      if (tipo && c.tipo !== tipo) return false;
      if (aud && String(c.audio ? 1 : 0) !== aud) return false;
      for (var i = 0; i < terminos.length; i++) if (c._txt.indexOf(terminos[i]) < 0) return false;
      return true;
    });
    if (q) { // código exacto primero, luego los que empiezan por el código
      estado.filtrados.sort(function (a, b) {
        var pa = a.codigo === qCodigo ? 0 : (a.codigo.indexOf(qCodigo) === 0 ? 1 : 2);
        var pb = b.codigo === qCodigo ? 0 : (b.codigo.indexOf(qCodigo) === 0 ? 1 : 2);
        return pa - pb;
      });
    }
    if (!mantenerPagina) estado.mostrados = POR_PAGINA;
    pintarLista();
  }

  function pintarLista() {
    var lista = $('cc-lista'), total = estado.filtrados.length;
    var conAudio = estado.filtrados.reduce(function (n, c) { return n + (c.audio ? 1 : 0); }, 0);
    var ocultos = tiene('ocultarClientes') ? estado.clientes.reduce(function (n, c) { return n + (c.oculto ? 1 : 0); }, 0) : 0;
    $('cc-conteo').textContent = estado.clientes.length
      ? miles(total) + (total === 1 ? ' cliente' : ' clientes') + ' · ' + miles(conAudio) + ' con nota de voz' +
        (ocultos && $('cc-f-oculto').value === 'vis' ? ' · ' + miles(ocultos) + ' ocultos (no se muestran)' : '')
      : 'Todavía no hay clientes cargados.' + (puede('editor') ? ' Usa la sección Cargar para importarlos.' : '');
    lista.innerHTML = '';
    var frag = document.createDocumentFragment();
    estado.filtrados.slice(0, estado.mostrados).forEach(function (c) {
      var info = [c.tipo, c.sector, c.ruta ? 'Ruta ' + c.ruta : '', c.vendedor].filter(Boolean).join(' · ');
      var li = h('li', {}, [
        h('button', { type: 'button', class: 'cc-cliente' + (c.oculto ? ' cc-cliente-oculto' : ''), onclick: function () { abrirFicha(c.codigo); } }, [
          h('span', { class: 'cc-estado-audio ' + (c.audio ? 'si' : 'no'), title: c.audio ? 'Tiene nota de voz' : 'Sin nota de voz' },
            [icono(c.audio ? 'play' : 'mic')]),
          h('span', { class: 'cc-cliente-txt' }, [
            h('strong', {}, [c.nombre || '(sin nombre)', c.oculto ? h('span', { class: 'cc-tag-oculto', text: 'Oculto' }) : null]),
            h('small', {}, [h('span', { class: 'cc-codigo', text: c.codigo }), info ? ' ' + info : '']),
            c.direccion ? h('small', { class: 'cc-dir', text: c.direccion }) : null
          ])
        ])
      ]);
      frag.appendChild(li);
    });
    if (estado.clientes.length && !total) frag.appendChild(h('li', { class: 'cc-vacio', text: 'Ningún cliente coincide con la búsqueda.' }));
    lista.appendChild(frag);
    $('cc-mas').hidden = total <= estado.mostrados;
    $('cc-mas').textContent = 'Mostrar más (' + miles(total - estado.mostrados) + ' restantes)';
  }

  /* ---------------- módulo: ficha del cliente ---------------- */

  function telefonoWhatsApp(tel) {
    var d = String(tel || '').replace(/\D/g, '');
    if (d.length === 10 && /^(809|829|849)/.test(d)) d = '1' + d;      // República Dominicana
    if (d.length === 11 && d[0] === '1') return d;
    return d.length >= 8 ? d : '';
  }

  function abrirFicha(codigo) {
    var c = estado.porCodigo[codigo];
    if (!c) return;
    var cuerpo = h('div', { class: 'cc-ficha' });
    var m = modal(c.nombre || c.codigo, cuerpo, { sub: 'Código ' + c.codigo, ancho: true });
    pintarFicha(c, cuerpo, m);
  }

  function pintarFicha(c, cuerpo, m) {
    cuerpo.innerHTML = '';
    m.titulo.textContent = c.nombre || c.codigo;
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    function redibujar(nuevo) { c = nuevo; pintarFicha(c, cuerpo, m); }

    // Aviso de cliente oculto
    var banner = c.oculto ? h('p', { class: 'cc-banner-oculto' }, [h('strong', { text: 'Cliente oculto. ' }),
      'Solo lo ven los administradores; no aparece en la lista, la búsqueda ni el resumen de los demás.']) : null;

    // Nota de voz
    var zonaAudio = h('div', { class: 'cc-audio' });
    function pintarAudio() {
      zonaAudio.innerHTML = '';
      if (c.audio) {
        var escuchar = h('button', { type: 'button', class: 'pa-btn pa-btn-bloque cc-btn-escuchar' }, [icono('play'), 'Escuchar nota de voz']);
        escuchar.addEventListener('click', function () { reproducir(c, zonaAudio, escuchar, msg); });
        zonaAudio.appendChild(escuchar);
        zonaAudio.appendChild(h('small', { class: 'cc-suave', text: 'Guardada ' + (c.audioFecha || '') + (c.audioPor ? ' por ' + c.audioPor : '') }));
        if (estado.cacheAudio[c.codigo]) reproducir(c, zonaAudio, escuchar, msg, true);
      } else {
        zonaAudio.appendChild(h('p', { class: 'cc-sin-audio' }, [icono('mic'), 'Este cliente todavía no tiene nota de voz.']));
      }
      // Agregar nota: editores en adelante (si no tiene). Cambiarla por otra: solo quien tiene "reemplazarAudio".
      var puedeGrabar = c.audio ? tiene('reemplazarAudio') : tiene('agregarAudio');
      var botones = [];
      if (puedeGrabar) {
        botones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', onclick: function () { grabar(c, redibujar); } },
          [icono('mic'), c.audio ? 'Grabar otra (reemplazar)' : 'Grabar nota de voz']));
        botones.push(h('label', { class: 'pa-btn pa-btn-sec pa-btn-sm cc-label-archivo' }, [icono('subir'), c.audio ? 'Reemplazar con archivo' : 'Subir archivo',
          h('input', { type: 'file', accept: 'audio/*,.opus,.ogg,.amr,.m4a', onchange: function (e) {
            var f = e.target.files[0]; e.target.value = '';
            if (!f) return;
            if (c.audio && !confirm('¿Cambiar la nota de voz actual de ' + c.nombre + ' por "' + f.name + '"?\nLa anterior se guarda en la carpeta "Notas reemplazadas" de Drive.')) return;
            subirUno(c, f, c.audio, msg).then(function (nuevo) { if (nuevo) redibujar(nuevo); });
          } })]));
      }
      if (c.audio && tiene('eliminarAudio')) {
        botones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar nota de voz', onclick: function (e) {
          if (!confirm('¿Eliminar la nota de voz de ' + c.nombre + '? Quedará 30 días en la papelera de Drive.')) return;
          var b = e.currentTarget; ocupado(b, true, 'Eliminando…');
          api('eliminarAudio', { codigo: c.codigo }).then(function (res) {
            ocupado(b, false);
            if (!res.ok) { mensaje(msg, res.error); return; }
            aviso(res.mensaje); redibujar(actualizarLocal(res.cliente));
          });
        } }));
      }
      if (botones.length) zonaAudio.appendChild(h('div', { class: 'pa-fila' }, botones));
      if (c.audio && tiene('agregarAudio') && !tiene('reemplazarAudio')) {
        zonaAudio.appendChild(h('small', { class: 'cc-suave', text: 'Solo un administrador puede cambiar esta nota de voz por otra.' }));
      }
    }
    pintarAudio();

    // Datos
    var datos = h('dl', { class: 'cc-datos' });
    ['tipo', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono'].forEach(function (k) {
      datos.appendChild(h('div', {}, [h('dt', { text: ETIQUETAS[k] }), h('dd', { text: c[k] || '—' })]));
    });

    // Contacto
    var tel = String(c.telefono || '').replace(/[^\d+]/g, '');
    var wa = telefonoWhatsApp(c.telefono);
    var consultaMapa = [c.direccion, c.sector, 'República Dominicana'].filter(Boolean).join(', ');
    var contacto = h('div', { class: 'cc-contacto' }, [
      tel ? h('a', { class: 'pa-btn pa-btn-sec', href: 'tel:' + tel }, [icono('tel'), 'Llamar']) : null,
      wa ? h('a', { class: 'pa-btn pa-btn-sec cc-wa', href: 'https://wa.me/' + wa, target: '_blank', rel: 'noopener' }, [icono('wa'), 'WhatsApp']) : null,
      c.direccion ? h('a', { class: 'pa-btn pa-btn-sec', href: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(consultaMapa), target: '_blank', rel: 'noopener' }, [icono('mapa'), 'Ver en el mapa']) : null
    ]);

    // Formularios e historial
    var zonaForms = h('div', { class: 'cc-seccion' });
    pintarFormulariosFicha(c, zonaForms);

    // Administración
    var acciones = [];
    if (tiene('editarClientes')) {
      acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Editar datos', onclick: function () {
        formularioCliente(c, redibujar);
      } }));
    }
    if (tiene('ocultarClientes')) {
      acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: c.oculto ? 'Volver a mostrar cliente' : 'Ocultar cliente', onclick: function (e) {
        if (!c.oculto && !confirm('¿Ocultar a ' + c.codigo + ' — ' + c.nombre + '?\n\nDejará de aparecer para usuarios, editores y supervisores. No se borra nada (datos, nota de voz e historial se conservan) y puedes volver a mostrarlo cuando quieras.')) return;
        var b = e.currentTarget; ocupado(b, true, 'Guardando…');
        api('ocultarCliente', { codigo: c.codigo, oculto: !c.oculto }).then(function (res) {
          ocupado(b, false);
          if (!res.ok) { mensaje(msg, res.error); return; }
          aviso(res.mensaje); redibujar(actualizarLocal(res.cliente));
        });
      } }));
    }
    if (tiene('eliminarClientes')) {
      acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar cliente', onclick: function (e) {
        if (!confirm('¿Eliminar el cliente ' + c.codigo + ' — ' + c.nombre + '?' + (c.audio ? ' Su nota de voz irá a la papelera de Drive.' : '') + ' Esta acción no se puede deshacer.\n\nSi solo quieres que no aparezca, usa "Ocultar cliente".')) return;
        var b = e.currentTarget; ocupado(b, true, 'Eliminando…');
        api('eliminarCliente', { codigo: c.codigo }).then(function (res) {
          ocupado(b, false);
          if (!res.ok) { mensaje(msg, res.error); return; }
          quitarLocal(c.codigo); aviso(res.mensaje); m.cerrar();
        });
      } }));
    }
    var admin = acciones.length ? h('div', { class: 'pa-fila cc-ficha-admin' }, acciones) : null;

    if (banner) cuerpo.appendChild(banner);
    cuerpo.appendChild(zonaAudio);
    cuerpo.appendChild(msg);
    if (contacto.childNodes.length) cuerpo.appendChild(contacto);
    cuerpo.appendChild(datos);
    cuerpo.appendChild(zonaForms);
    if (admin) cuerpo.appendChild(admin);
  }

  function reproducir(c, zona, boton, msg, autoSilencioso) {
    function mostrar(url) {
      var audio = h('audio', { controls: true, preload: 'auto', src: url, class: 'cc-reproductor' });
      var vel = h('div', { class: 'cc-velocidad', role: 'group', 'aria-label': 'Velocidad' }, ['1', '1.5', '2'].map(function (v) {
        var b = h('button', { type: 'button', class: v === '1' ? 'activo' : '', text: v + '×' });
        b.addEventListener('click', function () {
          audio.playbackRate = Number(v);
          [].slice.call(vel.children).forEach(function (x) { x.classList.toggle('activo', x === b); });
        });
        return b;
      }));
      audio.addEventListener('error', function () {
        mensaje(msg, 'Este teléfono no puede reproducir este formato de audio. Pide que la conviertan a MP3 o grábala de nuevo desde la app.');
      });
      boton.replaceWith(h('div', { class: 'cc-reproductor-cont' }, [audio, vel]));
      if (!autoSilencioso) audio.play().catch(function () { /* el usuario pulsa play */ });
    }
    if (estado.cacheAudio[c.codigo]) { mostrar(estado.cacheAudio[c.codigo].url); return; }
    ocupado(boton, true, 'Cargando audio…');
    api('audio', { codigo: c.codigo }).then(function (res) {
      if (!res.ok) { ocupado(boton, false); mensaje(msg, res.error); return; }
      var url = URL.createObjectURL(base64ABlob(res.base64, res.mime));
      var claves = Object.keys(estado.cacheAudio);
      if (claves.length >= 20) { URL.revokeObjectURL(estado.cacheAudio[claves[0]].url); delete estado.cacheAudio[claves[0]]; }
      estado.cacheAudio[c.codigo] = { url: url, mime: res.mime };
      mostrar(url);
    });
  }

  /** Sube un archivo de audio para un cliente. Devuelve el cliente actualizado o null. */
  function subirUno(c, archivo, reemplazar, msg, grabado) {
    if (archivo.size > MAX_MB * 1024 * 1024) {
      if (msg) mensaje(msg, 'El audio pesa más de ' + MAX_MB + ' MB.');
      return Promise.resolve(null);
    }
    if (msg) mensaje(msg, 'Subiendo nota de voz…', 'info');
    return leerBase64(archivo).then(function (b64) {
      return api('subirAudio', { codigo: c.codigo, nombreArchivo: archivo.name || 'grabacion', mime: archivo.type,
                                 base64: b64, reemplazar: !!reemplazar, grabado: !!grabado });
    }).then(function (res) {
      if (!res.ok && res.codigo === 'YA_TIENE' && tiene('reemplazarAudio')) {
        if (confirm('El cliente ya tiene nota de voz. ¿Reemplazarla por la nueva?')) return subirUno(c, archivo, true, msg, grabado);
        if (msg) mensaje(msg, '');
        return null;
      }
      if (!res.ok) { if (msg) mensaje(msg, res.error); return null; }
      if (msg) mensaje(msg, '');
      aviso(res.mensaje);
      return actualizarLocal(res.cliente);
    });
  }

  /* ---------------- módulo: grabar nota de voz ---------------- */

  function grabar(c, alGuardar) {
    if (!navigator.mediaDevices || !window.MediaRecorder) {
      alert('Este navegador no permite grabar audio. Usa Chrome en Android o Safari actualizado en iPhone, o sube un archivo.');
      return;
    }
    var tipos = ['audio/webm;codecs=opus', 'audio/mp4', 'audio/ogg;codecs=opus', 'audio/webm'];
    var tipo = tipos.filter(function (t) { return MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(t); })[0] || '';
    var MAX_SEG = 180;
    var flujo = null, grabador = null, trozos = [], inicio = 0, reloj = null, blob = null;
    var tiempo = h('div', { class: 'cc-tiempo', text: '0:00' });
    var punto = h('span', { class: 'cc-punto' });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var principal = h('button', { type: 'button', class: 'cc-btn-grabar', 'aria-label': 'Empezar a grabar' }, [icono('mic')]);
    var ayuda = h('p', { class: 'cc-suave cc-centro', text: 'Pulsa el micrófono y describe cómo llegar al cliente. Máximo 3 minutos.' +
      (c.audio ? ' Al guardar, esta grabación reemplaza la nota actual (la anterior se guarda en "Notas reemplazadas").' : '') });
    var vista = h('div', { class: 'cc-vista-previa' });
    var guardar = h('button', { type: 'button', class: 'pa-btn', text: 'Guardar nota de voz', disabled: true });
    var repetir = h('button', { type: 'button', class: 'pa-btn pa-btn-sec', text: 'Grabar de nuevo', hidden: true });
    var cuerpo = h('div', { class: 'cc-grabador' }, [h('div', { class: 'cc-centro' }, [principal, h('div', { class: 'cc-estado-grab' }, [punto, tiempo])]),
      ayuda, vista, msg, h('div', { class: 'pa-fila pa-fin' }, [repetir, guardar])]);
    var m = modal(c.audio ? 'Reemplazar nota de voz' : 'Grabar nota de voz', cuerpo, { sub: c.codigo + ' — ' + c.nombre, alCerrar: detenerTodo });

    function detenerTodo() {
      clearInterval(reloj);
      try { if (grabador && grabador.state !== 'inactive') grabador.stop(); } catch (e) { /* nada */ }
      if (flujo) flujo.getTracks().forEach(function (t) { t.stop(); });
    }
    function iniciar() {
      mensaje(msg, '');
      navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true } }).then(function (s) {
        flujo = s; trozos = []; blob = null; vista.innerHTML = ''; guardar.disabled = true; repetir.hidden = true;
        grabador = tipo ? new MediaRecorder(s, { mimeType: tipo, audioBitsPerSecond: 32000 }) : new MediaRecorder(s);
        grabador.ondataavailable = function (e) { if (e.data && e.data.size) trozos.push(e.data); };
        grabador.onstop = function () {
          clearInterval(reloj);
          flujo.getTracks().forEach(function (t) { t.stop(); });
          cuerpo.classList.remove('grabando');
          blob = new Blob(trozos, { type: (grabador.mimeType || tipo || 'audio/webm').split(';')[0] });
          vista.innerHTML = '';
          vista.appendChild(h('audio', { controls: true, src: URL.createObjectURL(blob), class: 'cc-reproductor' }));
          principal.setAttribute('aria-label', 'Grabar de nuevo');
          guardar.disabled = false; repetir.hidden = false;
          ayuda.textContent = 'Escucha la grabación. Si está bien, pulsa Guardar.';
        };
        grabador.start(1000);
        inicio = Date.now();
        cuerpo.classList.add('grabando');
        principal.setAttribute('aria-label', 'Detener grabación');
        ayuda.textContent = 'Grabando… pulsa de nuevo para detener.';
        reloj = setInterval(function () {
          var s = Math.floor((Date.now() - inicio) / 1000);
          tiempo.textContent = Math.floor(s / 60) + ':' + ('0' + (s % 60)).slice(-2);
          if (s >= MAX_SEG) grabador.stop();
        }, 250);
      }).catch(function () {
        mensaje(msg, 'No se pudo usar el micrófono. Revisa que el navegador tenga permiso para usarlo.');
      });
    }
    principal.addEventListener('click', function () {
      if (grabador && grabador.state === 'recording') grabador.stop(); else iniciar();
    });
    repetir.addEventListener('click', iniciar);
    guardar.addEventListener('click', function () {
      if (!blob) return;
      var ext = blob.type.indexOf('mp4') >= 0 ? 'm4a' : (blob.type.indexOf('ogg') >= 0 ? 'ogg' : 'webm');
      var archivo = new File([blob], c.codigo + '.' + ext, { type: blob.type });
      ocupado(guardar, true, 'Guardando…');
      subirUno(c, archivo, c.audio ? true : false, msg, true).then(function (nuevo) {
        ocupado(guardar, false);
        if (nuevo) { m.cerrar(); if (alGuardar) alGuardar(nuevo); }
      });
    });
  }

  /* ---------------- módulo: crear / editar cliente ---------------- */

  function formularioCliente(c, alGuardar) {
    var nuevo = !c;
    c = c || {};
    var entradas = {};
    var campos = EDITABLES.map(function (k) {
      var input = h(k === 'direccion' ? 'textarea' : 'input', { class: 'pa-input' + (k === 'direccion' ? ' cc-textarea' : ''), rows: k === 'direccion' ? '2' : null,
        type: k === 'telefono' ? 'tel' : (k === 'direccion' ? null : 'text'), maxlength: k === 'direccion' ? '250' : (k === 'codigo' ? '40' : '120'),
        autocapitalize: k === 'codigo' ? 'characters' : null, list: CON_LISTA.indexOf(k) >= 0 ? 'cc-dl-' + k : null,
        placeholder: k === 'tipo' ? 'Ej.: Colmado, Supermercado, Farmacia' : null });
      input.value = c[k] || '';
      entradas[k] = input;
      return h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: ETIQUETAS[k] + (k === 'codigo' || k === 'nombre' ? ' *' : '') }), input]);
    });
    var listas = CON_LISTA.map(function (k) {
      return h('datalist', { id: 'cc-dl-' + k }, valoresUnicos(k).slice(0, 500).map(function (v) { return h('option', { value: v }); }));
    });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn', text: nuevo ? 'Crear cliente' : 'Guardar cambios' });
    var form = h('form', { class: 'pa-form', novalidate: true },
      [h('div', { class: 'pa-rejilla' }, campos)].concat(listas, [msg, h('div', { class: 'pa-fila pa-fin' }, [guardar])]));
    var m = modal(nuevo ? 'Nuevo cliente' : 'Editar cliente', form, { sub: nuevo ? '' : c.codigo });
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var datos = {};
      EDITABLES.forEach(function (k) { datos[k] = entradas[k].value.trim(); });
      if (!datos.codigo || !datos.nombre) { mensaje(msg, 'El código y el nombre son obligatorios.'); return; }
      ocupado(guardar, true, 'Guardando…');
      api('guardarCliente', { cliente: datos, codigoOriginal: nuevo ? '' : c.codigo }).then(function (res) {
        ocupado(guardar, false);
        if (!res.ok) { mensaje(msg, res.error); return; }
        var actualizado = actualizarLocal(res.cliente, nuevo ? null : c.codigo);
        aviso(res.mensaje);
        m.cerrar();
        if (alGuardar) alGuardar(actualizado); else abrirFicha(actualizado.codigo);
      });
    });
    setTimeout(function () { entradas[nuevo ? 'codigo' : 'nombre'].focus(); }, 50);
  }

  $('cc-nuevo').addEventListener('click', function () { formularioCliente(null); });

  /* ---------------- exportar a Excel ---------------- */

  $('cc-exportar').addEventListener('click', function (e) {
    var b = e.currentTarget;
    ocupado(b, true, 'Preparando…');
    cargarScript(XLSX_URL).then(function () {
      var conOculto = tiene('ocultarClientes');
      var filas = [['Código', 'Nombre', 'Tipo de cliente', 'Dirección', 'Sector', 'Ruta', 'Vendedor/Chofer', 'Teléfono', 'Nota de voz', 'Fecha nota de voz'].concat(conOculto ? ['Oculto'] : [])];
      estado.filtrados.forEach(function (c) {
        filas.push([c.codigo, c.nombre, c.tipo, c.direccion, c.sector, c.ruta, c.vendedor, c.telefono, c.audio ? 'Sí' : 'No', c.audioFecha || '']
          .concat(conOculto ? [c.oculto ? 'Sí' : 'No'] : []));
      });
      var ws = XLSX.utils.aoa_to_sheet(filas);
      ws['!cols'] = [12, 32, 18, 40, 18, 10, 18, 14, 11, 16].map(function (w) { return { wch: w }; });
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Clientes');
      var hoy = new Date().toISOString().slice(0, 10);
      XLSX.writeFile(wb, 'clientes-' + hoy + '.xlsx');
      ocupado(b, false);
    }).catch(function () { ocupado(b, false); aviso('No se pudo preparar el Excel. Revisa tu conexión.', 'error'); });
  });

  /* ---------------- módulo: resumen ---------------- */

  [].slice.call(document.querySelectorAll('.cc-segmentos [data-grupo]')).forEach(function (b) {
    b.addEventListener('click', function () {
      estado.grupo = b.getAttribute('data-grupo');
      [].slice.call(document.querySelectorAll('.cc-segmentos [data-grupo]')).forEach(function (x) {
        x.classList.toggle('activo', x === b); x.setAttribute('aria-selected', x === b ? 'true' : 'false');
      });
      pintarResumen();
    });
  });

  function pintarResumen() {
    var visibles = estado.clientes.filter(function (c) { return !c.oculto; }); // los ocultos no cuentan
    var total = visibles.length;
    var con = visibles.reduce(function (n, c) { return n + (c.audio ? 1 : 0); }, 0);
    var kpis = $('cc-kpis');
    kpis.innerHTML = '';
    [['Clientes', miles(total), ''], ['Con nota de voz', miles(con), 'ok'], ['Sin nota de voz', miles(total - con), 'pend'],
     ['Cobertura', pct(con, total) + '%', 'cob']].forEach(function (k) {
      kpis.appendChild(h('div', { class: 'cc-kpi ' + k[2] }, [h('span', { text: k[0] }), h('strong', { text: k[1] })]));
    });
    var grupos = {};
    visibles.forEach(function (c) {
      var g = c[estado.grupo] || '(sin ' + estado.grupo + ')';
      grupos[g] = grupos[g] || { total: 0, con: 0 };
      grupos[g].total++;
      if (c.audio) grupos[g].con++;
    });
    var nombres = Object.keys(grupos).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); });
    var cont = $('cc-cobertura');
    cont.innerHTML = '';
    if (!nombres.length) { cont.appendChild(h('p', { class: 'pa-vacio', text: 'Todavía no hay clientes.' })); return; }
    var etq = { ruta: 'Ruta', vendedor: 'Vendedor / Chofer', sector: 'Sector', tipo: 'Tipo de cliente' }[estado.grupo];
    var cuerpo = h('tbody');
    nombres.forEach(function (n) {
      var g = grupos[n], p = pct(g.con, g.total);
      var fila = h('tr', { class: 'cc-fila-clic', tabindex: '0', title: 'Ver clientes sin nota de voz' }, [
        h('td', { text: n }), h('td', { class: 'num', text: miles(g.total) }), h('td', { class: 'num', text: miles(g.con) }),
        h('td', { class: 'num', text: miles(g.total - g.con) }),
        h('td', { class: 'cc-barra-celda' }, [h('span', { class: 'cc-barra' }, [h('span', { style: 'width:' + p + '%' })]), h('span', { class: 'cc-pct', text: p + '%' })])
      ]);
      function ir() {
        $('cc-buscar').value = '';
        ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector', 'cc-f-tipo'].forEach(function (id) { $(id).value = ''; });
        var sel = $('cc-f-' + estado.grupo);
        if (grupos[n] && sel.querySelector('option[value="' + CSS.escape(n) + '"]')) sel.value = n;
        $('cc-f-audio').value = '0';
        $('cc-f-oculto').value = 'vis';
        filtrar();
        irA('clientes');
      }
      fila.addEventListener('click', ir);
      fila.addEventListener('keydown', function (e) { if (e.key === 'Enter') ir(); });
      cuerpo.appendChild(fila);
    });
    cont.appendChild(h('table', { class: 'pa-tabla cc-tabla' }, [
      h('thead', {}, [h('tr', {}, [h('th', { text: etq }), h('th', { class: 'num', text: 'Clientes' }), h('th', { class: 'num', text: 'Con nota' }),
        h('th', { class: 'num', text: 'Sin nota' }), h('th', { text: 'Cobertura' })])]),
      cuerpo
    ]));
  }

  /* ---------------- módulo: importar clientes ---------------- */

  var SINONIMOS = {
    codigo: ['codigo', 'cod', 'cod cliente', 'codigo cliente', 'id cliente', 'id', 'cliente id', 'numero cliente', 'no cliente', 'cardcode', 'code'],
    nombre: ['nombre', 'cliente', 'nombre cliente', 'razon social', 'negocio', 'nombre comercial', 'cardname', 'name'],
    tipo: ['tipo', 'tipo cliente', 'tipo de cliente', 'tipo negocio', 'tipo de negocio', 'categoria', 'canal', 'segmento', 'clase', 'grupo cliente', 'groupname'],
    direccion: ['direccion', 'dir', 'domicilio', 'calle', 'address', 'ubicacion'],
    sector: ['sector', 'zona', 'barrio', 'localidad', 'municipio', 'provincia', 'ciudad'],
    ruta: ['ruta', 'route', 'ruta venta', 'ruta de venta', 'ruta entrega'],
    vendedor: ['vendedor', 'chofer', 'repartidor', 'preventa', 'preventista', 'vendedor chofer', 'slpname'],
    telefono: ['telefono', 'tel', 'celular', 'cel', 'movil', 'whatsapp', 'phone', 'contacto']
  };

  function detectarColumna(campo, titulos) {
    var cand = SINONIMOS[campo];
    var norm = titulos.map(function (t) { return normalizar(t).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); });
    for (var i = 0; i < norm.length; i++) if (cand.indexOf(norm[i]) >= 0) return i;
    for (var j = 0; j < norm.length; j++) {
      // "Tipo de cliente" no debe confundirse con el nombre o el código del cliente
      if (campo !== 'tipo' && /\btipo\b/.test(norm[j])) continue;
      for (var k = 0; k < cand.length; k++) if (cand[k].length > 3 && norm[j].indexOf(cand[k]) >= 0) return j;
    }
    return -1;
  }

  $('cc-imp-archivo').addEventListener('change', function (e) {
    var archivo = e.target.files[0];
    e.target.value = '';
    if (!archivo) return;
    var zona = $('cc-imp-zona');
    zona.innerHTML = '';
    zona.appendChild(h('p', { class: 'pa-msg pa-msg-info', text: 'Leyendo ' + archivo.name + '…' }));
    cargarScript(XLSX_URL).then(function () { return archivo.arrayBuffer(); }).then(function (buf) {
      var libro = XLSX.read(buf, { type: 'array', raw: false, codepage: 65001 });
      var hoja = libro.Sheets[libro.SheetNames[0]];
      var filas = XLSX.utils.sheet_to_json(hoja, { header: 1, raw: false, defval: '' })
        .filter(function (f) { return f.some(function (v) { return String(v).trim() !== ''; }); });
      if (filas.length < 2) throw new Error('El archivo no tiene filas de datos.');
      prepararImportacion(archivo.name, filas[0].map(String), filas.slice(1));
    }).catch(function (err) {
      zona.innerHTML = '';
      zona.appendChild(h('p', { class: 'pa-msg pa-msg-error', text: 'No se pudo leer el archivo: ' + err.message }));
    });
  });

  function prepararImportacion(nombre, titulos, datos) {
    var zona = $('cc-imp-zona');
    zona.innerHTML = '';
    var selects = {};
    var mapeo = h('div', { class: 'pa-rejilla' }, EDITABLES.map(function (k) {
      var sel = h('select', { class: 'pa-input' }, [h('option', { value: '-1', text: '— No importar —' })].concat(
        titulos.map(function (t, i) { return h('option', { value: String(i), text: t || ('Columna ' + (i + 1)) }); })));
      sel.value = String(detectarColumna(k, titulos));
      sel.addEventListener('change', previsualizar);
      selects[k] = sel;
      return h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: ETIQUETAS[k] + (k === 'codigo' ? ' *' : '') }), sel]);
    }));
    var previa = h('div', { class: 'cc-tabla-cont' });
    var resumen = h('p', { class: 'pa-msg' });
    var barra = h('div', { class: 'cc-progreso', hidden: true }, [h('span')]);
    var importar = h('button', { type: 'button', class: 'pa-btn', text: 'Importar' });
    var resultado = h('div');
    zona.appendChild(h('div', { class: 'cc-panel' }, [
      h('h3', { text: nombre + ' · ' + miles(datos.length) + ' filas' }),
      h('p', { class: 'cc-suave', text: 'Revisa qué columna del archivo corresponde a cada dato.' }),
      mapeo, h('h3', { text: 'Vista previa' }), previa, resumen, barra, h('div', { class: 'pa-fila pa-fin' }, [importar]), resultado
    ]));

    function filasMapeadas() {
      return datos.map(function (f) {
        var o = {};
        EDITABLES.forEach(function (k) { var i = Number(selects[k].value); o[k] = i >= 0 ? String(f[i] == null ? '' : f[i]).trim() : ''; });
        return o;
      });
    }
    function previsualizar() {
      var mapeadas = filasMapeadas();
      previa.innerHTML = '';
      var usados = EDITABLES.filter(function (k) { return selects[k].value !== '-1'; });
      previa.appendChild(h('table', { class: 'pa-tabla cc-tabla' }, [
        h('thead', {}, [h('tr', {}, usados.map(function (k) { return h('th', { text: ETIQUETAS[k] }); }))]),
        h('tbody', {}, mapeadas.slice(0, 5).map(function (o) { return h('tr', {}, usados.map(function (k) { return h('td', { text: o[k] }); })); }))
      ]));
      var sinCodigo = mapeadas.filter(function (o) { return !o.codigo; }).length;
      var existentes = mapeadas.filter(function (o) { return o.codigo && estado.porCodigo[o.codigo.toUpperCase()]; }).length;
      if (selects.codigo.value === '-1') { mensaje(resumen, 'Elige la columna del código del cliente.'); importar.disabled = true; return; }
      importar.disabled = false;
      mensaje(resumen, miles(mapeadas.length - existentes - sinCodigo) + ' clientes nuevos · ' + miles(existentes) +
        ' ya existen (se actualizarán sus datos)' + (sinCodigo ? ' · ' + miles(sinCodigo) + ' filas sin código se omitirán' : ''), 'info');
      importar.textContent = 'Importar ' + miles(mapeadas.length) + ' filas';
    }
    previsualizar();

    importar.addEventListener('click', function () {
      var mapeadas = filasMapeadas();
      var LOTE = 1000, total = { nuevos: 0, actualizados: 0, sinCambios: 0, omitidos: [], totalOmitidos: 0 };
      ocupado(importar, true, 'Importando…');
      barra.hidden = false;
      var i = 0;
      function siguiente() {
        if (i >= mapeadas.length) return Promise.resolve();
        var lote = mapeadas.slice(i, i + LOTE);
        return api('importarClientes', { filas: lote, desde: i }).then(function (res) {
          if (!res.ok) throw new Error(res.error);
          total.nuevos += res.nuevos; total.actualizados += res.actualizados; total.sinCambios += res.sinCambios;
          total.totalOmitidos += res.totalOmitidos; total.omitidos = total.omitidos.concat(res.omitidos);
          i += LOTE;
          barra.firstChild.style.width = Math.min(100, Math.round(i * 100 / mapeadas.length)) + '%';
          return siguiente();
        });
      }
      siguiente().then(function () {
        ocupado(importar, false);
        importar.disabled = true;
        resultado.innerHTML = '';
        resultado.appendChild(h('div', { class: 'pa-resultado' }, [
          h('strong', { text: 'Importación terminada' }),
          h('span', { text: miles(total.nuevos) + ' nuevos · ' + miles(total.actualizados) + ' actualizados · ' + miles(total.sinCambios) + ' sin cambios · ' + miles(total.totalOmitidos) + ' omitidos' }),
          total.omitidos.length ? h('details', {}, [h('summary', { text: 'Ver filas omitidas' }), h('ul', {}, total.omitidos.slice(0, 50).map(function (t) { return h('li', { text: t }); }))]) : null
        ]));
        aviso('Clientes importados');
        cargarClientes();
      }).catch(function (err) {
        ocupado(importar, false);
        mensaje(resumen, 'La importación se detuvo en la fila ' + miles(i + 1) + ': ' + err.message + ' Lo anterior ya quedó guardado; puedes volver a importar el archivo completo sin duplicar.');
      });
    });
  }

  /* ---------------- módulo: subir varias notas de voz ---------------- */

  /** Igual que en el servidor: busca el código del cliente dentro del nombre del archivo. */
  function codigoDesdeNombre(nombre) {
    var base = String(nombre).replace(/\.[^.]+$/, '').trim().toUpperCase();
    if (estado.porCodigo[base]) return base;
    var partes = base.split(/[^A-Z0-9]+/).filter(Boolean);
    for (var i = 0; i < partes.length; i++) if (estado.porCodigo[partes[i]]) return partes[i];
    var sinCeros = {};
    estado.clientes.forEach(function (c) {
      var s = c.codigo.replace(/^0+(?=.)/, '');
      sinCeros[s] = sinCeros[s] === undefined ? c.codigo : null; // null = ambiguo
    });
    for (var j = 0; j < partes.length; j++) {
      var s2 = partes[j].replace(/^0+(?=.)/, '');
      if (/\d/.test(s2) && sinCeros[s2]) return sinCeros[s2];
    }
    return '';
  }

  $('cc-aud-archivos').addEventListener('change', function (e) {
    var archivos = [].slice.call(e.target.files);
    e.target.value = '';
    if (!archivos.length) return;
    var zona = $('cc-aud-zona');
    zona.innerHTML = '';
    var filas = archivos.map(function (f) {
      var codigo = codigoDesdeNombre(f.name);
      var entrada = h('input', { class: 'pa-input pa-input-sm', value: codigo, placeholder: 'Código', 'aria-label': 'Código para ' + f.name, autocapitalize: 'characters' });
      var cliente = h('small', { class: 'cc-suave' });
      var estadoCelda = h('td', { class: 'cc-estado-celda' });
      function revisar() {
        var c = estado.porCodigo[entrada.value.trim().toUpperCase()];
        cliente.textContent = c ? c.nombre + (c.audio ? ' · ya tiene nota' : '') : (entrada.value ? 'Código no encontrado' : 'Escribe el código');
        cliente.className = c ? 'cc-suave' : 'cc-error';
      }
      entrada.addEventListener('input', revisar);
      revisar();
      var tr = h('tr', {}, [h('td', { class: 'cc-nombre-archivo', text: f.name }),
        h('td', { class: 'num', text: (f.size / 1048576).toFixed(1) + ' MB' }),
        h('td', {}, [entrada, h('br'), cliente]), estadoCelda]);
      return { archivo: f, entrada: entrada, tr: tr, estado: estadoCelda };
    });
    var detectados = filas.filter(function (f) { return f.entrada.value; }).length;
    var reemplazar = h('input', { type: 'checkbox' });
    var puedeReemplazar = tiene('reemplazarAudio');
    var subir = h('button', { type: 'button', class: 'pa-btn', text: 'Subir ' + filas.length + ' notas de voz' });
    var barra = h('div', { class: 'cc-progreso', hidden: true }, [h('span')]);
    var res = h('p', { class: 'pa-msg' });
    zona.appendChild(h('div', { class: 'cc-panel' }, [
      h('p', { class: 'pa-msg pa-msg-info', text: miles(detectados) + ' de ' + miles(filas.length) + ' archivos se asociaron solos a un cliente. Revisa o escribe el código de los demás.' }),
      h('div', { class: 'cc-tabla-cont' }, [h('table', { class: 'pa-tabla cc-tabla' }, [
        h('thead', {}, [h('tr', {}, [h('th', { text: 'Archivo' }), h('th', { class: 'num', text: 'Tamaño' }), h('th', { text: 'Cliente' }), h('th', { text: 'Estado' })])]),
        h('tbody', {}, filas.map(function (f) { return f.tr; }))])]),
      puedeReemplazar ? h('label', { class: 'cc-check' }, [reemplazar, ' Reemplazar la nota de voz si el cliente ya tiene una']) : null,
      barra, res, h('div', { class: 'pa-fila pa-fin' }, [subir])
    ]));
    subir.addEventListener('click', function () {
      ocupado(subir, true, 'Subiendo…');
      barra.hidden = false;
      var ok = 0, mal = 0, omitidos = 0, i = 0;
      function siguiente() {
        if (i >= filas.length) return Promise.resolve();
        var f = filas[i++];
        var c = estado.porCodigo[f.entrada.value.trim().toUpperCase()];
        barra.firstChild.style.width = Math.round(i * 100 / filas.length) + '%';
        if (!c) { f.estado.textContent = 'Sin cliente'; f.estado.className = 'cc-estado-celda cc-error'; omitidos++; return siguiente(); }
        if (f.archivo.size > MAX_MB * 1048576) { f.estado.textContent = 'Pesa más de ' + MAX_MB + ' MB'; f.estado.className = 'cc-estado-celda cc-error'; mal++; return siguiente(); }
        f.estado.textContent = 'Subiendo…'; f.estado.className = 'cc-estado-celda';
        return leerBase64(f.archivo).then(function (b64) {
          return api('subirAudio', { codigo: c.codigo, nombreArchivo: f.archivo.name, mime: f.archivo.type, base64: b64, reemplazar: reemplazar.checked });
        }).then(function (r) {
          if (r.ok) { actualizarLocal(r.cliente); f.estado.textContent = '✓ Guardada'; f.estado.className = 'cc-estado-celda cc-ok'; ok++; }
          else if (r.codigo === 'YA_TIENE' || r.codigo === 'SIN_PERMISO') { f.estado.textContent = 'Ya tenía nota (no se reemplazó)'; f.estado.className = 'cc-estado-celda cc-suave'; omitidos++; }
          else { f.estado.textContent = r.error; f.estado.className = 'cc-estado-celda cc-error'; mal++; }
          return siguiente();
        });
      }
      siguiente().then(function () {
        ocupado(subir, false);
        subir.disabled = true;
        mensaje(res, miles(ok) + ' guardadas · ' + miles(omitidos) + ' omitidas · ' + miles(mal) + ' con error', mal ? 'error' : 'ok');
      });
    });
  });

  /* ---------------- módulo: sincronizar carpeta de Drive ---------------- */

  function pintarCarpeta() {
    var el = $('cc-carpeta-enlace');
    el.innerHTML = '';
    if (estado.carpeta) el.appendChild(h('a', { href: estado.carpeta.url, target: '_blank', rel: 'noopener', text: 'Abrir la carpeta "' + estado.carpeta.nombre + '"' }));
  }

  $('cc-sinc').addEventListener('click', function (e) {
    var boton = e.currentTarget, zona = $('cc-sinc-zona');
    var reemplazar = tiene('reemplazarAudio') && $('cc-sinc-reemplazar').checked;
    zona.innerHTML = '';
    var progreso = h('p', { class: 'pa-msg pa-msg-info', text: 'Revisando la carpeta… (puede tardar varios minutos con miles de audios; no cierres esta página)' });
    zona.appendChild(progreso);
    ocupado(boton, true, 'Sincronizando…');
    var tot = { procesados: 0, asociados: 0, yaAsociados: 0, ignorados: 0, sinAsociar: [], conAudio: [] };
    function tanda(continuar) {
      return api('sincronizarAudios', { continuar: continuar || '', reemplazar: reemplazar }).then(function (r) {
        if (!r.ok) throw new Error(r.error);
        ['procesados', 'asociados', 'yaAsociados', 'ignorados'].forEach(function (k) { tot[k] += r[k]; });
        tot.sinAsociar = tot.sinAsociar.concat(r.sinAsociar);
        tot.conAudio = tot.conAudio.concat(r.conAudio);
        progreso.textContent = 'Revisados ' + miles(tot.procesados) + ' archivos · ' + miles(tot.asociados) + ' asociados' + (r.continuar ? '… seguimos' : '');
        if (r.continuar) return tanda(r.continuar);
      });
    }
    tanda('').then(function () {
      ocupado(boton, false);
      mensaje(progreso, 'Listo: ' + miles(tot.procesados) + ' archivos revisados · ' + miles(tot.asociados) + ' asociados ahora · ' +
        miles(tot.yaAsociados) + ' ya estaban asociados · ' + miles(tot.sinAsociar.length) + ' sin asociar · ' +
        miles(tot.conAudio.length) + ' de clientes que ya tenían nota' + (tot.ignorados ? ' · ' + miles(tot.ignorados) + ' archivos que no son audio' : ''), 'ok');
      pintarSueltos(zona, tot.sinAsociar, 'Audios sin asociar', 'No se encontró ningún código de cliente en el nombre. Escribe el código para asignarlo.', false);
      if (tiene('reemplazarAudio')) {
        pintarSueltos(zona, tot.conAudio, 'Clientes que ya tenían nota de voz', 'Estos audios llevan el código de un cliente que ya tiene nota. Pulsa Reemplazar si el nuevo es el correcto.', true);
      } else if (tot.conAudio.length) {
        zona.appendChild(h('p', { class: 'pa-msg pa-msg-info', text: miles(tot.conAudio.length) + ' audios son de clientes que ya tienen nota de voz. Solo un administrador puede reemplazarlas.' }));
      }
      cargarClientes();
    }).catch(function (err) {
      ocupado(boton, false);
      mensaje(progreso, 'La sincronización se detuvo: ' + err.message + ' Lo avanzado quedó guardado; pulsa Sincronizar de nuevo para continuar.');
    });
  });

  function pintarSueltos(zona, lista, titulo, explicacion, reemplazar) {
    if (!lista.length) return;
    var cuerpo = h('tbody');
    lista.slice(0, 500).forEach(function (a) {
      var entrada = h('input', { class: 'pa-input pa-input-sm', value: a.codigo || '', placeholder: 'Código', autocapitalize: 'characters', 'aria-label': 'Código para ' + a.nombre });
      var estadoCelda = h('td', { class: 'cc-estado-celda' });
      var boton = h('button', { type: 'button', class: 'pa-btn pa-btn-sm' + (reemplazar ? ' pa-btn-sec' : ''), text: reemplazar ? 'Reemplazar' : 'Asignar' });
      boton.addEventListener('click', function () {
        var codigo = entrada.value.trim().toUpperCase();
        if (!estado.porCodigo[codigo]) { estadoCelda.textContent = 'Código no encontrado'; estadoCelda.className = 'cc-estado-celda cc-error'; return; }
        ocupado(boton, true, '…');
        api('asignarAudio', { fileId: a.id, codigo: codigo, reemplazar: reemplazar }).then(function (r) {
          ocupado(boton, false);
          if (r.ok) { actualizarLocal(r.cliente); estadoCelda.textContent = '✓ Asignado'; estadoCelda.className = 'cc-estado-celda cc-ok'; boton.disabled = true; }
          else if (r.codigo === 'YA_TIENE' && confirm('Ese cliente ya tiene nota de voz. ¿Reemplazarla por ' + a.nombre + '?')) {
            api('asignarAudio', { fileId: a.id, codigo: codigo, reemplazar: true }).then(function (r2) {
              if (r2.ok) { actualizarLocal(r2.cliente); estadoCelda.textContent = '✓ Reemplazado'; estadoCelda.className = 'cc-estado-celda cc-ok'; boton.disabled = true; }
              else { estadoCelda.textContent = r2.error; estadoCelda.className = 'cc-estado-celda cc-error'; }
            });
          } else if (!r.ok) { estadoCelda.textContent = r.error; estadoCelda.className = 'cc-estado-celda cc-error'; }
        });
      });
      cuerpo.appendChild(h('tr', {}, [h('td', { class: 'cc-nombre-archivo', text: a.nombre }), h('td', {}, [entrada]), h('td', {}, [boton]), estadoCelda]));
    });
    zona.appendChild(h('div', { class: 'cc-panel' }, [
      h('h3', { text: titulo + ' (' + miles(lista.length) + ')' }),
      h('p', { class: 'cc-suave', text: explicacion + (lista.length > 500 ? ' Se muestran los primeros 500.' : '') }),
      h('div', { class: 'cc-tabla-cont' }, [h('table', { class: 'pa-tabla cc-tabla' }, [
        h('thead', {}, [h('tr', {}, [h('th', { text: 'Archivo' }), h('th', { text: 'Código' }), h('th', { text: '' }), h('th', { text: 'Estado' })])]), cuerpo])])
    ]));
  }

  /* ---------------- módulo: formularios e historial ---------------- */

  var TIPOS = [
    ['texto', 'Texto corto'], ['parrafo', 'Párrafo (texto largo)'], ['numero', 'Número'], ['sino', 'Sí / No'],
    ['opcion', 'Una opción de una lista'], ['varias', 'Varias opciones de una lista'], ['fecha', 'Fecha'], ['calificacion', 'Calificación 1 a 5']
  ];
  var NOMBRE_TIPO = {};
  TIPOS.forEach(function (t) { NOMBRE_TIPO[t[0]] = t[1]; });

  function vacio(v) { return v === '' || v === null || v === undefined || (Array.isArray(v) && !v.length); }

  function formatoValor(tipo, v) {
    if (vacio(v)) return '—';
    if (tipo === 'sino') return v === 'si' ? 'Sí' : 'No';
    if (tipo === 'varias') return v.join(', ');
    if (tipo === 'calificacion') return '★'.repeat(v) + '☆'.repeat(5 - v) + '  (' + v + '/5)';
    if (tipo === 'fecha') { var p = String(v).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
    if (tipo === 'numero') return Number(v).toLocaleString('es-DO', { maximumFractionDigits: 6 });
    return String(v);
  }

  function cargarFormularios() {
    return api('listarFormularios').then(function (r) {
      if (r.ok) {
        estado.formularios = r.formularios || [];
        estado.conteoRoles = r.conteoRoles || {};
        if (r.maxPreguntas) estado.maxPreguntas = r.maxPreguntas;
      }
      return r;
    });
  }

  /** Formularios que este usuario puede llenar para este cliente (activos, usuario autorizado y tipo de cliente que aplica). */
  function formulariosPara(c) {
    return estado.formularios.filter(function (f) { return f.activo && f.puedoLlenar !== false && aplicaATipo(f, c.tipo); });
  }

  /** Texto corto de a quién aplica el formulario. */
  function textoTipos(f) {
    return f.tiposCliente && f.tiposCliente.length ? 'Tipos: ' + f.tiposCliente.join(', ') : 'Todos los tipos de cliente';
  }
  var ROLES = [['usuario', 'Usuario'], ['editor', 'Editor'], ['supervisor', 'Supervisor'], ['admin', 'Administrador']];
  var NOMBRE_ROL = {};
  ROLES.forEach(function (r) { NOMBRE_ROL[r[0]] = r[1]; });
  function textoRoles(f) {
    var roles = f.roles || [];
    return roles.length ? 'Lo llenan: ' + roles.map(function (r) { return NOMBRE_ROL[r] || r; }).join(', ') : 'Lo llenan: todos los roles';
  }

  /** Sección de la ficha: botones para llenar formularios + historial de respuestas del cliente. */
  function pintarFormulariosFicha(c, zona) {
    zona.innerHTML = '';
    var botones = h('div', { class: 'cc-form-botones' });
    var hist = h('div', { class: 'cc-historial' }, [h('div', { class: 'pa-cargando', 'aria-hidden': 'true' })]);
    var info = h('p', { class: 'cc-suave' });
    var respondidos = {};

    function pintarBotones() {
      botones.innerHTML = '';
      if (!tiene('responderFormularios')) return;
      var activos = formulariosPara(c);
      if (!activos.length) {
        var hayActivos = estado.formularios.some(function (f) { return f.activo && f.puedoLlenar !== false; });
        botones.appendChild(h('p', { class: 'cc-suave', text: hayActivos
          ? 'No hay formularios para ' + (c.tipo ? 'clientes de tipo "' + c.tipo + '".' : 'clientes sin tipo asignado.')
          : (tiene('disenarFormularios') ? 'No hay formularios activos. Créalos en el menú Formularios.' : 'No tienes formularios asignados para llenar.') }));
        return;
      }
      activos.forEach(function (f) {
        var hecho = f.unaVez && respondidos[f.id];
        botones.appendChild(h('button', { type: 'button', class: 'pa-btn pa-btn-sm ' + (hecho ? 'pa-btn-sec' : ''), disabled: hecho ? true : null,
          title: hecho ? 'Ya respondido el ' + respondidos[f.id] + ' (se responde una sola vez)' : (f.descripcion || ''),
          onclick: function () { llenarFormulario(c, f, function () { cargarHistorial(); }); } },
          [hecho ? '✓ ' + f.titulo : 'Llenar: ' + f.titulo]));
      });
    }

    function cargarHistorial() {
      api('historialCliente', { codigo: c.codigo }).then(function (r) {
        hist.innerHTML = '';
        if (!r.ok) { hist.appendChild(h('p', { class: 'pa-msg pa-msg-error', text: r.error })); return; }
        respondidos = r.respondidos || {};
        pintarBotones();
        info.textContent = r.completo ? '' : 'Ves solo las respuestas que registraste tú.' +
          (r.ocultas ? ' Hay ' + r.ocultas + ' más de otros usuarios (las ven supervisores y administradores).' : '');
        if (!r.historial.length) { hist.appendChild(h('p', { class: 'pa-vacio', text: 'Todavía no hay respuestas de formularios para este cliente.' })); return; }
        r.historial.forEach(function (x, i) { hist.appendChild(entradaHistorial(x, i === 0, cargarHistorial)); });
      });
    }

    zona.appendChild(h('h3', { class: 'cc-seccion-tit', text: 'Formularios e historial' }));
    zona.appendChild(botones);
    zona.appendChild(info);
    zona.appendChild(hist);
    pintarBotones();
    cargarHistorial();
  }

  function entradaHistorial(x, abierta, alCambiar) {
    var lista = h('dl', { class: 'cc-resp' });
    (x.respuestas || []).forEach(function (q) {
      lista.appendChild(h('div', {}, [h('dt', { text: q.p }), h('dd', { class: q.t === 'parrafo' ? 'cc-pre' : '', text: formatoValor(q.t, q.v) })]));
    });
    var pie = null;
    if (tiene('eliminarRespuestas')) {
      pie = h('div', { class: 'pa-fila pa-fin' }, [h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar respuesta', onclick: function (e) {
        if (!confirm('¿Eliminar esta respuesta de "' + x.formulario + '" del ' + x.fecha + '? No se puede deshacer.')) return;
        var b = e.currentTarget; ocupado(b, true, 'Eliminando…');
        api('eliminarRespuesta', { id: x.id }).then(function (r) {
          ocupado(b, false);
          if (!r.ok) { aviso(r.error, 'error'); return; }
          aviso(r.mensaje); if (alCambiar) alCambiar();
        });
      } })]);
    }
    return h('details', { class: 'cc-hist-item', open: abierta ? true : null }, [
      h('summary', {}, [h('strong', { text: x.formulario }), h('span', { class: 'cc-suave', text: x.fecha + ' · ' + x.usuario })]),
      lista, pie
    ]);
  }

  /** Ventana para llenar un formulario de un cliente. */
  function llenarFormulario(c, f, alGuardar) {
    var lectores = {};
    var bloques = f.preguntas.map(function (q, i) {
      var etq = h('span', { class: 'pa-etq', text: (i + 1) + '. ' + q.texto + (q.requerida ? ' *' : '') });
      var control, nombre = 'q-' + q.id + '-' + Math.random().toString(36).slice(2, 7);
      function radios(valores, textos, clase) {
        return h('div', { class: 'cc-opciones ' + (clase || ''), role: 'radiogroup' }, valores.map(function (v, j) {
          return h('label', { class: 'cc-opcion' }, [h('input', { type: 'radio', name: nombre, value: String(v) }), h('span', { text: textos[j] })]);
        }));
      }
      function marcado() { var el = control.querySelector('input:checked'); return el ? el.value : ''; }
      switch (q.tipo) {
        case 'parrafo':
          control = h('textarea', { class: 'pa-input cc-textarea', rows: '3', maxlength: '2000' });
          lectores[q.id] = function () { return control.value.trim(); }; break;
        case 'numero':
          control = h('input', { class: 'pa-input', type: 'text', inputmode: 'decimal', maxlength: '20' });
          lectores[q.id] = function () { return control.value.trim(); }; break;
        case 'fecha':
          control = h('input', { class: 'pa-input', type: 'date' });
          lectores[q.id] = function () { return control.value; }; break;
        case 'sino':
          control = radios(['si', 'no'], ['Sí', 'No'], 'cc-pastillas'); lectores[q.id] = marcado; break;
        case 'calificacion':
          control = radios([1, 2, 3, 4, 5], ['1', '2', '3', '4', '5'], 'cc-pastillas cc-calif');
          lectores[q.id] = function () { var v = marcado(); return v ? Number(v) : ''; }; break;
        case 'opcion':
          if (q.opciones.length > 8) {
            control = h('select', { class: 'pa-input' }, [h('option', { value: '', text: '— Elige —' })].concat(q.opciones.map(function (o) { return h('option', { value: o, text: o }); })));
            lectores[q.id] = function () { return control.value; };
          } else { control = radios(q.opciones, q.opciones); lectores[q.id] = marcado; }
          break;
        case 'varias':
          control = h('div', { class: 'cc-opciones' }, q.opciones.map(function (o) {
            return h('label', { class: 'cc-opcion' }, [h('input', { type: 'checkbox', value: o }), h('span', { text: o })]);
          }));
          lectores[q.id] = function () { return [].slice.call(control.querySelectorAll('input:checked')).map(function (x) { return x.value; }); };
          break;
        default:
          control = h('input', { class: 'pa-input', type: 'text', maxlength: '300' });
          lectores[q.id] = function () { return control.value.trim(); };
      }
      var envoltura = q.tipo === 'sino' || q.tipo === 'calificacion' || q.tipo === 'varias' || (q.tipo === 'opcion' && q.opciones.length <= 8) ? 'div' : 'label';
      return h(envoltura, { class: 'pa-campo cc-pregunta', 'data-q': q.id }, [etq, control]);
    });
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn', text: 'Guardar respuestas' });
    var form = h('form', { class: 'pa-form', novalidate: true }, [
      f.descripcion ? h('p', { class: 'cc-suave cc-pre', text: f.descripcion }) : null
    ].concat(bloques, [msg, h('div', { class: 'pa-fila pa-fin' }, [guardar])]));
    var m = modal(f.titulo, form, { sub: c.codigo + ' — ' + c.nombre + (c.tipo ? ' · ' + c.tipo : '') });
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var respuestas = {}, falta = null, alguna = false;
      f.preguntas.forEach(function (q) {
        var v = lectores[q.id]();
        respuestas[q.id] = v;
        if (!vacio(v)) alguna = true;
        if (q.requerida && vacio(v) && !falta) falta = q;
        if (q.tipo === 'numero' && v && !isFinite(Number(String(v).replace(/\s/g, '').replace(',', '.'))) && !falta) falta = q;
      });
      [].slice.call(form.querySelectorAll('.cc-pregunta')).forEach(function (el) { el.classList.toggle('cc-falta', !!falta && el.getAttribute('data-q') === falta.id); });
      if (falta) { mensaje(msg, (falta.tipo === 'numero' && respuestas[falta.id] ? 'Escribe un número válido en: ' : 'Falta responder: ') + falta.texto); return; }
      if (!alguna) { mensaje(msg, 'Responde al menos una pregunta.'); return; }
      ocupado(guardar, true, 'Guardando…');
      api('responderFormulario', { formularioId: f.id, codigo: c.codigo, respuestas: respuestas }).then(function (r) {
        ocupado(guardar, false);
        if (!r.ok) { mensaje(msg, r.error); return; }
        aviso(r.mensaje); m.cerrar();
        if (alGuardar) alGuardar(r.respuesta);
      });
    });
    setTimeout(function () { var el = form.querySelector('input,textarea,select'); if (el) el.focus(); }, 50);
  }

  /* ---- vista Formularios (supervisor: ver respuestas · admin: además diseñar) ---- */

  $('cc-form-nuevo').addEventListener('click', function () { editarFormulario(null); });
  var tempForm = null;
  $('cc-form-buscar').addEventListener('input', function () { clearTimeout(tempForm); tempForm = setTimeout(pintarListaFormularios, 150); });
  $('cc-form-f-tipo').addEventListener('change', pintarListaFormularios);

  /** Llena el filtro "tipo de cliente" de la vista Formularios con los tipos usados en clientes y formularios. */
  function llenarFiltroTiposForm() {
    var sel = $('cc-form-f-tipo'), actual = sel.value, vistos = {};
    valoresUnicos('tipo').forEach(function (t) { vistos[normalizar(t)] = t; });
    estado.formularios.forEach(function (f) { (f.tiposCliente || []).forEach(function (t) { if (!vistos[normalizar(t)]) vistos[normalizar(t)] = t; }); });
    var tipos = Object.keys(vistos).map(function (k) { return vistos[k]; }).sort(function (a, b) { return a.localeCompare(b, 'es'); });
    sel.innerHTML = '';
    sel.appendChild(h('option', { value: '', text: 'Todos los tipos de cliente' }));
    tipos.forEach(function (t) { sel.appendChild(h('option', { value: t, text: 'Para: ' + t })); });
    sel.value = actual;
    sel.hidden = !tipos.length;
  }

  function pintarListaFormularios() {
    var cont = $('cc-form-lista');
    cont.innerHTML = '';
    var todos = estado.formularios;
    $('cc-form-filtros').hidden = todos.length < 2;
    llenarFiltroTiposForm();
    var q = normalizar($('cc-form-buscar').value), tipo = $('cc-form-f-tipo').value;
    var lista = todos.filter(function (f) {
      if (q && normalizar(f.titulo + ' ' + (f.descripcion || '') + ' ' + (f.tiposCliente || []).join(' ')).indexOf(q) < 0) return false;
      if (tipo && !aplicaATipo(f, tipo)) return false;
      return true;
    });
    if (!todos.length) {
      cont.appendChild(h('p', { class: 'pa-vacio', text: tiene('disenarFormularios') ? 'Todavía no hay formularios. Pulsa "+ Nuevo formulario".' : 'No tienes formularios asignados para llenar.' }));
      return;
    }
    if (!lista.length) { cont.appendChild(h('p', { class: 'pa-vacio', text: 'Ningún formulario coincide con el filtro.' })); return; }
    var ul = h('ul', { class: 'cc-form-lista' });
    lista.forEach(function (f) {
      var acciones = [];
      if (f.activo && f.puedoLlenar) acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sm', text: 'Llenar', onclick: function () { elegirCliente(f); } }));
      if (tiene('verRespuestas')) acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sm' + (f.activo && f.puedoLlenar ? ' pa-btn-sec' : ''), text: 'Ver respuestas', onclick: function () { verRespuestas(f); } }));
      if (tiene('disenarFormularios')) {
        acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Editar', onclick: function () { editarFormulario(f); } }));
        acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Duplicar', title: 'Crear una copia (por ejemplo, para otro tipo de cliente)', onclick: function () { editarFormulario(f, true); } }));
        acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: f.activo ? 'Desactivar' : 'Activar', onclick: function (e) {
          var b = e.currentTarget; ocupado(b, true, '…');
          var copia = JSON.parse(JSON.stringify(f)); copia.activo = !f.activo;
          api('guardarFormulario', { formulario: copia }).then(function (r) {
            ocupado(b, false);
            if (!r.ok) { aviso(r.error, 'error'); return; }
            aviso(copia.activo ? 'Formulario activado' : 'Formulario desactivado: ya no aparece en las fichas');
            cargarFormularios().then(pintarListaFormularios);
          });
        } }));
        acciones.push(h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar', onclick: function (e) {
          if (!confirm('¿Eliminar el formulario "' + f.titulo + '"?\n\nLas ' + miles(f.respuestas || 0) + ' respuestas ya guardadas se conservan en el historial de cada cliente. Si solo quieres que no se siga llenando, usa "Desactivar".')) return;
          var b = e.currentTarget; ocupado(b, true, '…');
          api('eliminarFormulario', { id: f.id }).then(function (r) {
            ocupado(b, false);
            if (!r.ok) { aviso(r.error, 'error'); return; }
            aviso(r.mensaje); cargarFormularios().then(pintarListaFormularios);
          });
        } }));
      }
      ul.appendChild(h('li', { class: 'cc-form-item' + (f.activo ? '' : ' cc-inactivo') }, [
        h('div', { class: 'cc-form-info' }, [
          h('strong', { text: f.titulo }),
          h('div', { class: 'cc-etiquetas' }, [
            h('span', { class: 'pa-tag ' + (f.activo ? 'pa-tag-editor' : ''), text: f.activo ? 'Activo' : 'Inactivo' }),
            f.unaVez ? h('span', { class: 'pa-tag', text: 'Una vez por cliente' }) : null,
            f.activo && !f.puedoLlenar ? h('span', { class: 'pa-tag pa-tag-yo', text: 'No es para tu rol' }) : null,
            h('span', { class: 'cc-suave', text: f.preguntas.length + (f.preguntas.length === 1 ? ' pregunta' : ' preguntas') +
              (f.respuestas !== undefined ? ' · ' + miles(f.respuestas) + (f.respuestas === 1 ? ' respuesta' : ' respuestas') : '') + (f.actualizado ? ' · editado ' + f.actualizado : '') })
          ]),
          h('div', { class: 'cc-etiquetas cc-alcance' }, [
            h('span', { class: 'cc-chip cc-chip-tipo', text: textoTipos(f) }),
            h('span', { class: 'cc-chip cc-chip-usu', text: textoRoles(f) })
          ]),
          f.descripcion ? h('small', { class: 'cc-suave', text: f.descripcion }) : null
        ]),
        h('div', { class: 'pa-fila cc-form-acciones' }, acciones)
      ]));
    });
    cont.appendChild(ul);
  }

  /**
   * Lista de casillas con buscador opcional. items: [{valor, texto, detalle}]. Devuelve {el, elegidos(), refrescar()}.
   */
  function grupoCasillas(items, marcados, opciones) {
    opciones = opciones || {};
    var set = {};
    (marcados || []).forEach(function (v) { set[v] = true; });
    var cont = h('div', { class: 'cc-casillas' });
    var buscar = opciones.buscador ? h('input', { class: 'pa-input pa-input-sm', type: 'search', placeholder: opciones.buscador }) : null;
    var cuenta = h('small', { class: 'cc-suave' });
    function pintar() {
      cont.innerHTML = '';
      var q = buscar ? normalizar(buscar.value) : '';
      var vis = items.filter(function (it) { return !q || normalizar(it.texto + ' ' + (it.detalle || '') + ' ' + it.valor).indexOf(q) >= 0; });
      vis.forEach(function (it) {
        var cb = h('input', { type: 'checkbox', checked: set[it.valor] ? true : null });
        cb.addEventListener('change', function () { if (cb.checked) set[it.valor] = true; else delete set[it.valor]; contar(); });
        cont.appendChild(h('label', { class: 'cc-casilla' + (it.apagado ? ' cc-apagado' : '') }, [cb, h('span', {}, [it.texto, it.detalle ? h('small', { text: ' ' + it.detalle }) : null])]));
      });
      if (!vis.length) cont.appendChild(h('p', { class: 'cc-suave', text: items.length ? 'Nada coincide con la búsqueda.' : (opciones.vacio || 'No hay opciones.') }));
      contar();
      return vis;
    }
    function contar() { var n = Object.keys(set).length; cuenta.textContent = n + (n === 1 ? ' elegido' : ' elegidos'); }
    var marcarVis = h('button', { type: 'button', class: 'pa-enlace', text: 'Marcar los visibles', onclick: function () {
      var q = buscar ? normalizar(buscar.value) : '';
      items.forEach(function (it) { if (!q || normalizar(it.texto + ' ' + (it.detalle || '') + ' ' + it.valor).indexOf(q) >= 0) set[it.valor] = true; });
      pintar();
    } });
    var quitar = h('button', { type: 'button', class: 'pa-enlace', text: 'Quitar todos', onclick: function () { set = {}; pintar(); } });
    if (buscar) buscar.addEventListener('input', pintar);
    pintar();
    return {
      el: h('div', { class: 'cc-casillas-cont' }, [h('div', { class: 'cc-casillas-barra' }, [buscar, cuenta, marcarVis, quitar]), cont]),
      elegidos: function () { return items.map(function (it) { return it.valor; }).filter(function (v) { return set[v]; }); },
      agregar: function (it) {
        if (!items.some(function (x) { return normalizar(x.valor) === normalizar(it.valor); })) items.push(it);
        var existente = items.filter(function (x) { return normalizar(x.valor) === normalizar(it.valor); })[0];
        set[existente.valor] = true; pintar();
      }
    };
  }

  /** Selector "Todos" / "Solo los elegidos" con su lista de casillas. */
  function selectorAlcance(nombre, textoTodos, textoSolo, soloInicial, contenido) {
    var radioTodos = h('input', { type: 'radio', name: nombre, value: 'todos', checked: soloInicial ? null : true });
    var radioSolo = h('input', { type: 'radio', name: nombre, value: 'solo', checked: soloInicial ? true : null });
    var zona = h('div', { class: 'cc-alcance-zona' }, [contenido]);
    function ver() { zona.hidden = !radioSolo.checked; }
    radioTodos.addEventListener('change', ver); radioSolo.addEventListener('change', ver); ver();
    return {
      el: h('div', { class: 'cc-alcance-sel' }, [
        h('div', { class: 'cc-opciones cc-pastillas', role: 'radiogroup' }, [
          h('label', { class: 'cc-opcion' }, [radioTodos, h('span', { text: textoTodos })]),
          h('label', { class: 'cc-opcion' }, [radioSolo, h('span', { text: textoSolo })])
        ]), zona]),
      solo: function () { return radioSolo.checked; }
    };
  }

  /** Editor de formularios (crear, modificar o duplicar). */
  function editarFormulario(f, duplicar) {
    var nuevo = !f || !!duplicar;
    f = f ? JSON.parse(JSON.stringify(f)) : { titulo: '', descripcion: '', activo: true, unaVez: false, preguntas: [], tiposCliente: [], roles: [] };
    if (duplicar) {
      f.id = ''; f.titulo = (f.titulo + ' (copia)').slice(0, 100); f.respuestas = 0;
      f.preguntas.forEach(function (q) { delete q.id; });
    }
    f.tiposCliente = f.tiposCliente || [];
    f.roles = f.roles || [];
    var titulo = h('input', { class: 'pa-input', type: 'text', maxlength: '100', placeholder: 'Ej.: Visita de supervisión' });
    titulo.value = f.titulo;
    var desc = h('textarea', { class: 'pa-input cc-textarea', rows: '2', maxlength: '500', placeholder: 'Instrucciones para quien lo llena (opcional)' });
    desc.value = f.descripcion || '';
    var activo = h('input', { type: 'checkbox', checked: f.activo !== false ? true : null });
    var unaVez = h('input', { type: 'checkbox', checked: f.unaVez ? true : null });
    var zonaPreg = h('div', { class: 'cc-preguntas' });
    var filas = [];

    function filaPregunta(q) {
      q = q || { texto: '', tipo: 'texto', requerida: false, opciones: [] };
      var texto = h('input', { class: 'pa-input', type: 'text', maxlength: '200', placeholder: 'Escribe la pregunta' });
      texto.value = q.texto || '';
      var tipo = h('select', { class: 'pa-input' }, TIPOS.map(function (t) { return h('option', { value: t[0], text: t[1] }); }));
      tipo.value = q.tipo || 'texto';
      var req = h('input', { type: 'checkbox', checked: q.requerida ? true : null });
      var ops = h('textarea', { class: 'pa-input cc-textarea', rows: '3', placeholder: 'Una opción por línea' });
      ops.value = (q.opciones || []).join('\n');
      var campoOps = h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Opciones (una por línea)' }), ops]);
      function verOps() { campoOps.hidden = ['opcion', 'varias'].indexOf(tipo.value) < 0; }
      tipo.addEventListener('change', verOps); verOps();
      var num = h('span', { class: 'cc-preg-num' });
      var fila = { q: q, el: null };
      function mover(d) {
        var i = filas.indexOf(fila), j = i + d;
        if (j < 0 || j >= filas.length) return;
        filas.splice(i, 1); filas.splice(j, 0, fila); repintar();
      }
      fila.el = h('div', { class: 'cc-preg' }, [
        h('div', { class: 'cc-preg-cab' }, [num,
          h('div', { class: 'cc-preg-btns' }, [
            h('button', { type: 'button', class: 'pa-btn-icono', title: 'Subir', 'aria-label': 'Subir pregunta', text: '↑', onclick: function () { mover(-1); } }),
            h('button', { type: 'button', class: 'pa-btn-icono', title: 'Bajar', 'aria-label': 'Bajar pregunta', text: '↓', onclick: function () { mover(1); } }),
            h('button', { type: 'button', class: 'pa-btn-icono cc-quitar', title: 'Quitar', 'aria-label': 'Quitar pregunta', text: '✕', onclick: function () {
              if (texto.value.trim() && !confirm('¿Quitar la pregunta "' + texto.value.trim() + '"?')) return;
              filas.splice(filas.indexOf(fila), 1); repintar();
            } })
          ])]),
        h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Pregunta' }), texto]),
        h('div', { class: 'cc-preg-fila' }, [
          h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Tipo de respuesta' }), tipo]),
          h('label', { class: 'cc-check' }, [req, ' Obligatoria'])
        ]),
        campoOps
      ]);
      fila.num = num;
      fila.leer = function () {
        return { id: q.id, texto: texto.value.trim(), tipo: tipo.value, requerida: req.checked,
                 opciones: ops.value.split('\n').map(function (s) { return s.trim(); }).filter(Boolean) };
      };
      fila.enfocar = function () { texto.focus(); };
      return fila;
    }
    var conteoPreg = h('small', { class: 'cc-suave' });
    function repintar() {
      zonaPreg.innerHTML = '';
      filas.forEach(function (fl, i) { fl.num.textContent = 'Pregunta ' + (i + 1); zonaPreg.appendChild(fl.el); });
      if (!filas.length) zonaPreg.appendChild(h('p', { class: 'cc-suave', text: 'Agrega la primera pregunta.' }));
      conteoPreg.textContent = filas.length + ' de ' + estado.maxPreguntas + ' preguntas posibles';
    }
    f.preguntas.forEach(function (q) { filas.push(filaPregunta(q)); });
    if (!filas.length) filas.push(filaPregunta());
    repintar();

    var agregar = h('button', { type: 'button', class: 'pa-btn pa-btn-sec', text: '+ Agregar pregunta', onclick: function () {
      if (filas.length >= estado.maxPreguntas) { mensaje(msg, 'Máximo ' + estado.maxPreguntas + ' preguntas por formulario.'); return; }
      var fl = filaPregunta(); filas.push(fl); repintar(); fl.enfocar();
    } });

    // ---- ¿Para qué tipos de cliente? ----
    var conteoTipo = {};
    estado.clientes.forEach(function (c) { if (c.tipo && !c.oculto) conteoTipo[c.tipo] = (conteoTipo[c.tipo] || 0) + 1; });
    var itemsTipo = Object.keys(conteoTipo).sort(function (a, b) { return a.localeCompare(b, 'es'); })
      .map(function (t) { return { valor: t, texto: t, detalle: '(' + miles(conteoTipo[t]) + (conteoTipo[t] === 1 ? ' cliente)' : ' clientes)') }; });
    f.tiposCliente.forEach(function (t) {
      var ya = itemsTipo.filter(function (x) { return normalizar(x.valor) === normalizar(t); })[0];
      if (!ya) itemsTipo.push({ valor: t, texto: t, detalle: '(sin clientes todavía)' });
    });
    var marcadosTipo = itemsTipo.filter(function (x) { return f.tiposCliente.some(function (t) { return normalizar(t) === normalizar(x.valor); }); })
      .map(function (x) { return x.valor; });
    var casTipos = grupoCasillas(itemsTipo, marcadosTipo, { vacio: 'Todavía no hay clientes con tipo. Escribe un tipo abajo o asígnalo en la ficha / importación de clientes.' });
    var nuevoTipo = h('input', { class: 'pa-input pa-input-sm', type: 'text', maxlength: '120', placeholder: 'Otro tipo (ej.: Farmacia)' });
    function agregarTipo() { var t = nuevoTipo.value.trim(); if (!t) return; casTipos.agregar({ valor: t, texto: t, detalle: '(nuevo)' }); nuevoTipo.value = ''; nuevoTipo.focus(); }
    nuevoTipo.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); agregarTipo(); } });
    var selTipos = selectorAlcance('alc-tipo-' + Date.now(), 'Todos los tipos de cliente', 'Solo estos tipos', f.tiposCliente.length > 0,
      h('div', {}, [casTipos.el, h('div', { class: 'pa-fila' }, [nuevoTipo, h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Agregar tipo', onclick: agregarTipo })])]));

    // ---- ¿Qué usuarios lo pueden llenar? ----
    var itemsRol = ROLES.map(function (r) {
      var n = estado.conteoRoles[r[0]];
      return { valor: r[0], texto: r[1], detalle: n === undefined ? '' : '(' + miles(n) + (n === 1 ? ' persona)' : ' personas)') };
    });
    var casRol = grupoCasillas(itemsRol, f.roles, {});
    var selRol = selectorAlcance('alc-rol-' + Date.now(), 'Todos los roles', 'Solo estos roles', f.roles.length > 0, casRol.el);
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });
    var guardar = h('button', { type: 'submit', class: 'pa-btn', text: nuevo ? 'Crear formulario' : 'Guardar cambios' });
    var form = h('form', { class: 'pa-form', novalidate: true }, [
      h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Título *' }), titulo]),
      h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: 'Descripción' }), desc]),
      h('div', { class: 'cc-preg-fila' }, [
        h('label', { class: 'cc-check' }, [activo, ' Activo (aparece en las fichas para llenarlo)']),
        h('label', { class: 'cc-check' }, [unaVez, ' Cada cliente lo responde una sola vez'])
      ]),
      !nuevo && f.respuestas ? h('p', { class: 'pa-msg pa-msg-info', text: 'Este formulario ya tiene ' + miles(f.respuestas) + ' respuestas. Si cambias las preguntas, las respuestas anteriores no cambian: cada una conserva las preguntas tal como estaban.' }) : null,
      h('h3', { class: 'cc-seccion-tit', text: '¿Para qué tipos de cliente?' }),
      h('p', { class: 'cc-suave', text: 'El formulario solo aparece en la ficha de los clientes de los tipos elegidos.' }),
      selTipos.el,
      h('h3', { class: 'cc-seccion-tit', text: '¿Quién lo puede llenar?' }),
      h('p', { class: 'cc-suave', text: 'Solo las personas con los roles elegidos lo verán para llenarlo. El rol de cada persona se cambia en el menú Usuarios. Supervisores y administradores pueden ver las respuestas aunque su rol no esté elegido.' }),
      selRol.el,
      h('div', { class: 'cc-tarjeta-cab' }, [h('h3', { class: 'cc-seccion-tit', text: 'Preguntas' }), conteoPreg]),
      zonaPreg, h('div', {}, [agregar]),
      msg, h('div', { class: 'pa-fila pa-fin' }, [guardar])
    ]);
    repintar();
    var m = modal(duplicar ? 'Duplicar formulario' : (nuevo ? 'Nuevo formulario' : 'Editar formulario'), form, { ancho: true });
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var preguntas = filas.map(function (fl) { return fl.leer(); });
      var tiposElegidos = selTipos.solo() ? casTipos.elegidos() : [];
      var rolesElegidos = selRol.solo() ? casRol.elegidos() : [];
      var err = !titulo.value.trim() ? 'Escribe el título del formulario.' : (!preguntas.length ? 'Agrega al menos una pregunta.' : '');
      if (!err && selTipos.solo() && !tiposElegidos.length) err = 'Elige al menos un tipo de cliente, o marca "Todos los tipos de cliente".';
      if (!err && selRol.solo() && !rolesElegidos.length) err = 'Elige al menos un rol, o marca "Todos los roles".';
      if (!err && preguntas.length > estado.maxPreguntas) err = 'Máximo ' + estado.maxPreguntas + ' preguntas por formulario.';
      preguntas.forEach(function (q, i) {
        if (err) return;
        if (!q.texto) err = 'La pregunta ' + (i + 1) + ' no tiene texto.';
        else if ((q.tipo === 'opcion' || q.tipo === 'varias') && q.opciones.length < 2) err = 'La pregunta ' + (i + 1) + ' necesita al menos 2 opciones (una por línea).';
      });
      if (err) { mensaje(msg, err); return; }
      ocupado(guardar, true, 'Guardando…');
      api('guardarFormulario', { formulario: { id: nuevo ? '' : f.id, titulo: titulo.value.trim(), descripcion: desc.value.trim(),
        activo: activo.checked, unaVez: unaVez.checked, preguntas: preguntas,
        tiposCliente: tiposElegidos, roles: rolesElegidos, soloTipos: selTipos.solo(), soloRoles: selRol.solo() } }).then(function (r) {
        ocupado(guardar, false);
        if (!r.ok) { mensaje(msg, r.error); return; }
        aviso(r.mensaje); m.cerrar();
        cargarFormularios().then(pintarListaFormularios);
      });
    });
    setTimeout(function () { titulo.focus(); }, 50);
  }

  /**
   * "Llenar" desde la vista Formularios: elige el cliente (solo los tipos a los que aplica el formulario)
   * y abre el formulario. La ventana queda abierta para seguir llenando a otros clientes.
   */
  function elegirCliente(f) {
    var aptos = estado.clientes.filter(function (c) { return !c.oculto && aplicaATipo(f, c.tipo); });
    var llenados = {};
    var buscar = h('input', { class: 'pa-input', type: 'search', placeholder: 'Buscar cliente por código, nombre, dirección o teléfono…', autocomplete: 'off' });
    var selRuta = h('select', { class: 'pa-input', 'aria-label': 'Filtrar por ruta' });
    var selTipo = h('select', { class: 'pa-input', 'aria-label': 'Filtrar por tipo de cliente' });
    function llenarSelect(sel, campo, todos) {
      var vistos = {};
      aptos.forEach(function (c) { if (c[campo]) vistos[c[campo]] = true; });
      sel.innerHTML = '';
      sel.appendChild(h('option', { value: '', text: todos }));
      Object.keys(vistos).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); })
        .forEach(function (v) { sel.appendChild(h('option', { value: v, text: v })); });
      sel.hidden = sel.options.length <= 2;
    }
    llenarSelect(selRuta, 'ruta', 'Todas las rutas');
    llenarSelect(selTipo, 'tipo', 'Todos los tipos');
    var info = h('p', { class: 'cc-suave' });
    var ul = h('ul', { class: 'cc-lista cc-lista-elegir' });
    function pintar() {
      var q = normalizar(buscar.value), terminos = q ? q.split(' ') : [], ruta = selRuta.value, tipo = selTipo.value;
      var vis = aptos.filter(function (c) {
        if (ruta && c.ruta !== ruta) return false;
        if (tipo && c.tipo !== tipo) return false;
        for (var i = 0; i < terminos.length; i++) if (c._txt.indexOf(terminos[i]) < 0) return false;
        return true;
      });
      info.textContent = miles(vis.length) + (vis.length === 1 ? ' cliente' : ' clientes') +
        (f.tiposCliente && f.tiposCliente.length ? ' de tipo ' + f.tiposCliente.join(', ') : '') +
        (vis.length > 50 ? ' · se muestran 50, escribe para buscar' : '');
      ul.innerHTML = '';
      vis.slice(0, 50).forEach(function (c) {
        var detalle = [c.tipo, c.ruta ? 'Ruta ' + c.ruta : '', c.sector].filter(Boolean).join(' · ');
        ul.appendChild(h('li', {}, [h('button', { type: 'button', class: 'cc-cliente' + (llenados[c.codigo] ? ' cc-llenado' : ''), onclick: function () {
          llenarFormulario(c, f, function () { llenados[c.codigo] = true; pintar(); });
        } }, [
          h('span', { class: 'cc-estado-audio ' + (llenados[c.codigo] ? 'si' : 'no'), text: llenados[c.codigo] ? '✓' : '' }),
          h('span', { class: 'cc-cliente-txt' }, [h('strong', { text: c.nombre || '(sin nombre)' }),
            h('small', {}, [h('span', { class: 'cc-codigo', text: c.codigo }), detalle ? ' ' + detalle : '']),
            c.direccion ? h('small', { class: 'cc-dir', text: c.direccion }) : null])
        ])]));
      });
      var tiposF = f.tiposCliente || [];
      if (!vis.length) ul.appendChild(h('li', { class: 'cc-vacio', text: aptos.length ? 'Ningún cliente coincide con la búsqueda.'
        : (tiposF.length ? 'No hay clientes de ' + (tiposF.length > 1 ? 'los tipos ' : 'tipo ') + tiposF.join(', ') + '. Asigna el tipo en la ficha del cliente o en la importación.'
                         : 'Todavía no hay clientes cargados.') }));
    }
    var temp = null;
    buscar.addEventListener('input', function () { clearTimeout(temp); temp = setTimeout(pintar, 120); });
    selRuta.addEventListener('change', pintar);
    selTipo.addEventListener('change', pintar);
    modal('Llenar: ' + f.titulo, h('div', { class: 'cc-elegir' }, [
      h('p', { class: 'cc-suave', text: 'Elige el cliente. Al guardar, puedes seguir con el siguiente.' }),
      buscar, h('div', { class: 'cc-filtros' }, [selTipo, selRuta]), info, ul
    ]), { ancho: true, sub: textoTipos(f) });
    pintar();
    setTimeout(function () { buscar.focus(); }, 50);
  }

  /** Tabla con todas las respuestas de un formulario + descarga a Excel. */
  function verRespuestas(f) {
    var cont = h('div', { class: 'cc-ver-resp' }, [h('div', { class: 'pa-cargando', 'aria-hidden': 'true' })]);
    modal('Respuestas: ' + f.titulo, cont, { ancho: true });
    api('respuestasFormulario', { id: f.id }).then(function (r) {
      cont.innerHTML = '';
      if (!r.ok) { cont.appendChild(h('p', { class: 'pa-msg pa-msg-error', text: r.error })); return; }
      var lista = r.respuestas;
      if (!lista.length) { cont.appendChild(h('p', { class: 'pa-vacio', text: 'Todavía nadie ha llenado este formulario.' })); return; }
      // Columnas: preguntas actuales en orden y luego las que ya no existen (de respuestas viejas).
      var cols = [], porId = {};
      function agregarCol(id, texto, tipo) { if (porId[id]) return; porId[id] = { id: id, texto: texto, tipo: tipo }; cols.push(porId[id]); }
      f.preguntas.forEach(function (q) { agregarCol(q.id, q.texto, q.tipo); });
      lista.forEach(function (x) { (x.respuestas || []).forEach(function (q) { agregarCol(q.id, q.p, q.t); }); });
      function valor(x, col) {
        var q = (x.respuestas || []).filter(function (a) { return a.id === col.id; })[0];
        return q ? formatoValor(q.t, q.v) : '';
      }
      /** Tipo del cliente al momento de responder (o el actual, en respuestas anteriores a este cambio). */
      function tipoDe(x) { return x.tipoCliente || (estado.porCodigo[x.codigo] || {}).tipo || ''; }
      var buscar = h('input', { class: 'pa-input', type: 'search', placeholder: 'Filtrar por cliente, código, tipo, usuario o respuesta…' });
      var excel = h('button', { type: 'button', class: 'pa-btn pa-btn-sec', text: 'Descargar Excel' });
      var info = h('p', { class: 'cc-suave' });
      var tablaCont = h('div', { class: 'cc-tabla-cont' });
      function filtradas() {
        var q = normalizar(buscar.value);
        if (!q) return lista;
        return lista.filter(function (x) {
          return normalizar([x.fecha, x.codigo, x.cliente, tipoDe(x), x.usuario].concat(cols.map(function (c) { return valor(x, c); })).join(' ')).indexOf(q) >= 0;
        });
      }
      function pintar() {
        var filas = filtradas();
        info.textContent = miles(filas.length) + ' de ' + miles(lista.length) + ' respuestas' + (filas.length > 300 ? ' · se muestran las 300 más recientes (el Excel trae todas)' : '');
        tablaCont.innerHTML = '';
        tablaCont.appendChild(h('table', { class: 'pa-tabla cc-tabla cc-tabla-resp' }, [
          h('thead', {}, [h('tr', {}, ['Fecha', 'Código', 'Cliente', 'Tipo', 'Usuario'].concat(cols.map(function (c) { return c.texto; })).map(function (t) { return h('th', { text: t }); }))]),
          h('tbody', {}, filas.slice(0, 300).map(function (x) {
            var tr = h('tr', { class: estado.porCodigo[x.codigo] ? 'cc-fila-clic' : '', title: estado.porCodigo[x.codigo] ? 'Abrir la ficha del cliente' : '' },
              [h('td', { 'data-etq': 'Fecha', text: x.fecha }), h('td', { 'data-etq': 'Código' }, [h('span', { class: 'cc-codigo', text: x.codigo })]),
               h('td', { 'data-etq': 'Cliente', text: x.cliente }), h('td', { 'data-etq': 'Tipo', text: tipoDe(x) }),
               h('td', { 'data-etq': 'Usuario', text: x.usuario })]
                .concat(cols.map(function (c) { return h('td', { 'data-etq': c.texto, text: valor(x, c) }); })));
            if (estado.porCodigo[x.codigo]) tr.addEventListener('click', function () { abrirFicha(x.codigo); });
            return tr;
          }))
        ]));
      }
      var temp = null;
      buscar.addEventListener('input', function () { clearTimeout(temp); temp = setTimeout(pintar, 200); });
      excel.addEventListener('click', function () {
        ocupado(excel, true, 'Preparando…');
        cargarScript(XLSX_URL).then(function () {
          var aoa = [['Fecha', 'Código', 'Cliente', 'Tipo de cliente', 'Ruta', 'Vendedor/Chofer', 'Usuario'].concat(cols.map(function (c) { return c.texto; }))];
          filtradas().forEach(function (x) {
            var cli = estado.porCodigo[x.codigo] || {};
            aoa.push([x.fecha, x.codigo, x.cliente, tipoDe(x), cli.ruta || '', cli.vendedor || '', x.usuario].concat(cols.map(function (c) {
              var q = (x.respuestas || []).filter(function (a) { return a.id === c.id; })[0];
              if (!q || vacio(q.v)) return '';
              if (q.t === 'numero' || q.t === 'calificacion') return Number(q.v);
              return formatoValor(q.t, q.v);
            })));
          });
          var ws = XLSX.utils.aoa_to_sheet(aoa);
          ws['!cols'] = aoa[0].map(function (t, i) { return { wch: i < 7 ? [16, 12, 30, 16, 10, 18, 14][i] : Math.min(40, Math.max(12, String(t).length)) }; });
          var wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'Respuestas');
          var nombre = normalizar(f.titulo).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'formulario';
          XLSX.writeFile(wb, 'respuestas-' + nombre + '-' + new Date().toISOString().slice(0, 10) + '.xlsx');
          ocupado(excel, false);
        }).catch(function () { ocupado(excel, false); aviso('No se pudo preparar el Excel. Revisa tu conexión.', 'error'); });
      });
      cont.appendChild(h('div', { class: 'pa-fila pa-entre' }, [buscar, excel]));
      cont.appendChild(info);
      cont.appendChild(tablaCont);
      pintar();
    });
  }

  /* ---------------- módulo: actividad (supervisor y admin) ---------------- */

  var tempAct = null;
  $('cc-act-filtro').addEventListener('input', function () { clearTimeout(tempAct); tempAct = setTimeout(cargarActividad, 400); });
  $('cc-act-actualizar').addEventListener('click', cargarActividad);

  function cargarActividad() {
    if (!tiene('verActividad')) return;
    var cont = $('cc-act-tabla');
    cont.innerHTML = '';
    cont.appendChild(h('div', { class: 'pa-cargando', 'aria-hidden': 'true' }));
    api('listarActividad', { filtro: $('cc-act-filtro').value, limite: 300 }).then(function (r) {
      cont.innerHTML = '';
      if (!r.ok) { cont.appendChild(h('p', { class: 'pa-msg pa-msg-error', text: r.error })); return; }
      $('cc-act-info').textContent = 'Mostrando ' + miles(r.actividad.length) + ' movimientos más recientes de ' + miles(r.total) + ' registrados.';
      if (!r.actividad.length) { cont.appendChild(h('p', { class: 'pa-vacio', text: 'Sin resultados.' })); return; }
      cont.appendChild(h('table', { class: 'pa-tabla cc-tabla cc-tabla-act' }, [
        h('thead', {}, [h('tr', {}, ['Fecha', 'Usuario', 'Acción', 'Detalle'].map(function (t) { return h('th', { text: t }); }))]),
        h('tbody', {}, r.actividad.map(function (f) {
          return h('tr', {}, [h('td', { 'data-etq': 'Fecha', text: f[0] }), h('td', { 'data-etq': 'Usuario', text: f[1] }),
                              h('td', { 'data-etq': 'Acción', text: f[2] }), h('td', { 'data-etq': 'Detalle', text: f[3] })]);
        }))
      ]));
    });
  }

  /* ---------------- inicio ---------------- */

  document.addEventListener('pa:sesion', function (e) {
    estado.usuario = e.detail;
    estado.cacheAudio = {};
    $('cc-f-oculto').value = 'vis';
    var vista = (location.hash || '#clientes').slice(1);
    irA(vista === 'usuarios' ? 'clientes' : vista);
    cargarClientes();
    cargarFormularios();
  });
})();
