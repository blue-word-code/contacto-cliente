/*
 * Contacto-Cliente — módulos de la app (después de iniciar sesión)
 *   Clientes · Ficha del cliente con nota de voz · Grabar/Subir audio
 *   Resumen de cobertura · Cargar (Excel, audios, sincronizar Drive) · Actividad
 * Depende de assets/auth.js (window.PA y el evento "pa:sesion").
 */
(function () {
  'use strict';

  var CAMPOS = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono', 'audio', 'audioFecha', 'audioPor'];
  var ETIQUETAS = { codigo: 'Código', nombre: 'Nombre', direccion: 'Dirección', sector: 'Sector / Zona', ruta: 'Ruta',
                    vendedor: 'Vendedor / Chofer', telefono: 'Teléfono' };
  var EDITABLES = ['codigo', 'nombre', 'direccion', 'sector', 'ruta', 'vendedor', 'telefono'];
  var NIVEL = { usuario: 0, editor: 1, admin: 2 };
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
    cacheAudio: {}      // codigo -> {url, mime}
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

  function normalizar(s) {
    return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
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
    c._txt = normalizar([c.codigo, c.nombre, c.direccion, c.sector, c.ruta, c.vendedor].join(' ')) + ' ' +
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
     ['cc-f-sector', 'sector', 'Todos los sectores']].forEach(function (f) {
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
  ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector', 'cc-f-audio'].forEach(function (id) {
    $(id).addEventListener('change', function () { filtrar(); });
  });
  $('cc-limpiar').addEventListener('click', function () {
    $('cc-buscar').value = '';
    ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector', 'cc-f-audio'].forEach(function (id) { $(id).value = ''; });
    filtrar();
    $('cc-buscar').focus();
  });
  $('cc-mas').addEventListener('click', function () { estado.mostrados += POR_PAGINA; pintarLista(); });

  function filtrar(mantenerPagina) {
    var q = normalizar($('cc-buscar').value);
    var terminos = q ? q.split(' ') : [];
    var ruta = $('cc-f-ruta').value, vend = $('cc-f-vendedor').value, sector = $('cc-f-sector').value, aud = $('cc-f-audio').value;
    var qCodigo = q.toUpperCase();
    estado.filtrados = estado.clientes.filter(function (c) {
      if (ruta && c.ruta !== ruta) return false;
      if (vend && c.vendedor !== vend) return false;
      if (sector && c.sector !== sector) return false;
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
    $('cc-conteo').textContent = estado.clientes.length
      ? miles(total) + (total === 1 ? ' cliente' : ' clientes') + ' · ' + miles(conAudio) + ' con nota de voz'
      : 'Todavía no hay clientes cargados.' + (puede('editor') ? ' Usa la sección Cargar para importarlos.' : '');
    lista.innerHTML = '';
    var frag = document.createDocumentFragment();
    estado.filtrados.slice(0, estado.mostrados).forEach(function (c) {
      var info = [c.sector, c.ruta ? 'Ruta ' + c.ruta : '', c.vendedor].filter(Boolean).join(' · ');
      var li = h('li', {}, [
        h('button', { type: 'button', class: 'cc-cliente', onclick: function () { abrirFicha(c.codigo); } }, [
          h('span', { class: 'cc-estado-audio ' + (c.audio ? 'si' : 'no'), title: c.audio ? 'Tiene nota de voz' : 'Sin nota de voz' },
            [icono(c.audio ? 'play' : 'mic')]),
          h('span', { class: 'cc-cliente-txt' }, [
            h('strong', { text: c.nombre || '(sin nombre)' }),
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
    var m = modal(c.nombre || c.codigo, cuerpo, { sub: 'Código ' + c.codigo });
    pintarFicha(c, cuerpo, m);
  }

  function pintarFicha(c, cuerpo, m) {
    cuerpo.innerHTML = '';
    m.titulo.textContent = c.nombre || c.codigo;
    var msg = h('p', { class: 'pa-msg', 'aria-live': 'polite' });

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
      if (puede('editor')) {
        var acciones = h('div', { class: 'pa-fila' }, [
          h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', onclick: function () { grabar(c, function (nuevo) { c = nuevo; pintarFicha(c, cuerpo, m); }); } },
            [icono('mic'), c.audio ? 'Grabar de nuevo' : 'Grabar nota de voz']),
          h('label', { class: 'pa-btn pa-btn-sec pa-btn-sm cc-label-archivo' }, [icono('subir'), 'Subir archivo',
            h('input', { type: 'file', accept: 'audio/*,.opus,.ogg,.amr,.m4a', onchange: function (e) {
              var f = e.target.files[0]; e.target.value = '';
              if (f) subirUno(c, f, false, msg).then(function (nuevo) { if (nuevo) { c = nuevo; pintarFicha(c, cuerpo, m); } });
            } })]),
          c.audio ? h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar nota de voz', onclick: function (e) {
            if (!confirm('¿Eliminar la nota de voz de ' + c.nombre + '? Quedará 30 días en la papelera de Drive.')) return;
            var b = e.currentTarget; ocupado(b, true, 'Eliminando…');
            api('eliminarAudio', { codigo: c.codigo }).then(function (res) {
              ocupado(b, false);
              if (!res.ok) { mensaje(msg, res.error); return; }
              c = actualizarLocal(res.cliente); aviso(res.mensaje); pintarFicha(c, cuerpo, m);
            });
          } }) : null
        ]);
        zonaAudio.appendChild(acciones);
      }
    }
    pintarAudio();

    // Datos
    var datos = h('dl', { class: 'cc-datos' });
    ['direccion', 'sector', 'ruta', 'vendedor', 'telefono'].forEach(function (k) {
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

    // Administración
    var admin = null;
    if (puede('editor')) {
      admin = h('div', { class: 'pa-fila cc-ficha-admin' }, [
        h('button', { type: 'button', class: 'pa-btn pa-btn-sec pa-btn-sm', text: 'Editar datos', onclick: function () {
          formularioCliente(c, function (nuevo) { c = nuevo; pintarFicha(c, cuerpo, m); });
        } }),
        puede('admin') ? h('button', { type: 'button', class: 'pa-btn pa-btn-peligro pa-btn-sm', text: 'Eliminar cliente', onclick: function (e) {
          if (!confirm('¿Eliminar el cliente ' + c.codigo + ' — ' + c.nombre + '?' + (c.audio ? ' Su nota de voz irá a la papelera de Drive.' : '') + ' Esta acción no se puede deshacer.')) return;
          var b = e.currentTarget; ocupado(b, true, 'Eliminando…');
          api('eliminarCliente', { codigo: c.codigo }).then(function (res) {
            ocupado(b, false);
            if (!res.ok) { mensaje(msg, res.error); return; }
            quitarLocal(c.codigo); aviso(res.mensaje); m.cerrar();
          });
        } }) : null
      ]);
    }

    cuerpo.appendChild(zonaAudio);
    cuerpo.appendChild(msg);
    if (contacto.childNodes.length) cuerpo.appendChild(contacto);
    cuerpo.appendChild(datos);
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
      if (!res.ok && res.codigo === 'YA_TIENE') {
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
    var ayuda = h('p', { class: 'cc-suave cc-centro', text: 'Pulsa el micrófono y describe cómo llegar al cliente. Máximo 3 minutos.' });
    var vista = h('div', { class: 'cc-vista-previa' });
    var guardar = h('button', { type: 'button', class: 'pa-btn', text: 'Guardar nota de voz', disabled: true });
    var repetir = h('button', { type: 'button', class: 'pa-btn pa-btn-sec', text: 'Grabar de nuevo', hidden: true });
    var cuerpo = h('div', { class: 'cc-grabador' }, [h('div', { class: 'cc-centro' }, [principal, h('div', { class: 'cc-estado-grab' }, [punto, tiempo])]),
      ayuda, vista, msg, h('div', { class: 'pa-fila pa-fin' }, [repetir, guardar])]);
    var m = modal('Grabar nota de voz', cuerpo, { sub: c.codigo + ' — ' + c.nombre, alCerrar: detenerTodo });

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
        autocapitalize: k === 'codigo' ? 'characters' : null, list: ['sector', 'ruta', 'vendedor'].indexOf(k) >= 0 ? 'cc-dl-' + k : null });
      input.value = c[k] || '';
      entradas[k] = input;
      return h('label', { class: 'pa-campo' }, [h('span', { class: 'pa-etq', text: ETIQUETAS[k] + (k === 'codigo' || k === 'nombre' ? ' *' : '') }), input]);
    });
    var listas = ['sector', 'ruta', 'vendedor'].map(function (k) {
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
      var filas = [['Código', 'Nombre', 'Dirección', 'Sector', 'Ruta', 'Vendedor/Chofer', 'Teléfono', 'Nota de voz', 'Fecha nota de voz']];
      estado.filtrados.forEach(function (c) {
        filas.push([c.codigo, c.nombre, c.direccion, c.sector, c.ruta, c.vendedor, c.telefono, c.audio ? 'Sí' : 'No', c.audioFecha || '']);
      });
      var ws = XLSX.utils.aoa_to_sheet(filas);
      ws['!cols'] = [12, 32, 40, 18, 10, 18, 14, 11, 16].map(function (w) { return { wch: w }; });
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
    var total = estado.clientes.length;
    var con = estado.clientes.reduce(function (n, c) { return n + (c.audio ? 1 : 0); }, 0);
    var kpis = $('cc-kpis');
    kpis.innerHTML = '';
    [['Clientes', miles(total), ''], ['Con nota de voz', miles(con), 'ok'], ['Sin nota de voz', miles(total - con), 'pend'],
     ['Cobertura', pct(con, total) + '%', 'cob']].forEach(function (k) {
      kpis.appendChild(h('div', { class: 'cc-kpi ' + k[2] }, [h('span', { text: k[0] }), h('strong', { text: k[1] })]));
    });
    var grupos = {};
    estado.clientes.forEach(function (c) {
      var g = c[estado.grupo] || '(sin ' + (estado.grupo === 'vendedor' ? 'vendedor' : estado.grupo) + ')';
      grupos[g] = grupos[g] || { total: 0, con: 0 };
      grupos[g].total++;
      if (c.audio) grupos[g].con++;
    });
    var nombres = Object.keys(grupos).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); });
    var cont = $('cc-cobertura');
    cont.innerHTML = '';
    if (!nombres.length) { cont.appendChild(h('p', { class: 'pa-vacio', text: 'Todavía no hay clientes.' })); return; }
    var etq = { ruta: 'Ruta', vendedor: 'Vendedor / Chofer', sector: 'Sector' }[estado.grupo];
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
        ['cc-f-ruta', 'cc-f-vendedor', 'cc-f-sector'].forEach(function (id) { $(id).value = ''; });
        var sel = $('cc-f-' + estado.grupo);
        if (grupos[n] && sel.querySelector('option[value="' + CSS.escape(n) + '"]')) sel.value = n;
        $('cc-f-audio').value = '0';
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
    var subir = h('button', { type: 'button', class: 'pa-btn', text: 'Subir ' + filas.length + ' notas de voz' });
    var barra = h('div', { class: 'cc-progreso', hidden: true }, [h('span')]);
    var res = h('p', { class: 'pa-msg' });
    zona.appendChild(h('div', { class: 'cc-panel' }, [
      h('p', { class: 'pa-msg pa-msg-info', text: miles(detectados) + ' de ' + miles(filas.length) + ' archivos se asociaron solos a un cliente. Revisa o escribe el código de los demás.' }),
      h('div', { class: 'cc-tabla-cont' }, [h('table', { class: 'pa-tabla cc-tabla' }, [
        h('thead', {}, [h('tr', {}, [h('th', { text: 'Archivo' }), h('th', { class: 'num', text: 'Tamaño' }), h('th', { text: 'Cliente' }), h('th', { text: 'Estado' })])]),
        h('tbody', {}, filas.map(function (f) { return f.tr; }))])]),
      h('label', { class: 'cc-check' }, [reemplazar, ' Reemplazar la nota de voz si el cliente ya tiene una']),
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
          else if (r.codigo === 'YA_TIENE') { f.estado.textContent = 'Ya tenía nota (no se reemplazó)'; f.estado.className = 'cc-estado-celda cc-suave'; omitidos++; }
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
    var reemplazar = $('cc-sinc-reemplazar').checked;
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
      pintarSueltos(zona, tot.conAudio, 'Clientes que ya tenían nota de voz', 'Estos audios llevan el código de un cliente que ya tiene nota. Pulsa Reemplazar si el nuevo es el correcto.', true);
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

  /* ---------------- módulo: actividad (admin) ---------------- */

  var tempAct = null;
  $('cc-act-filtro').addEventListener('input', function () { clearTimeout(tempAct); tempAct = setTimeout(cargarActividad, 400); });
  $('cc-act-actualizar').addEventListener('click', cargarActividad);

  function cargarActividad() {
    if (!puede('admin')) return;
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
    var vista = (location.hash || '#clientes').slice(1);
    irA(vista === 'usuarios' ? 'clientes' : vista);
    cargarClientes();
  });
})();
