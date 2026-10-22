// Shows the task of a Teams notification. Teams passes the notification's signed link as context.page.subPageId;
// outside Teams (testing in a browser) it is read from ?t=. A notification of consegsa-notificaciones-api carries
// "m:<base64url of {s: tema, t: titulo, m: mensaje, u: url}>" (shown as is, only as text) or, in its older format,
// "u:<url>" (only an "Abrir" button).
(async function () {
  const contenido = document.getElementById('contenido');

  const tema = (t) => document.documentElement.setAttribute('data-tema', t || 'default');
  const mensaje = (texto) => {
    contenido.replaceChildren(el('p', 'mensaje', texto));
  };

  function el(tag, clase, texto) {
    const nodo = document.createElement(tag);
    if (clase) nodo.className = clase;
    if (texto !== undefined && texto !== null) nodo.textContent = texto;
    return nodo;
  }

  function fecha(valor) {
    if (!valor) return null;
    const d = new Date(valor);
    if (isNaN(d)) return valor;
    return d.toLocaleString('es', { dateStyle: 'medium', timeStyle: 'short' });
  }

  function texto(valor) {
    return valor ? String(valor).replaceAll('_', ' ').toLowerCase().replace(/^./, (c) => c.toUpperCase()) : null;
  }

  // Teams: initialize (removes the loading indicator) and read the context. Fails fast outside Teams
  async function contextoTeams() {
    if (!window.microsoftTeams) return null;
    try {
      await Promise.race([
        microsoftTeams.app.initialize(),
        new Promise((_, rechazar) => setTimeout(() => rechazar(new Error('fuera de Teams')), 3000)),
      ]);
      const contexto = await microsoftTeams.app.getContext();
      microsoftTeams.app.notifySuccess();
      microsoftTeams.app.registerOnThemeChangeHandler(tema);
      return contexto;
    } catch {
      return null;
    }
  }

  function mostrar(tarea) {
    const tarjeta = el('article', 'tarjeta');
    tarjeta.append(el('p', 'numero', `Tarea #${tarea.id}${tarea.tipo_tarea_nombre ? ' · ' + tarea.tipo_tarea_nombre : ''}`));
    tarjeta.append(el('h1', null, tarea.titulo || 'Sin título'));

    const chips = el('div', 'chips');
    [texto(tarea.estado), tarea.urgencia ? 'Urgencia ' + texto(tarea.urgencia).toLowerCase() : null]
      .filter(Boolean)
      .forEach((c) => chips.append(el('span', 'chip', c)));
    tarjeta.append(chips);

    const responsables = (tarea.responsables || []).map((r) => r.usuario_nombre).filter(Boolean).join(', ');
    const filas = [
      ['Creada por', tarea.creador_nombre],
      ['Registrada', fecha(tarea.fecha_registro)],
      ['Asignada a', tarea.asignado_a_nombre || 'Sin asignar'],
      ['Responsables', responsables],
      ['Vence (SLA)', fecha(tarea.fecha_vencimiento_sla)],
      ['Autorización', texto(tarea.autorizacion_estado)],
      ['Finalizada por', tarea.finalizada_por_nombre],
      ['Comentario de cierre', tarea.comentario_cierre],
    ];
    const dl = el('dl');
    filas.filter(([, v]) => v).forEach(([k, v]) => dl.append(el('dt', null, k), el('dd', null, v)));
    tarjeta.append(dl);

    if (tarea.descripcion) {
      tarjeta.append(el('h2', null, 'Descripción'), el('p', 'descripcion', tarea.descripcion));
    }
    contenido.replaceChildren(tarjeta);
    document.title = `Tarea #${tarea.id} · Consegsa`;
  }

  // Only https links to consegsa.com (and subdomains), the same rule the notifications service applies
  function urlPermitida(valor) {
    try {
      const url = new URL(valor);
      const host = url.hostname.toLowerCase();
      return url.protocol === 'https:' && !url.username && (host === 'consegsa.com' || host.endsWith('.consegsa.com')) ? url.href : null;
    } catch {
      return null;
    }
  }

  // base64url → UTF-8 JSON; null when it is not a valid notification
  function leerContenido(valor) {
    try {
      const base64 = valor.replaceAll('-', '+').replaceAll('_', '/');
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      const datos = JSON.parse(new TextDecoder().decode(bytes));
      return datos && typeof datos.t === 'string' ? datos : null;
    } catch {
      return null;
    }
  }

  function mostrarNotificacion(datos) {
    const tarjeta = el('article', 'tarjeta');
    tarjeta.append(el('p', 'numero', typeof datos.s === 'string' && datos.s ? datos.s : 'Notificación de Consegsa'));
    tarjeta.append(el('h1', null, datos.t));
    if (typeof datos.m === 'string' && datos.m) {
      tarjeta.append(el('p', 'descripcion', datos.m));
    }
    const url = typeof datos.u === 'string' ? urlPermitida(datos.u) : null;
    if (url) {
      const boton = el('a', 'boton', 'Abrir');
      boton.href = url;
      boton.target = '_blank';
      boton.rel = 'noopener noreferrer';
      tarjeta.append(boton);
    }
    contenido.replaceChildren(tarjeta);
    document.title = `${datos.t} · Consegsa`;
  }

  function mostrarEnlace(url) {
    const tarjeta = el('article', 'tarjeta');
    tarjeta.append(el('p', 'numero', 'Notificación de Consegsa'));
    tarjeta.append(el('h1', null, 'Esta notificación tiene un enlace'));
    tarjeta.append(el('p', 'descripcion', new URL(url).hostname));
    const boton = el('a', 'boton', 'Abrir');
    boton.href = url;
    boton.target = '_blank';
    boton.rel = 'noopener noreferrer';
    tarjeta.append(boton);
    contenido.replaceChildren(tarjeta);
  }

  const contexto = await contextoTeams();
  tema(contexto?.app?.theme);
  const enlace = contexto?.page?.subPageId || new URLSearchParams(location.search).get('t');

  if (!enlace) {
    mensaje('Abre una notificación de Consegsa para ver su detalle.');
    return;
  }

  if (enlace.startsWith('m:')) {
    const datos = leerContenido(enlace.slice(2));
    datos ? mostrarNotificacion(datos) : mensaje('El contenido de esta notificación no es válido.');
    return;
  }

  if (enlace.startsWith('u:')) {
    const url = urlPermitida(enlace.slice(2));
    url ? mostrarEnlace(url) : mensaje('El enlace de esta notificación no es válido.');
    return;
  }

  try {
    const r = await fetch('api/tarea?t=' + encodeURIComponent(enlace));
    const cuerpo = await r.json();
    if (cuerpo.code === 200 && cuerpo.data?.length) {
      mostrar(cuerpo.data[0]);
    } else if (cuerpo.code === 401) {
      mensaje('El enlace de esta notificación no es válido o ya venció.');
    } else if (cuerpo.code === 404) {
      mensaje('La tarea no existe o ya no tienes acceso a ella.');
    } else {
      mensaje(cuerpo.server_message || 'No se pudo cargar la tarea.');
    }
  } catch {
    mensaje('No se pudo cargar la tarea. Intenta de nuevo más tarde.');
  }
})();
