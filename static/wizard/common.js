// Compartido por todas las paginas del asistente. Pide el API key una sola
// vez (prompt del navegador, se guarda en localStorage) -- NO es un sistema
// de login real, es el mismo API key estatico de la Fase 4.

const MODULOS_NOMBRES = [
  "Documentos legales",
  "Riesgo y ARL",
  "Sedes",
  "Responsable SST",
  "Fechas y opciones",
  "Descripción de la empresa",
  "Cargos",
];

// Envuelve todo el contenido de la pagina en un shell de aplicacion: barra
// lateral fija (marca + los 7 pasos) + panel principal con el contenido
// original. Se corre una sola vez aqui en vez de repetir el HTML en los
// 9 archivos del asistente.
(function inicializarAppShell() {
  // index.html es una landing page aparte (hero + CTA), no un paso del
  // asistente -- no se le aplica el shell con la barra lateral.
  if (/\/portal\/?(index\.html)?$/.test(window.location.pathname)) return;

  // Paso actual: se lee de <p class="progreso">Modulo X de N -- Nombre</p>
  // si existe; finalizar.html no tiene ninguno (los 7 ya quedaron
  // confirmados, se muestran todos como hechos).
  let actual = 0;
  const esFinalizar = /finalizar\.html/.test(window.location.pathname);
  if (esFinalizar) {
    actual = MODULOS_NOMBRES.length + 1;
  } else {
    const progresoEl = document.querySelector(".progreso");
    if (progresoEl) {
      const m = /(\d+)\s*de\s*(\d+)/.exec(progresoEl.textContent.trim());
      if (m) actual = parseInt(m[1], 10);
    }
  }

  // En este punto del parseo solo existen los elementos ANTES de este mismo
  // <script src="common.js">: el <div class="card">...</div> de la pagina.
  // El <script> inline de cada pagina (el que sigue despues) todavia no se
  // ha parseado -- no tocarlo aqui, ni recrear/mover ningun <script> (eso
  // haria que el navegador lo re-ejecute, y en el caso de este mismo
  // archivo se volveria a correr esta funcion infinitamente).
  const contenidoOriginal = Array.from(document.body.children).filter(
    (el) => el.tagName !== "SCRIPT"
  );

  const sidebar = document.createElement("aside");
  sidebar.className = "sidebar";
  const nav = MODULOS_NOMBRES.map((nombre, i) => {
    const n = i + 1;
    const estado = n < actual ? "done" : n === actual ? "active" : "";
    const marca = n < actual ? "✓" : String(n);
    return `<div class="sidebar-item ${estado}"><span class="dot">${marca}</span><span class="label">${nombre}</span></div>`;
  }).join("");
  sidebar.innerHTML = `
    <div class="sidebar-brand">
      <span class="logo-chip"><img src="logo-mark.png" alt="SG-SST Empresarial"></span>
      <div class="text"><b>SG-SST EMPRESARIAL</b><span>Portal de autoservicio</span></div>
    </div>
    <nav class="sidebar-nav">${nav}</nav>
  `;

  const main = document.createElement("main");
  main.className = "main";
  const mainInner = document.createElement("div");
  mainInner.className = "main-inner";
  for (const el of contenidoOriginal) mainInner.appendChild(el);
  main.appendChild(mainInner);

  const app = document.createElement("div");
  app.className = "app";
  app.appendChild(sidebar);
  app.appendChild(main);

  // prepend (no innerHTML="") -- no molesta al parser, que sigue agregando
  // los <script> que faltan por leer al final de <body> con normalidad.
  document.body.prepend(app);

  // El "Modulo X de 7 -- Nombre" ya lo muestra la barra lateral, no hace
  // falta repetirlo dentro de la tarjeta.
  const progresoEl = mainInner.querySelector(".progreso");
  if (progresoEl) progresoEl.remove();
})();

function getToken() {
  let t = localStorage.getItem("wizard_token");
  if (!t) {
    t = prompt("Ingresa el API key de la API (Authorization Bearer token):");
    if (t) localStorage.setItem("wizard_token", t.trim());
  }
  return t ? t.trim() : "";
}

async function api(path, options = {}) {
  const token = getToken();
  const opts = { ...options };
  opts.headers = { ...(options.headers || {}), Authorization: `Bearer ${token}` };
  if (opts.body && !(opts.body instanceof FormData)) {
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(opts.body);
  }
  const res = await fetch(`/wizard${path}`, opts);
  if (res.status === 401) {
    localStorage.removeItem("wizard_token");
    throw new Error("API key incorrecto. Recarga la pagina para intentar de nuevo.");
  }
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || `Error ${res.status}`);
  }
  return res.json();
}

function getBorradorId() {
  const id = new URLSearchParams(window.location.search).get("borrador_id");
  if (!id) {
    alert("Falta el borrador_id en la URL. Vuelve a empezar desde index.html.");
  }
  return id;
}

function irAModulo(n, borradorId) {
  window.location.href = `modulo${n}.html?borrador_id=${borradorId}`;
}

function irAFinalizar(borradorId) {
  window.location.href = `finalizar.html?borrador_id=${borradorId}`;
}

function llenarProgreso(moduloActualTexto) {
  const el = document.getElementById("progreso");
  if (el) el.textContent = moduloActualTexto;
}

function rellenarFormulario(form, datos) {
  for (const el of form.elements) {
    if (!el.name || datos[el.name] === undefined) continue;
    if (el.type === "checkbox") el.checked = !!datos[el.name];
    else el.value = datos[el.name];
  }
}
