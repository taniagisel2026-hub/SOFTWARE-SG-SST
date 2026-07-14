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

// Deshabilita el boton, le pone una ruedita + texto de "cargando" mientras
// corre fn (async), y lo restaura al terminar (exito o error) para que el
// usuario nunca se quede sin saber si el clic hizo algo.
async function conBotonCargando(boton, textoCargando, fn) {
  const textoOriginal = boton.textContent;
  boton.disabled = true;
  boton.innerHTML = `<span class="spinner"></span>${textoCargando}`;
  try {
    return await fn();
  } finally {
    boton.disabled = false;
    boton.textContent = textoOriginal;
  }
}

function escaparHtml(valor) {
  return String(valor).replace(/&/g, "&amp;").replace(/</g, "&lt;");
}

// Arma las filas "campo: valor" para el resumen de confirmacion de cada
// modulo. pares es un array de [etiqueta, valor]; los valores undefined se
// omiten (campos que no aplican, ej. sitio web vacio).
function resumenFilas(pares) {
  return pares
    .filter(([, valor]) => valor !== undefined)
    .map(([label, valor]) => {
      let texto;
      if (valor === true) texto = "Sí";
      else if (valor === false) texto = "No";
      else if (valor === "" || valor === null) texto = "<em>(sin dato)</em>";
      else texto = escaparHtml(valor);
      return `<div class="resumen-fila"><span class="resumen-label">${label}</span><span class="resumen-valor">${texto}</span></div>`;
    }).join("");
}

// Pantalla de "revisa y confirma" antes de avanzar al siguiente modulo.
// contenidoHtml ya viene armado (resumenFilas() o HTML propio para listas
// como sedes/cargos). onConfirmar hace el POST real + navega; si truena,
// el error se muestra ahi mismo y se puede reintentar sin perder el resumen.
function mostrarResumenConfirmacion(contenidoHtml, onConfirmar) {
  const card = document.querySelector(".card");
  const ocultos = Array.from(card.children).filter((el) => el.tagName !== "SCRIPT");
  ocultos.forEach((el) => {
    el.dataset.resumenDisplayPrevio = el.style.display;
    el.style.display = "none";
  });

  const resumen = document.createElement("div");
  resumen.className = "resumen-confirmacion";
  resumen.innerHTML = `
    <h2>✓ Revisa la información antes de continuar</h2>
    <p class="ayuda">¿Estás de acuerdo con esta información?</p>
    <div class="resumen-lista">${contenidoHtml}</div>
    <div class="resumen-botones">
      <button type="button" class="secundario" id="resumen-editar">Editar</button>
      <button type="button" id="resumen-confirmar">Confirmar y continuar</button>
    </div>
    <p id="resumen-error" class="error"></p>
  `;
  card.appendChild(resumen);

  resumen.querySelector("#resumen-editar").onclick = () => {
    resumen.remove();
    ocultos.forEach((el) => { el.style.display = el.dataset.resumenDisplayPrevio || ""; });
  };
  resumen.querySelector("#resumen-confirmar").onclick = async () => {
    const btn = resumen.querySelector("#resumen-confirmar");
    btn.disabled = true;
    btn.textContent = "Guardando...";
    try {
      await onConfirmar();
    } catch (e) {
      resumen.querySelector("#resumen-error").textContent = e.message;
      btn.disabled = false;
      btn.textContent = "Confirmar y continuar";
    }
  };
}
