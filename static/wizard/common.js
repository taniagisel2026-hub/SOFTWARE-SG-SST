// Compartido por todas las paginas del asistente. Pide el API key una sola
// vez (prompt del navegador, se guarda en localStorage) -- NO es un sistema
// de login real, es el mismo API key estatico de la Fase 4.

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
