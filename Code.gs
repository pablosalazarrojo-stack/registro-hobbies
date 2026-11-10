/**
 * Registro de Hobbies — el "cerebro" que vive en la hoja de Google.
 * Guarda los registros, las parejas y si ya se liberaron.
 * El usuario y la contraseña del administrador viven aquí (Propiedades del script), no en la página.
 *
 * Instalación (una sola vez):
 *  1. Hoja de Google nueva → Extensiones → Apps Script → pegar este archivo.
 *  2. Ejecutar la función `setup` (pide permisos de tu cuenta).
 *  3. Implementar → Nueva implementación → Aplicación web
 *     · Ejecutar como: yo   · Quién tiene acceso: cualquier persona
 *  4. Copiar la URL que termina en /exec y pegarla en CONFIG.SHEET_URL de index.html.
 */

const SHEET_ID = "1tyGYel58AcGE2HjyEo-Obr94sECDk6VeBV-ReuubwWc";
function libro() { return SpreadsheetApp.openById(SHEET_ID); }
const HOJA_REG = "Registros";
const HOJA_PAR = "Parejas";
const COLS = ["id", "fecha", "nombre", "nombreCompleto", "whatsapp", "hobby1", "hobby2", "hobby3", "carrera"];

function setup() {
  const ss = libro();
  let r = ss.getSheetByName(HOJA_REG) || ss.insertSheet(HOJA_REG);
  r.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  r.getRange("E:E").setNumberFormat("@"); // WhatsApp como texto
  r.setFrozenRows(1);
  const s1 = ss.getSheetByName("Sheet1") || ss.getSheetByName("Hoja 1");
  let p = ss.getSheetByName(HOJA_PAR) || ss.insertSheet(HOJA_PAR);
  p.getRange(1, 1, 1, 5).setValues([["grupo", "id", "nombreCompleto", "carrera", "enComun"]]).setFontWeight("bold");
  p.setFrozenRows(1);
  if (s1 && ss.getSheets().length > 2) ss.deleteSheet(s1);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("ADMIN_USER")) props.setProperty("ADMIN_USER", "pablo");
  if (!props.getProperty("ADMIN_PASS")) props.setProperty("ADMIN_PASS", "1234");
  if (!props.getProperty("LIBERADO")) props.setProperty("LIBERADO", "false");
}

function doGet() { return salida({ ok: true, msg: "Registro de Hobbies activo" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "register": out = registrar(d.registro); break;
      case "status": out = estado(d); break;
      case "adminList": admin(d); out = { ok: true, registros: leerRegistros(), grupos: leerGrupos(), liberado: liberado() }; break;
      case "savePairs": admin(d); guardarGrupos(d.grupos); out = { ok: true }; break;
      case "release":
        admin(d);
        if (d.liberar && !leerGrupos().length) throw new Error("Primero forma las parejas.");
        PropertiesService.getScriptProperties().setProperty("LIBERADO", d.liberar ? "true" : "false");
        out = { ok: true }; break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function admin(d) {
  const p = PropertiesService.getScriptProperties();
  if (String(d.user || "") !== p.getProperty("ADMIN_USER") || String(d.pass || "") !== p.getProperty("ADMIN_PASS")) {
    Utilities.sleep(1500); // frena a quien intente adivinar
    throw new Error("Usuario o contraseña incorrectos.");
  }
}

function liberado() { return PropertiesService.getScriptProperties().getProperty("LIBERADO") === "true"; }

function hoja(n) { return libro().getSheetByName(n); }

function leerRegistros() {
  const v = hoja(HOJA_REG).getDataRange().getValues();
  return v.slice(1).filter((r) => r[0]).map((r) => {
    const o = {}; COLS.forEach((c, i) => (o[c] = String(r[i])));
    return o;
  });
}

function limpio(s) { return String(s || "").trim().replace(/\s+/g, " ").slice(0, 80); }

function registrar(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const tel = String(r.whatsapp || "").replace(/\D/g, "");
    if (!/^\d{10}$/.test(tel)) throw new Error("WhatsApp inválido.");
    if (["LEM", "LDE", "BGB", "LAF", "CPF", "LAE"].indexOf(r.carrera) < 0) throw new Error("Carrera inválida.");
    const sh = hoja(HOJA_REG);
    const regs = leerRegistros();
    const previo = regs.findIndex((x) => x.whatsapp === tel);
    const id = previo >= 0 ? regs[previo].id : Utilities.getUuid().slice(0, 8);
    const fila = [id, new Date().toISOString(), limpio(r.nombre), limpio(r.nombreCompleto), tel, limpio(r.hobby1), limpio(r.hobby2), limpio(r.hobby3), r.carrera];
    if (previo >= 0) sh.getRange(previo + 2, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function leerGrupos() {
  const v = hoja(HOJA_PAR).getDataRange().getValues().slice(1);
  const m = {};
  v.forEach((r) => { if (r[1]) (m[r[0]] = m[r[0]] || []).push(String(r[1])); });
  return Object.keys(m).sort((a, b) => a - b).map((k) => m[k]);
}

function guardarGrupos(grupos) {
  const porId = {}; leerRegistros().forEach((r) => (porId[r.id] = r));
  const filas = [];
  grupos.forEach((g, i) => g.forEach((id) => {
    const r = porId[id]; if (!r) return;
    const otros = g.filter((x) => x !== id).map((x) => porId[x]).filter(Boolean);
    const com = [].concat.apply([], otros.map((o) => comunes(r, o)));
    filas.push([i + 1, id, r.nombreCompleto, r.carrera, com.join(", ")]);
  }));
  const sh = hoja(HOJA_PAR);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (filas.length) sh.getRange(2, 1, filas.length, 5).setValues(filas);
  PropertiesService.getScriptProperties().setProperty("LIBERADO", "false");
}

function norm(s) { return limpio(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
function hobbiesDe(r) { return [r.hobby1, r.hobby2, r.hobby3].filter(String); }
function comunes(a, b) {
  const deB = {}; hobbiesDe(b).forEach((h) => (deB[norm(h)] = true));
  const vistos = {}, out = [];
  hobbiesDe(a).forEach((h) => { const k = norm(h); if (deB[k] && !vistos[k]) { vistos[k] = true; out.push(h); } });
  return out;
}

function estado(d) {
  const regs = leerRegistros();
  const tel = String(d.whatsapp || "").replace(/\D/g, "");
  const yo = regs.find((x) => (d.id && x.id === d.id) || (tel && x.whatsapp === tel));
  if (!yo) return { ok: true, registrado: false };
  const lib = liberado();
  let pareja = [];
  if (lib) {
    const g = leerGrupos().find((g) => g.indexOf(yo.id) >= 0);
    if (g) pareja = g.filter((id) => id !== yo.id).map((id) => regs.find((r) => r.id === id)).filter(Boolean)
      .map((o) => ({ nombreCompleto: o.nombreCompleto, carrera: o.carrera, whatsapp: o.whatsapp, enComun: comunes(yo, o) }));
  }
  return { ok: true, registrado: true, id: yo.id, yo: { nombre: yo.nombre, nombreCompleto: yo.nombreCompleto, whatsapp: yo.whatsapp, hobby1: yo.hobby1, hobby2: yo.hobby2, hobby3: yo.hobby3, carrera: yo.carrera }, liberado: lib, pareja: pareja };
}
