'use strict';
/* =====================================================================
   Hucha — control de gastos personal.
   Todo se guarda SOLO en este dispositivo (localStorage). Sin servidor.
   ===================================================================== */

const STORE_KEY = 'hucha.v1';
const TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

/* ---------- Categorías ---------- */
const CATS = [
  { id: 'super', name: 'Supermercado', short: 'Súper', emoji: '🛒', type: 'need', color: '#22a06b' },
  { id: 'casa', name: 'Casa y facturas', short: 'Casa', emoji: '🏠', type: 'need', color: '#3b82f6' },
  { id: 'transporte', name: 'Transporte', short: 'Transporte', emoji: '🚌', type: 'need', color: '#0891b2' },
  { id: 'salud', name: 'Salud', short: 'Salud', emoji: '💊', type: 'need', color: '#0d9488' },
  { id: 'comer', name: 'Comer fuera', short: 'Comer fuera', emoji: '🍽️', type: 'want', color: '#f97316' },
  { id: 'delivery', name: 'Comida a domicilio', short: 'Domicilio', emoji: '🛵', type: 'want', color: '#e11d48' },
  { id: 'cafe', name: 'Cafés y snacks', short: 'Cafés', emoji: '☕', type: 'want', color: '#92400e' },
  { id: 'ocio', name: 'Ocio y planes', short: 'Ocio', emoji: '🎉', type: 'want', color: '#9333ea' },
  { id: 'compras', name: 'Compras', short: 'Compras', emoji: '🛍️', type: 'want', color: '#db2777' },
  { id: 'subs', name: 'Suscripciones', short: 'Suscrip.', emoji: '📺', type: 'want', color: '#4f46e5' },
  { id: 'viajes', name: 'Viajes', short: 'Viajes', emoji: '✈️', type: 'want', color: '#ca8a04' },
  { id: 'otros', name: 'Otros', short: 'Otros', emoji: '📦', type: 'want', color: '#64748b' },
];
const CAT = Object.fromEntries(CATS.map(c => [c.id, c]));
const catOf = id => CAT[id] || CAT.otros;

/* Palabras clave → categoría (comercios habituales en España). El orden importa. */
const RULES = [
  ['delivery', ['glovo', 'just eat', 'justeat', 'uber eats', 'ubereats', 'deliveroo', 'telepizza', 'dominos', 'foodora']],
  ['subs', ['netflix', 'spotify', 'hbo', 'max', 'disney', 'prime video', 'amazon prime', 'apple com bill', 'icloud', 'youtube', 'chatgpt', 'openai', 'claude', 'anthropic', 'gimnasio', 'gym', 'basic fit', 'basicfit', 'dazn', 'movistar plus', 'filmin', 'audible', 'crunchyroll', 'patreon', 'duolingo', 'google one', 'dropbox', 'adobe', 'canva', 'notion', 'twitch', 'suscripcion']],
  ['super', ['mercadona', 'carrefour', 'lidl', 'aldi', 'dia', 'alcampo', 'eroski', 'consum', 'hipercor', 'ahorramas', 'bonpreu', 'caprabo', 'condis', 'froiz', 'gadis', 'spar', 'supercor', 'masymas', 'supermercado', 'super', 'fruteria', 'carniceria', 'panaderia', 'mercado', 'coviran', 'family cash', 'bonarea']],
  ['cafe', ['starbucks', 'cafe', 'cafes', 'cafeteria', 'granier', 'costa coffee', 'tim hortons', 'churreria', 'pasteleria', 'vending', 'heladeria']],
  ['comer', ['restaurante', 'bar', 'burger', 'mcdonalds', 'mc donalds', 'kfc', 'taberna', 'vips', 'foster', 'goiko', 'pizzeria', 'sushi', 'tapas', 'cerveceria', 'five guys', 'montaditos', 'rodilla', 'taco bell', 'popeyes', 'wok', 'kebab', 'tagliatella', 'ginos', 'asador', 'meson']],
  ['transporte', ['renfe', 'metro', 'emt', 'tmb', 'cabify', 'uber', 'bolt', 'freenow', 'free now', 'taxi', 'gasolinera', 'gasolina', 'repsol', 'cepsa', 'galp', 'bp', 'shell', 'parking', 'aparcamiento', 'bicimad', 'bicing', 'alsa', 'peaje', 'autopista', 'blablacar', 'ouigo', 'iryo', 'abono transporte']],
  ['casa', ['alquiler', 'comunidad', 'endesa', 'iberdrola', 'naturgy', 'holaluz', 'totalenergies', 'movistar', 'vodafone', 'orange', 'digi', 'yoigo', 'masmovil', 'pepephone', 'lowi', 'simyo', 'finetwork', 'jazztel', 'agua', 'canal de isabel', 'seguro hogar', 'leroy merlin', 'bauhaus', 'bricomart', 'luz', 'gas natural', 'internet', 'fibra']],
  ['salud', ['farmacia', 'clinica', 'dentista', 'hospital', 'medico', 'optica', 'fisio', 'fisioterapia', 'sanitas', 'adeslas', 'dkv', 'asisa', 'psicologo', 'psicologa', 'parafarmacia']],
  ['ocio', ['cine', 'cines', 'yelmo', 'cinesa', 'kinepolis', 'entradas', 'ticketmaster', 'steam', 'playstation', 'nintendo', 'xbox', 'teatro', 'museo', 'concierto', 'discoteca', 'bolera', 'escape room', 'casa del libro', 'libreria']],
  ['viajes', ['iberia', 'vueling', 'ryanair', 'easyjet', 'air europa', 'booking', 'airbnb', 'hotel', 'hostal', 'expedia', 'edreams', 'volotea', 'binter']],
  ['compras', ['amazon', 'amzn', 'zara', 'primark', 'shein', 'aliexpress', 'temu', 'corte ingles', 'decathlon', 'ikea', 'fnac', 'mediamarkt', 'media markt', 'pccomponentes', 'mango', 'pull bear', 'bershka', 'stradivarius', 'sephora', 'druni', 'primor', 'miniso', 'tiger', 'wallapop', 'vinted', 'h m', 'uniqlo', 'nike', 'adidas']],
];

const BADGES = [
  { id: 'first', emoji: '🌱', name: 'Primer paso', desc: 'Apunta tu primer gasto' },
  { id: 'streak3', emoji: '🔥', name: '3 días', desc: 'Racha de 3 días' },
  { id: 'streak7', emoji: '⚡', name: 'Una semana', desc: 'Racha de 7 días' },
  { id: 'streak30', emoji: '🏆', name: 'Un mes', desc: 'Racha de 30 días' },
  { id: 'nospend1', emoji: '✨', name: 'Día zen', desc: 'Un día sin gastar' },
  { id: 'nospend5', emoji: '🧘', name: '5 días zen', desc: '5 días sin gastar en un mes' },
  { id: 'capture', emoji: '📸', name: 'Fotógrafo', desc: 'Apunta desde una captura' },
  { id: 'goal', emoji: '🎯', name: 'Con rumbo', desc: 'Crea una meta' },
  { id: 'goaldone', emoji: '🥳', name: '¡Conseguido!', desc: 'Completa una meta' },
  { id: 'resist', emoji: '🛡️', name: 'Escudo', desc: 'Resiste una compra' },
  { id: 'resist100', emoji: '💪', name: 'Voluntad', desc: '100 € sin gastar gracias a la lista de espera' },
  { id: 'underbudget', emoji: '👑', name: 'Mes redondo', desc: 'Termina un mes dentro del presupuesto' },
  { id: 'backup', emoji: '💾', name: 'Precavido', desc: 'Haz una copia de seguridad' },
];

const SRC_ICON = { ia: ' · ✨', ocr: ' · 📸', fijo: ' · 🔁 fijo', csv: ' · 🏦', paste: ' · 📋', wish: ' · ⏳' };
const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const MONTHS_S = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const WEEKDAYS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const GOAL_EMOJIS = ['🛟', '🏖️', '✈️', '💻', '🚗', '🏠', '🎓', '🎁', '🎸', '🐶', '💍', '🌱'];

/* =====================================================================
   Utilidades
   ===================================================================== */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const sum = arr => arr.reduce((a, e) => a + (e.amount || 0), 0);
const fmt2 = new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' });
const fmt0 = new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR', minimumFractionDigits: 0, maximumFractionDigits: 0 });
const money = n => fmt2.format(Math.round((n || 0) * 100) / 100);
const moneyR = n => fmt0.format(Math.round(n || 0));
const pad = n => String(n).padStart(2, '0');
const toISO = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromISO = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const todayISO = () => toISO(new Date());
const addDays = (iso, n) => { const d = fromISO(iso); d.setDate(d.getDate() + n); return toISO(d); };
const diffDays = (a, b) => Math.round((fromISO(b) - fromISO(a)) / 86400000);
const clampDay = (y, m, day) => Math.min(day, new Date(y, m + 1, 0).getDate());
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const words = s => norm(s).split(/[^a-z0-9]+/).filter(Boolean);
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const isStandalone = () => navigator.standalone === true || matchMedia('(display-mode: standalone)').matches;
const isTouch = () => matchMedia('(pointer: coarse)').matches;

function niceDate(iso) {
  const t = todayISO();
  if (iso === t) return 'Hoy';
  if (iso === addDays(t, -1)) return 'Ayer';
  const d = fromISO(iso);
  return `${cap(WEEKDAYS[d.getDay()])} ${d.getDate()} ${MONTHS_S[d.getMonth()]}`;
}

/** Convierte "1.234,56", "-12,50 €", "12.50" o "(3,20)" en número. */
function parseAmount(raw) {
  if (raw == null) return NaN;
  if (typeof raw === 'number') return raw;
  let s = String(raw).replace(/eur/gi, '').replace(/[€\s ]/g, '').replace(/^\+/, '').replace(/[−–]/g, '-');
  const neg = /^-/.test(s) || /-$/.test(s) || /^\(.*\)$/.test(s);
  s = s.replace(/[-()]/g, '');
  if (!/^\d[\d.,]*$/.test(s)) return NaN;
  const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (lc > -1) { const parts = s.split(','); const dec = parts.pop(); s = parts.join('') + '.' + dec; }
  else if (ld > -1) {
    const dots = (s.match(/\./g) || []).length;
    if (dots > 1 || s.length - ld - 1 === 3) s = s.replace(/\./g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : (neg ? -n : n);
}

function validISO(y, m, d) {
  if (y < 100) y += 2000;
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return null;
  const dt = new Date(y, m - 1, d);
  if (dt.getMonth() !== m - 1) return null;
  return toISO(dt);
}

/** Fechas "fuertes": 2026-11-13, 13/11/2026, 13-11-26, 13.11.2026 */
function parseDateStrict(s) {
  s = String(s || '').trim();
  let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return validISO(+m[1], +m[2], +m[3]);
  m = s.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/);
  if (m) return validISO(+m[3], +m[2], +m[1]);
  return null;
}
const MONTH_DATE_RE = /\b(\d{1,2})\s*(?:de\s+)?(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\.?(?:\s*(?:de\s+)?(\d{4}))?/i;
/** Además de las fuertes: "13 nov", "13 de noviembre", "13/11" */
function parseDateLoose(s) {
  const strict = parseDateStrict(s);
  if (strict) return strict;
  const now = new Date();
  const fix = iso => (iso && iso > addDays(todayISO(), 1)) ? addDays(iso, -365) : iso;
  let m = norm(s).match(MONTH_DATE_RE);
  if (m) {
    const mi = MONTHS_S.indexOf(m[2] === 'set' ? 'sep' : m[2]) + 1;
    const iso = validISO(m[3] ? +m[3] : now.getFullYear(), mi, +m[1]);
    return m[3] ? iso : fix(iso);
  }
  m = String(s).match(/\b(\d{1,2})\/(\d{1,2})\b(?!\/)/);
  if (m) return fix(validISO(now.getFullYear(), +m[2], +m[1]));
  return null;
}

function cleanNote(s) {
  let t = String(s || '')
    .replace(/\b(compra|pago)\s+(con\s+)?(en\s+)?(tarj(eta)?\.?|tj\.?|m[oó]vil)?(\s+(de\s+)?(cr[eé]dito|d[eé]bito))?/gi, ' ')
    .replace(/\btransacci[oó]n\b|\brecibo\b|\badeudo\b|\btarj(eta)?\b\.?\s*\d*|\b\d{4,}\b|\b\d*x{3,}\d*\b|\*+\d+/gi, ' ')
    .replace(/,?\s+s\.?\s?[al]\.?(u\.?)?\s*$/i, '')
    .replace(/[|_]+/g, ' ').replace(/\s+/g, ' ').trim();
  for (let i = 0; i < 3; i++) t = t.replace(/^(en|por|con|de|a)\s+|\s+(en|por|con|de|a)$/i, '').trim();
  if (t && t === t.toUpperCase()) t = t.toLowerCase().replace(/(^|\s)\S/g, x => x.toUpperCase());
  return cap(t.slice(0, 40));
}

function merchantKey(s) { return words(s).filter(x => x.length > 1 && !/^\d+$/.test(x)).slice(0, 2).join(' '); }

function autoCat(text) {
  if (!text) return null;
  const k = merchantKey(text);
  if (k && state.merchantMap[k]) return state.merchantMap[k];
  const w = words(text), ws = ' ' + w.join(' ') + ' ', joined = w.join('');
  for (const [cat, keys] of RULES) {
    for (const key of keys) {
      if (key.includes(' ')) { if (ws.includes(' ' + key + ' ')) return cat; }
      else if (w.includes(key) || (key.length >= 6 && joined.includes(key))) return cat;
    }
  }
  return null;
}
function learn(note, cat) { const k = merchantKey(note); if (k) state.merchantMap[k] = cat; }

const loadedScripts = {};
function loadScript(src) {
  if (!loadedScripts[src]) {
    loadedScripts[src] = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = src; s.onload = res;
      s.onerror = () => { delete loadedScripts[src]; rej(new Error('sin conexión')); };
      document.head.appendChild(s);
    });
  }
  return loadedScripts[src];
}

/* =====================================================================
   Datos
   ===================================================================== */
const DEFAULT = () => ({
  v: 1,
  settings: { income: 0, saveGoal: 0, startDay: 1, hoursMonth: 160, onboarded: false, lastBackup: null },
  expenses: [],   // {id, amount, cat, note, date, ts, source, recurringId?}
  recurring: [],  // {id, name, amount, cat, day, created}
  goals: [],      // {id, name, emoji, target, saved, created}
  wishlist: [],   // {id, name, price, created, waitH, status: waiting|bought|resisted, decided}
  noSpend: [],    // fechas ISO marcadas como "hoy no he gastado"
  merchantMap: {},
  badges: {},
  resistedTotal: 0,
  dismissed: {},
});

function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return DEFAULT();
    const s = JSON.parse(raw), d = DEFAULT();
    return { ...d, ...s, settings: { ...d.settings, ...(s.settings || {}) } };
  } catch (e) { return DEFAULT(); }
}
let state = load();
function persist() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
  catch (e) { toast('⚠️ No se ha podido guardar (¿memoria llena?)'); }
}
function save() {
  persist();
  syncSoon();
}

const ui = { tab: 'hoy', monthOffset: 0, cat: null, day: null };

/* =====================================================================
   Lógica: periodos, presupuesto, racha, gastos fijos
   ===================================================================== */
function periodRange(offset = 0, refISO = todayISO()) {
  const S = state.settings.startDay || 1;
  const r = fromISO(refISO);
  let y = r.getFullYear(), m = r.getMonth();
  if (r.getDate() < clampDay(y, m, S)) m -= 1;
  const base = new Date(y, m + offset, 1);
  y = base.getFullYear(); m = base.getMonth();
  const start = toISO(new Date(y, m, clampDay(y, m, S)));
  const nx = new Date(y, m + 1, 1);
  const end = addDays(toISO(new Date(nx.getFullYear(), nx.getMonth(), clampDay(nx.getFullYear(), nx.getMonth(), S))), -1);
  const label = S === 1
    ? `${cap(MONTHS[m])} ${y}`
    : `${fromISO(start).getDate()} ${MONTHS_S[fromISO(start).getMonth()]} – ${fromISO(end).getDate()} ${MONTHS_S[fromISO(end).getMonth()]}`;
  return { start, end, days: diffDays(start, end) + 1, key: start, label };
}
const inRange = p => state.expenses.filter(e => e.date >= p.start && e.date <= p.end);

function recurringDateIn(r, p) {
  const s = fromISO(p.start);
  for (let i = 0; i < 2; i++) {
    const b = new Date(s.getFullYear(), s.getMonth() + i, 1);
    const iso = toISO(new Date(b.getFullYear(), b.getMonth(), clampDay(b.getFullYear(), b.getMonth(), r.day)));
    if (iso >= p.start && iso <= p.end) return iso;
  }
  return null;
}
/** Apunta solos los gastos fijos cuando llega su día. */
function applyRecurring() {
  const t = todayISO();
  let added = 0;
  for (const off of [-1, 0]) {
    const p = periodRange(off);
    for (const r of state.recurring) {
      const iso = recurringDateIn(r, p);
      if (!iso || iso > t || (r.created && iso < r.created)) continue;
      if (state.expenses.some(e => e.recurringId === r.id && e.date === iso)) continue;
      state.expenses.push({ id: uid(), amount: r.amount, cat: r.cat, note: r.name, date: iso, ts: Date.now(), source: 'fijo', recurringId: r.id });
      added++;
    }
  }
  if (added) save();
}
function pendingRecurring(p) {
  const t = todayISO();
  return state.recurring.reduce((a, r) => {
    const iso = recurringDateIn(r, p);
    return a + (iso && iso > t ? r.amount : 0);
  }, 0);
}

/** Presupuesto de un periodo. El primer mes (si empiezas a usar la app a mitad)
    es proporcional a los días que quedan: lo de antes ya lo gastaste sin apuntarlo. */
function periodBudget(p) {
  const s = state.settings;
  const full = s.income > 0 ? Math.max(0, s.income - (s.saveGoal || 0)) : 0;
  const st = s.startedOn;
  if (full && st && st > p.start && st <= p.end) return full * (diffDays(st, p.end) + 1) / p.days;
  return full;
}

function budgetInfo() {
  const s = state.settings, p = periodRange(0), t = todayISO();
  const ex = inRange(p), spent = sum(ex), spentToday = sum(ex.filter(e => e.date === t));
  const budget = periodBudget(p);
  const firstMonth = !!(s.startedOn && s.startedOn > p.start && s.startedOn <= p.end);
  const pending = pendingRecurring(p);
  const daysLeft = diffDays(t, p.end) + 1;
  const remaining = budget - spent - pending;
  const startOfDay = budget - (spent - spentToday) - pending;
  const daily = startOfDay / daysLeft;
  return { p, budget, spent, spentToday, pending, daysLeft, remaining, daily, todayLeft: daily - spentToday, hasBudget: budget > 0, firstMonth };
}

const HABIT_SOURCES = ['manual', 'ocr', 'paste', 'wish'];
function loggedDays() {
  const set = new Set(state.noSpend);
  for (const e of state.expenses) if (HABIT_SOURCES.includes(e.source)) set.add(e.date);
  return set;
}
/** Racha "con perdón": un día suelto sin apuntar no la rompe; dos seguidos, sí. */
function streak() {
  const days = loggedDays();
  let d = todayISO();
  if (!days.has(d)) d = addDays(d, -1);
  let count = 0, miss = 0;
  for (let i = 0; i < 2000; i++) {
    if (days.has(d)) { count++; miss = 0; }
    else if (++miss >= 2) break;
    d = addDays(d, -1);
  }
  return count;
}
function noSpendDays() {
  const spentDays = new Set(state.expenses.filter(e => e.source !== 'fijo').map(e => e.date));
  return state.noSpend.filter(d => !spentDays.has(d));
}

/* =====================================================================
   Detector de ahorro
   ===================================================================== */
function detectSubscriptions() {
  const res = new Map();
  for (const r of state.recurring) if (r.cat === 'subs') res.set('r' + r.id, { name: r.name, amount: r.amount });
  const since = periodRange(-3).start;
  const groups = {};
  for (const e of state.expenses) {
    if (e.date < since || e.source === 'fijo' || catOf(e.cat).type !== 'want') continue;
    const k = merchantKey(e.note);
    if (k) (groups[k] = groups[k] || []).push(e);
  }
  for (const [k, arr] of Object.entries(groups)) {
    const months = new Set(arr.map(e => periodRange(0, e.date).key));
    if (months.size < 2 || arr.length > months.size * 1.5) continue;
    const avg = sum(arr) / arr.length;
    if (avg >= 1.5 && arr.every(e => Math.abs(e.amount - avg) <= avg * 0.12)) res.set('k' + k, { name: arr[arr.length - 1].note, amount: avg });
  }
  return [...res.values()];
}

function computeInsights() {
  const out = [];
  if (state.expenses.length < 5) return out;
  const t = todayISO(), p = periodRange(0), cur = inRange(p);
  const variable = cur.filter(e => e.source !== 'fijo');
  const elapsed = diffDays(p.start, t) + 1, frac = elapsed / p.days;
  const monthly = tot => (elapsed >= 7 ? tot / frac : tot);
  const b = budgetInfo();

  // Ritmo del mes
  if (b.hasBudget && elapsed >= 5 && variable.length >= 3) {
    const fixedTot = sum(cur.filter(e => e.source === 'fijo')) + b.pending;
    const proj = fixedTot + sum(variable) / frac;
    if (proj > b.budget * 1.03) {
      const perDay = (proj - b.budget) / b.daysLeft;
      out.push({ id: 'pace', tone: 'warn', emoji: '🧭', title: `A este ritmo te pasarías ≈ ${moneyR(proj - b.budget)}`, text: `Si gastas unos ${moneyR(Math.max(perDay, 1))} menos al día lo que queda de mes, cuadras. Sin dramas: un ajuste pequeño basta.` });
    } else if (proj < b.budget * 0.95) {
      out.push({ id: 'pace', tone: 'good', emoji: '🌤️', title: 'Vas bien este mes', text: `A este ritmo acabarías con ≈ ${moneyR(b.budget - proj)} de margen. ¿Lo mandas a una meta cuando acabe el mes?` });
    }
  }

  // Suscripciones y pagos repetidos
  const subs = detectSubscriptions();
  if (subs.length) {
    const m = subs.reduce((a, s) => a + s.amount, 0);
    const cheapest = Math.min(...subs.map(s => s.amount));
    out.push({ id: 'subs', tone: 'tip', emoji: '📺', title: `${subs.length} ${subs.length === 1 ? 'pago que se repite' : 'pagos que se repiten'}: ${money(m)}/mes`, text: `Son ${moneyR(m * 12)} al año: ${subs.map(s => `${esc(s.name)} (${money(s.amount)})`).join(', ')}. ¿Usas cada una al menos una vez por semana? Si no, cancélala y vuelve cuando la eches de menos.`, save: cheapest * 12, saveText: `Cancelar la más barata: ${moneyR(cheapest * 12)}/año` });
  }

  // Comida a domicilio
  const del = variable.filter(e => e.cat === 'delivery');
  if (del.length >= 2) {
    const mo = monthly(sum(del)), s = mo * 0.5 * 0.65 * 12;
    out.push({ id: 'delivery', tone: 'tip', emoji: '🛵', title: `${del.length} pedidos a domicilio: ${money(sum(del))}`, text: `Truco TDAH: ten 2–3 "cenas de emergencia" en casa (pasta, huevos, congelados) para los días sin energía. Cambiar la mitad de los pedidos por eso ahorra mucho.`, save: s, saveText: `≈ ${moneyR(s)}/año` });
  }

  // Cafés
  const cafes = variable.filter(e => e.cat === 'cafe');
  if (cafes.length >= 6) {
    const mo = monthly(sum(cafes)), s = mo * 0.5 * 12;
    out.push({ id: 'cafe', tone: 'tip', emoji: '☕', title: `${cafes.length} cafés y snacks: ${money(sum(cafes))}`, text: `No hace falta dejarlos: prueba a llevar un termo dos días por semana o elige "días de café" fijos.`, save: s, saveText: `La mitad: ${moneyR(s)}/año` });
  }

  // Gastos hormiga
  const small = variable.filter(e => e.amount <= 8 && catOf(e.cat).type === 'want' && e.cat !== 'cafe');
  if (small.length >= 5) {
    const mo = monthly(sum(small)), s = mo * 0.5 * 12;
    out.push({ id: 'goteo', tone: 'tip', emoji: '💧', title: `Gastos hormiga: ${small.length} compras pequeñas`, text: `Compras de menos de 8 € que suman ${money(sum(small))} este mes. Una a una no se notan; juntas, sí.`, save: s, saveText: `Recortando la mitad: ${moneyR(s)}/año` });
  }

  // Categorías que se disparan
  // Solo meses completos (el primer mes que usas la app suele estar a medias)
  const firstDate = state.expenses.reduce((m, e) => (e.date < m ? e.date : m), t);
  const prev = [1, 2, 3].map(o => periodRange(-o)).filter(pp => pp.start >= firstDate)
    .map(pp => inRange(pp).filter(e => e.source !== 'fijo')).filter(a => a.length >= 5);
  if (prev.length) {
    for (const c of CATS) {
      const avg = prev.reduce((a, arr) => a + sum(arr.filter(e => e.cat === c.id)), 0) / prev.length;
      const mine = variable.filter(e => e.cat === c.id), now = sum(mine);
      // Solo se proyecta a fin de mes lo que se compra a menudo (no un pago mensual suelto)
      const frequent = mine.length >= 4 && c.id !== 'subs' && c.id !== 'casa';
      const proj = elapsed >= 10 && frequent ? now / frac : now;
      if (avg >= 20 && proj > avg * 1.3 && proj - avg >= 25) {
        out.push({ id: 'spike-' + c.id, tone: 'warn', emoji: c.emoji, title: `${c.name}: más que de costumbre`, text: `Llevas ${money(now)} y sueles gastar ≈ ${moneyR(avg)} al mes.${proj !== now ? ` A este ritmo acabarías en ≈ ${moneyR(proj)}.` : ''} ¿Ha pasado algo especial o se ha colado la costumbre?` });
      }
    }
  }

  // Fines de semana
  if (variable.length >= 10) {
    const we = variable.filter(e => [0, 6].includes(fromISO(e.date).getDay()));
    const share = sum(we) / Math.max(sum(variable), 1);
    if (share > 0.5) out.push({ id: 'weekend', tone: 'tip', emoji: '📅', title: `Los findes se llevan el ${Math.round(share * 100)} %`, text: 'De tus gastos variables, más de la mitad cae en sábado y domingo. Prueba a decidir el viernes cuánto quieres gastar el finde (y apúntalo en una nota).' });
  }

  // Sitio más repetido
  const byM = {};
  for (const e of variable) { const k = merchantKey(e.note); if (k) (byM[k] = byM[k] || []).push(e); }
  const covered = new Set(out.map(i => i.id));
  const top = Object.values(byM).filter(a => a.length >= 4 && a[0].cat !== 'super' && !covered.has(a[0].cat)).sort((a, b) => sum(b) - sum(a))[0];
  if (top) out.push({ id: 'top', tone: 'tip', emoji: '📍', title: `${esc(top[top.length - 1].note)}: ${top.length} veces este mes`, text: `Suman ${money(sum(top))}. No está mal ni bien: solo para que lo sepas. ¿Te compensa?` });

  // 50/30/20
  if (state.settings.income > 0 && elapsed >= 10) {
    const wants = sum(cur.filter(e => catOf(e.cat).type === 'want'));
    const share = monthly(wants) / state.settings.income;
    if (share > 0.35) out.push({ id: '5030', tone: 'warn', emoji: '⚖️', title: `Caprichos: ${Math.round(share * 100)} % de tus ingresos`, text: 'La regla 50/30/20 sugiere ≈ 30 % para caprichos. No es una ley, es una brújula: mira qué categoría de "caprichos" te aporta menos y empieza por ahí.' });
  }

  if (state.resistedTotal > 0) out.push({ id: 'resisted', tone: 'good', emoji: '🛡️', title: `Has evitado gastar ${moneyR(state.resistedTotal)}`, text: 'Gracias a la lista de espera. Eso es dinero que sigue siendo tuyo.' });

  const order = { warn: 0, tip: 1, good: 2 };
  return out.filter(i => !state.dismissed['ins:' + i.id + ':' + p.key])
    .sort((a, b) => order[a.tone] - order[b.tone] || (b.save || 0) - (a.save || 0));
}

/* =====================================================================
   Logros
   ===================================================================== */
const BADGE_RULES = {
  first: () => state.expenses.some(e => HABIT_SOURCES.includes(e.source)),
  streak3: () => streak() >= 3,
  streak7: () => streak() >= 7,
  streak30: () => streak() >= 30,
  nospend1: () => noSpendDays().length >= 1,
  nospend5: () => { const p = periodRange(0); return noSpendDays().filter(d => d >= p.start && d <= p.end).length >= 5; },
  capture: () => state.expenses.some(e => e.source === 'ocr' || e.source === 'ia'),
  goal: () => state.goals.length > 0,
  goaldone: () => state.goals.some(g => g.saved >= g.target),
  resist: () => state.wishlist.some(w => w.status === 'resisted'),
  resist100: () => state.resistedTotal >= 100,
  underbudget: () => {
    const pp = periodRange(-1), b = periodBudget(pp); if (!b) return false;
    const ex = inRange(pp);
    return ex.filter(e => e.source !== 'fijo').length >= 10 && sum(ex) <= b;
  },
  backup: () => !!state.settings.lastBackup,
};
function checkBadges() {
  const newly = BADGES.filter(b => !state.badges[b.id] && BADGE_RULES[b.id]());
  if (!newly.length) return;
  newly.forEach(b => { state.badges[b.id] = todayISO(); });
  save();
  newly.forEach((b, i) => setTimeout(() => celebrate(b), 700 + i * 3000));
}

/* =====================================================================
   UI básica: render, hojas, avisos, confeti
   ===================================================================== */
function render() {
  const views = { hoy: viewHoy, mes: viewMes, ahorro: viewAhorro, metas: viewMetas };
  $('#view').innerHTML = views[ui.tab]();
  $$('.tabbar [data-tab]').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === ui.tab);
    b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false');
  });
  const h = new Date().getHours();
  const titles = {
    hoy: `${h < 6 ? 'Buenas noches' : h < 14 ? 'Buenos días' : h < 21 ? 'Buenas tardes' : 'Buenas noches'}<small>${cap(WEEKDAYS[new Date().getDay()])} ${new Date().getDate()} de ${MONTHS[new Date().getMonth()]}</small>`,
    mes: 'Tu mes<small>Dónde se va tu dinero</small>',
    ahorro: 'Ahorro<small>Pistas, sin sermones</small>',
    metas: 'Metas<small>Para qué ahorras</small>',
  };
  $('#topTitle').innerHTML = titles[ui.tab];
}

let sheetKind = null;
function openSheet(html, kind = '') {
  const root = $('#sheetRoot');
  root.innerHTML = `<div class="overlay" data-action="overlay"></div><div class="sheet" role="dialog" aria-modal="true">${html}</div>`;
  root.classList.add('open');
  sheetKind = kind;
  document.body.classList.add('noscroll');
  initWheels();
}
function setSheet(html) { const s = $('.sheet'); if (s) { const top = s.scrollTop; s.innerHTML = html; s.scrollTop = top; initWheels(); } }
function closeSheet() {
  const root = $('#sheetRoot');
  root.classList.remove('open');
  root.innerHTML = '';
  sheetKind = null; add = null;
  document.body.classList.remove('noscroll');
}
const sheetHead = (title = '') => `<div class="sheet-head"><h2 style="margin:0">${title}</h2><button class="icon-btn" data-action="closeSheet" aria-label="Cerrar">✕</button></div>`;

let toastTimer;
function toast(msg, opts = {}) {
  const el = $('#toast');
  el.innerHTML = `<span>${esc(msg)}</span>${opts.undo ? '<button type="button">Deshacer</button>' : ''}`;
  el.classList.add('show');
  if (opts.undo) el.querySelector('button').onclick = () => { opts.undo(); save(); el.classList.remove('show'); render(); };
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), opts.undo ? 5500 : 3200);
}

function confetti() {
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const wrap = document.createElement('div');
  wrap.className = 'confetti';
  const colors = ['#4f46e5', '#22a06b', '#f97316', '#e11d48', '#eab308', '#0891b2', '#db2777'];
  for (let i = 0; i < 46; i++) {
    const s = document.createElement('span');
    s.style.left = Math.random() * 100 + '%';
    s.style.background = colors[i % colors.length];
    s.style.animationDuration = 1.4 + Math.random() * 1.2 + 's';
    s.style.animationDelay = Math.random() * 0.3 + 's';
    s.style.transform = `rotate(${Math.random() * 360}deg)`;
    wrap.appendChild(s);
  }
  document.body.appendChild(wrap);
  setTimeout(() => wrap.remove(), 3000);
}
function celebrate(b) {
  const el = $('#celebrate');
  el.innerHTML = `<div class="e">${b.emoji}</div><div><b>¡Logro desbloqueado!</b><div>${esc(b.name)} — <span class="muted">${esc(b.desc)}</span></div></div>`;
  el.classList.add('show');
  confetti();
  setTimeout(() => el.classList.remove('show'), 3200);
}
function shake(sel) { const el = $(sel); if (!el) return; el.classList.remove('shake'); void el.offsetWidth; el.classList.add('shake'); }

function listRows(arr) {
  return `<ul class="rows">${arr.map(e => {
    const c = catOf(e.cat);
    return `<li><button class="row" data-action="editExpense" data-id="${e.id}">
      <span class="ico" style="background:${c.color}22">${c.emoji}</span>
      <span class="row-main"><span class="row-title">${esc(e.note || c.name)}</span><span class="row-sub">${esc(c.name)}${SRC_ICON[e.source] || ''}</span></span>
      <span class="row-amt">${money(e.amount)}</span></button></li>`;
  }).join('')}</ul>`;
}
const byRecent = (a, b) => (b.date.localeCompare(a.date)) || (b.ts - a.ts);

/* =====================================================================
   Vista: HOY
   ===================================================================== */
function viewHoy() {
  const b = budgetInfo(), t = todayISO();
  const todays = state.expenses.filter(e => e.date === t).sort(byRecent);
  const manualToday = todays.filter(e => e.source !== 'fijo');
  let html = '';

  if (isIOS() && !isStandalone() && !state.dismissed.install) {
    html += `<div class="card hint">📲 <b>Instálala en tu iPhone</b> para que no se borren tus datos: toca <b>Compartir</b> (□↑) → <b>Añadir a pantalla de inicio</b>.<button class="x" data-action="dismiss" data-k="install" aria-label="Cerrar aviso">✕</button></div>`;
  }

  if (b.hasBudget) {
    const pct = Math.min(100, b.budget ? b.spent / b.budget * 100 : 0);
    const monthOver = b.remaining < 0 && b.budget - (b.spent - b.spentToday) - b.pending <= 0;
    let label, big, sub, cls = '';
    if (monthOver) {
      label = 'Presupuesto del mes superado'; big = moneyR(-b.remaining); cls = 'is-warn';
      sub = 'por encima. Respira: lo importante es saberlo. Intenta que lo que queda de mes sean días de gasto mínimo 💪';
    } else if (b.todayLeft < 0) {
      label = 'Hoy te has pasado'; big = moneyR(-b.todayLeft); cls = 'is-warn';
      sub = `No pasa nada: se reparte entre los próximos días. Quedan ${moneyR(Math.max(b.remaining, 0))} para ${b.daysLeft} ${b.daysLeft === 1 ? 'día' : 'días'}.`;
    } else {
      label = 'Hoy puedes gastar'; big = moneyR(b.todayLeft); cls = 'is-good';
      sub = `Quedan <b>${moneyR(b.remaining)}</b> para ${b.daysLeft} ${b.daysLeft === 1 ? 'día' : 'días'}${b.pending ? ` · ${moneyR(b.pending)} guardados para fijos` : ''}`
        + (b.firstMonth ? '<br><small>Primer mes: el presupuesto es proporcional a los días que quedan.</small>' : '');
    }
    html += `<section class="card hero ${cls}">
      <div class="label">${label}</div>
      <div class="big-number">${big}</div>
      <div class="sub">${sub}</div>
      <div class="bar"><i style="width:${pct}%;${pct > 90 ? 'background:var(--warn)' : ''}"></i></div>
      <div class="bar-legend"><span>Gastado ${moneyR(b.spent)}</span><span>de ${moneyR(b.budget)}</span></div>
    </section>`;
  } else {
    html += `<section class="card hero"><div class="label">Este mes llevas</div><div class="big-number">${moneyR(b.spent)}</div>
      <div class="sub">Dime cuánto cobras y te diré cuánto puedes gastar cada día.</div>
      <button class="btn primary w100" data-action="openSettings">Poner mis ingresos</button></section>`;
  }

  // Racha
  const st = streak(), days = loggedDays();
  const mon = addDays(t, -((fromISO(t).getDay() + 6) % 7));
  const dots = [...Array(7)].map((_, i) => {
    const d = addDays(mon, i);
    return `<div class="dot ${days.has(d) ? 'on' : ''} ${d === t ? 'today' : ''}" title="${niceDate(d)}">${'LMXJVSD'[i]}</div>`;
  }).join('');
  const marked = state.noSpend.includes(t);
  html += `<section class="card">
    <div class="streak-n">🔥 ${st} ${st === 1 ? 'día' : 'días'} de racha</div>
    <div class="muted small">Cuenta cada día que apuntas algo. Saltarte un día suelto no la rompe.</div>
    <div class="week">${dots}</div>
    ${manualToday.length === 0 && !marked ? '<button class="btn ghost w100" data-action="noSpendToday">✨ Hoy no he gastado nada</button>' : ''}
    ${manualToday.length === 0 && marked ? '<div class="pill-note">✨ Hoy, día sin gastos (de momento)</div>' : ''}
  </section>`;

  if (todays.length) html += `<section class="card"><h3>Hoy · ${money(sum(todays))}</h3>${listRows(todays)}</section>`;
  else html += `<section class="card empty"><div class="empty-emoji">🪙</div><p>Aún no has apuntado nada hoy.<br>Toca <b>+</b> y lo tienes en 2 toques.</p></section>`;

  // Empujoncitos (uno cada vez, para no saturar)
  const tip = computeInsights()[0];
  if (b.hasBudget && !state.recurring.length && !state.dismissed.fixed) {
    html += `<div class="card hint">🔁 <b>¿Tienes gastos fijos?</b> (alquiler, móvil, gimnasio…) Añádelos una vez y se apuntarán solos cada mes.
      <div style="margin-top:10px"><button class="btn sm primary" data-action="openSettings">Añadir fijos</button></div>
      <button class="x" data-action="dismiss" data-k="fixed" aria-label="Cerrar aviso">✕</button></div>`;
  } else if (tip) {
    html += `<section class="card"><button class="tip-card" data-tab="ahorro"><span class="emoji">${tip.emoji}</span><span><b>${tip.title}</b><br><span class="muted small">Toca para ver más pistas de ahorro →</span></span></button></section>`;
  }

  const last = state.settings.lastBackup;
  if (state.expenses.length >= 15 && (!last || diffDays(last, t) > 30)) {
    html += `<div class="card hint">💾 <b>Haz una copia</b> (10 segundos). Tus datos solo viven en este dispositivo.
      <div style="margin-top:10px"><button class="btn sm primary" data-action="exportBackup">Guardar copia</button></div></div>`;
  }
  return html;
}

/* =====================================================================
   Vista: MES
   ===================================================================== */
function donut(parts, total) {
  const R = 60, C = 2 * Math.PI * R;
  let off = 0;
  const segs = parts.map(p => {
    const len = p.v / total * C;
    const s = `<circle r="${R}" cx="80" cy="80" fill="none" stroke="${p.c.color}" stroke-width="24" stroke-dasharray="${len} ${C - len}" stroke-dashoffset="${-off}" transform="rotate(-90 80 80)"/>`;
    off += len;
    return s;
  }).join('');
  return `<svg viewBox="0 0 160 160" class="donut" role="img" aria-label="Gasto por categoría"><circle r="${R}" cx="80" cy="80" fill="none" stroke="var(--track)" stroke-width="24"/>${segs}</svg>`;
}

function viewMes() {
  const off = ui.monthOffset, p = periodRange(off), t = todayISO();
  const ex = inRange(p), total = sum(ex);
  const prevP = periodRange(off - 1), prevEx = inRange(prevP);
  let html = `<div class="month-nav"><button data-action="month" data-d="-1" aria-label="Mes anterior">‹</button><b>${p.label}</b><button data-action="month" data-d="1" aria-label="Mes siguiente" ${off >= 0 ? 'disabled' : ''}>›</button></div>`;

  let cmp = '';
  if (prevEx.length) {
    const base = off === 0 ? sum(prevEx.filter(e => e.date <= addDays(prevP.start, diffDays(p.start, t)))) : sum(prevEx);
    if (base > 0) {
      const d = (total - base) / base * 100;
      cmp = `<span class="cmp ${d > 0 ? 'up' : 'down'}">${d > 0 ? '↑' : '↓'} ${Math.abs(Math.round(d))} % ${off === 0 ? 'vs. el mes pasado a estas alturas' : 'vs. el mes anterior'}</span>`;
    }
  }
  const budget = periodBudget(p);
  html += `<section class="card hero"><div class="label">Gastado</div><div class="big-number">${money(total)}</div>${cmp}
    ${budget ? `<div class="bar" style="margin-top:14px"><i style="width:${Math.min(100, total / budget * 100)}%;${total > budget ? 'background:var(--warn)' : ''}"></i></div><div class="bar-legend"><span>${Math.round(total / budget * 100)} % del presupuesto</span><span>${moneyR(budget)}</span></div>` : ''}
  </section>`;

  if (!ex.length) return html + `<section class="card empty"><div class="empty-emoji">📭</div><p>No hay gastos en este periodo.</p></section>`;

  const byCat = CATS.map(c => ({ c, v: sum(ex.filter(e => catOf(e.cat).id === c.id)) })).filter(x => x.v > 0).sort((a, b) => b.v - a.v);
  html += `<section class="card"><h3>¿En qué se va?</h3>
    <div class="donut-wrap">${donut(byCat, total)}<div class="donut-center"><div><b>${moneyR(total)}</b><small>${ex.length} ${ex.length === 1 ? 'gasto' : 'gastos'}</small></div></div></div>
    <ul class="legend">${byCat.map(x => `<li><button data-action="filterCat" data-cat="${x.c.id}" class="${ui.cat === x.c.id ? 'on' : ''}">
      <span class="sw" style="background:${x.c.color}22">${x.c.emoji}</span><span>${x.c.name}</span><span class="pct">${Math.round(x.v / total * 100)} %</span><b class="num">${money(x.v)}</b></button></li>`).join('')}</ul>
    <p class="muted small" style="margin:8px 0 0">Toca una categoría para ver solo esos gastos.</p></section>`;

  const needs = sum(ex.filter(e => catOf(e.cat).type === 'need')), wants = total - needs;
  html += `<section class="card"><h3>Necesidades vs. caprichos</h3>
    <div class="split"><i style="width:${needs / total * 100}%;background:#22a06b"></i><i style="width:${wants / total * 100}%;background:#f97316"></i></div>
    <div class="split-legend"><span><i class="sw-dot" style="background:#22a06b"></i>Necesidades <b>${moneyR(needs)}</b></span><span><i class="sw-dot" style="background:#f97316"></i>Caprichos <b>${moneyR(wants)}</b></span></div>
    <p class="muted small" style="margin:8px 0 0">"Caprichos" no es malo: es lo que puedes ajustar si lo necesitas.</p></section>`;

  html += calendarHTML(p, ex, t);

  const list = ex.filter(e => (!ui.cat || catOf(e.cat).id === ui.cat) && (!ui.day || e.date === ui.day)).sort(byRecent);
  const groups = {};
  for (const e of list) (groups[e.date] = groups[e.date] || []).push(e);
  const filters = [
    ui.day ? `<button class="pill" data-action="pickDay" data-d="${ui.day}">${niceDate(ui.day)} ✕</button>` : '',
    ui.cat ? `<button class="pill" data-action="filterCat" data-cat="${ui.cat}">${catOf(ui.cat).emoji} ${catOf(ui.cat).name} ✕</button>` : '',
  ].join(' ');
  html += `<section class="card" id="movs"><h3>Movimientos ${filters}</h3>
    ${list.length ? Object.entries(groups).map(([d, arr]) => `<div class="day-head"><span>${niceDate(d)}</span><span>${money(sum(arr))}</span></div>${listRows(arr)}`).join('')
      : `<p class="muted">${ui.day && ui.day <= t ? 'Ese día no gastaste nada ✨' : 'No hay gastos con este filtro.'}</p>`}
  </section>`;
  return html;
}

/* ---------- Calendario: qué días gastas más ---------- */
function calendarHTML(p, ex, t) {
  const perDay = {};
  for (const e of ex) if (e.source !== 'fijo') perDay[e.date] = (perDay[e.date] || 0) + e.amount;
  const max = Math.max(0, ...Object.values(perDay));
  const level = v => (v > 0 && max > 0 ? Math.max(1, Math.ceil(Math.sqrt(v / max) * 4)) : 0);
  const short = v => (v >= 1000 ? (v / 1000).toFixed(1).replace('.', ',') + 'k' : Math.round(v)) + '€';
  const blanks = (fromISO(p.start).getDay() + 6) % 7; // lunes primero
  let cells = '<span class="cal-blank"></span>'.repeat(blanks);
  for (let i = 0; i < p.days; i++) {
    const d = addDays(p.start, i), dt = fromISO(d), v = perDay[d] || 0;
    const label = dt.getDate() + (dt.getDate() === 1 || i === 0 ? ` ${MONTHS_S[dt.getMonth()]}` : '');
    cells += `<button class="cal-d l${level(v)}${d === t ? ' today' : ''}${d > t ? ' future' : ''}${ui.day === d ? ' sel' : ''}" data-action="pickDay" data-d="${d}" aria-label="${niceDate(d)}: ${money(v)}">
      <span class="n">${label}</span>${v ? `<span class="v">${short(v)}</span>` : ''}</button>`;
  }
  return `<section class="card"><h3>Calendario <span class="muted small">(sin gastos fijos)</span></h3>
    <div class="cal-head">${['L', 'M', 'X', 'J', 'V', 'S', 'D'].map(x => `<span>${x}</span>`).join('')}</div>
    <div class="cal">${cells}</div>
    <div class="cal-legend">Menos ${[1, 2, 3, 4].map(l => `<i class="l${l}"></i>`).join('')} Más</div>
    ${weekdayInsight(t)}
    <p class="muted small" style="margin:6px 0 0">Toca un día para ver qué gastaste.</p></section>`;
}
/** Qué día de la semana sueles gastar más (últimos 90 días, sin fijos). */
function weekdayInsight(t) {
  const from = addDays(t, -89);
  const ex = state.expenses.filter(e => e.source !== 'fijo' && e.date >= from && e.date <= t);
  if (!ex.length) return '';
  const first = ex.reduce((m, e) => (e.date < m ? e.date : m), t);
  const span = diffDays(first, t) + 1;
  if (span < 14) return '';
  const tot = Array(7).fill(0), n = Array(7).fill(0);
  for (let i = 0; i < span; i++) n[fromISO(addDays(first, i)).getDay()]++;
  for (const e of ex) tot[fromISO(e.date).getDay()] += e.amount;
  const avg = tot.map((x, i) => (n[i] ? x / n[i] : 0));
  const top = avg.indexOf(Math.max(...avg));
  if (!avg[top]) return '';
  const rest = avg.filter((_, i) => i !== top && n[i]);
  const restAvg = rest.reduce((a, b) => a + b, 0) / (rest.length || 1);
  return `<p class="cal-insight">📌 Tu día más caro suele ser el <b>${WEEKDAYS[top]}</b>: de media <b>${money(avg[top])}</b> (el resto de días, ${money(restAvg)}).</p>`;
}

/* =====================================================================
   Vista: AHORRO
   ===================================================================== */
const TIPS = [
  ['💸', '<b>Págate primero.</b> Programa en tu banco una transferencia automática a otra cuenta el mismo día que cobras. Lo que no ves, no lo gastas.'],
  ['⏳', '<b>Regla de las 48 horas.</b> ¿Un capricho de más de 30 €? A la lista de espera (pestaña Metas). Si en 2 días sigues queriéndolo, adelante.'],
  ['🧱', '<b>Pon fricción.</b> Borra las tarjetas guardadas en Amazon, Shein y apps de comida. Tener que sacar la tarjeta frena muchas compras impulsivas.'],
  ['🔕', '<b>Silencia las tentaciones.</b> Date de baja de newsletters de tiendas y quita notificaciones de apps de compras y delivery.'],
  ['🫙', '<b>Botes separados.</b> Una cuenta para fijos, otra para el día a día. Si la del día a día se vacía, sabes que toca frenar.'],
  ['📝', '<b>Lista antes del súper</b> y nunca con hambre. Compra online si en tienda te dispersas.'],
  ['⏱️', '<b>¿Cuántas horas de trabajo cuesta?</b> Al apuntar un gasto te lo enseño si pones tus ingresos.'],
  ['📆', '<b>Revisión de 10 minutos</b> el primer domingo de cada mes: mira la pestaña Mes y elige UNA cosa que mejorar. Solo una.'],
];

function viewAhorro() {
  const ins = computeInsights();
  let html = '';
  const total = ins.reduce((a, i) => a + (i.save || 0), 0);
  if (total > 0) {
    html += `<section class="card total-save"><div class="label muted">Ahorro posible detectado</div><div class="big-number">${moneyR(total)}<span style="font-size:22px">/año</span></div><div class="muted small">Sin dejar de vivir: solo con pequeños ajustes de abajo.</div></section>`;
  }
  if (!ins.length) {
    html += `<section class="card empty"><div class="empty-emoji">🔍</div><p>${state.expenses.length < 5 ? 'Apunta tus gastos unos días y aquí aparecerán pistas de dónde puedes ahorrar.' : 'Ahora mismo no veo nada raro. ¡Bien hecho! 👏'}</p></section>`;
  }
  const p = periodRange(0);
  html += ins.map(i => `<section class="card insight ${i.tone}">
    <div class="insight-head"><span class="emoji">${i.emoji}</span><b>${i.title}</b></div>
    <p class="muted" style="margin-bottom:6px">${i.text}</p>
    ${i.saveText ? `<span class="save-tag">💰 ${i.saveText}</span>` : ''}
    <button class="x" data-action="dismiss" data-k="ins:${i.id}:${p.key}" aria-label="Ocultar esta pista este mes">✕</button>
  </section>`).join('');
  html += `<section class="card"><h3>Ideas que funcionan con TDAH</h3><ul class="tips">${TIPS.map(([e, t]) => `<li><span class="emoji">${e}</span><span>${t}</span></li>`).join('')}</ul></section>`;
  return html;
}

/* =====================================================================
   Vista: METAS (metas, lista de espera y logros)
   ===================================================================== */
function waitLeft(w) { return w.created + w.waitH * 3600e3 - Date.now(); }
function fmtLeft(ms) {
  const h = Math.ceil(ms / 3600e3);
  if (h >= 48) return `${Math.ceil(h / 24)} días`;
  if (h > 1) return `${h} h`;
  return `${Math.max(1, Math.ceil(ms / 60000))} min`;
}

function viewMetas() {
  let html = '';
  // Metas
  html += `<div class="section-title" style="margin-top:4px">Mis metas</div>`;
  if (!state.goals.length) {
    const b = budgetInfo();
    const emergency = b.hasBudget ? Math.round(b.budget * 3 / 100) * 100 : 0;
    html += `<section class="card empty"><div class="empty-emoji">🎯</div><p>Ahorrar "por ahorrar" cuesta. Ahorrar <b>para algo</b> motiva mucho más.</p>
      ${emergency ? `<p class="small">Idea para empezar: un <b>colchón de emergencia</b> de 3 meses de gastos ≈ ${moneyR(emergency)}.</p>` : ''}
      <button class="btn primary" data-action="newGoal">Crear mi primera meta</button></section>`;
  } else {
    html += state.goals.map(g => {
      const pct = Math.min(100, g.saved / g.target * 100);
      const done = g.saved >= g.target;
      return `<section class="card goal">
        <div class="goal-top"><span class="goal-emoji">${g.emoji}</span><div><div class="goal-name">${esc(g.name)}</div><div class="muted small num">${money(g.saved)} de ${money(g.target)}</div></div>
        <button class="goal-menu" data-action="goalMenu" data-id="${g.id}" aria-label="Opciones de la meta">⋯</button></div>
        <div class="bar"><i style="width:${pct}%"></i></div>
        <div class="bar-legend"><span>${done ? '🥳 ¡Conseguida!' : `${Math.round(pct)} %`}</span><span>${done ? '' : `faltan ${money(g.target - g.saved)}`}</span></div>
        ${done ? '' : `<div class="goal-btns">${[10, 25, 50].map(n => `<button class="btn good" data-action="goalAdd" data-id="${g.id}" data-n="${n}">+${n} €</button>`).join('')}<button class="btn ghost" data-action="goalOther" data-id="${g.id}">Otra</button></div>`}
      </section>`;
    }).join('');
    html += `<button class="btn ghost w100" data-action="newGoal" style="margin-bottom:8px">＋ Nueva meta</button>
      <p class="muted small" style="text-align:center">Apunta aquí lo que pasas a tu cuenta de ahorro.</p>`;
  }

  // Lista de espera
  const waiting = state.wishlist.filter(w => w.status === 'waiting').sort((a, b) => waitLeft(a) - waitLeft(b));
  html += `<div class="section-title">Lista de espera ⏳</div><section class="card">
    <p class="muted small">¿Te ha dado un antojo? Déjalo aquí en vez de comprarlo. Cuando pase el tiempo, decides con la cabeza fría.</p>
    ${state.resistedTotal > 0 ? `<div class="resisted"><span class="e">🛡️</span><span>Has evitado gastar ${money(state.resistedTotal)}</span></div>` : ''}
    ${waiting.map(w => {
      const left = waitLeft(w), ready = left <= 0;
      const pct = ready ? 100 : Math.min(100, (1 - left / (w.waitH * 3600e3)) * 100);
      return `<div class="wish ${ready ? 'ready' : ''}">
        <div class="wish-top"><span>${esc(w.name)}</span><span class="num">${money(w.price)}</span></div>
        <div class="bar"><i style="width:${pct}%"></i></div>
        <div class="muted small">${ready ? '¿Todavía lo quieres? Decide ahora.' : `Podrás decidir en ${fmtLeft(left)}. ¿Te sigue apeteciendo tanto?`}</div>
        <div class="btn-row">
          ${ready ? `<button class="btn sm ghost" data-action="wishBuy" data-id="${w.id}">Sí, lo compro</button>` : ''}
          <button class="btn sm good" data-action="wishResist" data-id="${w.id}">Ya no lo quiero 🎉</button>
        </div></div>`;
    }).join('')}
    <button class="btn primary w100" data-action="newWish" style="margin-top:12px">＋ Añadir antojo</button>
  </section>`;

  // Logros
  const got = BADGES.filter(b => state.badges[b.id]).length;
  html += `<div class="section-title">Logros · ${got}/${BADGES.length}</div><section class="card"><div class="badges">
    ${BADGES.map(b => `<div class="badge ${state.badges[b.id] ? '' : 'off'}" title="${esc(b.desc)}"><span class="e">${b.emoji}</span>${esc(b.name)}<div class="muted" style="font-weight:500;font-size:11px">${esc(b.desc)}</div></div>`).join('')}
  </div></section>`;
  return html;
}

/* =====================================================================
   Hoja: APUNTAR GASTO (2 toques: importe → categoría)
   ===================================================================== */
let add = null;
const amountToStr = n => { const r = Math.round(n * 100) / 100; return (Number.isInteger(r) ? String(r) : r.toFixed(2)).replace('.', ','); };
const addAmountNum = () => parseFloat((add?.amount || '0').replace(',', '.')) || 0;

function openAdd(pre = {}) {
  add = {
    amount: pre.amount > 0 ? amountToStr(pre.amount) : '',
    cat: pre.cat || null, note: pre.note || '', date: pre.date || todayISO(),
    editId: pre.editId || null, source: pre.source || 'manual', wishId: pre.wishId || null,
  };
  openSheet(addHTML(), 'add');
  updateAmt(); updateSuggest();
}

function addHTML() {
  const t = todayISO(), y = addDays(t, -1);
  const custom = add.date !== t && add.date !== y;
  return `
    <div class="sheet-head">
      <div class="add-tools">
        ${add.editId ? '<b>Editar gasto</b>' : '<button class="pill ai-pill" data-action="aiOpen">✨ Asistente</button><button class="pill" data-action="capture">📸 Captura</button><button class="pill" data-action="pasteText">📋 Pegar</button>'}
      </div>
      <button class="icon-btn" data-action="closeSheet" aria-label="Cerrar">✕</button>
    </div>
    <div class="amt-wrap" id="amtWrap"><span class="amt" id="amt">0</span><span class="amt-cur">€</span><div class="hours" id="hours"></div></div>
    <input id="note" class="note-input" placeholder="¿Dónde o qué? (opcional)" value="${esc(add.note)}" autocomplete="off" enterkeyhint="done">
    <div class="date-chips">
      <button class="chip ${add.date === t ? 'on' : ''}" data-action="setDate" data-d="${t}">Hoy</button>
      <button class="chip ${add.date === y ? 'on' : ''}" data-action="setDate" data-d="${y}">Ayer</button>
      <label class="chip ${custom ? 'on' : ''}">📅 ${custom ? niceDate(add.date) : 'Otro día'}<input type="date" id="addDate" value="${add.date}" max="${t}"></label>
    </div>
    <div class="keypad">
      ${['1', '2', '3', '4', '5', '6', '7', '8', '9', ',', '0', 'del'].map(k => `<button class="key" data-action="key" data-k="${k}" aria-label="${k === 'del' ? 'Borrar' : k}">${k === 'del' ? '⌫' : k}</button>`).join('')}
    </div>
    <div class="cat-label">${add.editId ? 'Toca una categoría para guardar los cambios' : 'Toca una categoría y listo ✓'}</div>
    <div class="catgrid">
      ${CATS.map(c => `<button class="cat ${add.cat === c.id ? 'selected' : ''}" data-action="pickCat" data-cat="${c.id}"><span class="e">${c.emoji}</span>${c.short}</button>`).join('')}
    </div>
    ${add.editId ? `<div class="btn-row" style="margin-top:14px"><button class="btn danger" data-action="deleteExpense">🗑️ Borrar gasto</button></div>` : ''}`;
}

function updateAmt() {
  const el = $('#amt'); if (!el || !add) return;
  el.textContent = add.amount || '0';
  $('#amtWrap').classList.toggle('empty', !add.amount);
  $('#hours').textContent = hoursText(addAmountNum());
}
function hoursText(n) {
  const s = state.settings;
  if (!s.income || !n) return '';
  const mins = Math.round(n / (s.income / (s.hoursMonth || 160)) * 60);
  if (mins < 1) return '';
  const txt = mins >= 60 ? `${Math.floor(mins / 60)} h${mins % 60 ? ' ' + (mins % 60) + ' min' : ''}` : `${mins} min`;
  return `≈ ${txt} de tu trabajo`;
}
function updateSuggest() {
  if (!add) return;
  const sug = add.editId ? null : autoCat($('#note')?.value || '');
  $$('.catgrid .cat').forEach(b => b.classList.toggle('suggested', b.dataset.cat === sug && add.cat !== sug));
}
function keyPress(k) {
  if (!add) return;
  let a = add.amount;
  if (k === 'del') a = a.slice(0, -1);
  else if (k === ',') { if (!a.includes(',')) a = (a || '0') + ','; }
  else {
    if (a.includes(',')) { if (a.split(',')[1].length >= 2) return; }
    else if (a.replace(/^0+/, '').length >= 6) return;
    a = a === '0' ? k : a + k;
  }
  add.amount = a;
  updateAmt();
}

function pickCat(cat) {
  const amount = addAmountNum();
  if (!(amount > 0)) { shake('#amtWrap'); toast('Primero escribe el importe 👆'); return; }
  const note = ($('#note')?.value || '').trim();
  const { date, editId, source, wishId } = add;
  if (note) learn(note, cat);
  if (editId) {
    const e = state.expenses.find(x => x.id === editId);
    const before = { ...e };
    Object.assign(e, { amount, cat, note, date });
    save(); closeSheet(); render();
    toast('Cambios guardados', { undo: () => Object.assign(e, before) });
    return;
  }
  const firstToday = date === todayISO() && !loggedDays().has(date);
  const e = { id: uid(), amount, cat, note, date, ts: Date.now(), source };
  state.expenses.push(e);
  const wish = wishId && state.wishlist.find(x => x.id === wishId);
  if (wish) { wish.status = 'bought'; wish.decided = Date.now(); }
  save(); closeSheet(); render();
  const c = catOf(cat);
  toast(`${money(amount)} en ${c.emoji} ${c.short}${firstToday ? ` · racha ${streak()} 🔥` : ''}`, {
    undo: () => { state.expenses = state.expenses.filter(x => x.id !== e.id); if (wish) wish.status = 'waiting'; },
  });
  checkBadges();
}

function deleteExpense() {
  const e = state.expenses.find(x => x.id === add.editId);
  if (!e) return;
  state.expenses = state.expenses.filter(x => x !== e);
  save(); closeSheet(); render();
  toast('Gasto borrado', { undo: () => state.expenses.push(e) });
}

/* =====================================================================
   Capturas (OCR en el propio móvil), texto pegado e importación del banco
   ===================================================================== */
const AMT_RE = /([-−–+]\s?)?(?:€\s?)?(\d{1,3}(?:[.\s]\d{3})+|\d+)[,.](\d{2})(?!\d)(?!\s?%)(\s?(?:€|eur))?/gi;
const DATE_STRIP = [/\b\d{4}-\d{2}-\d{2}\b/g, /\b\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\b/g, /\b\d{1,2}\/\d{1,2}\b/g, /\b\d{1,2}:\d{2}(:\d{2})?\b/g];

/** Convierte texto (de una captura, notificación o email) en posibles gastos. */
function textToItems(text) {
  const lines = String(text).split(/\r?\n/).map(l => l.replace(/\s+/g, ' ').trim()).filter(l => l.length > 1);
  const entries = [];
  let lastDate = null, lastText = '', firstDate = null;
  for (const line of lines) {
    const d = parseDateLoose(line);
    if (d) { lastDate = d; firstDate = firstDate || d; }
    let clean = line;
    DATE_STRIP.forEach(re => { clean = clean.replace(re, ' '); });
    clean = clean.replace(new RegExp(MONTH_DATE_RE.source, 'gi'), ' ');
    const ms = [...clean.matchAll(AMT_RE)];
    const textPart = s => s.replace(/[^\p{L}\d&'’ .-]/gu, ' ').replace(/\s+/g, ' ').trim();
    if (!ms.length) {
      const txt = textPart(clean);
      if (/\p{L}{3,}/u.test(txt)) lastText = txt;
      continue;
    }
    const m = ms.find(x => x[1] || x[4]) || ms[ms.length - 1];
    const val = parseFloat(m[2].replace(/[.\s]/g, '') + '.' + m[3]);
    let txt = textPart(clean.replace(AMT_RE, ' '));
    const n = norm(line);
    // El OCR a veces se come el "−": sin signo = gasto. Ingreso solo con "+" o palabras claras.
    const income = (m[1] || '').trim() === '+' || /nomina|ingreso|abono(?! (de )?transporte)|transferencia recibida|bizum recibido|devolucion|reembolso/.test(norm(line + ' ' + lastText));
    const isTotal = /\btotal\b/.test(n) && !/sub ?total|base|iva/.test(n);
    const isNoise = /saldo|disponible|\biva\b|base imp|cambio|entregado|limite/.test(n);
    if (!/\p{L}{3,}/u.test(txt)) txt = lastText;
    entries.push({ amount: income ? val : -val, income, note: txt, date: lastDate, isTotal, isNoise });
    lastText = '';
  }

  // ¿Es un ticket? → un único gasto con el TOTAL
  const totals = entries.filter(e => e.isTotal);
  if (totals.length) {
    const tot = totals.reduce((a, e) => (Math.abs(e.amount) > Math.abs(a.amount) ? e : a));
    const merchant = lines.find(l => /\p{L}{3,}/u.test(l) && !/ticket|factura|simplificada|nif|cif|c\.i\.f|tel[eé]f|www|fecha|hora/i.test(l)) || '';
    return finalizeItems([{ amount: -Math.abs(tot.amount), note: cleanNote(merchant), date: firstDate }]);
  }
  let raw = entries.filter(e => !e.isNoise).map(e => ({ amount: e.amount, income: e.income, note: cleanNote(e.note), date: e.date }));

  // Último recurso: "12 € Mercadona", "Mercadona 12€"
  if (!raw.length) {
    const m = String(text).match(/(-)?(\d+(?:[.,]\d{1,2})?)\s?(?:€|eur)|€\s?(\d+(?:[.,]\d{1,2})?)/i);
    if (m) {
      const amount = parseAmount(m[2] || m[3]);
      const note = cleanNote(String(text).replace(m[0], ' ').replace(/[^\p{L}\d&'’ .-]/gu, ' '));
      raw = [{ amount: -Math.abs(amount), note, date: firstDate }];
    } else {
      const only = String(text).trim().match(/^(\d+(?:[.,]\d{1,2})?)\s+(.+)$|^(.+?)\s+(\d+(?:[.,]\d{1,2})?)$/);
      if (only) raw = [{ amount: -Math.abs(parseAmount(only[1] || only[4])), note: cleanNote(only[2] || only[3]), date: null }];
    }
  }
  return finalizeItems(raw);
}

/** Separa ingresos, marca posibles duplicados y propone categoría. */
function finalizeItems(raw) {
  raw = raw.filter(x => isFinite(x.amount) && x.amount !== 0);
  const hasNeg = raw.some(x => x.amount < 0);
  const existing = new Set(state.expenses.map(e => e.date + '|' + e.amount.toFixed(2)));
  let incomes = 0;
  const items = [];
  for (const x of raw) {
    if (x.income || (hasNeg && x.amount > 0)) { incomes++; continue; }
    const amount = Math.abs(x.amount), date = x.date || todayISO();
    const dup = existing.has(date + '|' + amount.toFixed(2));
    items.push({ amount, note: x.note || '', date, cat: autoCat(x.note) || 'otros', dup, checked: !dup });
  }
  return { items, incomes };
}

function loadingSheet(msg) {
  openSheet(`${sheetHead('')}<div class="ob"><div class="ob-emoji">🔎</div><h2>${msg}</h2>
    <p class="muted" id="loadMsg">Se hace en tu propio dispositivo: la imagen no se sube a ningún sitio.</p>
    <div class="progress"><i id="loadBar"></i></div></div>`, 'loading');
}

async function runOCR(file) {
  loadingSheet('Leyendo la captura…');
  try {
    await loadScript(TESSERACT_URL);
    const worker = await Tesseract.createWorker('spa', 1, {
      logger: m => {
        const bar = $('#loadBar');
        if (!bar) return;
        if (m.status === 'recognizing text') bar.style.width = Math.round(m.progress * 100) + '%';
        else if (/loading|initializ/.test(m.status)) { bar.style.width = '8%'; const lm = $('#loadMsg'); if (lm) lm.textContent = 'La primera vez descarga el "diccionario" (unos segundos)…'; }
      },
    });
    const { data } = await worker.recognize(file);
    await worker.terminate();
    if (sheetKind !== 'loading') return; // lo ha cancelado
    handleParsed(textToItems(data.text), 'ocr');
  } catch (err) {
    if (sheetKind !== 'loading') return;
    setSheet(`${sheetHead('')}<div class="ob"><div class="ob-emoji">😕</div><h2>No he podido leerla</h2>
      <p class="muted">La primera vez hace falta conexión para descargar el lector. También puedes escribir el importe a mano.</p>
      <button class="btn primary w100" data-action="openAddBlank">Apuntar a mano</button></div>`);
  }
}

function handleParsed({ items, incomes }, source) {
  if (!items.length) {
    closeSheet();
    openAdd({ source });
    toast(incomes ? 'Solo he visto ingresos, no gastos' : 'No he encontrado importes. Escríbelo tú 🙂');
    return;
  }
  if (items.length === 1 && source !== 'csv') {
    const it = items[0];
    closeSheet();
    openAdd({ amount: it.amount, note: it.note, date: it.date, source });
    if (it.cat && it.cat !== 'otros') $$('.catgrid .cat').forEach(b => b.classList.toggle('suggested', b.dataset.cat === it.cat));
    toast('Revisa el importe y toca la categoría');
    return;
  }
  openReview(items, source, incomes);
}

/* ---------- Revisión de varios movimientos ---------- */
let review = null;
function openReview(items, source, incomes = 0) {
  review = { items: items.slice(0, 600), source, incomes };
  openSheet(reviewHTML(), 'review');
}
function reviewHTML() {
  const n = review.items.filter(i => i.checked).length;
  const dups = review.items.filter(i => i.dup).length;
  return `${sheetHead(`He encontrado ${review.items.length} ${review.items.length === 1 ? 'gasto' : 'gastos'}`)}
    ${review.source === 'ia' ? '<button class="btn link" data-action="aiOpen" style="margin:-6px 0 6px;padding-left:0">← Volver a la conversación</button>' : ''}
    <p class="muted small">Revisa y corrige lo que haga falta. Solo se guardan los marcados.
    ${review.incomes ? `He ignorado ${review.incomes} ${review.incomes === 1 ? 'ingreso' : 'ingresos'}.` : ''} ${dups ? `<b>${dups} parecen repetidos</b> (ya estaban apuntados) y los he desmarcado.` : ''}</p>
    <div class="btn-row" style="margin-bottom:10px"><button class="btn sm ghost" data-action="rvAll" data-v="1">Marcar todos</button><button class="btn sm ghost" data-action="rvAll" data-v="0">Ninguno</button></div>
    ${review.items.map((it, i) => `<div class="rv ${it.checked ? '' : 'off'}" data-i="${i}">
      <div class="rv-top"><input type="checkbox" data-rv="checked" ${it.checked ? 'checked' : ''} aria-label="Guardar este gasto">
        <input class="rv-note" data-rv="note" value="${esc(it.note)}" placeholder="Concepto">
        <input class="rv-amt" data-rv="amount" inputmode="decimal" value="${amountToStr(it.amount)}" aria-label="Importe"></div>
      <div class="rv-bottom"><select data-rv="cat" aria-label="Categoría">${CATS.map(c => `<option value="${c.id}" ${c.id === it.cat ? 'selected' : ''}>${c.emoji} ${c.name}</option>`).join('')}</select>
        <input type="date" data-rv="date" value="${it.date}" aria-label="Fecha"></div>
      ${it.dup ? '<div class="dup">⚠️ Puede que ya esté apuntado</div>' : ''}
    </div>`).join('')}
    <div class="sticky-save"><button class="btn primary w100" data-action="rvSave" ${n ? '' : 'disabled'}>Guardar ${n} ${n === 1 ? 'gasto' : 'gastos'}</button></div>`;
}
function reviewSave() {
  const chosen = review.items.filter(i => i.checked && i.amount > 0);
  const ids = [];
  for (const it of chosen) {
    const e = { id: uid(), amount: it.amount, cat: it.cat, note: it.note, date: it.date || todayISO(), ts: Date.now(), source: review.source };
    state.expenses.push(e); ids.push(e.id);
    if (it.note) learn(it.note, it.cat);
  }
  if (review.source === 'ia') ai = null;   // conversación terminada
  save(); closeSheet(); render();
  toast(`${chosen.length} ${chosen.length === 1 ? 'gasto guardado' : 'gastos guardados'}`, { undo: () => { state.expenses = state.expenses.filter(e => !ids.includes(e.id)); } });
  checkBadges();
}

/* ---------- Pegar texto ---------- */
async function pasteText() {
  let text = '';
  try { if (navigator.clipboard?.readText) text = await navigator.clipboard.readText(); } catch (e) { /* sin permiso */ }
  if (text && text.trim()) {
    const res = textToItems(text);
    if (res.items.length) { handleParsed(res, 'paste'); return; }
  }
  openSheet(`${sheetHead('Pegar texto')}
    <p class="muted small">Pega una notificación del banco, un email de compra o una lista de movimientos. Yo saco los importes.</p>
    <label class="field"><textarea id="pasteArea" placeholder="Ej.: Compra en MERCADONA por 23,45 €">${esc(text)}</textarea></label>
    <button class="btn primary w100" data-action="pasteGo">Buscar gastos</button>`, 'paste');
}

/* ---------- Importar movimientos del banco (CSV / Excel) ---------- */
async function readTextSmart(file) {
  const buf = await file.arrayBuffer();
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch (e) { return new TextDecoder('windows-1252').decode(buf); }
}
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const head = text.split(/\r?\n/).slice(0, 20).join('\n');
  const delim = [';', ',', '\t'].map(d => [d, head.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = [];
  let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cell); cell = ''; }
    else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
    else if (ch !== '\r') cell += ch;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows.map(r => r.map(c => c.trim())).filter(r => r.some(c => c));
}
function cellDate(c) {
  if (c instanceof Date && !isNaN(c)) return toISO(new Date(c.getTime() + 12 * 3600e3));
  if (typeof c === 'number' && c > 30000 && c < 80000) {
    const d = new Date(Math.round((c - 25569) * 86400e3));
    return validISO(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  return parseDateStrict(c);
}
const cellNum = c => (typeof c === 'number' ? c : parseAmount(String(c ?? '')));

function rowsToItems(rows) {
  const N = c => norm(String(c ?? '')).trim();
  let hi = -1, cols = null;
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const h = rows[i].map(N);
    const find = re => h.findIndex(x => re.test(x));
    let date = find(/fecha.*oper|f\.? ?oper/);
    if (date < 0) date = find(/fecha|^date|^f\.? ?valor/);
    const amount = h.findIndex(x => /importe|amount|cantidad|^monto/.test(x) && !/saldo|balance/.test(x));
    const debit = find(/^cargo|^debe|debit|^gasto/), credit = find(/^abono|^haber|credit|^ingreso/);
    const desc = find(/concepto|descrip|detalle|movimiento|comercio|description|payee|beneficiario|establecimiento/);
    if (date >= 0 && (amount >= 0 || debit >= 0)) { hi = i; cols = { date, amount, debit, credit, desc }; break; }
  }
  if (!cols) cols = guessCols(rows);
  if (!cols) return null;
  const raw = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const date = cellDate(r[cols.date]);
    if (!date) continue;
    let amt;
    if (cols.amount >= 0) amt = cellNum(r[cols.amount]);
    else {
      const d = cellNum(r[cols.debit]), c = cols.credit >= 0 ? cellNum(r[cols.credit]) : NaN;
      amt = isFinite(d) && d !== 0 ? -Math.abs(d) : (isFinite(c) ? Math.abs(c) : NaN);
    }
    if (!isFinite(amt) || amt === 0) continue;
    const note = cols.desc >= 0 ? String(r[cols.desc] ?? '')
      : r.filter((c, j) => j !== cols.date && j !== cols.amount && typeof c === 'string' && !/^[-\d.,\s€]+$/.test(c)).sort((a, b) => b.length - a.length)[0] || '';
    raw.push({ date, amount: amt, note: cleanNote(note) });
  }
  return finalizeItems(raw);
}
function guessCols(rows) {
  const sample = rows.slice(0, 60), width = Math.max(...sample.map(r => r.length));
  let best = { date: -1, amount: -1, desc: -1 }, ds = 0, as = 0, ls = 0;
  for (let j = 0; j < width; j++) {
    const vals = sample.map(r => r[j]).filter(v => v !== '' && v != null);
    if (!vals.length) continue;
    const dScore = vals.filter(v => cellDate(v)).length / vals.length;
    const aScore = vals.filter(v => typeof v === 'number' || /^[-+]?[\d.,]+\s?€?$/.test(String(v).trim()) && /[.,]\d{2}\s?€?$/.test(String(v).trim())).length / vals.length;
    const len = vals.reduce((a, v) => a + String(v).length, 0) / vals.length;
    if (dScore > 0.6 && dScore > ds) { ds = dScore; best.date = j; }
    else if (aScore > 0.6 && aScore > as && best.amount < 0) { as = aScore; best.amount = j; }
    else if (len > ls && dScore < 0.3 && aScore < 0.3) { ls = len; best.desc = j; }
  }
  return best.date >= 0 && best.amount >= 0 ? { ...best, debit: -1, credit: -1 } : null;
}

async function importBank(file) {
  loadingSheet('Leyendo tus movimientos…');
  try {
    let rows;
    if (/\.xlsx?$/i.test(file.name)) {
      await loadScript(XLSX_URL);
      const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' })
        .map(r => r.map(c => (typeof c === 'string' ? c.trim() : c)))
        .filter(r => r.some(c => c !== ''));
    } else {
      rows = parseCSV(await readTextSmart(file));
    }
    const res = rowsToItems(rows);
    if (!res || !res.items.length) throw new Error('vacío');
    openReview(res.items, 'csv', res.incomes);
  } catch (err) {
    setSheet(`${sheetHead('')}<div class="ob"><div class="ob-emoji">🤔</div><h2>No entiendo este archivo</h2>
      <p class="muted">Necesito un CSV o Excel con columnas de <b>fecha</b>, <b>concepto</b> e <b>importe</b>. Descárgalo desde la web de tu banco (Movimientos → Exportar).</p></div>`);
  }
}

/* =====================================================================
   Ajustes, gastos fijos y copias de seguridad
   ===================================================================== */
function settingsHTML() {
  const s = state.settings;
  const opts = id => CATS.map(c => `<option value="${c.id}" ${c.id === id ? 'selected' : ''}>${c.emoji} ${c.name}</option>`).join('');
  return `${sheetHead('Ajustes')}
    <div class="section-title">Tu mes</div>
    <div class="grid2">
      <label class="field"><span>Ingresos al mes</span><input data-setting="income" inputmode="decimal" placeholder="1500" value="${s.income || ''}"></label>
      <label class="field"><span>Quiero ahorrar</span><input data-setting="saveGoal" inputmode="decimal" placeholder="150" value="${s.saveGoal || ''}"></label>
      <label class="field"><span>Día de cobro</span><input data-setting="startDay" inputmode="numeric" placeholder="1" value="${s.startDay || 1}"></label>
      <label class="field"><span>Horas de trabajo/mes</span><input data-setting="hoursMonth" inputmode="numeric" placeholder="160" value="${s.hoursMonth || 160}"></label>
    </div>
    <p class="muted small">Presupuesto para gastar = ingresos − ahorro. Los gastos fijos salen de ahí.</p>

    <div class="section-title">Gastos fijos (se apuntan solos)</div>
    <ul class="list-plain">${state.recurring.map(r => `<li><span style="font-size:22px">${catOf(r.cat).emoji}</span><span class="grow"><b>${esc(r.name)}</b><br><span class="muted small">Día ${r.day} · ${money(r.amount)}</span></span><button class="del" data-action="delRecurring" data-id="${r.id}" aria-label="Quitar">✕</button></li>`).join('') || '<li class="muted small">Aún no tienes ninguno. Alquiler, móvil, gimnasio, Netflix…</li>'}</ul>
    <div class="grid2">
      <label class="field"><span>Nombre</span><input id="rName" placeholder="Alquiler"></label>
      <label class="field"><span>Importe</span><input id="rAmount" inputmode="decimal" placeholder="650"></label>
      <label class="field"><span>Día del mes</span><input id="rDay" inputmode="numeric" placeholder="1"></label>
      <label class="field"><span>Categoría</span><select id="rCat">${opts('casa')}</select></label>
    </div>
    <button class="btn ghost w100" data-action="addRecurring">＋ Añadir gasto fijo</button>
    <p class="muted small">Se empieza a apuntar a partir de hoy (no duplica lo de este mes si ya ha pasado su día).</p>

    <div class="section-title">Importar del banco</div>
    <p class="muted small">Descarga tus movimientos en CSV o Excel desde la web de tu banco y súbelos aquí. Detecto repetidos para que no se dupliquen.</p>
    <button class="btn ghost w100" data-action="importBank">🏦 Importar movimientos (CSV / Excel)</button>

    <div class="section-title">Asistente con IA ✨</div>
    ${aiCfg.key
      ? `<p class="muted small">Activo en este dispositivo. Gastado hasta ahora: <b>${aiSpentText()}</b> (cálculo aproximado; el real lo ves en <a href="${AI_CONSOLE}/settings/billing" target="_blank" rel="noopener">platform.claude.com</a>).</p>
        <div class="btn-row"><button class="btn ghost" data-action="aiOpen">✨ Abrir asistente</button><button class="btn ghost" data-action="aiForget">Quitar llave</button></div>`
      : `<p class="muted small">Mándale capturas del banco y explícale de palabra qué es cada cosa; te pregunta lo que no sepa. Cuesta céntimos al mes.</p>
        <button class="btn ghost w100" data-action="aiOpen">✨ Activar asistente</button>`}

    ${syncSettingsBlock()}

    <div class="section-title">Copias de seguridad</div>
    <p class="muted small">${sync.cfg ? 'Tus datos están en tus dispositivos sincronizados. Aun así, una copia de vez en cuando no sobra.' : 'Tus datos solo están en este dispositivo. Guarda una copia en iCloud Drive / Archivos de vez en cuando.'}${s.lastBackup ? ` Última: ${niceDate(s.lastBackup).toLowerCase()}.` : ''}</p>
    <div class="btn-row">
      <button class="btn primary" data-action="exportBackup">💾 Guardar copia</button>
      <button class="btn ghost" data-action="importBackup">📂 Restaurar</button>
    </div>
    <button class="btn link" data-action="exportCSV">Exportar gastos a Excel (CSV)</button>

    <div class="section-title">Aspecto</div>
    <div class="btn-row">${[['auto', 'Automático'], ['light', 'Claro'], ['dark', 'Oscuro']].map(([v, l]) => `<button class="btn sm ${(s.theme || 'auto') === v ? 'primary' : 'ghost'}" data-action="theme" data-v="${v}">${l}</button>`).join('')}</div>

    <div class="section-title">Zona delicada</div>
    <button class="btn danger w100" data-action="wipe">Borrar todos mis datos</button>
    <p class="muted small" style="text-align:center;margin-top:16px">Hucha · ${sync.cfg ? 'tus datos viajan cifrados entre tus dispositivos 🔒' : 'tus datos no salen de este dispositivo 🔒'}</p>`;
}

function addRecurring() {
  const name = $('#rName').value.trim(), amount = parseAmount($('#rAmount').value), day = parseInt($('#rDay').value, 10) || 1, cat = $('#rCat').value;
  if (!name || !(amount > 0)) { toast('Pon nombre e importe'); return; }
  state.recurring.push({ id: uid(), name, amount: Math.abs(amount), cat, day: Math.min(31, Math.max(1, day)), created: todayISO() });
  applyRecurring(); save(); setSheet(settingsHTML()); render();
  toast(`${name} se apuntará solo cada mes`);
}

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
async function shareOrDownload(blob, name, type) {
  const file = new File([blob], name, { type });
  if (isTouch() && navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return true; }
    catch (e) { if (e.name === 'AbortError') return false; }
  }
  download(blob, name);
  return true;
}
async function exportBackup() {
  const ok = await shareOrDownload(new Blob([JSON.stringify(state)], { type: 'application/json' }), `hucha-copia-${todayISO()}.json`, 'application/json');
  if (!ok) return;
  state.settings.lastBackup = todayISO();
  save(); render();
  if (sheetKind === 'settings') setSheet(settingsHTML());
  toast('Copia guardada 💾');
  checkBadges();
}
async function importBackupFile(file) {
  try {
    const data = JSON.parse(await file.text());
    if (!data || !Array.isArray(data.expenses)) throw new Error('formato');
    if (!confirm(`Esta copia tiene ${data.expenses.length} gastos. Se sustituirán tus datos actuales. ¿Seguro?`)) return;
    const d = DEFAULT();
    state = { ...d, ...data, settings: { ...d.settings, ...(data.settings || {}), onboarded: true } };
    save(); applyTheme(); closeSheet(); render();
    toast('Copia restaurada ✅');
  } catch (e) { toast('Ese archivo no es una copia de Hucha'); }
}
function exportCSV() {
  const rows = [['Fecha', 'Categoría', 'Concepto', 'Importe', 'Origen']].concat(
    [...state.expenses].sort((a, b) => a.date.localeCompare(b.date)).map(e => [e.date, catOf(e.cat).name, e.note || '', e.amount.toFixed(2).replace('.', ','), { manual: 'A mano', ocr: 'Captura', paste: 'Texto pegado', csv: 'Banco', fijo: 'Gasto fijo', wish: 'Lista de espera' }[e.source] || e.source]));
  const csv = '﻿' + rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(';')).join('\r\n');
  shareOrDownload(new Blob([csv], { type: 'text/csv' }), `hucha-gastos-${todayISO()}.csv`, 'text/csv');
}
function applyTheme() {
  const t = state.settings.theme || 'auto';
  if (t === 'auto') document.documentElement.removeAttribute('data-theme');
  else document.documentElement.setAttribute('data-theme', t);
}

/* =====================================================================
   Metas y lista de espera (hojas)
   ===================================================================== */
let goalDraft = null;
function goalSheetHTML() {
  const g = goalDraft;
  return `${sheetHead(g.id ? 'Editar meta' : 'Nueva meta')}
    <div class="emoji-pick">${GOAL_EMOJIS.map(e => `<button data-action="goalEmoji" data-v="${e}" class="${g.emoji === e ? 'on' : ''}">${e}</button>`).join('')}</div>
    <label class="field"><span>¿Para qué ahorras?</span><input id="gName" value="${esc(g.name)}" placeholder="Colchón de emergencia"></label>
    <label class="field"><span>¿Cuánto necesitas?</span><input id="gTarget" inputmode="decimal" value="${g.target ? amountToStr(g.target) : ''}" placeholder="1000"></label>
    ${g.id ? `<label class="field"><span>Llevas ahorrado</span><input id="gSaved" inputmode="decimal" value="${amountToStr(g.saved)}"></label>` : ''}
    <button class="btn primary w100" data-action="goalSave">${g.id ? 'Guardar' : 'Crear meta'}</button>
    ${g.id ? '<button class="btn link" data-action="goalDelete" style="color:var(--warn)">Borrar meta</button>' : ''}`;
}
function goalAdd(g, n) {
  const before = g.saved;
  const oldPct = g.saved / g.target;
  g.saved = Math.round((g.saved + n) * 100) / 100;
  const newPct = g.saved / g.target;
  save(); render();
  const crossed = [0.25, 0.5, 0.75, 1].find(m => oldPct < m && newPct >= m);
  if (crossed) {
    confetti();
    toast(crossed === 1 ? `🥳 ¡Meta "${g.name}" conseguida!` : `¡${Math.round(crossed * 100)} % de "${g.name}"! Sigue así`, { undo: () => { g.saved = before; } });
  } else toast(`+${money(n)} a ${g.emoji} ${g.name}`, { undo: () => { g.saved = before; } });
  checkBadges();
}
function askNumber(title, onOk) {
  askNumber.cb = onOk;
  openSheet(`${sheetHead(title)}<div class="big-input"><input id="askNum" inputmode="decimal" placeholder="0"><span>€</span></div>
    <button class="btn primary w100" data-action="askOk">Añadir</button>`, 'ask');
  setTimeout(() => $('#askNum')?.focus(), 50);
}
let wishWait = 48;
function wishSheetHTML() {
  return `${sheetHead('Nuevo antojo')}
    <p class="muted small">Apúntalo aquí en lugar de comprarlo ahora. Te aviso cuando pase el tiempo para que decidas con calma.</p>
    <label class="field"><span>¿Qué quieres comprar?</span><input id="wName" placeholder="Zapatillas nuevas"></label>
    <label class="field"><span>¿Cuánto cuesta?</span><input id="wPrice" inputmode="decimal" placeholder="80"></label>
    <div class="field"><span>¿Cuánto esperamos?</span><div class="btn-row">${[[24, '24 h'], [48, '48 h'], [168, '1 semana'], [720, '30 días']].map(([h, l]) => `<button class="btn sm ${wishWait === h ? 'primary' : 'ghost'}" data-action="wishWait" data-v="${h}">${l}</button>`).join('')}</div></div>
    <button class="btn primary w100" data-action="wishSave">Dejar en espera ⏳</button>`;
}

/* =====================================================================
   Bienvenida (3 preguntas)
   ===================================================================== */
let ob = null;
function obHTML() {
  const s = state.settings;
  const steps = [
    () => `<div class="ob"><div class="ob-emoji">🐷</div><h2>¡Hola! Soy tu Hucha</h2>
      <p>Apuntar un gasto te llevará <b>2 toques</b>. Yo hago las cuentas, busco dónde puedes ahorrar y celebro tus avances.</p>
      <p class="muted small">🔒 Todo se guarda en tu dispositivo. Si quieres usarla también en otro, puedes sincronizarla cifrada.</p>
      <button class="btn primary w100" data-action="obNext">Empezar (30 segundos)</button>
      <button class="btn link" data-action="syncJoin">¿Ya usas Hucha en otro dispositivo? Vincular</button></div>`,
    () => `<div class="ob"><div class="ob-step">1 de 3</div><h2>¿Cuánto cobras al mes?</h2>
      <p class="muted">Más o menos. Sirve para calcular cuánto puedes gastar cada día.</p>
      <div class="big-input"><input id="obIncome" inputmode="decimal" placeholder="1500" value="${s.income || ''}"><span>€</span></div>
      <button class="btn primary w100" data-action="obNext">Siguiente</button>
      <button class="btn link" data-action="obSkip">Prefiero no decirlo ahora</button></div>`,
    () => {
      const inc = s.income || 0, v = s.saveGoal || 0;
      return `<div class="ob"><div class="ob-step">2 de 3</div><h2>¿Cuánto quieres ahorrar al mes?</h2>
        <p class="muted">Lo apartamos primero, como si fuera un gasto fijo. Puedes cambiarlo cuando quieras.</p>
        <div class="big-input"><input id="obSave" inputmode="decimal" placeholder="0" value="${v || ''}" aria-label="Ahorro al mes"><span>€</span></div>
        ${inc ? `<input type="range" id="obSlider" class="slider" min="0" max="${saveSliderMax(inc)}" step="5" value="${Math.min(v, saveSliderMax(inc))}" aria-label="Ahorro al mes">
          <p class="save-info" id="obSaveInfo">${saveInfo(v, inc)}</p>
          <div class="chips" id="obChips">${savePresets(inc).map(([l, x]) => `<button class="chip ${x === v ? 'on' : ''}" data-action="obSave" data-v="${x}">${l}<small>${moneyR(x)}</small></button>`).join('')}</div>
          <p class="muted small">Escribe la cantidad que quieras o mueve la barra. Los botones son solo atajos.</p>`
        : '<p class="muted small">Escribe la cantidad que quieras. Puede ser 0.</p>'}
        <button class="btn primary w100" data-action="obNext">Siguiente</button></div>`;
    },
    () => `<div class="ob"><div class="ob-step">3 de 3</div><h2>¿Qué día cobras?</h2>
      <p class="muted">Tu "mes" empezará ese día. Desliza la ruleta.</p>
      ${dayWheelHTML(s.startDay || 1, 'ob')}
      <p class="muted small" id="wheelInfo">${payDayInfo(s.startDay || 1)}</p>
      <button class="btn primary w100" data-action="obFinish">¡Listo!</button></div>`,
  ];
  return steps[ob.step]();
}
/* ---------- Ahorro: barra + cantidad libre ---------- */
const saveSliderMax = inc => Math.max(50, Math.round(inc * 0.5 / 10) * 10);
const savePresets = inc => [['Nada', 0], ['5 %', .05], ['10 %', .1], ['20 %', .2]].map(([l, f]) => [l, Math.round(inc * f / 10) * 10]);
function saveInfo(v, inc) {
  if (!inc) return '';
  if (v >= inc) return '⚠️ Es todo lo que cobras: no te quedaría nada para vivir.';
  return `<b>${Math.round(v / inc * 100)} %</b> de lo que cobras · te quedan <b>${moneyR(inc - v)}</b> al mes para tus gastos`;
}
function obSaveChanged(v, from) {
  v = Math.max(0, Math.round((v || 0) * 100) / 100);
  const inc = state.settings.income || 0;
  const input = $('#obSave'), slider = $('#obSlider'), info = $('#obSaveInfo');
  if (input && from !== 'input') input.value = v || (from === 'chip' ? '0' : '');
  if (slider && from !== 'slider') slider.value = Math.min(v, Number(slider.max));
  if (info) info.innerHTML = saveInfo(v, inc);
  $$('#obChips .chip').forEach(c => c.classList.toggle('on', Number(c.dataset.v) === v));
}

/* ---------- Ruleta de días (1–31) ---------- */
function dayWheelHTML(v, kind) {
  return `<div class="wheel">
    <div class="wheel-band" aria-hidden="true"><span>día</span></div>
    <div class="wheel-list" data-wheel="${kind}" data-value="${v}" tabindex="0" role="listbox" aria-label="Día de cobro">
      <div class="wheel-pad"></div>
      ${Array.from({ length: 31 }, (_, i) => `<div class="wheel-item" role="option" data-day="${i + 1}">${i + 1}</div>`).join('')}
      <div class="wheel-pad"></div>
    </div></div>`;
}
function payDayInfo(d) {
  if (d === 1) return 'Tu mes irá del día 1 al último día de cada mes.';
  return `Tu mes irá del día ${d} al ${d - 1} del mes siguiente.${d > 28 ? ` Si un mes no tiene día ${d}, empieza el último día de ese mes.` : ''}`;
}
function initWheels() {
  $$('[data-wheel]').forEach(list => {
    if (list.dataset.ready) return;
    list.dataset.ready = '1';
    const items = $$('.wheel-item', list);
    const h = () => items[0].offsetHeight || 44;
    const mark = d => {
      items.forEach(it => { const on = Number(it.dataset.day) === d; it.classList.toggle('on', on); it.setAttribute('aria-selected', on); });
    };
    const start = Number(list.dataset.value) || 1;
    list.scrollTop = (start - 1) * h();
    mark(start);
    let raf = 0;
    list.addEventListener('scroll', () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        const d = Math.min(31, Math.max(1, Math.round(list.scrollTop / h()) + 1));
        if (d === Number(list.dataset.value)) return;
        list.dataset.value = d; mark(d);
        if (list.dataset.wheel === 'ob') { state.settings.startDay = d; const i = $('#wheelInfo'); if (i) i.textContent = payDayInfo(d); }
      });
    });
    list.addEventListener('click', e => {
      const it = e.target.closest('.wheel-item');
      if (it) list.scrollTo({ top: (Number(it.dataset.day) - 1) * h(), behavior: 'smooth' });
    });
    list.addEventListener('keydown', e => {
      if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
      e.preventDefault();
      list.scrollBy({ top: e.key === 'ArrowUp' ? -h() : h(), behavior: 'smooth' });
    });
  });
}

function obReadInputs() {
  const inc = $('#obIncome'), sv = $('#obSave');
  if (inc) { const v = parseAmount(inc.value); state.settings.income = v > 0 ? v : 0; }
  if (sv) { const v = parseAmount(sv.value); state.settings.saveGoal = v > 0 ? v : 0; }
}

/* =====================================================================
   Sincronización entre dispositivos (gratis, sin servidor)
   Los datos se guardan en un Gist SECRETO de tu cuenta de GitHub, CIFRADOS
   con AES-256 en tu dispositivo. GitHub solo ve texto ilegible; la clave
   va dentro del "código de vinculación" y nunca sale de tus dispositivos.
   ===================================================================== */
const SYNC_KEY = 'hucha.sync';          // {t: token, g: gistId, k: clave}  (solo en este dispositivo)
const SYNC_BASE_KEY = 'hucha.syncBase'; // última versión común: sirve para saber qué se ha borrado
const GIST_FILE = 'hucha.json';
const GH_API = 'https://api.github.com';
const TOKEN_URL = 'https://github.com/settings/tokens/new?scopes=gist&description=Hucha';
const LOCAL_ONLY = ['theme'];           // ajustes que cada dispositivo decide por su cuenta

const sync = { cfg: readJSON(SYNC_KEY), busy: false, again: false, timer: null, last: 0, error: '' };

function readJSON(k) { try { return JSON.parse(localStorage.getItem(k)) || null; } catch (e) { return null; } }
function writeJSON(k, v) { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sin espacio */ } }
const clone = o => JSON.parse(JSON.stringify(o));

/* ---------- Cifrado ---------- */
function bytesToB64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
const b64ToBytes = b => Uint8Array.from(atob(b), c => c.charCodeAt(0));
const cryptoKey = (k, use) => crypto.subtle.importKey('raw', b64ToBytes(k), 'AES-GCM', false, [use]);
async function encryptState(obj, k) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const buf = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await cryptoKey(k, 'encrypt'), new TextEncoder().encode(JSON.stringify(obj)));
  return JSON.stringify({ app: 'hucha', v: 1, iv: bytesToB64(iv), data: bytesToB64(new Uint8Array(buf)) });
}
async function decryptState(text, k) {
  const p = JSON.parse(text);
  const buf = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64ToBytes(p.iv) }, await cryptoKey(k, 'decrypt'), b64ToBytes(p.data));
  return JSON.parse(new TextDecoder().decode(buf));
}

/* ---------- GitHub ---------- */
function syncErr(code) { const e = new Error(code); e.code = code; return e; }
async function gh(path, token, opts = {}) {
  let res;
  try {
    res = await fetch(GH_API + path, {
      ...opts, cache: 'no-store',
      headers: { Accept: 'application/vnd.github+json', Authorization: `Bearer ${token}`, ...(opts.body ? { 'Content-Type': 'application/json' } : {}) },
    });
  } catch (e) { throw syncErr('offline'); }
  if (res.status === 401) throw syncErr('auth');
  if (res.status === 403 || res.status === 404) throw syncErr(path.startsWith('/gists/') ? 'notfound' : 'scope');
  if (!res.ok) throw syncErr('server');
  return res.json();
}
async function readRemote(cfg) {
  const g = await gh(`/gists/${cfg.g}`, cfg.t);
  const f = g.files && g.files[GIST_FILE];
  if (!f) return null;
  let text = f.content;
  if (f.truncated) text = await (await fetch(f.raw_url, { cache: 'no-store' })).text();
  if (!text || !text.trim().startsWith('{')) return null;
  try { return await decryptState(text, cfg.k); } catch (e) { throw syncErr('key'); }
}
async function writeRemote(cfg, data) {
  await gh(`/gists/${cfg.g}`, cfg.t, { method: 'PATCH', body: JSON.stringify({ files: { [GIST_FILE]: { content: await encryptState(data, cfg.k) } } }) });
}

/* ---------- Mezcla de tres vías: une los cambios de los dos dispositivos ----------
   base = lo último que ambos tenían en común. Si algo cambió solo en un lado,
   gana ese lado; si se borró en un lado (y el otro no lo tocó), se borra. */
function stable(v) {
  if (Array.isArray(v)) return '[' + v.map(stable).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).filter(k => v[k] !== undefined).sort().map(k => JSON.stringify(k) + ':' + stable(v[k])).join(',') + '}';
  return JSON.stringify(v);
}
const same = (a, b) => stable(a) === stable(b);
const isObj = v => v && typeof v === 'object' && !Array.isArray(v);
const hasIds = a => Array.isArray(a) && a.every(x => isObj(x) && x.id != null);

function merge3(base, local, remote) {
  if (same(local, remote)) return local;
  if (same(local, base)) return remote;
  if (same(remote, base)) return local;
  // Los dos lados han cambiado
  if (hasIds(local) && hasIds(remote)) {
    const B = new Map((hasIds(base) ? base : []).map(x => [x.id, x]));
    const L = new Map(local.map(x => [x.id, x])), R = new Map(remote.map(x => [x.id, x]));
    const ids = [...new Set([...remote.map(x => x.id), ...local.map(x => x.id)])];
    return ids.map(id => merge3(B.get(id), L.get(id), R.get(id))).filter(x => x !== undefined);
  }
  if (Array.isArray(local) && Array.isArray(remote)) {
    const B = Array.isArray(base) ? base : [];
    return [...new Set([...remote, ...local])].filter(x => (local.includes(x) && remote.includes(x)) || !B.includes(x));
  }
  if (isObj(local) && isObj(remote)) {
    const B = isObj(base) ? base : {}, out = {};
    for (const k of new Set([...Object.keys(remote), ...Object.keys(local)])) {
      const v = merge3(B[k], local[k], remote[k]);
      if (v !== undefined) out[k] = v;
    }
    return out;
  }
  return local; // mismo dato cambiado en los dos sitios: gana este dispositivo
}
function mergeStates(base, local, remote) {
  const d = DEFAULT();
  const out = merge3(base || d, local, { ...d, ...remote, settings: { ...d.settings, ...(remote.settings || {}) } });
  // El total "evitado" es un contador: se suman los incrementos de cada lado
  const b = (base || d).resistedTotal || 0;
  out.resistedTotal = Math.round(((local.resistedTotal || 0) + (remote.resistedTotal || 0) - b) * 100) / 100;
  for (const k of LOCAL_ONLY) { if (local.settings[k] === undefined) delete out.settings[k]; else out.settings[k] = local.settings[k]; }
  out.settings.onboarded = !!(local.settings.onboarded || (remote.settings || {}).onboarded);
  // Un gasto fijo pudo apuntarse a la vez en los dos dispositivos: quitar el repetido
  const seen = new Set();
  out.expenses = (out.expenses || []).filter(e => {
    if (!e.recurringId) return true;
    const k = e.recurringId + '|' + e.date;
    if (seen.has(k)) return false;
    seen.add(k); return true;
  });
  return { ...d, ...out, settings: { ...d.settings, ...out.settings } };
}

/* ---------- Ciclo de sincronización ---------- */
function syncSoon(ms = 2500) {
  if (!sync.cfg) return;
  clearTimeout(sync.timer);
  sync.timer = setTimeout(() => syncNow(), ms);
}
const SYNC_MSG = {
  offline: 'Sin conexión. Se sincronizará sola cuando vuelva internet.',
  auth: 'Tu llave de GitHub ha caducado o la has borrado. Desvincula y vuelve a vincular con una llave nueva.',
  scope: 'Esa llave no tiene permiso de "gist". Crea otra marcando la casilla "gist".',
  notfound: 'No encuentro tus datos en GitHub (¿se borró el Gist?). Desvincula y vuelve a activar.',
  key: 'El código de vinculación no coincide con los datos guardados. Vuelve a copiarlo del otro dispositivo.',
  server: 'GitHub no responde ahora mismo. Lo intentaré luego.',
  nocrypto: 'Para sincronizar, abre Hucha desde su dirección https:// (no desde un archivo).',
};
async function syncNow() {
  if (!sync.cfg) return false;
  if (sync.busy) { sync.again = true; return false; }
  clearTimeout(sync.timer);
  sync.busy = true; sync.error = ''; refreshSyncUI();
  let ok = false;
  try {
    if (!window.crypto?.subtle) throw syncErr('nocrypto');
    const cfg = sync.cfg;
    const snap = clone(state);
    const remote = await readRemote(cfg);
    const merged = remote ? mergeStates(readJSON(SYNC_BASE_KEY), snap, remote) : snap;
    if (!remote || !same(merged, remote)) await writeRemote(cfg, merged);
    if (sync.cfg !== cfg) return false; // se desvinculó mientras tanto
    writeJSON(SYNC_BASE_KEY, merged);
    // Aplicar lo que llega sin pisar lo que hayas tocado mientras tanto
    const cur = clone(state);
    const next = same(cur, snap) ? merged : mergeStates(snap, cur, merged);
    if (!same(next, cur)) {
      state = next; persist(); applyTheme();
      render();
    }
    if (!same(next, merged)) sync.again = true;
    sync.last = Date.now(); ok = true;
  } catch (e) {
    sync.error = e.code || 'server';
    if (sync.error === 'offline' || sync.error === 'server') syncSoon(60000);
  } finally {
    sync.busy = false; refreshSyncUI();
    if (sync.again) { sync.again = false; syncSoon(500); }
  }
  return ok;
}
function refreshSyncUI() {
  if (sheetKind === 'settings') setSheet(settingsHTML());
  const st = $('#syncState'); if (st) st.innerHTML = syncStateText();
}
function syncStateText() {
  if (sync.busy) return '🔄 Sincronizando…';
  if (sync.error) return `⚠️ ${esc(SYNC_MSG[sync.error] || SYNC_MSG.server)}`;
  if (sync.last) return `✅ Sincronizado ${Date.now() - sync.last < 60000 ? 'ahora mismo' : 'a las ' + new Date(sync.last).toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' })}`;
  return '✅ Activada';
}

/* ---------- Activar / vincular / desvincular ---------- */
const syncCode = cfg => 'HUCHA-' + btoa(JSON.stringify(cfg)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
function parseSyncCode(s) {
  try {
    const b = String(s).replace(/\s+/g, '').replace(/^HUCHA-/i, '').replace(/-/g, '+').replace(/_/g, '/');
    const c = JSON.parse(atob(b));
    return c && c.t && c.g && c.k ? { t: c.t, g: c.g, k: c.k } : null;
  } catch (e) { return null; }
}
function syncSettingsBlock() {
  if (!sync.cfg) return `
    <div class="section-title">Sincronizar móvil ↔ ordenador</div>
    <p class="muted small">Gratis, usando tu cuenta de GitHub. Los datos viajan <b>cifrados</b>: ni GitHub puede leerlos.</p>
    <div class="btn-row">
      <button class="btn primary" data-action="syncSetup">☁️ Activar (primera vez)</button>
      <button class="btn ghost" data-action="syncJoin">🔗 Tengo un código</button>
    </div>`;
  return `
    <div class="section-title">Sincronizar móvil ↔ ordenador</div>
    <p class="small" id="syncState">${syncStateText()}</p>
    <div class="btn-row">
      <button class="btn ghost" data-action="syncNowBtn">🔄 Sincronizar ahora</button>
      <button class="btn ghost" data-action="syncCopy">📋 Copiar código</button>
    </div>
    <p class="muted small">Para añadir otro dispositivo: copia el código, pásatelo (AirDrop, Notas, o copiar en el Mac y pegar en el iPhone) y en el otro dispositivo toca <b>Ajustes → Tengo un código</b>. Guárdalo como una contraseña: quien lo tenga puede ver tus gastos.</p>
    <button class="btn link" data-action="syncUnlink">Desvincular este dispositivo</button>`;
}
function syncSetupHTML(found) {
  return `${sheetHead('Activar sincronización')}
    <ol class="steps">
      <li>Entra en GitHub (o crea una cuenta gratis) y abre <a href="${TOKEN_URL}" target="_blank" rel="noopener"><b>este enlace</b></a>.</li>
      <li>Ya viene marcada la casilla <b>gist</b>. En <b>Expiration</b> elige <b>No expiration</b>.</li>
      <li>Baja del todo, toca <b>Generate token</b> y copia la llave (empieza por <code>ghp_</code>).</li>
      <li>Pégala aquí:</li>
    </ol>
    <label class="field"><span>Llave de GitHub</span><input id="syncToken" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="ghp_…"></label>
    ${found ? `<p class="small">⚠️ Ya tienes una Hucha sincronizada en esta cuenta. Si es tuya, usa mejor <b>Tengo un código</b> con el código del otro dispositivo, o perderías la conexión entre ellos.</p>
      <div class="btn-row"><button class="btn primary" data-action="syncJoin">🔗 Tengo un código</button><button class="btn ghost" data-action="syncCreateAnyway">Crear otra nueva</button></div>`
    : '<button class="btn primary w100" data-action="syncCreate">Activar</button>'}
    <p class="muted small">La llave solo da permiso para guardar "gists" (notas), nada más de tu cuenta. Se queda guardada en este dispositivo.</p>`;
}
function syncJoinHTML() {
  return `${sheetHead('Vincular con otro dispositivo')}
    <p class="muted">En el dispositivo donde ya usas Hucha: <b>Ajustes → Sincronizar → 📋 Copiar código</b>. Pégalo aquí:</p>
    <label class="field"><span>Código de vinculación</span><textarea id="syncCodeIn" rows="4" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="HUCHA-…"></textarea></label>
    <button class="btn primary w100" data-action="syncJoinGo">Vincular</button>
    <p class="muted small">Si este dispositivo ya tenía gastos, se juntan con los del otro (no se pierde nada).</p>`;
}
async function syncCreate(force) {
  const input = $('#syncToken'), btn = $('[data-action="syncCreate"],[data-action="syncCreateAnyway"]');
  const token = (input?.value || syncCreate.token || '').trim();
  if (!token) { toast('Pega primero la llave de GitHub'); return; }
  if (!window.crypto?.subtle) { toast(SYNC_MSG.nocrypto); return; }
  syncCreate.token = token;
  if (btn) { btn.disabled = true; btn.textContent = 'Conectando…'; }
  try {
    if (!force) {
      const list = await gh('/gists?per_page=100', token);
      if (list.some(g => g.files && g.files[GIST_FILE])) { setSheet(syncSetupHTML(true)); $('#syncToken').value = token; return; }
    }
    const k = bytesToB64(crypto.getRandomValues(new Uint8Array(32)));
    const g = await gh('/gists', token, { method: 'POST', body: JSON.stringify({ description: 'Hucha · datos cifrados (no borrar)', public: false, files: { [GIST_FILE]: { content: await encryptState(state, k) } } }) });
    sync.cfg = { t: token, g: g.id, k };
    writeJSON(SYNC_KEY, sync.cfg); writeJSON(SYNC_BASE_KEY, clone(state));
    syncCreate.token = '';
    sync.last = Date.now(); sync.error = '';
    openSheet(settingsHTML(), 'settings');
    toast('☁️ Sincronización activada. Ahora copia el código para el otro dispositivo');
  } catch (e) {
    toast(e.code === 'auth' ? 'Esa llave no es válida. Cópiala otra vez' : SYNC_MSG[e.code] || SYNC_MSG.server);
    if (btn) { btn.disabled = false; btn.textContent = force ? 'Crear otra nueva' : 'Activar'; }
  }
}
async function syncJoinGo() {
  const cfg = parseSyncCode($('#syncCodeIn').value);
  if (!cfg) { toast('Ese código no parece de Hucha. Cópialo entero'); return; }
  if (!window.crypto?.subtle) { toast(SYNC_MSG.nocrypto); return; }
  const btn = $('[data-action="syncJoinGo"]'); btn.disabled = true; btn.textContent = 'Vinculando…';
  try {
    await readRemote(cfg); // comprueba llave, Gist y clave antes de guardar nada
    sync.cfg = cfg; writeJSON(SYNC_KEY, cfg);
    writeJSON(SYNC_BASE_KEY, null); // primera vez: se juntan los datos de los dos lados
    const ok = await syncNow();
    if (!ok) throw syncErr(sync.error || 'server');
    if (!state.settings.startedOn && state.settings.onboarded) state.settings.startedOn = todayISO();
    applyRecurring(); save();
    closeSheet(); render(); confetti();
    if (!state.settings.onboarded) { ob = { step: 0 }; openSheet(obHTML(), 'onboard'); }
    toast(`🔗 ¡Vinculado! ${state.expenses.length} gastos sincronizados`);
  } catch (e) {
    if (e.code !== 'offline' && e.code !== 'server') { sync.cfg = null; writeJSON(SYNC_KEY, null); }
    toast(SYNC_MSG[e.code] || SYNC_MSG.server);
    btn.disabled = false; btn.textContent = 'Vincular';
  }
}
function closeSyncSheet() {
  if (!state.settings.onboarded) { ob = { step: 0 }; openSheet(obHTML(), 'onboard'); }
  else openSheet(settingsHTML(), 'settings');
}

/* =====================================================================
   Asistente con IA (Claude Haiku): capturas + explicación → gastos.
   La llave se guarda SOLO en este dispositivo y solo se envía a Anthropic.
   ===================================================================== */
const AI_STORE = 'hucha.ai';                 // { key, usd } (no se sincroniza)
const AI_MODEL = 'claude-haiku-5-5';
const AI_SDK_URL = 'https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@1.3.0/+esm';
const AI_CONSOLE = 'https://platform.claude.com';
const AI_PRICE = { in: 0.10, out: 0.50 };    // $ por millón de tokens (Haiku 5.5, peticiones ≤ 100K)
const AI_MAX_IMGS = 6;
const aiCfg = readJSON(AI_STORE) || {};
let ai = null;                               // conversación abierta
let aiSdkPromise = null;

const AI_CAT_HINTS = {
  super: 'Mercadona, Spar, Suma, Lidl, fruterías, panaderías',
  casa: 'alquiler, luz, agua, internet, móvil, cosas para la casa',
  transporte: 'gasolina, guagua/bus, taxi, parking, Uber',
  salud: 'farmacia, médico, dentista, óptica, fisio',
  comer: 'restaurantes, bares, cenas fuera',
  delivery: 'Glovo, Just Eat, Uber Eats',
  cafe: 'cafeterías, snacks, heladerías',
  ocio: 'cine, conciertos, planes, juegos, libros',
  compras: 'ropa, Amazon, bazares, tecnología, peluquería, regalos',
  subs: 'suscripciones mensuales, gimnasio, cursos online pagados a plazos',
  viajes: 'vuelos, hoteles, Airbnb',
  otros: 'lo que no encaje (copias, trámites, comisiones…)',
};
const AI_SYSTEM = `Eres el asistente de Hucha, una app de control de gastos pensada para personas con TDAH. El usuario te manda capturas (app del banco, tickets, notificaciones) y/o te explica sus gastos con palabras, a menudo dictadas (puede haber errores de dictado).

Tu trabajo: sacar la lista de GASTOS para apuntarlos.
- Solo gastos (dinero que sale). Deja fuera ingresos, abonos, transferencias recibidas, nóminas y devoluciones, y dilo en el mensaje en pocas palabras.
- importe: positivo, en euros (62.03).
- fecha: AAAA-MM-DD. Las apps del banco ponen "Hoy", "Lunes 16" o "Viernes 13" sin año: deduce el mes y el año con la fecha de hoy y las cabeceras (p. ej. "Noviembre"). Nunca pongas fechas futuras.
- Si las capturas se solapan, no repitas un movimiento (misma fecha, importe y comercio = uno solo).
- concepto: corto y claro. Quita "Compra Tarj.", números de tarjeta (5402..2010) y la ciudad. Ej.: "Mercadona", "Peluquería Figueroa y Guerra", "Bizum a Josep (cena)".
- categoria: una de la lista. Lo que diga el usuario manda sobre lo que tú supongas.
- Pregunta SOLO lo que no puedas deducir y cambie la categoría o el importe: típico en Bizums enviados, pagos financiados o a plazos (Sequra, Aplazame, Klarna…), comercios que no reconoces o importes ilegibles. Máximo 3 preguntas a la vez, cada una con 2–4 opciones muy cortas (nombres de categoría, "No es un gasto"…). Mientras, incluye ese gasto con tu mejor suposición.
- Si el usuario dice que algo no es un gasto, quítalo de la lista.
- Devuelve SIEMPRE la lista completa y actualizada de toda la conversación, no solo lo nuevo.
- mensaje: 1–2 frases en español de España, tono cercano, sin sermones ni markdown.

Categorías (id: nombre — ejemplos):
${CATS.map(c => `${c.id}: ${c.name} — ${AI_CAT_HINTS[c.id] || ''}`).join('\n')}`;

const AI_SCHEMA = {
  type: 'object', additionalProperties: false, required: ['mensaje', 'gastos', 'preguntas'],
  properties: {
    mensaje: { type: 'string' },
    gastos: { type: 'array', items: {
      type: 'object', additionalProperties: false, required: ['fecha', 'importe', 'concepto', 'categoria'],
      properties: { fecha: { type: 'string' }, importe: { type: 'number' }, concepto: { type: 'string' }, categoria: { type: 'string', enum: CATS.map(c => c.id) } },
    } },
    preguntas: { type: 'array', items: {
      type: 'object', additionalProperties: false, required: ['pregunta', 'opciones'],
      properties: { pregunta: { type: 'string' }, opciones: { type: 'array', items: { type: 'string' } } },
    } },
  },
};

function aiSdk() {
  if (!aiSdkPromise) aiSdkPromise = import(AI_SDK_URL).then(m => m.default).catch(e => { aiSdkPromise = null; throw e; });
  return aiSdkPromise;
}
async function aiClient(key = aiCfg.key) {
  const Anthropic = await aiSdk();
  return { Anthropic, client: new Anthropic({ apiKey: key, dangerouslyAllowBrowser: true, maxRetries: 2 }) };
}
function aiErrorText(err, Anthropic) {
  if (!Anthropic) return 'No he podido cargar el asistente. ¿Tienes conexión?';
  if (err instanceof Anthropic.AuthenticationError) return 'La llave no es válida. Revisa que la has copiado entera.';
  if (err instanceof Anthropic.PermissionDeniedError) return 'Esa llave no tiene permiso. Crea una nueva en platform.claude.com.';
  if (err instanceof Anthropic.RateLimitError) return 'Demasiadas peticiones seguidas. Espera un minuto y vuelve a probar.';
  if (err instanceof Anthropic.BadRequestError) return 'Anthropic ha rechazado la petición. Si es una cuenta nueva, revisa que tengas saldo (Billing).';
  if (err instanceof Anthropic.APIConnectionError) return 'Sin conexión. Prueba otra vez cuando tengas internet.';
  if (err instanceof Anthropic.APIError) return 'Anthropic no responde ahora mismo. Prueba en un rato.';
  return 'Algo ha fallado. Prueba otra vez.';
}
function aiAddUsage(u) {
  if (!u) return;
  const inTok = (u.input_tokens || 0) + (u.cache_creation_input_tokens || 0) * 1.25 + (u.cache_read_input_tokens || 0) * 0.1;
  aiCfg.usd = (aiCfg.usd || 0) + (inTok * AI_PRICE.in + (u.output_tokens || 0) * AI_PRICE.out) / 1e6;
  writeJSON(AI_STORE, aiCfg);
}
const aiSpentText = () => {
  const usd = aiCfg.usd || 0;
  return usd < 0.01 ? 'menos de 1 céntimo' : `≈ ${usd.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} $`;
};

/** Reduce la captura (lado largo ≤ 1568 px, JPEG) para que pese y cueste menos. */
async function shrinkImage(file, max = 1568) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(img, 0, 0, c.width, c.height);
    const dataUrl = c.toDataURL('image/jpeg', 0.85);
    return { dataUrl, b64: dataUrl.split(',')[1] };
  } finally { URL.revokeObjectURL(url); }
}

function aiOpen() {
  if (!aiCfg.key) { openSheet(aiSetupHTML(), 'ai'); return; }
  if (!ai) ai = { msgs: [], log: [], pending: [], gastos: [], preguntas: [], answers: {}, draft: '', busy: false, error: '' };
  openSheet(aiHTML(), 'ai');
  aiScroll();
}
function aiRender() { if (sheetKind === 'ai' && ai) { setSheet(aiHTML()); aiScroll(); } }
function aiScroll() { const s = $('.sheet'); if (s) s.scrollTop = s.scrollHeight; }

function aiSetupHTML() {
  return `${sheetHead('✨ Activar el asistente')}
    <p class="muted small">Le mandas capturas, le explicas qué es cada cosa y él te pregunta lo que no sepa. Usa Claude Haiku, de Anthropic: pagas tú directamente lo que gastes (céntimos al mes), sin suscripción.</p>
    <ol class="steps">
      <li>Entra en <a href="${AI_CONSOLE}" target="_blank" rel="noopener"><b>platform.claude.com</b></a> y crea una cuenta. Es distinta de la suscripción de Claude.</li>
      <li>En <a href="${AI_CONSOLE}/settings/billing" target="_blank" rel="noopener"><b>Billing</b></a>, añade tu tarjeta y compra un poco de saldo (con 5 $ tienes para muchísimo).</li>
      <li>En <a href="${AI_CONSOLE}/settings/keys" target="_blank" rel="noopener"><b>API keys</b></a>, toca <b>Create key</b>, llámala "Hucha" y copia la llave (empieza por <code>sk-ant-</code>).</li>
      <li>Pégala aquí:</li>
    </ol>
    <label class="field"><span>Llave de Anthropic</span><input id="aiKeyIn" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="sk-ant-…"></label>
    <button class="btn primary w100" data-action="aiSaveKey">Activar</button>
    <p class="muted small">🔒 La llave se guarda solo en este dispositivo (no se sincroniza) y solo se envía a Anthropic. Las capturas y lo que escribas también se envían a Anthropic para leerlas.</p>`;
}
async function aiSaveKey() {
  const key = ($('#aiKeyIn')?.value || '').trim();
  if (!/^sk-ant-/.test(key)) { toast('La llave empieza por sk-ant-'); return; }
  const btn = $('[data-action="aiSaveKey"]');
  btn.disabled = true; btn.textContent = 'Comprobando…';
  let Anthropic = null;
  try {
    const c = await aiClient(key);
    Anthropic = c.Anthropic;
    await c.client.models.retrieve(AI_MODEL);   // no gasta saldo: solo comprueba que la llave funciona
    aiCfg.key = key; writeJSON(AI_STORE, aiCfg);
    toast('✨ Asistente activado');
    aiOpen();
  } catch (err) {
    toast(aiErrorText(err, Anthropic));
    btn.disabled = false; btn.textContent = 'Activar';
  }
}

function aiHTML() {
  const a = ai;
  const bubbles = a.log.map(m => `<div class="ai-msg ${m.me ? 'me' : 'bot'}">${m.imgs?.length ? `<div class="ai-mini">${m.imgs.map(u => `<img src="${u}" alt="">`).join('')}</div>` : ''}${m.text ? `<div>${esc(m.text).replace(/\n/g, '<br>')}</div>` : ''}</div>`).join('');
  const qs = !a.busy ? a.preguntas.map((q, i) => `<div class="ai-q"><p>${esc(q.pregunta)}</p><div class="ai-opts">${q.opciones.map((o, j) => `<button class="chip ${a.answers[i] === o ? 'on' : ''}" data-action="aiOpt" data-q="${i}" data-o="${j}">${esc(o)}</button>`).join('')}</div></div>`).join('') : '';
  const n = a.gastos.length;
  return `${sheetHead('✨ Asistente')}
    <div class="ai-log" id="aiLog">
      ${a.log.length ? '' : `<div class="ai-msg bot">Mándame capturas del banco o de un ticket y, si quieres, explícame qué es cada cosa. Si algo no lo tengo claro, te pregunto.</div>`}
      ${bubbles}${qs}
      ${a.busy ? '<div class="ai-msg bot typing">Mirando…</div>' : ''}
      ${a.error ? `<div class="ai-msg bot err">😕 ${esc(a.error)}</div>` : ''}
    </div>
    ${n && !a.busy ? `<div class="ai-found"><span><b>${n} ${n === 1 ? 'gasto listo' : 'gastos listos'}</b><br><span class="muted small">${money(a.gastos.reduce((s, g) => s + g.importe, 0))} en total</span></span><button class="btn primary" data-action="aiReview">Revisar y guardar</button></div>` : ''}
    <div class="ai-compose">
      ${a.pending.length ? `<div class="ai-thumbs">${a.pending.map((p, i) => `<span><img src="${p.dataUrl}" alt="Captura ${i + 1}"><button data-action="aiDropImg" data-i="${i}" aria-label="Quitar captura">✕</button></span>`).join('')}</div>` : ''}
      <textarea id="aiText" rows="2" placeholder="${a.preguntas.length ? 'Respóndele o explícale algo más…' : 'Escribe o dicta con el 🎤 del teclado…'}">${esc(a.draft)}</textarea>
      <div class="btn-row">
        <button class="btn ghost" data-action="aiPick" ${a.busy ? 'disabled' : ''}>📎 Capturas</button>
        <button class="btn primary" data-action="aiSend" ${a.busy ? 'disabled' : ''}>Enviar</button>
      </div>
      <p class="muted small ai-foot">Gastado con el asistente: ${aiSpentText()}${a.log.length ? ' · <button class="btn link" data-action="aiReset">Empezar de nuevo</button>' : ''}</p>
    </div>`;
}

async function aiAddFiles(files) {
  if (!ai) return;
  const room = AI_MAX_IMGS - ai.pending.length;
  if (files.length > room) toast(`Como mucho ${AI_MAX_IMGS} capturas a la vez`);
  for (const f of files.slice(0, Math.max(0, room))) {
    try { ai.pending.push(await shrinkImage(f)); } catch (e) { toast('No he podido abrir una imagen'); }
  }
  aiRender();
}

function aiTodayText() {
  const t = todayISO(), d = fromISO(t);
  return `Hoy es ${WEEKDAYS[d.getDay()]}, ${d.getDate()} de ${MONTHS[d.getMonth()]} de ${d.getFullYear()} (${t}).`;
}

async function aiSend() {
  const a = ai;
  if (!a || a.busy) return;
  a.draft = $('#aiText')?.value ?? a.draft;
  const answered = a.preguntas.map((q, i) => (a.answers[i] ? `- ${q.pregunta} → ${a.answers[i]}` : '')).filter(Boolean);
  if (!a.pending.length && !a.draft.trim() && !answered.length) { toast('Añade una captura o escribe algo'); return; }

  let text = a.msgs.length ? '' : aiTodayText() + '\n';
  if (answered.length) text += 'Respuestas:\n' + answered.join('\n') + '\n';
  if (a.draft.trim()) text += a.draft.trim();
  if (!a.draft.trim() && !answered.length) text += 'Saca los gastos de estas capturas.';
  const content = [
    ...a.pending.map(p => ({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: p.b64 } })),
    { type: 'text', text: text.trim() },
  ];
  const shown = { me: true, imgs: a.pending.map(p => p.dataUrl), text: [answered.map(l => l.replace(/^- .*? → /, '')).join(', '), a.draft.trim()].filter(Boolean).join('\n') };
  const sent = { pending: a.pending, draft: a.draft, answers: a.answers };
  a.log.push(shown);
  a.pending = []; a.draft = ''; a.answers = {}; a.busy = true; a.error = '';
  aiRender();

  let Anthropic = null;
  try {
    const c = await aiClient();
    Anthropic = c.Anthropic;
    const msgs = [...a.msgs, { role: 'user', content }];
    const res = await c.client.messages.create({
      model: AI_MODEL,
      max_tokens: 16000,
      system: AI_SYSTEM,
      messages: msgs,
      cache_control: { type: 'ephemeral' },
      output_config: { effort: 'low', format: { type: 'json_schema', schema: AI_SCHEMA } },
    });
    aiAddUsage(res.usage);
    if (res.stop_reason === 'refusal') throw Object.assign(new Error('refusal'), { hucha: 'Esto no lo puedo procesar. Prueba con otra captura o escríbemelo.' });
    if (res.stop_reason === 'max_tokens') throw Object.assign(new Error('max_tokens'), { hucha: 'Eran demasiadas cosas de golpe. Prueba con menos capturas.' });
    const block = res.content.find(b => b.type === 'text');
    const out = JSON.parse(block?.text || '');
    a.msgs = [...msgs, { role: 'assistant', content: res.content }];
    a.gastos = (out.gastos || []).filter(g => isFinite(g.importe) && Math.abs(g.importe) > 0).map(g => ({ ...g, importe: Math.round(Math.abs(g.importe) * 100) / 100 }));
    a.preguntas = (out.preguntas || []).filter(q => q.pregunta && q.opciones?.length).slice(0, 3);
    a.log.push({ me: false, text: out.mensaje || (a.gastos.length ? 'Listo.' : 'No he encontrado gastos.') });
  } catch (err) {
    a.log.pop();
    Object.assign(a, sent);
    a.error = err.hucha || (err instanceof SyntaxError ? 'No he entendido la respuesta. Prueba otra vez.' : aiErrorText(err, Anthropic));
    if (Anthropic && err instanceof Anthropic.AuthenticationError) { delete aiCfg.key; writeJSON(AI_STORE, aiCfg); }
  }
  a.busy = false;
  if (ai === a) aiRender();
}

function aiOpt(el) {
  const a = ai, q = Number(el.dataset.q), o = a.preguntas[q]?.opciones[Number(el.dataset.o)];
  if (o == null || a.busy) return;
  a.draft = $('#aiText')?.value ?? a.draft;
  if (a.answers[q] === o) delete a.answers[q]; else a.answers[q] = o;
  // Respondidas todas y sin nada escrito → se envía sola (un toque menos)
  if (a.preguntas.every((_, i) => a.answers[i]) && !a.draft.trim() && !a.pending.length) aiSend();
  else aiRender();
}

function aiReview() {
  const t = todayISO();
  const existing = new Set(state.expenses.map(e => e.date + '|' + e.amount.toFixed(2)));
  const items = ai.gastos.map(g => {
    const iso = parseDateStrict(g.fecha);
    const date = iso && iso <= t ? iso : t;
    const dup = existing.has(date + '|' + g.importe.toFixed(2));
    return { amount: g.importe, note: g.concepto || '', date, cat: CAT[g.categoria] ? g.categoria : 'otros', dup, checked: !dup };
  });
  openReview(items, 'ia');
}

/* =====================================================================
   Eventos (delegados)
   ===================================================================== */
const actions = {
  openAdd: () => openAdd(),
  openAddBlank: () => { closeSheet(); openAdd(); },
  overlay: () => { if (sheetKind === 'sync') closeSyncSheet(); else if (sheetKind !== 'onboard') closeSheet(); },
  closeSheet: () => (sheetKind === 'sync' ? closeSyncSheet() : closeSheet()),
  // Sincronización
  syncSetup: () => openSheet(syncSetupHTML(false), 'sync'),
  syncCreate: () => syncCreate(false),
  syncCreateAnyway: () => syncCreate(true),
  syncJoin: () => openSheet(syncJoinHTML(), 'sync'),
  syncJoinGo: () => syncJoinGo(),
  syncNowBtn: () => syncNow(),
  syncCopy: async () => {
    const code = syncCode(sync.cfg);
    try { await navigator.clipboard.writeText(code); toast('📋 Código copiado. Pégalo en el otro dispositivo'); }
    catch (e) { prompt('Copia este código:', code); }
  },
  syncUnlink: () => {
    if (!confirm('¿Desvincular este dispositivo? Tus datos se quedan aquí y en la nube, pero dejarán de sincronizarse.')) return;
    sync.cfg = null; sync.last = 0; sync.error = ''; clearTimeout(sync.timer);
    writeJSON(SYNC_KEY, null); writeJSON(SYNC_BASE_KEY, null);
    setSheet(settingsHTML()); toast('Desvinculado');
  },
  key: el => keyPress(el.dataset.k),
  pickCat: el => pickCat(el.dataset.cat),
  setDate: el => { add.date = el.dataset.d; add.note = $('#note').value; setSheet(addHTML()); updateAmt(); updateSuggest(); },
  editExpense: el => {
    const e = state.expenses.find(x => x.id === el.dataset.id);
    if (e) openAdd({ amount: e.amount, cat: e.cat, note: e.note, date: e.date, editId: e.id, source: e.source });
  },
  deleteExpense: () => deleteExpense(),
  capture: () => $('#fileImg').click(),
  pasteText: () => pasteText(),
  pasteGo: () => { const r = textToItems($('#pasteArea').value); handleParsed(r, 'paste'); },
  rvAll: el => { review.items.forEach(i => { i.checked = el.dataset.v === '1'; }); setSheet(reviewHTML()); },
  rvSave: () => reviewSave(),
  // Asistente con IA
  aiOpen: () => aiOpen(),
  aiSaveKey: () => aiSaveKey(),
  aiSend: () => aiSend(),
  aiOpt: el => aiOpt(el),
  aiPick: () => { if (ai) ai.draft = $('#aiText')?.value ?? ai.draft; $('#fileAi').click(); },
  aiDropImg: el => { ai.draft = $('#aiText')?.value ?? ai.draft; ai.pending.splice(Number(el.dataset.i), 1); aiRender(); },
  aiReview: () => aiReview(),
  aiReset: () => { ai = null; aiOpen(); },
  aiForget: () => {
    if (!confirm('¿Quitar la llave del asistente de este dispositivo? Para volver a usarlo tendrás que pegarla otra vez.')) return;
    delete aiCfg.key; writeJSON(AI_STORE, aiCfg); ai = null;
    setSheet(settingsHTML()); toast('Llave quitada');
  },
  noSpendToday: () => {
    const t = todayISO();
    if (!state.noSpend.includes(t)) state.noSpend.push(t);
    save(); render(); confetti();
    toast('✨ ¡Día sin gastos! Eso también cuenta para la racha', { undo: () => { state.noSpend = state.noSpend.filter(d => d !== t); } });
    checkBadges();
  },
  dismiss: el => { state.dismissed[el.dataset.k] = true; save(); render(); },
  month: el => { ui.monthOffset = Math.min(0, ui.monthOffset + Number(el.dataset.d)); ui.cat = null; ui.day = null; render(); },
  pickDay: el => {
    ui.day = ui.day === el.dataset.d ? null : el.dataset.d;
    render();
    if (ui.day) $('#movs')?.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'start' });
  },
  filterCat: el => { ui.cat = ui.cat === el.dataset.cat ? null : el.dataset.cat; render(); },
  openSettings: () => openSheet(settingsHTML(), 'settings'),
  addRecurring: () => addRecurring(),
  delRecurring: el => {
    const r = state.recurring.find(x => x.id === el.dataset.id);
    state.recurring = state.recurring.filter(x => x !== r);
    save(); setSheet(settingsHTML()); render();
    toast(`${r.name} ya no se apuntará solo`, { undo: () => { state.recurring.push(r); if (sheetKind === 'settings') setSheet(settingsHTML()); } });
  },
  importBank: () => $('#fileBank').click(),
  exportBackup: () => exportBackup(),
  importBackup: () => $('#fileJson').click(),
  exportCSV: () => exportCSV(),
  theme: el => { state.settings.theme = el.dataset.v; save(); applyTheme(); setSheet(settingsHTML()); },
  wipe: () => {
    if (!confirm('¿Borrar TODOS tus datos de este dispositivo? No se puede deshacer (salvo que tengas una copia).')) return;
    if (!confirm('¿Seguro del todo?')) return;
    sync.cfg = null; clearTimeout(sync.timer); writeJSON(SYNC_KEY, null); writeJSON(SYNC_BASE_KEY, null); // la copia en la nube no se toca
    delete aiCfg.key; delete aiCfg.usd; writeJSON(AI_STORE, null); ai = null;
    state = DEFAULT(); save(); closeSheet(); applyTheme(); render();
    ob = { step: 0 }; openSheet(obHTML(), 'onboard');
  },
  // Metas
  newGoal: () => { goalDraft = { name: '', emoji: '🛟', target: 0, saved: 0 }; openSheet(goalSheetHTML(), 'goal'); },
  goalMenu: el => { const g = state.goals.find(x => x.id === el.dataset.id); goalDraft = { ...g }; openSheet(goalSheetHTML(), 'goal'); },
  goalEmoji: el => { goalDraft.name = $('#gName').value; goalDraft.target = parseAmount($('#gTarget').value) || 0; if ($('#gSaved')) goalDraft.saved = parseAmount($('#gSaved').value) || 0; goalDraft.emoji = el.dataset.v; setSheet(goalSheetHTML()); },
  goalSave: () => {
    const name = $('#gName').value.trim(), target = parseAmount($('#gTarget').value);
    if (!name || !(target > 0)) { toast('Pon un nombre y una cantidad'); return; }
    if (goalDraft.id) {
      const g = state.goals.find(x => x.id === goalDraft.id);
      Object.assign(g, { name, target, emoji: goalDraft.emoji, saved: Math.max(0, parseAmount($('#gSaved').value) || 0) });
    } else state.goals.push({ id: uid(), name, target, emoji: goalDraft.emoji, saved: 0, created: todayISO() });
    save(); closeSheet(); render(); checkBadges();
  },
  goalDelete: () => {
    const g = state.goals.find(x => x.id === goalDraft.id);
    state.goals = state.goals.filter(x => x !== g);
    save(); closeSheet(); render();
    toast('Meta borrada', { undo: () => state.goals.push(g) });
  },
  goalAdd: el => goalAdd(state.goals.find(x => x.id === el.dataset.id), Number(el.dataset.n)),
  goalOther: el => {
    const g = state.goals.find(x => x.id === el.dataset.id);
    askNumber(`Añadir a ${g.emoji} ${g.name}`, n => goalAdd(g, n));
  },
  askOk: () => { const n = parseAmount($('#askNum').value); if (!(n > 0)) { toast('Escribe una cantidad'); return; } closeSheet(); askNumber.cb(n); },
  // Lista de espera
  newWish: () => { wishWait = 48; openSheet(wishSheetHTML(), 'wish'); },
  wishWait: el => { const n = $('#wName').value, p = $('#wPrice').value; wishWait = Number(el.dataset.v); setSheet(wishSheetHTML()); $('#wName').value = n; $('#wPrice').value = p; },
  wishSave: () => {
    const name = $('#wName').value.trim(), price = parseAmount($('#wPrice').value);
    if (!name || !(price > 0)) { toast('Pon qué es y cuánto cuesta'); return; }
    state.wishlist.push({ id: uid(), name, price, created: Date.now(), waitH: wishWait, status: 'waiting' });
    save(); closeSheet(); render();
    toast('⏳ En espera. Tu yo del futuro decidirá');
  },
  wishResist: el => {
    const w = state.wishlist.find(x => x.id === el.dataset.id);
    w.status = 'resisted'; w.decided = Date.now();
    state.resistedTotal = Math.round((state.resistedTotal + w.price) * 100) / 100;
    save(); render(); confetti();
    toast(`🛡️ ${money(w.price)} que siguen siendo tuyos`, { undo: () => { w.status = 'waiting'; state.resistedTotal -= w.price; } });
    checkBadges();
  },
  wishBuy: el => {
    // Se marca como comprado al guardar el gasto (si cierras la hoja, sigue en la lista)
    const w = state.wishlist.find(x => x.id === el.dataset.id);
    openAdd({ amount: w.price, note: w.name, source: 'wish', wishId: w.id });
    toast('Elige la categoría y queda apuntado');
  },
  // Bienvenida
  obNext: () => { obReadInputs(); ob.step++; setSheet(obHTML()); },
  obSkip: () => { ob.step = 3; setSheet(obHTML()); },
  obSave: el => obSaveChanged(Number(el.dataset.v), 'chip'),
  obFinish: () => {
    state.settings.onboarded = true;
    state.settings.startedOn = state.settings.startedOn || todayISO();
    save(); closeSheet(); render(); confetti();
    toast('¡Listo! Toca + para apuntar tu primer gasto');
  },
};

document.addEventListener('click', e => {
  if (e.target.id === 'addDate' && e.target.showPicker) { try { e.target.showPicker(); } catch (err) { /* iOS lo abre solo */ } }
  const tab = e.target.closest('[data-tab]');
  if (tab) { ui.tab = tab.dataset.tab; ui.cat = null; ui.day = null; render(); window.scrollTo({ top: 0 }); return; }
  const el = e.target.closest('[data-action]');
  if (el && actions[el.dataset.action]) actions[el.dataset.action](el);
});

document.addEventListener('input', e => {
  if (e.target.id === 'note') updateSuggest();
  if (e.target.id === 'obSave') obSaveChanged(parseAmount(e.target.value), 'input');
  if (e.target.id === 'obSlider') obSaveChanged(Number(e.target.value), 'slider');
  if (e.target.id === 'aiText' && ai) ai.draft = e.target.value;
  const rv = e.target.closest('[data-rv]');
  if (rv && review) {
    const it = review.items[Number(rv.closest('.rv').dataset.i)];
    const k = rv.dataset.rv;
    if (k === 'note') it.note = rv.value;
    if (k === 'amount') it.amount = Math.abs(parseAmount(rv.value)) || 0;
  }
});

document.addEventListener('change', e => {
  const t = e.target;
  if (t.id === 'addDate' && add && t.value) { add.date = t.value; add.note = $('#note').value; setSheet(addHTML()); updateAmt(); updateSuggest(); }
  if (t.dataset.setting) {
    const k = t.dataset.setting;
    let v = parseAmount(t.value);
    if (!(v >= 0)) v = 0;
    if (k === 'startDay') v = Math.min(31, Math.max(1, Math.round(v) || 1));
    if (k === 'hoursMonth') v = Math.round(v) || 160;
    state.settings[k] = v;
    save(); applyRecurring(); render();
  }
  const rv = t.closest?.('[data-rv]');
  if (rv && review) {
    const box = rv.closest('.rv'), it = review.items[Number(box.dataset.i)];
    const k = rv.dataset.rv;
    if (k === 'checked') { it.checked = rv.checked; box.classList.toggle('off', !rv.checked); const n = review.items.filter(i => i.checked).length; const btn = $('[data-action="rvSave"]'); btn.disabled = !n; btn.textContent = `Guardar ${n} ${n === 1 ? 'gasto' : 'gastos'}`; }
    if (k === 'cat') it.cat = rv.value;
    if (k === 'date') it.date = rv.value || todayISO();
  }
});

$('#fileImg').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) runOCR(f); });
$('#fileBank').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) importBank(f); });
$('#fileJson').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) importBackupFile(f); });
$('#fileAi').addEventListener('change', e => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) aiAddFiles(fs); });

// Teclado físico (Mac): números, coma, borrar, Enter en la nota, Escape
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && sheetKind && sheetKind !== 'onboard') { closeSheet(); return; }
  if (sheetKind === 'add' && add) {
    if (e.target.id === 'note') { if (e.key === 'Enter') e.target.blur(); return; }
    if (e.target.matches('input, textarea, select')) return;
    if (/^\d$/.test(e.key)) keyPress(e.key);
    else if (e.key === ',' || e.key === '.') keyPress(',');
    else if (e.key === 'Backspace') keyPress('del');
    return;
  }
  if (!sheetKind && (e.key === 'n' || e.key === '+') && !e.target.matches('input, textarea')) { e.preventDefault(); openAdd(); }
  if (sheetKind === 'ask' && e.key === 'Enter') actions.askOk();
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') { applyRecurring(); if (!sheetKind) render(); syncSoon(300); }
  else if (sync.timer) syncNow();
});
window.addEventListener('online', () => syncSoon(300));
setInterval(() => { if (document.visibilityState === 'visible' && !sync.busy && Date.now() - sync.last > 55000) syncSoon(0); }, 60000);
setInterval(() => { if (!sheetKind && ui.tab === 'metas') render(); }, 60000);

/* Atajo por URL: …/#add?importe=12,50&concepto=Mercadona (útil desde Atajos o el Mac) */
function handleHash() {
  if (!location.hash.startsWith('#add')) return;
  const q = new URLSearchParams(location.hash.split('?')[1] || '');
  const amount = Math.abs(parseAmount(q.get('importe') || q.get('i') || ''));
  openAdd({ amount: amount > 0 ? amount : 0, note: q.get('concepto') || q.get('c') || '' });
  history.replaceState(null, '', location.pathname + location.search);
}

/* ---------- Arranque ---------- */
function init() {
  applyTheme();
  if (state.settings.onboarded && !state.settings.startedOn) {
    state.settings.startedOn = state.expenses.reduce((m, e) => (e.date < m ? e.date : m), todayISO());
    save();
  }
  applyRecurring();
  if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});
  render();
  if (!state.settings.onboarded) { ob = { step: 0 }; openSheet(obHTML(), 'onboard'); }
  else handleHash();
  checkBadges();
  syncSoon(300);
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(() => {});
}
init();
