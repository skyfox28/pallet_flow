'use strict';

/* ---------------------------------------------------------------------- *
 * Storage
 * ---------------------------------------------------------------------- */

const STORAGE_KEYS = {
  zm19: 'pf_zm19_rows',
  vttk: 'pf_vttk_map',
  mouvements: 'pf_mouvements',
  retours: 'pf_retours',
  refSupports: 'pf_ref_supports',
  refTransporteurs: 'pf_ref_transporteurs',
  batches: 'pf_import_batches',
};

const DEFAULT_REF_SUPPORTS = {
  FE025: { label: 'Palette Europe', quantite: 1 },
  FE050: { label: '2 Palette Europe', quantite: 2 },
  FE075: { label: '3 Palette Europe', quantite: 3 },
  FE100: { label: '4 Palette Europe', quantite: 4 },
  FE125: { label: '5 Palette Europe', quantite: 5 },
  FE150: { label: '6 Palette Europe', quantite: 6 },
  FE175: { label: '7 Palette Europe', quantite: 7 },
  FE200: { label: '8 Palette Europe', quantite: 8 },
  FH017: { label: 'Demi Palette', quantite: 0.5 },
  FA000: { label: 'Sans Support', quantite: 0 },
};

function load(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (raw === null) return fallback;
    return JSON.parse(raw);
  } catch (e) {
    console.error('Erreur de lecture du stockage', key, e);
    return fallback;
  }
}

function save(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

let state = {
  zm19: load(STORAGE_KEYS.zm19, []),
  vttk: load(STORAGE_KEYS.vttk, {}),
  mouvements: load(STORAGE_KEYS.mouvements, []),
  retours: load(STORAGE_KEYS.retours, []),
  refSupports: load(STORAGE_KEYS.refSupports, JSON.parse(JSON.stringify(DEFAULT_REF_SUPPORTS))),
  refTransporteurs: load(STORAGE_KEYS.refTransporteurs, {}),
  batches: load(STORAGE_KEYS.batches, []),
};

function persist() {
  save(STORAGE_KEYS.zm19, state.zm19);
  save(STORAGE_KEYS.vttk, state.vttk);
  save(STORAGE_KEYS.mouvements, state.mouvements);
  save(STORAGE_KEYS.retours, state.retours);
  save(STORAGE_KEYS.refSupports, state.refSupports);
  save(STORAGE_KEYS.refTransporteurs, state.refTransporteurs);
  save(STORAGE_KEYS.batches, state.batches);
}

/* ---------------------------------------------------------------------- *
 * Excel helpers
 * ---------------------------------------------------------------------- */

function normalizeHeader(h) {
  return String(h || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function findColumn(headers, matcher) {
  for (let i = 0; i < headers.length; i++) {
    if (matcher(normalizeHeader(headers[i]))) return i;
  }
  return -1;
}

function formatDateCell(value) {
  if (value instanceof Date && !isNaN(value)) {
    const y = value.getUTCFullYear();
    const m = String(value.getUTCMonth() + 1).padStart(2, '0');
    const d = String(value.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(value || '').trim();
  return s;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sheetToRows(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
}

/* ---------------------------------------------------------------------- *
 * Import ZM19
 * ---------------------------------------------------------------------- */

async function importZM19(file) {
  const wb = await readWorkbook(file);
  const rows = sheetToRows(wb);
  if (!rows.length) throw new Error('Fichier vide.');

  const headers = rows[0];
  const idxSupport = findColumn(headers, (h) => h === 'ute stckage');
  const idxTypPal = findColumn(headers, (h) => h === 'typ pal');
  const idxDocument = findColumn(headers, (h) => h === 'document');
  const idxExp = findColumn(headers, (h) => h === 'exp');
  const idxDate = findColumn(headers, (h) => h === 'date');
  const idxPalletCount = findColumn(headers, (h) => h.includes('pallet count'));

  if (idxTypPal === -1 || idxDocument === -1 || idxExp === -1 || idxDate === -1) {
    throw new Error('Colonnes attendues introuvables (Typ Pal / Document / Exp / Date). Vérifie le fichier.');
  }

  const existingById = new Map(state.zm19.map((r) => [r.id, r]));
  let kept = 0, skipped = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;

    const typePal = String(row[idxTypPal] || '').trim().toUpperCase();
    const document = String(row[idxDocument] || '').trim();
    const transport = String(row[idxExp] || '').trim();
    const date = formatDateCell(row[idxDate]);
    const support = idxSupport !== -1 ? String(row[idxSupport] || '').trim() : '';
    const palletCount = idxPalletCount !== -1 ? Number(row[idxPalletCount]) || 1 : 1;

    if (!typePal || !document || !transport || !date) { skipped++; continue; }

    const id = support ? `su-${support}` : `zm19-${document}-${transport}-${typePal}-${i}`;
    existingById.set(id, { id, date, document, transport, typePal, palletCount });
    kept++;
  }

  state.zm19 = Array.from(existingById.values());
  state.batches.unshift({
    date: new Date().toISOString(),
    type: 'ZM19',
    filename: file.name,
    rows: kept,
    skipped,
  });
  persist();
  return { kept, skipped, total: state.zm19.length };
}

/* ---------------------------------------------------------------------- *
 * Import VTTK
 * ---------------------------------------------------------------------- */

async function importVTTK(file) {
  const wb = await readWorkbook(file);
  const rows = sheetToRows(wb);
  if (!rows.length) throw new Error('Fichier vide.');

  const headers = rows[0];
  const idxTransport = findColumn(headers, (h) => h.startsWith('n') && h.includes('transport') && h.includes('du'));
  const idxItineraire = findColumn(headers, (h) => h.includes('itineraire'));

  if (idxTransport === -1 || idxItineraire === -1) {
    throw new Error('Colonnes attendues introuvables (Nº du transport / Itinéraire transport). Vérifie le fichier.');
  }

  let kept = 0, skipped = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;

    const transport = String(row[idxTransport] || '').trim();
    const itineraire = String(row[idxItineraire] || '').trim().toUpperCase();

    if (!transport || !itineraire) { skipped++; continue; }

    state.vttk[transport] = itineraire;
    if (!state.refTransporteurs[itineraire]) {
      state.refTransporteurs[itineraire] = { nom: '' };
    }
    kept++;
  }

  state.batches.unshift({
    date: new Date().toISOString(),
    type: 'VTTK',
    filename: file.name,
    rows: kept,
    skipped,
  });
  persist();
  return { kept, skipped };
}

/* ---------------------------------------------------------------------- *
 * Rapprochement
 * ---------------------------------------------------------------------- */

function rebuildMouvements() {
  const mouvements = [];

  for (const row of state.zm19) {
    let support = state.refSupports[row.typePal];
    if (!support) {
      support = { label: 'À définir (code inconnu)', quantite: null };
      state.refSupports[row.typePal] = support;
    }
    const quantite = support.quantite != null ? support.quantite * (row.palletCount || 1) : null;

    const itineraire = state.vttk[row.transport] || null;
    let transporteurNom = 'Non identifié';
    let statut = 'attente';

    if (itineraire) {
      let ref = state.refTransporteurs[itineraire];
      if (!ref) {
        ref = { nom: '' };
        state.refTransporteurs[itineraire] = ref;
      }
      if (ref.nom) {
        transporteurNom = ref.nom;
        statut = quantite != null ? 'ok' : 'attente';
      } else {
        transporteurNom = `(à nommer : ${itineraire})`;
        statut = 'attente';
      }
    }

    mouvements.push({
      id: row.id,
      date: row.date,
      livraison: row.document,
      transport: row.transport,
      typePal: row.typePal,
      typePalLabel: support.label,
      quantite,
      itineraire,
      transporteur: transporteurNom,
      statut,
    });
  }

  mouvements.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  state.mouvements = mouvements;
  persist();

  const ok = mouvements.filter((m) => m.statut === 'ok').length;
  return { total: mouvements.length, ok, attente: mouvements.length - ok };
}

/* ---------------------------------------------------------------------- *
 * Retours manuels
 * ---------------------------------------------------------------------- */

function addRetour(date, transporteur, quantite, reference) {
  state.retours.push({
    id: `retour-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    date, transporteur, quantite, reference,
  });
  persist();
}

function deleteRetour(id) {
  state.retours = state.retours.filter((r) => r.id !== id);
  persist();
}

/* ---------------------------------------------------------------------- *
 * Solde
 * ---------------------------------------------------------------------- */

function computeSolde() {
  const buckets = new Map(); // key -> { nom, itineraires:Set, sorties, retours }

  function bucketFor(key, nom) {
    if (!buckets.has(key)) {
      buckets.set(key, { nom, itineraires: new Set(), sorties: 0, retours: 0 });
    }
    return buckets.get(key);
  }

  for (const m of state.mouvements) {
    if (m.quantite == null) continue;
    let key, nom;
    if (!m.itineraire) {
      key = 'NON_IDENTIFIE';
      nom = 'Non identifié';
    } else if (state.refTransporteurs[m.itineraire] && state.refTransporteurs[m.itineraire].nom) {
      nom = state.refTransporteurs[m.itineraire].nom;
      key = `nom:${nom}`;
    } else {
      key = `code:${m.itineraire}`;
      nom = `(à nommer : ${m.itineraire})`;
    }
    const b = bucketFor(key, nom);
    if (m.itineraire) b.itineraires.add(m.itineraire);
    b.sorties += m.quantite;
  }

  for (const r of state.retours) {
    const key = `nom:${r.transporteur}`;
    const b = bucketFor(key, r.transporteur);
    b.retours += Number(r.quantite) || 0;
  }

  return Array.from(buckets.values())
    .map((b) => ({
      transporteur: b.nom,
      itineraires: Array.from(b.itineraires).join(', '),
      sorties: b.sorties,
      retours: b.retours,
      solde: b.sorties - b.retours,
    }))
    .sort((a, b) => a.transporteur.localeCompare(b.transporteur));
}

/* ---------------------------------------------------------------------- *
 * Rendering — Import tab
 * ---------------------------------------------------------------------- */

function renderBatches() {
  const tbody = document.querySelector('#table-batches tbody');
  tbody.innerHTML = '';
  for (const b of state.batches) {
    const tr = document.createElement('tr');
    const d = new Date(b.date);
    tr.innerHTML = `
      <td>${d.toLocaleString('fr-FR')}</td>
      <td>${b.type}</td>
      <td>${b.filename}</td>
      <td>${b.rows}${b.skipped ? ` (+${b.skipped} ignorées)` : ''}</td>
    `;
    tbody.appendChild(tr);
  }
}

/* ---------------------------------------------------------------------- *
 * Rendering — Mouvements tab
 * ---------------------------------------------------------------------- */

function distinctSorted(arr) {
  return Array.from(new Set(arr)).sort();
}

function populateMouvementsFilters() {
  const years = distinctSorted(state.mouvements.map((m) => m.date.slice(0, 4)).filter(Boolean));
  const months = distinctSorted(state.mouvements.map((m) => m.date.slice(5, 7)).filter(Boolean));
  const transporteurs = distinctSorted(state.mouvements.map((m) => m.transporteur));

  fillSelect('#filter-annee', years, years[0] ? '' : '');
  fillSelect('#filter-mois', months);
  fillSelect('#filter-transporteur', transporteurs);
}

function fillSelect(selector, values, keepValue) {
  const select = document.querySelector(selector);
  const current = keepValue !== undefined ? keepValue : select.value;
  select.innerHTML = `<option value="">${select.id === 'filter-mois' ? 'Tous' : 'Toutes'}</option>`;
  for (const v of values) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  }
  if (values.includes(current)) select.value = current;
}

function getFilteredMouvements() {
  const annee = document.querySelector('#filter-annee').value;
  const mois = document.querySelector('#filter-mois').value;
  const transporteur = document.querySelector('#filter-transporteur').value;
  const statut = document.querySelector('#filter-statut').value;

  return state.mouvements.filter((m) => {
    if (annee && m.date.slice(0, 4) !== annee) return false;
    if (mois && m.date.slice(5, 7) !== mois) return false;
    if (transporteur && m.transporteur !== transporteur) return false;
    if (statut && m.statut !== statut) return false;
    return true;
  });
}

function renderMouvements() {
  populateMouvementsFilters();
  const rows = getFilteredMouvements();
  const tbody = document.querySelector('#table-mouvements tbody');
  tbody.innerHTML = '';

  let totalQte = 0, nbAttente = 0;

  for (const m of rows) {
    if (m.quantite != null) totalQte += m.quantite;
    if (m.statut === 'attente') nbAttente++;

    const tr = document.createElement('tr');
    const badge = m.statut === 'ok'
      ? '<span class="badge badge-ok">OK</span>'
      : '<span class="badge badge-warn">En attente</span>';
    tr.innerHTML = `
      <td>${m.date}</td>
      <td>${m.livraison}</td>
      <td>${m.transport}</td>
      <td>${m.typePalLabel} (${m.typePal})</td>
      <td>${m.quantite ?? '?'}</td>
      <td>${m.itineraire || '—'}</td>
      <td>${m.transporteur}</td>
      <td>${badge}</td>
    `;
    tbody.appendChild(tr);
  }

  document.querySelector('#mouvements-totals').textContent =
    `${rows.length} mouvement(s) — ${totalQte} palette(s) — ${nbAttente} en attente de rapprochement`;
}

/* ---------------------------------------------------------------------- *
 * Rendering — Retours tab
 * ---------------------------------------------------------------------- */

function namedTransporteurs() {
  return distinctSorted(
    Object.values(state.refTransporteurs).map((r) => r.nom).filter(Boolean)
  );
}

function renderRetoursForm() {
  const select = document.querySelector('#retour-transporteur');
  const current = select.value;
  select.innerHTML = '<option value="">Sélectionner…</option>';
  for (const nom of namedTransporteurs()) {
    const opt = document.createElement('option');
    opt.value = nom;
    opt.textContent = nom;
    select.appendChild(opt);
  }
  if (namedTransporteurs().includes(current)) select.value = current;
}

function renderRetoursTable() {
  const tbody = document.querySelector('#table-retours tbody');
  tbody.innerHTML = '';
  const sorted = [...state.retours].sort((a, b) => (a.date < b.date ? 1 : -1));
  for (const r of sorted) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.date}</td>
      <td>${r.transporteur}</td>
      <td>${r.quantite}</td>
      <td>${r.reference || ''}</td>
      <td class="row-actions"><button data-id="${r.id}">Supprimer</button></td>
    `;
    tr.querySelector('button').addEventListener('click', () => {
      deleteRetour(r.id);
      renderRetoursTable();
      renderSolde();
    });
    tbody.appendChild(tr);
  }
}

/* ---------------------------------------------------------------------- *
 * Rendering — Solde tab
 * ---------------------------------------------------------------------- */

function renderSolde() {
  const data = computeSolde();
  const tbody = document.querySelector('#table-solde tbody');
  tbody.innerHTML = '';
  for (const row of data) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.transporteur}</td>
      <td>${row.itineraires}</td>
      <td>${row.sorties}</td>
      <td>${row.retours}</td>
      <td><strong>${row.solde}</strong></td>
    `;
    tbody.appendChild(tr);
  }
}

/* ---------------------------------------------------------------------- *
 * Rendering — Référentiels tab
 * ---------------------------------------------------------------------- */

function renderRefSupports() {
  const tbody = document.querySelector('#table-ref-supports tbody');
  tbody.innerHTML = '';
  for (const code of Object.keys(state.refSupports).sort()) {
    const s = state.refSupports[code];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${code}</td>
      <td><input type="text" value="${escapeAttr(s.label)}" data-field="label"></td>
      <td><input type="number" step="0.5" value="${s.quantite ?? ''}" data-field="quantite"></td>
      <td class="row-actions"><button>Supprimer</button></td>
    `;
    tr.querySelector('[data-field="label"]').addEventListener('change', (e) => {
      s.label = e.target.value;
      persist();
      rebuildMouvements();
      renderAll();
    });
    tr.querySelector('[data-field="quantite"]').addEventListener('change', (e) => {
      s.quantite = e.target.value === '' ? null : Number(e.target.value);
      persist();
      rebuildMouvements();
      renderAll();
    });
    tr.querySelector('button').addEventListener('click', () => {
      delete state.refSupports[code];
      persist();
      rebuildMouvements();
      renderAll();
    });
    tbody.appendChild(tr);
  }
}

function renderRefTransporteurs() {
  const tbody = document.querySelector('#table-ref-transporteurs tbody');
  tbody.innerHTML = '';
  for (const code of Object.keys(state.refTransporteurs).sort()) {
    const t = state.refTransporteurs[code];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${code}</td>
      <td><input type="text" value="${escapeAttr(t.nom)}" placeholder="À renseigner" data-field="nom"></td>
      <td class="row-actions"><button>Supprimer</button></td>
    `;
    tr.querySelector('[data-field="nom"]').addEventListener('change', (e) => {
      t.nom = e.target.value.trim();
      persist();
      rebuildMouvements();
      renderAll();
    });
    tr.querySelector('button').addEventListener('click', () => {
      delete state.refTransporteurs[code];
      persist();
      rebuildMouvements();
      renderAll();
    });
    tbody.appendChild(tr);
  }
}

function escapeAttr(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

/* ---------------------------------------------------------------------- *
 * Export
 * ---------------------------------------------------------------------- */

function exportMouvementsXLSX() {
  const rows = getFilteredMouvements().map((m) => ({
    Date: m.date,
    Livraison: m.livraison,
    'N° transport': m.transport,
    'Type palette': m.typePalLabel,
    Code: m.typePal,
    Quantité: m.quantite,
    Itinéraire: m.itineraire || '',
    Transporteur: m.transporteur,
    Statut: m.statut,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Mouvements');
  XLSX.writeFile(wb, `mouvements_palettes_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function exportSoldeXLSX() {
  const rows = computeSolde().map((r) => ({
    Transporteur: r.transporteur,
    'Itinéraire(s)': r.itineraires,
    Sorties: r.sorties,
    Retours: r.retours,
    Solde: r.solde,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Solde');
  XLSX.writeFile(wb, `solde_palettes_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ---------------------------------------------------------------------- *
 * Tabs
 * ---------------------------------------------------------------------- */

function setupTabs() {
  const buttons = document.querySelectorAll('.tab-btn');
  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      buttons.forEach((b) => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach((p) => p.classList.remove('active'));
      btn.classList.add('active');
      document.querySelector(`#tab-${btn.dataset.tab}`).classList.add('active');
    });
  });
}

/* ---------------------------------------------------------------------- *
 * Wiring
 * ---------------------------------------------------------------------- */

function showResult(selector, html) {
  document.querySelector(selector).innerHTML = html;
}

function setupImport() {
  document.querySelector('#file-zm19').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const res = await importZM19(file);
      showResult('#zm19-import-result',
        `<span class="ok">${res.kept} ligne(s) importée(s)${res.skipped ? `, ${res.skipped} ignorée(s)` : ''}. Total en base : ${res.total}.</span>`);
      renderBatches();
    } catch (err) {
      showResult('#zm19-import-result', `<span class="err">${err.message}</span>`);
    }
    e.target.value = '';
  });

  document.querySelector('#file-vttk').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const res = await importVTTK(file);
      showResult('#vttk-import-result',
        `<span class="ok">${res.kept} ligne(s) importée(s)${res.skipped ? `, ${res.skipped} ignorée(s) (itinéraire manquant)` : ''}.</span>`);
      renderBatches();
      renderRefTransporteurs();
      renderRetoursForm();
    } catch (err) {
      showResult('#vttk-import-result', `<span class="err">${err.message}</span>`);
    }
    e.target.value = '';
  });

  document.querySelector('#btn-rapprocher').addEventListener('click', () => {
    const res = rebuildMouvements();
    showResult('#rapprochement-result',
      `<span class="ok">${res.total} mouvement(s) — ${res.ok} identifié(s), ${res.attente} en attente.</span>`);
    renderAll();
  });

  document.querySelector('#btn-reset-all').addEventListener('click', () => {
    if (!confirm('Supprimer toutes les données importées, mouvements, retours et référentiels ?')) return;
    Object.values(STORAGE_KEYS).forEach((k) => localStorage.removeItem(k));
    state = {
      zm19: [], vttk: {}, mouvements: [], retours: [],
      refSupports: JSON.parse(JSON.stringify(DEFAULT_REF_SUPPORTS)),
      refTransporteurs: {}, batches: [],
    };
    renderAll();
  });
}

function setupMouvements() {
  ['#filter-annee', '#filter-mois', '#filter-transporteur', '#filter-statut'].forEach((sel) => {
    document.querySelector(sel).addEventListener('change', renderMouvements);
  });
  document.querySelector('#btn-export-mouvements').addEventListener('click', exportMouvementsXLSX);
}

function setupRetours() {
  document.querySelector('#form-retour').addEventListener('submit', (e) => {
    e.preventDefault();
    const date = document.querySelector('#retour-date').value;
    const transporteur = document.querySelector('#retour-transporteur').value;
    const quantite = Number(document.querySelector('#retour-quantite').value);
    const reference = document.querySelector('#retour-reference').value.trim();
    if (!date || !transporteur || !quantite) return;
    addRetour(date, transporteur, quantite, reference);
    e.target.reset();
    renderRetoursTable();
    renderSolde();
  });
}

function setupSolde() {
  document.querySelector('#btn-export-solde').addEventListener('click', exportSoldeXLSX);
}

function setupReferentiels() {
  document.querySelector('#btn-add-support').addEventListener('click', () => {
    const code = prompt('Code du nouveau type de support (ex: FE300) :');
    if (!code) return;
    const key = code.trim().toUpperCase();
    if (!state.refSupports[key]) state.refSupports[key] = { label: '', quantite: null };
    persist();
    renderRefSupports();
  });

  document.querySelector('#btn-add-transporteur').addEventListener('click', () => {
    const code = prompt('Code itinéraire transport (ex: FRSEXP) :');
    if (!code) return;
    const key = code.trim().toUpperCase();
    if (!state.refTransporteurs[key]) state.refTransporteurs[key] = { nom: '' };
    persist();
    renderRefTransporteurs();
  });
}

/* ---------------------------------------------------------------------- *
 * Init
 * ---------------------------------------------------------------------- */

function renderAll() {
  renderBatches();
  renderMouvements();
  renderRetoursForm();
  renderRetoursTable();
  renderSolde();
  renderRefSupports();
  renderRefTransporteurs();
}

function init() {
  setupTabs();
  setupImport();
  setupMouvements();
  setupRetours();
  setupSolde();
  setupReferentiels();
  renderAll();
}

document.addEventListener('DOMContentLoaded', init);
