/* ============================================
   Colliers Denver Design Studio v3 — App Logic
   HTML-to-PDF output engine with CSS @page
   ============================================ */

// ============ STATE ============
const state = {
  docType: 'om',
  propType: 'industrial',
  currentStep: 1,
  uploadedFiles: {},
  photoDataURLs: [],
  apiKey: localStorage.getItem('colliers_api_key') || '',
  projects: JSON.parse(localStorage.getItem('colliers_projects') || '[]'),
  tenantRows: 0,
  pageSettings: { widthIn: 11, heightIn: 8.5, orientation: 'landscape' },
  design: {
    colors: { primary: '#001a4d', secondary: '#0057b8', accent: '#c8a96e', text: '#1a2332', bg: '#ffffff', rule: '#0057b8' },
    fonts: { heading: '', body: '', number: '' },
    layout: 'classic',
    aiChosenFonts: null,
  },
  narrativeContext: '',
};

// ============ NAVIGATION ============
function navigate(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const t = document.getElementById('page-' + page);
  if (t) t.classList.add('active');
  const n = document.querySelector(`[data-page="${page}"]`);
  if (n) n.classList.add('active');
  if (page === 'dashboard') refreshDashboard();
  if (page === 'settings') loadSettings();
  window.scrollTo(0, 0);
}
document.querySelectorAll('.nav-item:not(.coming-soon)').forEach(item => {
  item.addEventListener('click', e => { e.preventDefault(); navigate(item.dataset.page); });
});

// ============ STEPS ============
function goStep(n) {
  document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.step').forEach(s => {
    const sn = parseInt(s.dataset.step);
    s.classList.remove('active', 'completed');
    if (sn === n) s.classList.add('active');
    if (sn < n) s.classList.add('completed');
  });
  const panel = document.getElementById('step-panel-' + n);
  if (panel) panel.classList.add('active');
  state.currentStep = n;
  if (n === 6) buildReviewSummary();
  window.scrollTo(0, 0);
}

// ============ DOC/PROP TYPE ============
function selectDocType(t) { state.docType = t; document.querySelectorAll('.doc-type-card').forEach(c => c.classList.toggle('active', c.dataset.type === t)); }
function selectPropType(t) { state.propType = t; document.querySelectorAll('.prop-type-btn').forEach(b => b.classList.toggle('active', b.dataset.ptype === t)); }

// ============ PAGE SETTINGS ============
function setOrientation(o) {
  state.pageSettings.orientation = o;
  document.getElementById('btn-landscape').classList.toggle('active', o === 'landscape');
  document.getElementById('btn-portrait').classList.toggle('active', o === 'portrait');
  const { widthIn, heightIn } = state.pageSettings;
  if (o === 'landscape' && heightIn > widthIn) { state.pageSettings.widthIn = heightIn; state.pageSettings.heightIn = widthIn; }
  if (o === 'portrait' && widthIn > heightIn) { state.pageSettings.widthIn = heightIn; state.pageSettings.heightIn = widthIn; }
  updatePagePreview();
}
function applyPagePreset(val) {
  const cr = document.getElementById('custom-size-row');
  if (val === 'custom') { cr.style.display = 'flex'; return; }
  cr.style.display = 'none';
  const p = val.split('x').map(Number);
  state.pageSettings.widthIn = p[0]; state.pageSettings.heightIn = p[1];
  state.pageSettings.orientation = p[0] >= p[1] ? 'landscape' : 'portrait';
  document.getElementById('btn-landscape').classList.toggle('active', state.pageSettings.orientation === 'landscape');
  document.getElementById('btn-portrait').classList.toggle('active', state.pageSettings.orientation === 'portrait');
  updatePagePreview();
}
function updateCustomSize() {
  const w = parseFloat(document.getElementById('custom-width').value) || 11;
  const h = parseFloat(document.getElementById('custom-height').value) || 8.5;
  state.pageSettings.widthIn = w; state.pageSettings.heightIn = h;
  state.pageSettings.orientation = w >= h ? 'landscape' : 'portrait';
  updatePagePreview();
}
function updatePagePreview() {
  const { widthIn, heightIn } = state.pageSettings;
  const r = document.getElementById('page-preview-rect');
  const l = document.getElementById('page-preview-label');
  if (!r || !l) return;
  const scale = 38;
  r.style.width = Math.round(widthIn * scale) + 'px';
  r.style.height = Math.round(heightIn * scale) + 'px';
  l.textContent = `${widthIn}" × ${heightIn}" ${state.pageSettings.orientation}`;
}

// ============ COLOR PALETTE ============
function syncColorInput(key) {
  const picker = document.getElementById('color-' + key + '-picker');
  const text = document.getElementById('color-' + key);
  const preview = document.getElementById('color-preview-' + key);
  if (picker && text) text.value = picker.value;
  if (preview && picker) preview.style.background = picker.value;
  state.design.colors[key] = picker?.value || state.design.colors[key];
}
function syncColorPicker(key) {
  const text = document.getElementById('color-' + key);
  const picker = document.getElementById('color-' + key + '-picker');
  const preview = document.getElementById('color-preview-' + key);
  const val = text?.value;
  if (val && /^#[0-9a-fA-F]{6}$/.test(val)) {
    if (picker) picker.value = val;
    if (preview) preview.style.background = val;
    state.design.colors[key] = val;
  }
}
function setColor(key, val) {
  state.design.colors[key] = val;
  const picker = document.getElementById('color-' + key + '-picker');
  const text = document.getElementById('color-' + key);
  const preview = document.getElementById('color-preview-' + key);
  if (picker) picker.value = val;
  if (text) text.value = val;
  if (preview) preview.style.background = val;
}
function applyPalettePreset(name) {
  const presets = {
    colliers: { primary: '#001a4d', secondary: '#0057b8', accent: '#a8d4f5', text: '#0a1628', bg: '#ffffff', rule: '#0057b8' },
    dark:     { primary: '#0f0f0f', secondary: '#c8a96e', accent: '#e8d4a8', text: '#f0ece0', bg: '#1a1a1a', rule: '#c8a96e' },
    earth:    { primary: '#3b2a1a', secondary: '#7a4f2e', accent: '#c8a96e', text: '#2a1e10', bg: '#fdf8f0', rule: '#c8a96e' },
    slate:    { primary: '#1e2d3d', secondary: '#4a7fa5', accent: '#8fb8d4', text: '#1e2d3d', bg: '#f5f8fb', rule: '#4a7fa5' },
    forest:   { primary: '#1a3a2a', secondary: '#2d6e4e', accent: '#7ab894', text: '#1a3a2a', bg: '#f5faf7', rule: '#2d6e4e' },
    clear:    { primary: '#001a4d', secondary: '#0057b8', accent: '#c8a96e', text: '#1a2332', bg: '#ffffff', rule: '#0057b8' },
  };
  const p = presets[name];
  if (!p) return;
  Object.entries(p).forEach(([k, v]) => setColor(k, v));
}

// ============ FONTS ============
const GOOGLE_FONTS = ['Playfair Display', 'Merriweather', 'Lora', 'Cormorant Garamond', 'Montserrat', 'Raleway', 'Oswald', 'Bebas Neue', 'Inter', 'DM Sans', 'Source Sans 3', 'Lato', 'Open Sans', 'Nunito', 'Rajdhani', 'Barlow'];
const loadedFonts = new Set();

function loadGoogleFont(fontName) {
  if (!fontName || fontName === 'Georgia' || loadedFonts.has(fontName)) return;
  loadedFonts.add(fontName);
  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = `https://fonts.googleapis.com/css2?family=${fontName.replace(/ /g, '+')}:wght@300;400;500;600;700&display=swap`;
  document.head.appendChild(link);
}

function updateFontPreview() {
  const heading = document.getElementById('font-heading').value;
  const body = document.getElementById('font-body').value;
  const number = document.getElementById('font-number').value;
  state.design.fonts = { heading, body, number };
  if (heading) loadGoogleFont(heading);
  if (body) loadGoogleFont(body);
  if (number) loadGoogleFont(number);
  const fph = document.getElementById('fp-heading');
  const fpb = document.getElementById('fp-body');
  const fps = document.getElementById('fp-stats');
  if (fph) fph.style.fontFamily = heading ? `'${heading}', serif` : "'Playfair Display', serif";
  if (fpb) fpb.style.fontFamily = body ? `'${body}', sans-serif` : 'Inter, sans-serif';
  if (fps) fps.style.fontFamily = (number || heading) ? `'${number || heading}', serif` : "'Playfair Display', serif";
}

// ============ LAYOUT ============
function selectLayout(l) {
  state.design.layout = l;
  document.querySelectorAll('.layout-card').forEach(c => c.classList.toggle('active', c.dataset.layout === l));
}

// ============ TEMPLATE STYLE EXTRACTION ============
async function extractTemplateStyle() {
  const files = state.uploadedFiles['template'] || [];
  if (!files.length) { showToast('Upload a template file first'); return; }
  const statusEl = document.getElementById('extract-status-template');
  statusEl.innerHTML = '<span class="status-spinner" style="display:inline-block;width:11px;height:11px;border:2px solid rgba(0,87,184,0.2);border-top-color:#0057b8;border-radius:50%;animation:spin 0.8s linear infinite;margin-right:5px;vertical-align:middle;"></span>Analyzing template style with AI...';

  const file = files[0];
  let base64 = null;
  let mediaType = 'application/pdf';

  try {
    base64 = await fileToBase64(file);
    mediaType = file.type || 'application/pdf';
  } catch (e) {
    statusEl.textContent = '✗ Could not read file'; return;
  }

  const prompt = `Analyze this document's visual design and return ONLY a JSON object describing its style. Extract:
- Color palette (primary background color, accent/highlight color, text color, secondary color, rule/divider color)
- Typography (heading font name if identifiable, body font name, overall style: serif/sans/modern/classic/luxury/minimal)
- Layout approach: classic (wide content + sidebar), editorial (bold header, single column), magazine (two columns), or minimal (clean whitespace)
- Overall aesthetic description in 1 sentence

Return ONLY this JSON, no other text:
{
  "colors": { "primary": "#hexcode", "secondary": "#hexcode", "accent": "#hexcode", "text": "#hexcode", "bg": "#hexcode", "rule": "#hexcode" },
  "fonts": { "headingStyle": "serif|sans|display", "bodyStyle": "serif|sans", "suggestedHeading": "font name or empty string", "suggestedBody": "font name or empty string" },
  "layout": "classic|editorial|magazine|minimal",
  "aesthetic": "one sentence description"
}`;

  try {
    const messages = [{ role: 'user', content: [{ type: 'document', source: { type: 'base64', media_type: mediaType, data: base64 } }, { type: 'text', text: prompt }] }];
    const data = await callClaude({ model: 'claude-sonnet-4-20250514', max_tokens: 1000, messages });
    const text = data.content?.[0]?.text || '';
    const parsed = JSON.parse(text.replace(/```json|```/g, '').trim());

    if (parsed.colors) Object.entries(parsed.colors).forEach(([k, v]) => { if (/^#[0-9a-fA-F]{6}$/.test(v)) setColor(k, v); });

    const fontMap = { heading: parsed.fonts?.suggestedHeading, body: parsed.fonts?.suggestedBody };
    if (fontMap.heading) {
      const match = GOOGLE_FONTS.find(f => f.toLowerCase().includes(fontMap.heading.toLowerCase().split(' ')[0]));
      if (match) { document.getElementById('font-heading').value = match; }
    }
    if (fontMap.body) {
      const match = GOOGLE_FONTS.find(f => f.toLowerCase().includes(fontMap.body.toLowerCase().split(' ')[0]));
      if (match) { document.getElementById('font-body').value = match; }
    }
    if (parsed.layout) selectLayout(parsed.layout);
    updateFontPreview();

    statusEl.textContent = `✓ Style extracted — ${parsed.aesthetic || 'Design settings applied'}`;
    statusEl.className = 'extract-status ok';
    showToast('Template style applied to Design Settings');
  } catch (e) {
    statusEl.textContent = '✗ ' + (e.message || 'Unknown error');
    statusEl.className = 'extract-status err';
  }
}

// ============ FILE UPLOADS ============
function triggerStepUpload(key) { document.getElementById('step-file-' + key)?.click(); }

function handleStepUpload(key, input) {
  const files = Array.from(input.files);
  if (!files.length) return;
  state.uploadedFiles[key] = [...(state.uploadedFiles[key] || []), ...files];
  const zone = input.closest('.iupload');
  if (zone) zone.classList.add('has-files');
  if (key === 'photos') { handlePhotoUploads(files); return; }
  const listEl = document.getElementById('step-files-' + key);
  if (listEl) files.forEach(f => { const t = document.createElement('div'); t.className = 'ifile-tag'; t.innerHTML = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> ${f.name}`; listEl.appendChild(t); });
  const btnMap = { 'property-doc': 'btn-extract-property', 'financials': 'btn-extract-financials', 'narrative': 'btn-extract-narrative', 'template': 'btn-extract-template' };
  const btn = document.getElementById(btnMap[key]);
  if (btn) btn.style.display = 'inline-flex';
}

function handlePhotoUploads(files) {
  const grid = document.getElementById('photo-preview-grid');
  files.forEach(f => {
    if (!f.type.startsWith('image/')) return;
    const reader = new FileReader();
    reader.onload = e => { state.photoDataURLs.push(e.target.result); if (grid) { const img = document.createElement('img'); img.className = 'photo-thumb'; img.src = e.target.result; grid.appendChild(img); } };
    reader.readAsDataURL(f);
  });
}

async function fileToBase64(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = e => res(e.target.result.split(',')[1]); r.onerror = rej; r.readAsDataURL(file); });
}
async function readFileAsText(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = e => res(e.target.result); r.onerror = rej; r.readAsText(file); });
}

// ============ SHARED API CALL ============
async function callClaude(body) {
  const res = await fetch('/api/claude', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!res.ok) {
    const msg = data?.error?.message || data?.error || `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return data;
}

// ============ AI EXTRACTION ============
async function extractFromUpload(type) {
  const statusEl = document.getElementById('extract-status-' + type);
  if (statusEl) statusEl.innerHTML = '<span class="status-spinner" style="display:inline-block;width:11px;height:11px;border:2px solid rgba(0,87,184,0.2);border-top-color:#0057b8;border-radius:50%;animation:spin 0.8s linear infinite;margin-right:5px;vertical-align:middle;"></span>Extracting...';
  const keyMap = { property: 'property-doc', financials: 'financials', narrative: 'narrative' };
  const files = state.uploadedFiles[keyMap[type]] || [];
  if (!files.length) { if (statusEl) { statusEl.textContent = 'No file uploaded.'; statusEl.className = 'extract-status'; } return; }
  let fileContent = '';
  try { fileContent = await readFileAsText(files[0]); } catch (e) { fileContent = `[File: ${files[0].name}]`; }
  let prompt = '';
  if (type === 'property') {
    prompt = `Extract property info from this document. Return ONLY JSON with these keys (empty string if not found):
{"propName":"","propAddress":"","propCity":"","propState":"","propZip":"","propCounty":"","propYear":"","propSF":"","propAcres":"","propBuildings":"","propUnits":"","propZoning":"","propParking":"","propClearHeight":"","propDesc":"","propHighlights":"","brokerName":"","brokerTitle":"","brokerPhone":"","brokerEmail":"","brokerLicense":""}
Document: ${fileContent.slice(0, 6000)}`;
  } else if (type === 'financials') {
    prompt = `Extract financial data. Return ONLY JSON:
{"finPrice":"","finPpsf":"","finGpr":"","finVacancy":"","finEgi":"","finOpex":"","finNoi":"","finCaprate":"","finOccupancy":"","finWalt":"","rentRoll":[{"tenant":"","suite":"","sf":"","leaseStart":"","leaseEnd":"","annualRent":""}]}
Document: ${fileContent.slice(0, 6000)}`;
  } else {
    prompt = `Summarize key marketing points from this document in 3-4 paragraphs for use as CRE narrative context. Document: ${fileContent.slice(0, 6000)}`;
  }
  try {
    const data = await callClaude({ model: 'claude-sonnet-4-20250514', max_tokens: 2000, messages: [{ role: 'user', content: prompt }] });
    const text = data.content?.[0]?.text || '';
    if (type === 'narrative') { state.narrativeContext = text; if (statusEl) { statusEl.textContent = '✓ Narrative saved as AI context'; statusEl.className = 'extract-status ok'; } return; }
    const parsed = JSON.parse(text.replace(/```json|```/g, '').trim());
    if (type === 'property') {
      const m = { 'prop-name': parsed.propName, 'prop-address': parsed.propAddress, 'prop-city': parsed.propCity, 'prop-state': parsed.propState, 'prop-zip': parsed.propZip, 'prop-county': parsed.propCounty, 'prop-year': parsed.propYear, 'prop-sf': parsed.propSF, 'prop-acres': parsed.propAcres, 'prop-buildings': parsed.propBuildings, 'prop-units': parsed.propUnits, 'prop-zoning': parsed.propZoning, 'prop-parking': parsed.propParking, 'prop-clearheight': parsed.propClearHeight, 'prop-desc': parsed.propDesc, 'prop-highlights': parsed.propHighlights, 'broker-name': parsed.brokerName, 'broker-title': parsed.brokerTitle, 'broker-phone': parsed.brokerPhone, 'broker-email': parsed.brokerEmail, 'broker-license': parsed.brokerLicense };
      Object.entries(m).forEach(([id, val]) => { const el = document.getElementById(id); if (el && val) el.value = val; });
    }
    if (type === 'financials') {
      const m = { 'fin-price': parsed.finPrice, 'fin-ppsf': parsed.finPpsf, 'fin-gpr': parsed.finGpr, 'fin-vacancy': parsed.finVacancy, 'fin-egi': parsed.finEgi, 'fin-opex': parsed.finOpex, 'fin-noi': parsed.finNoi, 'fin-caprate': parsed.finCaprate, 'fin-occupancy': parsed.finOccupancy, 'fin-walt': parsed.finWalt };
      Object.entries(m).forEach(([id, val]) => { const el = document.getElementById(id); if (el && val) el.value = val; });
      if (parsed.rentRoll?.length) { const tb = document.getElementById('rent-roll-body'); tb.innerHTML = ''; parsed.rentRoll.forEach(r => { state.tenantRows++; const tr = document.createElement('tr'); tr.id = 'tenant-row-' + state.tenantRows; tr.innerHTML = buildTenantRowHTML(state.tenantRows, r); tb.appendChild(tr); }); }
      calcFinancials();
    }
    if (statusEl) { statusEl.textContent = type === 'property' ? '✓ Fields populated from document' : '✓ Financials and rent roll populated'; statusEl.className = 'extract-status ok'; }
  } catch (e) {
    if (statusEl) { statusEl.textContent = '✗ ' + (e.message || 'Unknown error'); statusEl.className = 'extract-status err'; }
  }
}

// ============ FINANCIALS ============
function calcFinancials() {
  const price = parseFloat(document.getElementById('fin-price')?.value) || 0;
  const sf = parseFloat(document.getElementById('prop-sf')?.value) || 0;
  const gpr = parseFloat(document.getElementById('fin-gpr')?.value) || 0;
  const vacancy = parseFloat(document.getElementById('fin-vacancy')?.value) || 0;
  const opex = parseFloat(document.getElementById('fin-opex')?.value) || 0;
  const egi = gpr * (1 - vacancy / 100);
  const noi = egi - opex;
  const capRate = price > 0 && noi > 0 ? (noi / price) * 100 : 0;
  const ppsf = sf > 0 && price > 0 ? price / sf : 0;
  const grm = gpr > 0 && price > 0 ? price / gpr : 0;
  const trySet = (id, val) => { const el = document.getElementById(id); if (el && !el.value && val > 0) el.value = (Number.isInteger(val) ? val : val.toFixed(2)); };
  trySet('fin-egi', egi); trySet('fin-noi', noi); trySet('fin-caprate', capRate); trySet('fin-ppsf', ppsf);
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  set('calc-ppsf', ppsf > 0 ? '$' + ppsf.toFixed(2) : '—');
  set('calc-noi', noi > 0 ? '$' + fmtNum(Math.round(noi)) : '—');
  set('calc-caprate', capRate > 0 ? capRate.toFixed(2) + '%' : '—');
  set('calc-grm', grm > 0 ? grm.toFixed(2) + 'x' : '—');
}

// ============ RENT ROLL ============
function buildTenantRowHTML(id, row = {}) {
  return `<td><input type="text" value="${row.tenant||''}" placeholder="Tenant name" /></td><td><input type="text" value="${row.suite||''}" placeholder="101" style="width:60px;" /></td><td><input type="number" value="${row.sf||''}" placeholder="5000" oninput="calcRowRentSF(${id})" /></td><td><input type="date" value="${row.leaseStart||''}" /></td><td><input type="date" value="${row.leaseEnd||''}" /></td><td><input type="number" value="${row.annualRent||''}" placeholder="60000" oninput="calcRowRentSF(${id})" /></td><td><span id="rpsf-${id}" style="color:var(--text-muted);font-size:12px;">—</span></td><td><button class="btn-del-row" onclick="deleteRow(${id})">×</button></td>`;
}
function addTenantRow() {
  const tb = document.getElementById('rent-roll-body');
  const er = tb.querySelector('.empty-row'); if (er) er.remove();
  const id = ++state.tenantRows;
  const tr = document.createElement('tr'); tr.id = 'tenant-row-' + id; tr.innerHTML = buildTenantRowHTML(id); tb.appendChild(tr);
}
function calcRowRentSF(id) {
  const row = document.getElementById('tenant-row-' + id); if (!row) return;
  const inputs = row.querySelectorAll('input[type="number"]');
  const sf = parseFloat(inputs[0]?.value) || 0, rent = parseFloat(inputs[1]?.value) || 0;
  const el = document.getElementById('rpsf-' + id); if (el) el.textContent = sf > 0 && rent > 0 ? '$' + (rent / sf).toFixed(2) : '—';
}
function deleteRow(id) {
  const row = document.getElementById('tenant-row-' + id); if (row) row.remove();
  const tb = document.getElementById('rent-roll-body');
  if (!tb.querySelector('tr:not(.empty-row)')) tb.innerHTML = '<tr class="empty-row"><td colspan="8" style="text-align:center;color:#6b7fa3;padding:20px;">No tenants added yet.</td></tr>';
}
function getRentRollData() {
  return Array.from(document.querySelectorAll('#rent-roll-body tr:not(.empty-row)')).map(row => {
    const i = row.querySelectorAll('input');
    return { tenant: i[0]?.value||'', suite: i[1]?.value||'', sf: i[2]?.value||'', leaseStart: i[3]?.value||'', leaseEnd: i[4]?.value||'', annualRent: i[5]?.value||'' };
  });
}

// ============ API KEY ============
function saveApiKey() {
  const key = document.getElementById('api-key-input')?.value.trim();
  const st = document.getElementById('api-key-status');
  if (key?.startsWith('sk-ant-')) { state.apiKey = key; localStorage.setItem('colliers_api_key', key); if (st) { st.textContent = '✓ Saved'; st.className = 'api-key-status ok'; } showToast('API key saved'); }
  else { if (st) { st.textContent = 'Invalid key format'; st.className = 'api-key-status err'; } }
}
function saveApiKeyFromSettings() {
  const key = document.getElementById('settings-api-key')?.value.trim();
  if (key?.startsWith('sk-ant-')) { state.apiKey = key; localStorage.setItem('colliers_api_key', key); showToast('API key saved'); }
  else showToast('Invalid key — should start with sk-ant-');
}
function loadSettings() {
  const key = localStorage.getItem('colliers_api_key') || '';
  if (key) { const el = document.getElementById('settings-api-key'); if (el) el.value = key; }
  const s = JSON.parse(localStorage.getItem('colliers_settings') || '{}');
  if (s.firm) document.getElementById('settings-firm').value = s.firm;
  if (s.city) document.getElementById('settings-city').value = s.city;
  if (s.phone) document.getElementById('settings-phone').value = s.phone;
  if (s.disclaimer) document.getElementById('settings-disclaimer').value = s.disclaimer;
}
function saveSettings() {
  const settings = { firm: document.getElementById('settings-firm')?.value, city: document.getElementById('settings-city')?.value, phone: document.getElementById('settings-phone')?.value, disclaimer: document.getElementById('settings-disclaimer')?.value };
  localStorage.setItem('colliers_settings', JSON.stringify(settings));
  showToast('Settings saved');
}

// ============ REVIEW ============
function buildReviewSummary() {
  const g = id => document.getElementById(id)?.value || '—';
  const d = state.design;
  const sections = [
    { title: 'Document', items: [['Type', state.docType === 'om' ? 'Offering Memorandum' : 'Broker Opinion of Value'], ['Property Type', state.propType], ['Page Size', `${state.pageSettings.widthIn}" × ${state.pageSettings.heightIn}"`]] },
    { title: 'Property', items: [['Name', g('prop-name')], ['Address', `${g('prop-address')}, ${g('prop-city')}, ${g('prop-state')}`], ['SF', g('prop-sf') !== '—' ? fmtNum(g('prop-sf')) + ' SF' : '—'], ['Year Built', g('prop-year')]] },
    { title: 'Financials', items: [['Price', g('fin-price') !== '—' ? '$' + fmtNum(g('fin-price')) : '—'], ['NOI', g('fin-noi') !== '—' ? '$' + fmtNum(g('fin-noi')) : '—'], ['Cap Rate', g('fin-caprate') !== '—' ? g('fin-caprate') + '%' : '—'], ['Occupancy', g('fin-occupancy') !== '—' ? g('fin-occupancy') + '%' : '—']] },
    { title: 'Design', items: [['Layout', d.layout], ['Heading Font', d.fonts.heading || 'AI choice'], ['Body Font', d.fonts.body || 'AI choice'], ['Primary Color', d.colors.primary]] },
    { title: 'Files', items: [['Property doc', (state.uploadedFiles['property-doc']?.length || 0) + ' file(s)'], ['Financials', (state.uploadedFiles['financials']?.length || 0) + ' file(s)'], ['Photos', state.photoDataURLs.length + ' photo(s)'], ['Template', (state.uploadedFiles['template']?.length || 0) + ' file(s)']] },
    { title: 'AI', items: [['API Key', state.apiKey ? '✓ Configured' : '✗ Not set'], ['Narrative context', state.narrativeContext ? '✓ Loaded' : '—']] },
  ];
  document.getElementById('review-summary').innerHTML = sections.map(s => `<div class="review-section"><div class="review-section-title">${s.title}</div>${s.items.map(([l,v]) => `<div class="review-item"><span class="review-item-label">${l}</span><span class="review-item-value">${v}</span></div>`).join('')}</div>`).join('');
}

// ============ SAVE DRAFT ============
function saveDraft() {
  const g = id => document.getElementById(id)?.value || '';
  const d = { id: Date.now(), docType: state.docType, propType: state.propType, name: g('prop-name') || 'Untitled', address: `${g('prop-address')}, ${g('prop-city')}, ${g('prop-state')}`, price: g('fin-price'), noi: g('fin-noi'), capRate: g('fin-caprate'), sf: g('prop-sf'), createdAt: new Date().toLocaleDateString(), status: 'draft' };
  state.projects = [d, ...state.projects.filter(p => p.name !== d.name)];
  localStorage.setItem('colliers_projects', JSON.stringify(state.projects));
  showToast('Draft saved');
}

// ============ GENERATE DOCUMENT ============
async function generateDocument() {
  const g = id => document.getElementById(id)?.value || '';
  const settings = JSON.parse(localStorage.getItem('colliers_settings') || '{}');
  const disclaimer = settings.disclaimer || 'The information contained herein has been obtained from sources believed to be reliable. Colliers International makes no guarantee, warranty, or representation about it.';
  const firmName = settings.firm || 'Colliers International';
  const officeCity = settings.city || 'Denver, Colorado';
  const docTypeLabel = state.docType === 'om' ? 'Offering Memorandum' : 'Broker Opinion of Value';

  const prop = { name: g('prop-name') || 'Subject Property', address: g('prop-address'), city: g('prop-city'), state: g('prop-state'), zip: g('prop-zip'), county: g('prop-county'), yearBuilt: g('prop-year'), sf: g('prop-sf'), acres: g('prop-acres'), buildings: g('prop-buildings'), units: g('prop-units'), zoning: g('prop-zoning'), parking: g('prop-parking'), clearHeight: g('prop-clearheight'), desc: g('prop-desc'), highlights: g('prop-highlights') };
  const fin = { price: g('fin-price'), ppsf: g('fin-ppsf'), gpr: g('fin-gpr'), vacancy: g('fin-vacancy'), egi: g('fin-egi'), opex: g('fin-opex'), noi: g('fin-noi'), capRate: g('fin-caprate'), occupancy: g('fin-occupancy'), walt: g('fin-walt') };
  const broker = { name: g('broker-name'), title: g('broker-title'), phone: g('broker-phone'), email: g('broker-email'), license: g('broker-license') };
  const rentRoll = getRentRollData();
  const selectedSections = Array.from(document.querySelectorAll('.sections-checklist input:checked')).map(i => i.dataset.section);

  const statusEl = document.getElementById('generate-status');
  statusEl.style.display = 'block';
  statusEl.innerHTML = '<span class="status-spinner"></span> Generating AI narratives...';

  let ai = {};
  let resolvedFonts = { heading: state.design.fonts.heading || 'Playfair Display', body: state.design.fonts.body || 'Inter', number: state.design.fonts.number || state.design.fonts.heading || 'Playfair Display' };

  try {
    // If fonts not chosen, ask AI to pick them based on context
    if (!state.design.fonts.heading || !state.design.fonts.body) {
      const fontPrompt = `You are a CRE design director. Pick professional Google Fonts for a ${docTypeLabel} for a ${state.propType} property.
Design palette primary color is ${state.design.colors.primary}, layout style is ${state.design.layout}.
Return ONLY JSON: {"heading": "font name", "body": "font name", "number": "font name"}
Choose from: Playfair Display, Merriweather, Lora, Cormorant Garamond, Montserrat, Raleway, Oswald, Inter, DM Sans, Source Sans 3, Lato, Bebas Neue, Barlow`;
      try {
        const fd = await callClaude({ model: 'claude-sonnet-4-20250514', max_tokens: 200, messages: [{ role: 'user', content: fontPrompt }] });
        const fp = JSON.parse((fd.content?.[0]?.text || '{}').replace(/```json|```/g, '').trim());
        if (fp.heading) resolvedFonts.heading = state.design.fonts.heading || fp.heading;
        if (fp.body) resolvedFonts.body = state.design.fonts.body || fp.body;
        if (fp.number) resolvedFonts.number = state.design.fonts.number || fp.number;
      } catch (e) { /* use defaults */ }
    }

    statusEl.innerHTML = '<span class="status-spinner"></span> Writing document narratives...';
    const narrativePrompt = buildNarrativePrompt(prop, fin, broker, rentRoll, docTypeLabel);
    const nd = await callClaude({ model: 'claude-sonnet-4-20250514', max_tokens: 4000, messages: [{ role: 'user', content: narrativePrompt }] });
    try { ai = JSON.parse((nd.content?.[0]?.text || '{}').replace(/```json|```/g, '').trim()); } catch (e) { ai = {}; }
  } catch (e) {
    statusEl.innerHTML = `<span style="color:#f5a623;">⚠ AI unavailable (${e.message}) — building document from entered data.</span>`;
    await new Promise(r => setTimeout(r, 2000));
  }

  // Load chosen fonts before rendering
  [resolvedFonts.heading, resolvedFonts.body, resolvedFonts.number].filter(Boolean).forEach(loadGoogleFont);
  await new Promise(r => setTimeout(r, 800)); // allow fonts to load

  statusEl.innerHTML = '<span class="status-spinner"></span> Building document pages...';
  await new Promise(r => setTimeout(r, 200));

  const html = buildDocumentHTML(prop, fin, broker, rentRoll, ai, selectedSections, disclaimer, docTypeLabel, firmName, officeCity, resolvedFonts, state.design);

  // Navigate to preview
  navigate('preview');
  document.getElementById('preview-title').textContent = `${prop.name} — ${docTypeLabel}`;
  document.getElementById('preview-subtitle').textContent = `${prop.address}, ${prop.city}, ${prop.state} ${prop.zip}`;
  document.getElementById('document-preview-container').innerHTML = html;

  // Inject @page CSS for correct print dimensions
  injectPrintCSS();

  statusEl.style.display = 'none';

  const project = { id: Date.now(), docType: state.docType, propType: state.propType, name: prop.name, address: `${prop.address}, ${prop.city}, ${prop.state}`, price: fin.price, noi: fin.noi, capRate: fin.capRate, sf: prop.sf, createdAt: new Date().toLocaleDateString(), status: 'complete' };
  state.projects = [project, ...state.projects];
  localStorage.setItem('colliers_projects', JSON.stringify(state.projects));
  showToast('Document ready — click "Print / Save as PDF" to export');
}

function injectPrintCSS() {
  const existing = document.getElementById('print-page-css');
  if (existing) existing.remove();
  const { widthIn, heightIn } = state.pageSettings;
  const style = document.createElement('style');
  style.id = 'print-page-css';
  style.textContent = `
    @media print {
      @page { size: ${widthIn}in ${heightIn}in; margin: 0; }
      .doc-page { width: ${widthIn}in !important; height: ${heightIn}in !important; min-height: ${heightIn}in !important; }
    }
  `;
  document.head.appendChild(style);
}

function buildNarrativePrompt(prop, fin, broker, rentRoll, docTypeLabel) {
  return `You are a senior CRE broker at Colliers International writing a ${docTypeLabel}.
${state.narrativeContext ? 'Use this existing narrative as context:\n' + state.narrativeContext + '\n\n' : ''}
PROPERTY: ${prop.name}, ${prop.address}, ${prop.city}, ${prop.state} ${prop.zip}
Type: ${prop.sf ? fmtNum(prop.sf) + ' SF ' : ''}${state.propType}, Built ${prop.yearBuilt||'N/A'}
Zoning: ${prop.zoning||'N/A'}, Clear Height: ${prop.clearHeight||'N/A'} ft
Description: ${prop.desc}
Highlights: ${prop.highlights}
FINANCIALS: Price $${fmtNum(fin.price)}, NOI $${fmtNum(fin.noi)}, Cap Rate ${fin.capRate}%, Occupancy ${fin.occupancy}%, WALT ${fin.walt} yrs
RENT ROLL: ${rentRoll.length ? JSON.stringify(rentRoll) : 'Not provided'}

Return ONLY JSON — no preamble, no markdown fences:
{"executiveSummary":"2-3 paragraphs","propertyDescription":"2-3 paragraphs","locationOverview":"2 paragraphs","investmentHighlights":["h1","h2","h3","h4","h5"],"tenantSummary":"1-2 paragraphs","valuationNarrative":"2 paragraphs"}`;
}

// ============ HTML DOCUMENT BUILDER ============
function buildDocumentHTML(prop, fin, broker, rentRoll, ai, sections, disclaimer, docTypeLabel, firmName, officeCity, fonts, design) {
  const { widthIn, heightIn } = state.pageSettings;
  const C = design.colors;
  const F = fonts;
  const fullAddr = `${prop.address}, ${prop.city}, ${prop.state} ${prop.zip}`;
  const layout = design.layout;

  const googleFontUrl = [...new Set([F.heading, F.body, F.number].filter(Boolean))].map(f => `family=${f.replace(/ /g, '+')}:wght@300;400;500;600;700`).join('&');

  // Shared CSS for all pages
  const sharedCSS = `
    @import url('https://fonts.googleapis.com/css2?${googleFontUrl}&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }
    .doc-page {
      width: ${widthIn}in;
      height: ${heightIn}in;
      min-height: ${heightIn}in;
      overflow: hidden;
      background: ${C.bg};
      position: relative;
      font-family: '${F.body}', Inter, sans-serif;
      color: ${C.text};
      font-size: ${widthIn >= 11 ? '10.5' : '9.5'}pt;
      line-height: 1.65;
    }
    .pg-footer {
      position: absolute; bottom: 0; left: 0; right: 0; height: 28px;
      background: ${C.primary}; display: flex; align-items: center;
      justify-content: space-between; padding: 0 ${widthIn >= 11 ? '0.55in' : '0.4in'};
    }
    .pg-footer-firm { font-size: 7pt; color: ${C.accent}; font-family: '${F.body}', sans-serif; letter-spacing: 0.5px; }
    .pg-footer-broker { font-size: 7pt; color: rgba(255,255,255,0.6); font-family: '${F.body}', sans-serif; }
    .section-rule { height: 2.5px; background: ${C.rule}; margin-bottom: 14px; }
    .section-eyebrow { font-size: 7.5pt; letter-spacing: 2px; text-transform: uppercase; color: ${C.secondary}; font-family: '${F.heading}', serif; font-weight: 600; margin-bottom: 6px; }
    .section-title { font-size: ${widthIn >= 11 ? '18' : '15'}pt; font-family: '${F.heading}', serif; font-weight: 500; color: ${C.primary}; margin-bottom: 14px; line-height: 1.2; }
    .body-text { font-size: ${widthIn >= 11 ? '10' : '9'}pt; line-height: 1.75; color: ${C.text}; }
    .stat-num { font-family: '${F.number || F.heading}', serif; }
    h2 { font-family: '${F.heading}', serif; font-size: ${widthIn >= 11 ? '13' : '11'}pt; font-weight: 600; color: ${C.primary}; margin-bottom: 8px; }
  `;

  function footer(brokerStr) {
    return `<div class="pg-footer">
      <span class="pg-footer-firm">${firmName} · ${officeCity}</span>
      ${brokerStr ? `<span class="pg-footer-broker">${brokerStr}</span>` : ''}
    </div>`;
  }

  function brokerLine() {
    return broker.name ? `${broker.name}${broker.title ? ' · ' + broker.title : ''}${broker.phone ? ' · ' + broker.phone : ''}` : '';
  }

  function pad() { return `padding: ${widthIn >= 11 ? '0.45in 0.55in 0.45in' : '0.35in 0.4in 0.35in'};`; }

  let pages = `<style>${sharedCSS}</style>`;

  // ===== COVER PAGE =====
  if (sections.includes('cover')) {
    const coverStats = [
      fin.price ? { label: 'Asking Price', value: '$' + fmtNum(fin.price) } : null,
      prop.sf ? { label: 'Building SF', value: fmtNum(prop.sf) + ' SF' } : null,
      fin.capRate ? { label: 'Cap Rate', value: fin.capRate + '%' } : null,
      fin.noi ? { label: 'NOI', value: '$' + fmtNum(fin.noi) } : null,
      fin.occupancy ? { label: 'Occupancy', value: fin.occupancy + '%' } : null,
    ].filter(Boolean);

    const heroPhoto = state.photoDataURLs[0] || null;

    pages += `<div class="doc-page" style="background:${C.primary};">
      ${heroPhoto ? `<div style="position:absolute;inset:0;background:url(${heroPhoto}) center/cover no-repeat;opacity:0.22;"></div>` : ''}
      <div style="position:absolute;inset:0;background:linear-gradient(135deg, ${C.primary} 0%, ${C.primary}cc 60%, ${C.secondary}44 100%);"></div>

      <div style="position:relative;z-index:2;height:100%;display:flex;flex-direction:column;${pad()}padding-bottom:0;">

        <div style="margin-bottom:auto;">
          <div style="font-size:8pt;letter-spacing:2.5px;text-transform:uppercase;color:${C.accent};font-family:'${F.heading}',serif;font-weight:600;margin-bottom:${widthIn >= 11 ? '0.3in' : '0.2in'};">${firmName} &nbsp;·&nbsp; ${docTypeLabel}</div>
          <div style="font-family:'${F.heading}',serif;font-size:${widthIn >= 11 ? '38' : '28'}pt;font-weight:500;color:#fff;line-height:1.1;margin-bottom:0.15in;max-width:${widthIn >= 11 ? '6in' : '5in'};">${prop.name}</div>
          <div style="font-size:${widthIn >= 11 ? '13' : '11'}pt;color:rgba(255,255,255,0.7);margin-bottom:0.1in;">${fullAddr}</div>
          ${prop.propType ? `<div style="display:inline-block;background:${C.accent};color:${C.primary};font-size:8pt;font-weight:600;letter-spacing:1px;text-transform:uppercase;padding:4px 12px;border-radius:2px;margin-top:8px;font-family:'${F.body}',sans-serif;">${state.propType.replace('-', ' / ')}</div>` : ''}
        </div>

        <div style="border-top:1.5px solid ${C.accent};padding-top:0.2in;padding-bottom:0.35in;display:flex;gap:${widthIn >= 11 ? '0.4in' : '0.25in'};flex-wrap:wrap;">
          ${coverStats.map(s => `<div>
            <div style="font-size:7pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.accent};opacity:0.8;margin-bottom:4px;font-family:'${F.body}',sans-serif;">${s.label}</div>
            <div class="stat-num" style="font-size:${widthIn >= 11 ? '22' : '17'}pt;color:#fff;font-weight:600;">${s.value}</div>
          </div>`).join('')}
        </div>

        ${broker.name ? `<div style="position:absolute;bottom:36px;right:${widthIn >= 11 ? '0.55in' : '0.4in'};text-align:right;">
          <div style="font-size:9pt;color:#fff;font-weight:500;font-family:'${F.body}',sans-serif;">${broker.name}</div>
          ${broker.title ? `<div style="font-size:8pt;color:${C.accent};font-family:'${F.body}',sans-serif;">${broker.title}</div>` : ''}
          ${broker.phone ? `<div style="font-size:8pt;color:rgba(255,255,255,0.6);font-family:'${F.body}',sans-serif;">${broker.phone}</div>` : ''}
        </div>` : ''}

        <div style="position:absolute;bottom:0;left:0;right:0;height:6px;background:${C.accent};"></div>
      </div>
    </div>`;
  }

  // ===== HELPERS FOR INTERIOR PAGES =====
  function pageWrap(content, pageNum) {
    const sidebarWidth = layout === 'classic' ? (widthIn >= 11 ? '2.1in' : '1.7in') : '0';
    return `<div class="doc-page">
      ${layout === 'classic' ? `<div style="position:absolute;top:0;right:0;width:${sidebarWidth};height:100%;background:${C.primary}08;border-left:3px solid ${C.primary}15;"></div>` : ''}
      ${layout === 'magazine' ? `<div style="position:absolute;top:0;left:0;right:0;height:5px;background:${C.secondary};"></div>` : ''}
      ${layout === 'editorial' ? `<div style="position:absolute;top:0;left:0;right:0;height:${widthIn >= 11 ? '1.1in' : '0.85in'};background:${C.primary};"></div>` : ''}
      <div style="position:relative;z-index:2;height:calc(100% - 28px);${pad()}overflow:hidden;">
        ${content}
      </div>
      ${footer(brokerLine())}
    </div>`;
  }

  function editorialHeader(title) {
    if (layout !== 'editorial') return '';
    return `<div style="margin:-${widthIn >= 11 ? '0.45in' : '0.35in'} -${widthIn >= 11 ? '0.55in' : '0.4in'} 0.3in;background:${C.primary};padding:${widthIn >= 11 ? '0.2in 0.55in' : '0.15in 0.4in'};">
      <div style="font-size:7pt;letter-spacing:2px;text-transform:uppercase;color:${C.accent};font-family:'${F.heading}',serif;margin-bottom:6px;">${docTypeLabel}</div>
      <div style="font-family:'${F.heading}',serif;font-size:${widthIn >= 11 ? '22' : '17'}pt;font-weight:500;color:#fff;">${title}</div>
    </div>`;
  }

  function interiorTitle(title) {
    if (layout === 'editorial') return editorialHeader(title);
    if (layout === 'minimal') return `<div style="margin-bottom:0.2in;"><div class="section-eyebrow">${docTypeLabel}</div><div style="border-top:2px solid ${C.rule};padding-top:10px;"><span style="font-family:'${F.heading}',serif;font-size:${widthIn >= 11 ? '20' : '16'}pt;font-weight:500;color:${C.primary};">${title}</span></div></div>`;
    if (layout === 'magazine') return `<div style="margin-bottom:0.18in;padding-bottom:10px;border-bottom:2.5px solid ${C.rule};display:flex;align-items:baseline;gap:12px;"><span style="font-family:'${F.heading}',serif;font-size:${widthIn >= 11 ? '20' : '16'}pt;font-weight:500;color:${C.primary};">${title}</span><span style="font-size:7pt;letter-spacing:2px;text-transform:uppercase;color:${C.secondary};font-family:'${F.body}',sans-serif;">${docTypeLabel}</span></div>`;
    return `<div style="margin-bottom:0.18in;"><div class="section-eyebrow">${docTypeLabel}</div><div class="section-rule"></div><div class="section-title">${title}</div></div>`;
  }

  function twoColWrap(main, side) {
    if (layout !== 'classic') return main;
    const sw = widthIn >= 11 ? '2.0in' : '1.6in';
    return `<div style="display:flex;gap:0.25in;height:100%;">
      <div style="flex:1;overflow:hidden;">${main}</div>
      <div style="width:${sw};flex-shrink:0;padding-left:0.1in;border-left:1.5px solid ${C.rule}20;">${side}</div>
    </div>`;
  }

  // ===== HIGHLIGHTS + EXEC SUMMARY =====
  if (sections.includes('highlights')) {
    const hiList = ai.investmentHighlights?.length ? ai.investmentHighlights : (prop.highlights ? prop.highlights.split('\n').filter(Boolean) : ['See property details for full information.']);

    const sideContent = layout === 'classic' ? `
      <div style="margin-bottom:0.15in;">
        <div style="font-size:7.5pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.secondary};font-family:'${F.heading}',serif;font-weight:600;margin-bottom:8px;">Key Metrics</div>
        ${[fin.price && ['Asking Price', '$'+fmtNum(fin.price)], fin.capRate && ['Cap Rate', fin.capRate+'%'], fin.noi && ['NOI', '$'+fmtNum(fin.noi)], fin.occupancy && ['Occupancy', fin.occupancy+'%'], fin.walt && ['WALT', fin.walt+' Yrs'], prop.sf && ['Building SF', fmtNum(prop.sf)+' SF']].filter(Boolean).map(([l,v]) => `
          <div style="padding:7px 0;border-bottom:1px solid ${C.rule}20;">
            <div style="font-size:7pt;color:${C.secondary};font-family:'${F.body}',sans-serif;">${l}</div>
            <div class="stat-num" style="font-size:13pt;color:${C.primary};font-weight:600;">${v}</div>
          </div>`).join('')}
      </div>` : '';

    const mainContent = `
      ${interiorTitle('Investment Highlights')}
      <div style="columns:${layout === 'magazine' ? '2' : '1'};column-gap:0.25in;margin-bottom:${ai.executiveSummary ? '0.2in' : '0'};">
        ${hiList.map(h => `<div style="display:flex;align-items:flex-start;gap:10px;padding:7px 0;border-bottom:1px solid ${C.rule}15;break-inside:avoid;">
          <div style="width:6px;height:6px;border-radius:50%;background:${C.secondary};margin-top:4px;flex-shrink:0;"></div>
          <span class="body-text">${h}</span>
        </div>`).join('')}
      </div>
      ${ai.executiveSummary ? `<div style="margin-top:0.15in;"><div style="font-size:7.5pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.secondary};font-family:'${F.heading}',serif;font-weight:600;margin-bottom:8px;">Executive Summary</div><div class="body-text">${ai.executiveSummary.slice(0, 800)}</div></div>` : ''}
    `;
    pages += pageWrap(layout === 'classic' ? twoColWrap(mainContent, sideContent) : mainContent);
  }

  // ===== PROPERTY DETAILS =====
  if (sections.includes('property')) {
    const propRows = [['Property Name', prop.name], ['Address', fullAddr], ['County', prop.county], ['Property Type', state.propType], ['Year Built', prop.yearBuilt], ['Building Size', prop.sf ? fmtNum(prop.sf) + ' SF' : ''], ['Lot Size', prop.acres ? prop.acres + ' Acres' : ''], ['Buildings', prop.buildings], ['Suites / Units', prop.units], ['Zoning', prop.zoning], ['Clear Height', prop.clearHeight ? prop.clearHeight + ' ft' : ''], ['Parking', prop.parking ? prop.parking + ' spaces' : '']].filter(([,v]) => v);

    const tableHTML = `<table style="width:100%;border-collapse:collapse;font-size:9.5pt;">
      ${propRows.map((r, i) => `<tr style="background:${i%2===0 ? C.primary+'08' : 'transparent'};">
        <td style="padding:8px 10px;font-weight:600;color:${C.primary};width:35%;font-family:'${F.body}',sans-serif;border-bottom:1px solid ${C.rule}15;">${r[0]}</td>
        <td style="padding:8px 10px;color:${C.text};border-bottom:1px solid ${C.rule}15;">${r[1]}</td>
      </tr>`).join('')}
    </table>`;

    const sideContent = layout === 'classic' && ai.propertyDescription ? `<div style="font-size:7.5pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.secondary};font-family:'${F.heading}',serif;font-weight:600;margin-bottom:8px;">Description</div><div class="body-text" style="font-size:8.5pt;">${ai.propertyDescription.slice(0, 500)}</div>` : '';

    const mainContent = `${interiorTitle('Property Details')}${ai.propertyDescription && layout !== 'classic' ? `<div class="body-text" style="margin-bottom:0.18in;">${ai.propertyDescription.slice(0, 400)}</div>` : ''}${tableHTML}`;
    pages += pageWrap(layout === 'classic' ? twoColWrap(mainContent, sideContent) : mainContent);
  }

  // ===== PHOTOS =====
  if (sections.includes('photos') && state.photoDataURLs.length > 0) {
    const photos = state.photoDataURLs.slice(0, 8);
    const perRow = widthIn >= 11 ? (photos.length >= 4 ? 3 : 2) : 2;
    const gapIn = 0.1;
    const availW = widthIn - 1.1 - (layout === 'classic' ? 2.2 : 0);
    const imgW = (availW - gapIn * (perRow - 1)) / perRow;
    const imgH = imgW * 0.65;

    const photoGrid = `<div style="display:grid;grid-template-columns:repeat(${perRow},1fr);gap:${gapIn}in;">
      ${photos.map((src, i) => `<div style="border-radius:3px;overflow:hidden;height:${imgH}in;"><img src="${src}" style="width:100%;height:100%;object-fit:cover;display:block;" /></div>`).join('')}
    </div>`;

    pages += pageWrap(`${interiorTitle('Property Photos')}${photoGrid}`);
  }

  // ===== LOCATION =====
  if (sections.includes('location')) {
    const locText = ai.locationOverview || `${prop.city}, ${prop.state} offers a strong commercial real estate market with excellent transportation infrastructure and a growing employment base. The subject property is ideally positioned within the ${prop.city} submarket, providing convenient access to major thoroughfares and a deep pool of regional tenants.`;
    const sideContent = layout === 'classic' ? `<div style="font-size:7.5pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.secondary};font-family:'${F.heading}',serif;font-weight:600;margin-bottom:10px;">Location</div><div style="background:${C.primary};border-radius:6px;padding:14px 12px;color:#fff;font-size:8.5pt;line-height:1.6;font-family:'${F.body}',sans-serif;">${fullAddr}<br><br>${prop.county ? prop.county + '<br>' : ''}${prop.zoning ? 'Zoning: ' + prop.zoning : ''}</div>` : '';
    const mainContent = `${interiorTitle('Location & Market Overview')}<div class="body-text">${locText}</div>`;
    pages += pageWrap(layout === 'classic' ? twoColWrap(mainContent, sideContent) : mainContent);
  }

  // ===== FINANCIAL SUMMARY =====
  if (sections.includes('financials')) {
    const cards = [fin.price && ['Asking Price', '$'+fmtNum(fin.price)], fin.ppsf && ['Price / SF', '$'+fin.ppsf], fin.noi && ['NOI', '$'+fmtNum(fin.noi)], fin.capRate && ['Cap Rate', fin.capRate+'%'], fin.occupancy && ['Occupancy', fin.occupancy+'%'], fin.walt && ['WALT', fin.walt+' Yrs'], fin.gpr && ['Gross Rent', '$'+fmtNum(fin.gpr)], fin.opex && ['Op. Expenses', '$'+fmtNum(fin.opex)]].filter(Boolean);
    const cols = Math.min(cards.length, widthIn >= 11 ? 4 : 3);

    const statsGrid = `<div style="display:grid;grid-template-columns:repeat(${cols},1fr);gap:10px;margin-bottom:0.2in;">
      ${cards.map(([l,v]) => `<div style="background:${C.primary};border-radius:6px;padding:${widthIn >= 11 ? '14px 12px' : '10px 10px'};text-align:center;">
        <div style="font-size:7pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.accent};font-family:'${F.body}',sans-serif;margin-bottom:5px;">${l}</div>
        <div class="stat-num" style="font-size:${widthIn >= 11 ? '18' : '15'}pt;color:#fff;font-weight:600;">${v}</div>
      </div>`).join('')}
    </div>`;

    const incomeTable = `<table style="width:100%;border-collapse:collapse;font-size:9.5pt;margin-top:0.1in;">
      <thead><tr style="background:${C.primary};"><th style="padding:8px 10px;color:#fff;text-align:left;font-family:'${F.body}',sans-serif;font-weight:500;font-size:8pt;">Income Item</th><th style="padding:8px 10px;color:#fff;text-align:right;font-family:'${F.body}',sans-serif;font-weight:500;font-size:8pt;">Amount</th></tr></thead>
      <tbody>
        ${fin.gpr ? `<tr><td style="padding:7px 10px;border-bottom:1px solid ${C.rule}20;">Gross Potential Rent</td><td style="padding:7px 10px;text-align:right;border-bottom:1px solid ${C.rule}20;">$${fmtNum(fin.gpr)}</td></tr>` : ''}
        ${fin.vacancy ? `<tr style="background:${C.primary}06;"><td style="padding:7px 10px;border-bottom:1px solid ${C.rule}20;">Less: Vacancy (${fin.vacancy}%)</td><td style="padding:7px 10px;text-align:right;border-bottom:1px solid ${C.rule}20;">($${fmtNum(Math.round(parseFloat(fin.gpr||0)*parseFloat(fin.vacancy||0)/100))})</td></tr>` : ''}
        ${fin.egi ? `<tr><td style="padding:7px 10px;border-bottom:1px solid ${C.rule}20;">Effective Gross Income</td><td style="padding:7px 10px;text-align:right;border-bottom:1px solid ${C.rule}20;">$${fmtNum(fin.egi)}</td></tr>` : ''}
        ${fin.opex ? `<tr style="background:${C.primary}06;"><td style="padding:7px 10px;border-bottom:1px solid ${C.rule}20;">Less: Operating Expenses</td><td style="padding:7px 10px;text-align:right;border-bottom:1px solid ${C.rule}20;">($${fmtNum(fin.opex)})</td></tr>` : ''}
        ${fin.noi ? `<tr style="background:${C.primary};"><td style="padding:8px 10px;color:#fff;font-weight:600;">Net Operating Income</td><td style="padding:8px 10px;text-align:right;color:${C.accent};font-weight:600;font-family:'${F.number||F.heading}',serif;font-size:11pt;">$${fmtNum(fin.noi)}</td></tr>` : ''}
      </tbody>
    </table>`;

    pages += pageWrap(`${interiorTitle('Financial Summary')}${statsGrid}${incomeTable}`);
  }

  // ===== RENT ROLL =====
  if (sections.includes('rentroll') && rentRoll.length > 0) {
    const rrTable = `<table style="width:100%;border-collapse:collapse;font-size:${widthIn >= 11 ? '9' : '8'}pt;">
      <thead><tr style="background:${C.primary};">
        ${['Tenant','Suite','SF','Lease Start','Lease End','Annual Rent','Rent/SF'].map(h => `<th style="padding:8px 10px;color:#fff;text-align:left;font-family:'${F.body}',sans-serif;font-weight:500;font-size:7.5pt;">${h}</th>`).join('')}
      </tr></thead>
      <tbody>
        ${rentRoll.map((r,i) => {
          const rpsf = r.sf && r.annualRent ? '$'+(parseFloat(r.annualRent)/parseFloat(r.sf)).toFixed(2) : '—';
          return `<tr style="background:${i%2===0 ? C.primary+'06' : 'transparent'};">
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;font-weight:500;color:${C.primary};">${r.tenant}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;">${r.suite}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;">${r.sf ? fmtNum(r.sf) : ''}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;">${r.leaseStart}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;">${r.leaseEnd}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;font-family:'${F.number||F.heading}',serif;">${r.annualRent ? '$'+fmtNum(r.annualRent) : ''}</td>
            <td style="padding:7px 10px;border-bottom:1px solid ${C.rule}15;font-family:'${F.number||F.heading}',serif;">${rpsf}</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>`;
    pages += pageWrap(`${interiorTitle('Rent Roll')}${rrTable}`);
  }

  // ===== TENANT SUMMARY =====
  if (sections.includes('tenants')) {
    const tenantText = ai.tenantSummary || `The property is currently ${fin.occupancy ? fin.occupancy + '% occupied' : 'occupied'}${fin.walt ? ' with a weighted average lease term of ' + fin.walt + ' years' : ''}. Please refer to the rent roll for complete tenant information.`;
    pages += pageWrap(`${interiorTitle('Tenant Summary')}<div class="body-text">${tenantText}</div>`);
  }

  // ===== VALUATION =====
  if (sections.includes('valuation')) {
    const valText = ai.valuationNarrative || `${docTypeLabel === 'Offering Memorandum' ? 'The Seller is offering the property' : 'Based on our analysis, the estimated value of the property is'} ${fin.price ? '$' + fmtNum(fin.price) : 'to be determined'}${fin.ppsf ? ' ($' + fin.ppsf + '/SF)' : ''}${fin.capRate ? ', representing a ' + fin.capRate + '% capitalization rate on an NOI of $' + fmtNum(fin.noi) : ''}.`;

    const valBox = fin.price ? `<div style="background:${C.primary};border-radius:8px;padding:${widthIn >= 11 ? '20px 24px' : '14px 18px'};margin-top:0.2in;display:flex;justify-content:space-between;align-items:center;">
      <div><div style="font-size:7pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.accent};font-family:'${F.body}',sans-serif;margin-bottom:4px;">Offered At</div><div class="stat-num" style="font-size:${widthIn >= 11 ? '28' : '22'}pt;color:#fff;font-weight:600;">$${fmtNum(fin.price)}</div></div>
      ${fin.ppsf ? `<div style="text-align:right;"><div style="font-size:7pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.accent};font-family:'${F.body}',sans-serif;margin-bottom:4px;">Price / SF</div><div class="stat-num" style="font-size:${widthIn >= 11 ? '22' : '17'}pt;color:#fff;font-weight:600;">$${fin.ppsf}</div></div>` : ''}
      ${fin.capRate ? `<div style="text-align:right;"><div style="font-size:7pt;letter-spacing:1.5px;text-transform:uppercase;color:${C.accent};font-family:'${F.body}',sans-serif;margin-bottom:4px;">Cap Rate</div><div class="stat-num" style="font-size:${widthIn >= 11 ? '22' : '17'}pt;color:#fff;font-weight:600;">${fin.capRate}%</div></div>` : ''}
    </div>` : '';

    pages += pageWrap(`${interiorTitle('Valuation & Pricing Guidance')}<div class="body-text">${valText}</div>${valBox}`);
  }

  // ===== DISCLAIMER =====
  if (sections.includes('disclaimer')) {
    pages += pageWrap(`${interiorTitle('Disclaimer')}<div style="font-size:8pt;color:${C.text};opacity:0.65;line-height:1.8;font-family:'${F.body}',sans-serif;">${disclaimer}</div>`);
  }

  return pages;
}

// ============ HELPERS ============
function fmtNum(n) { const num = parseFloat(n); if (isNaN(num)) return String(n||''); return num.toLocaleString('en-US', { maximumFractionDigits: 0 }); }
function showToast(msg) { const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t); setTimeout(() => t.remove(), 3500); }

// ============ DASHBOARD ============
function refreshDashboard() {
  const projects = JSON.parse(localStorage.getItem('colliers_projects') || '[]');
  state.projects = projects;
  document.getElementById('stat-docs').textContent = projects.filter(p => p.status === 'complete').length;
  document.getElementById('stat-props').textContent = projects.length;
  const thisMonth = new Date().toLocaleDateString('en-US', { month: 'numeric', year: 'numeric' });
  document.getElementById('stat-month').textContent = projects.filter(p => { try { return new Date(p.createdAt).toLocaleDateString('en-US', { month: 'numeric', year: 'numeric' }) === thisMonth; } catch(e) { return false; } }).length;
  const container = document.getElementById('recent-projects-list');
  if (!projects.length) { container.innerHTML = `<div class="empty-state"><svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg><p>No projects yet.</p><button class="btn-primary" onclick="navigate('om-builder')">Create Document</button></div>`; return; }
  container.innerHTML = `<table class="projects-table"><thead><tr><th>Property</th><th>Type</th><th>Price</th><th>Cap Rate</th><th>SF</th><th>Date</th><th>Status</th></tr></thead><tbody>${projects.map(p => `<tr><td><div style="font-weight:500;">${p.name}</div><div style="font-size:11px;color:var(--text-muted);">${p.address}</div></td><td><span class="doc-type-badge ${p.docType}">${p.docType==='om'?'OM':'BOV'}</span></td><td>${p.price?'$'+fmtNum(p.price):'—'}</td><td>${p.capRate?p.capRate+'%':'—'}</td><td>${p.sf?fmtNum(p.sf)+' SF':'—'}</td><td>${p.createdAt}</td><td><span style="font-size:11px;color:${p.status==='complete'?'#1a7a4a':'#b8720a'};font-weight:500;">${p.status}</span></td></tr>`).join('')}</tbody></table>`;
}

// ============ API DIAGNOSTIC ============
async function testAPIConnection() {
  const btn = document.getElementById('btn-api-test');
  const result = document.getElementById('api-test-result');
  btn.disabled = true;
  btn.textContent = 'Testing...';
  result.textContent = '';

  // Step 1: check if proxy endpoint is reachable
  try {
    const res = await fetch('/api/claude', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 20, messages: [{ role: 'user', content: 'Say "ok" and nothing else.' }] }),
    });
    const data = await res.json();

    if (res.status === 500 && data?.error === 'API key not configured on server') {
      result.innerHTML = '<span style="color:#cc3333;">✗ Proxy reached but <strong>ANTHROPIC_API_KEY</strong> environment variable is not set on Vercel. Add it in Vercel → Settings → Environment Variables, then redeploy.</span>';
    } else if (res.status === 401) {
      result.innerHTML = '<span style="color:#cc3333;">✗ Proxy reached but API key is invalid or expired. Check the key value in Vercel environment variables.</span>';
    } else if (res.status === 404) {
      result.innerHTML = '<span style="color:#cc3333;">✗ Proxy function not found. Make sure the <code>api/claude.js</code> file was included in your Vercel deployment.</span>';
    } else if (!res.ok) {
      result.innerHTML = `<span style="color:#cc3333;">✗ Error ${res.status}: ${data?.error?.message || data?.error || 'Unknown error'}</span>`;
    } else if (data?.content?.[0]?.text) {
      result.innerHTML = '<span style="color:#1a7a4a;">✓ API connection working. Claude responded successfully.</span>';
    } else {
      result.innerHTML = `<span style="color:#b8720a;">⚠ Unexpected response format: ${JSON.stringify(data).slice(0, 200)}</span>`;
    }
  } catch (e) {
    if (e.message.includes('Failed to fetch') || e.message.includes('NetworkError')) {
      result.innerHTML = '<span style="color:#cc3333;">✗ Could not reach <code>/api/claude</code> — this usually means you\'re running from a local file rather than a server. Open via VS Code Live Server or deploy to Vercel.</span>';
    } else {
      result.innerHTML = `<span style="color:#cc3333;">✗ ${e.message}</span>`;
    }
  }

  btn.disabled = false;
  btn.textContent = 'Test API Connection';
}

// ============ INIT ============
function init() {
  if (state.apiKey) { const el = document.getElementById('api-key-input'); if (el) el.value = state.apiKey; const st = document.getElementById('api-key-status'); if (st) { st.textContent = '✓ Key loaded'; st.className = 'api-key-status ok'; } }
  updatePagePreview();
  refreshDashboard();
}
init();
