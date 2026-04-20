/* ============================================
   Colliers Denver Design Studio — App Logic
   ============================================ */

// ============ STATE ============
const state = {
  docType: 'om',
  propType: 'industrial',
  currentStep: 1,
  uploadedFiles: { financials: [], narrative: [], photos: [], template: [] },
  apiKey: localStorage.getItem('colliers_api_key') || '',
  projects: JSON.parse(localStorage.getItem('colliers_projects') || '[]'),
  tenantRows: 0,
};

// ============ NAVIGATION ============
function navigate(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const target = document.getElementById('page-' + page);
  if (target) target.classList.add('active');
  const navItem = document.querySelector(`[data-page="${page}"]`);
  if (navItem) navItem.classList.add('active');
  if (page === 'dashboard') refreshDashboard();
  if (page === 'settings') loadSettings();
  window.scrollTo(0, 0);
}

document.querySelectorAll('.nav-item:not(.coming-soon)').forEach(item => {
  item.addEventListener('click', (e) => {
    e.preventDefault();
    navigate(item.dataset.page);
  });
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
  if (n === 5) buildReviewSummary();
  window.scrollTo(0, 0);
}

// ============ DOC TYPE ============
function selectDocType(type) {
  state.docType = type;
  document.querySelectorAll('.doc-type-card').forEach(c => {
    c.classList.toggle('active', c.dataset.type === type);
  });
}

function selectPropType(type) {
  state.propType = type;
  document.querySelectorAll('.prop-type-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.ptype === type);
  });
}

// ============ FINANCIAL CALCULATIONS ============
function calcFinancials() {
  const price = parseFloat(document.getElementById('fin-price').value) || 0;
  const sf = parseFloat(document.getElementById('prop-sf').value) || 0;
  const gpr = parseFloat(document.getElementById('fin-gpr').value) || 0;
  const vacancy = parseFloat(document.getElementById('fin-vacancy').value) || 0;
  const opex = parseFloat(document.getElementById('fin-opex').value) || 0;

  const egi = gpr * (1 - vacancy / 100);
  const noi = egi - opex;
  const capRate = price > 0 ? (noi / price) * 100 : 0;
  const ppsf = sf > 0 ? price / sf : 0;
  const grm = gpr > 0 ? price / gpr : 0;

  if (egi > 0) document.getElementById('fin-egi').value = Math.round(egi);
  if (noi > 0) document.getElementById('fin-noi').value = Math.round(noi);
  if (capRate > 0) document.getElementById('fin-caprate').value = capRate.toFixed(2);
  if (ppsf > 0) document.getElementById('fin-ppsf').value = ppsf.toFixed(2);

  document.getElementById('calc-ppsf').textContent = ppsf > 0 ? '$' + ppsf.toFixed(2) : '—';
  document.getElementById('calc-noi').textContent = noi > 0 ? '$' + formatNum(Math.round(noi)) : '—';
  document.getElementById('calc-caprate').textContent = capRate > 0 ? capRate.toFixed(2) + '%' : '—';
  document.getElementById('calc-grm').textContent = grm > 0 ? grm.toFixed(2) + 'x' : '—';
}

// ============ RENT ROLL ============
function addTenantRow() {
  const tbody = document.getElementById('rent-roll-body');
  const emptyRow = tbody.querySelector('.empty-row');
  if (emptyRow) emptyRow.remove();
  const id = ++state.tenantRows;
  const tr = document.createElement('tr');
  tr.id = 'tenant-row-' + id;
  tr.innerHTML = `
    <td><input type="text" placeholder="Tenant name" /></td>
    <td><input type="text" placeholder="101" style="width:60px;" /></td>
    <td><input type="number" placeholder="5000" oninput="calcRowRentSF(${id})" /></td>
    <td><input type="date" /></td>
    <td><input type="date" /></td>
    <td><input type="number" placeholder="60000" oninput="calcRowRentSF(${id})" /></td>
    <td><span id="rpsf-${id}" style="color:var(--text-muted);font-size:12px;">—</span></td>
    <td><button class="btn-del-row" onclick="deleteRow(${id})">×</button></td>
  `;
  tbody.appendChild(tr);
}

function calcRowRentSF(id) {
  const row = document.getElementById('tenant-row-' + id);
  if (!row) return;
  const inputs = row.querySelectorAll('input[type="number"]');
  const sf = parseFloat(inputs[0].value) || 0;
  const rent = parseFloat(inputs[1].value) || 0;
  const rpsf = sf > 0 && rent > 0 ? (rent / sf).toFixed(2) : '—';
  const el = document.getElementById('rpsf-' + id);
  if (el) el.textContent = rpsf !== '—' ? '$' + rpsf : '—';
}

function deleteRow(id) {
  const row = document.getElementById('tenant-row-' + id);
  if (row) row.remove();
  const tbody = document.getElementById('rent-roll-body');
  if (!tbody.querySelector('tr:not(.empty-row)')) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="8" style="text-align:center;color:#6b7fa3;padding:20px;">No tenants added yet.</td></tr>';
  }
}

function getRentRollData() {
  const rows = document.querySelectorAll('#rent-roll-body tr:not(.empty-row)');
  return Array.from(rows).map(row => {
    const inputs = row.querySelectorAll('input');
    return {
      tenant: inputs[0]?.value || '',
      suite: inputs[1]?.value || '',
      sf: inputs[2]?.value || '',
      leaseStart: inputs[3]?.value || '',
      leaseEnd: inputs[4]?.value || '',
      annualRent: inputs[5]?.value || '',
    };
  });
}

// ============ FILE UPLOADS ============
function triggerUpload(type) {
  document.getElementById('file-' + type).click();
}

function handleFileUpload(type, input) {
  const files = Array.from(input.files);
  if (!files.length) return;
  state.uploadedFiles[type] = [...state.uploadedFiles[type], ...files];
  const listEl = document.getElementById('files-' + type);
  const zone = document.getElementById('upload-' + type);
  zone.classList.add('has-files');
  files.forEach(f => {
    const div = document.createElement('div');
    div.className = 'uploaded-file';
    div.innerHTML = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> ${f.name}`;
    listEl.appendChild(div);
  });
}

// ============ API KEY ============
function saveApiKey() {
  const key = document.getElementById('api-key-input').value.trim();
  const status = document.getElementById('api-key-status');
  if (key.startsWith('sk-ant-')) {
    state.apiKey = key;
    localStorage.setItem('colliers_api_key', key);
    status.textContent = '✓ Saved';
    status.className = 'api-key-status ok';
    showToast('API key saved');
  } else {
    status.textContent = 'Invalid key format';
    status.className = 'api-key-status err';
  }
}

function saveApiKeyFromSettings() {
  const key = document.getElementById('settings-api-key').value.trim();
  if (key.startsWith('sk-ant-')) {
    state.apiKey = key;
    localStorage.setItem('colliers_api_key', key);
    showToast('API key saved');
  } else {
    showToast('Invalid key format — should start with sk-ant-');
  }
}

// ============ SETTINGS ============
function loadSettings() {
  const key = localStorage.getItem('colliers_api_key') || '';
  if (key) document.getElementById('settings-api-key').value = key;
  const settings = JSON.parse(localStorage.getItem('colliers_settings') || '{}');
  if (settings.firm) document.getElementById('settings-firm').value = settings.firm;
  if (settings.city) document.getElementById('settings-city').value = settings.city;
  if (settings.phone) document.getElementById('settings-phone').value = settings.phone;
  if (settings.disclaimer) document.getElementById('settings-disclaimer').value = settings.disclaimer;
}

function saveSettings() {
  const settings = {
    firm: document.getElementById('settings-firm').value,
    city: document.getElementById('settings-city').value,
    phone: document.getElementById('settings-phone').value,
    disclaimer: document.getElementById('settings-disclaimer').value,
  };
  localStorage.setItem('colliers_settings', JSON.stringify(settings));
  showToast('Settings saved');
}

// ============ REVIEW SUMMARY ============
function buildReviewSummary() {
  const g = id => document.getElementById(id)?.value || '—';
  const docLabel = state.docType === 'om' ? 'Offering Memorandum' : 'Broker Opinion of Value';
  const propTypeLabel = { industrial:'Industrial / Flex', office:'Office', retail:'Retail / Mixed-Use', multifamily:'Multifamily' }[state.propType];
  const noi = g('fin-noi');
  const cap = g('fin-caprate');

  const sections = [
    {
      title: 'Document',
      items: [
        ['Type', docLabel],
        ['Property Type', propTypeLabel],
      ]
    },
    {
      title: 'Property',
      items: [
        ['Name', g('prop-name')],
        ['Address', `${g('prop-address')}, ${g('prop-city')}, ${g('prop-state')} ${g('prop-zip')}`],
        ['Building SF', g('prop-sf') !== '—' ? formatNum(g('prop-sf')) + ' SF' : '—'],
        ['Year Built', g('prop-year')],
      ]
    },
    {
      title: 'Financials',
      items: [
        ['Asking Price', g('fin-price') !== '—' ? '$' + formatNum(g('fin-price')) : '—'],
        ['NOI', noi !== '—' ? '$' + formatNum(noi) : '—'],
        ['Cap Rate', cap !== '—' ? cap + '%' : '—'],
        ['Occupancy', g('fin-occupancy') !== '—' ? g('fin-occupancy') + '%' : '—'],
      ]
    },
    {
      title: 'Files Uploaded',
      items: [
        ['Financials', state.uploadedFiles.financials.length + ' file(s)'],
        ['Narrative', state.uploadedFiles.narrative.length + ' file(s)'],
        ['Photos', state.uploadedFiles.photos.length + ' file(s)'],
        ['Template', state.uploadedFiles.template.length + ' file(s)'],
      ]
    },
    {
      title: 'Broker',
      items: [
        ['Name', g('broker-name')],
        ['Title', g('broker-title')],
        ['Phone', g('broker-phone')],
        ['Email', g('broker-email')],
      ]
    },
    {
      title: 'AI',
      items: [
        ['API Key', state.apiKey ? '✓ Configured' : '✗ Not set'],
      ]
    },
  ];

  const container = document.getElementById('review-summary');
  container.innerHTML = sections.map(s => `
    <div class="review-section">
      <div class="review-section-title">${s.title}</div>
      ${s.items.map(([l, v]) => `
        <div class="review-item">
          <span class="review-item-label">${l}</span>
          <span class="review-item-value">${v}</span>
        </div>`).join('')}
    </div>`).join('');
}

// ============ SAVE DRAFT ============
function saveDraft() {
  const g = id => document.getElementById(id)?.value || '';
  const draft = {
    id: Date.now(),
    docType: state.docType,
    propType: state.propType,
    name: g('prop-name') || 'Untitled Property',
    address: `${g('prop-address')}, ${g('prop-city')}, ${g('prop-state')}`,
    price: g('fin-price'),
    noi: g('fin-noi'),
    capRate: g('fin-caprate'),
    sf: g('prop-sf'),
    createdAt: new Date().toLocaleDateString(),
    status: 'draft',
  };
  state.projects = [draft, ...state.projects.filter(p => p.name !== draft.name || p.address !== draft.address)];
  localStorage.setItem('colliers_projects', JSON.stringify(state.projects));
  showToast('Draft saved');
}

// ============ GENERATE DOCUMENT ============
async function generateDocument() {
  const g = id => document.getElementById(id)?.value || '';
  const settings = JSON.parse(localStorage.getItem('colliers_settings') || '{}');
  const disclaimer = settings.disclaimer || 'The information contained herein has been obtained from sources believed to be reliable. Colliers International makes no guarantee, warranty, or representation about it.';
  const docTypeLabel = state.docType === 'om' ? 'Offering Memorandum' : 'Broker Opinion of Value';

  const propData = {
    name: g('prop-name') || 'Subject Property',
    address: g('prop-address'),
    city: g('prop-city'),
    state: g('prop-state'),
    zip: g('prop-zip'),
    county: g('prop-county'),
    yearBuilt: g('prop-year'),
    sf: g('prop-sf'),
    acres: g('prop-acres'),
    buildings: g('prop-buildings'),
    units: g('prop-units'),
    zoning: g('prop-zoning'),
    parking: g('prop-parking'),
    clearHeight: g('prop-clearheight'),
    desc: g('prop-desc'),
    highlights: g('prop-highlights'),
    docType: docTypeLabel,
    propType: state.propType,
  };

  const finData = {
    price: g('fin-price'),
    ppsf: g('fin-ppsf'),
    gpr: g('fin-gpr'),
    vacancy: g('fin-vacancy'),
    egi: g('fin-egi'),
    opex: g('fin-opex'),
    noi: g('fin-noi'),
    capRate: g('fin-caprate'),
    occupancy: g('fin-occupancy'),
    walt: g('fin-walt'),
  };

  const brokerData = {
    name: g('broker-name'),
    title: g('broker-title'),
    phone: g('broker-phone'),
    email: g('broker-email'),
    license: g('broker-license'),
  };

  const rentRoll = getRentRollData();

  const selectedSections = Array.from(document.querySelectorAll('.sections-checklist input:checked')).map(i => i.dataset.section);
  const outputFormat = document.querySelector('input[name="output-format"]:checked')?.value || 'html';

  // Show status
  const statusEl = document.getElementById('generate-status');
  statusEl.style.display = 'block';
  statusEl.innerHTML = '<span class="status-spinner"></span> Generating document with AI...';

  let aiNarrative = null;

  if (state.apiKey) {
    try {
      statusEl.innerHTML = '<span class="status-spinner"></span> Calling Claude API — writing narratives and analysis...';

      const prompt = buildAIPrompt(propData, finData, brokerData, rentRoll, selectedSections);

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': state.apiKey, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 4000,
          messages: [{ role: 'user', content: prompt }]
        })
      });

      if (response.ok) {
        const data = await response.json();
        aiNarrative = data.content?.[0]?.text || null;
        statusEl.innerHTML = '<span class="status-spinner"></span> AI content generated — assembling document...';
      } else {
        statusEl.innerHTML = '⚠ API error — generating document with entered data only.';
      }
    } catch (err) {
      statusEl.innerHTML = '⚠ Could not reach API — generating document with entered data only.';
    }
  } else {
    statusEl.innerHTML = '<span class="status-spinner"></span> No API key — assembling document from entered data...';
  }

  await new Promise(r => setTimeout(r, 600));
  statusEl.style.display = 'none';

  // Save to projects
  const project = {
    id: Date.now(),
    docType: state.docType,
    propType: state.propType,
    name: propData.name,
    address: `${propData.address}, ${propData.city}, ${propData.state}`,
    price: finData.price,
    noi: finData.noi,
    capRate: finData.capRate,
    sf: propData.sf,
    createdAt: new Date().toLocaleDateString(),
    status: 'complete',
  };
  state.projects = [project, ...state.projects.filter(p => p.id !== project.id)];
  localStorage.setItem('colliers_projects', JSON.stringify(state.projects));

  // Navigate to output
  navigate('output');
  document.getElementById('output-title').textContent = `${propData.name} — ${docTypeLabel}`;
  document.getElementById('output-subtitle').textContent = `${propData.address}, ${propData.city}, ${propData.state} ${propData.zip}`;

  renderDocument(propData, finData, brokerData, rentRoll, aiNarrative, selectedSections, disclaimer, docTypeLabel, settings);
}

function buildAIPrompt(prop, fin, broker, rentRoll, sections) {
  return `You are a senior commercial real estate broker at Colliers International writing professional document content for a ${prop.docType}.

PROPERTY DETAILS:
- Name: ${prop.name}
- Address: ${prop.address}, ${prop.city}, ${prop.state} ${prop.zip}
- Property Type: ${prop.propType}
- Year Built: ${prop.yearBuilt}
- Building Size: ${formatNum(prop.sf)} SF
- Lot Size: ${prop.acres} acres
- Zoning: ${prop.zoning}
- Clear Height: ${prop.clearHeight ? prop.clearHeight + ' ft' : 'N/A'}
- Description: ${prop.desc}
- Investment Highlights: ${prop.highlights}

FINANCIAL SUMMARY:
- Asking Price: $${formatNum(fin.price)}
- Price/SF: $${fin.ppsf}
- NOI: $${formatNum(fin.noi)}
- Cap Rate: ${fin.capRate}%
- Occupancy: ${fin.occupancy}%
- WALT: ${fin.walt} years

RENT ROLL: ${rentRoll.length > 0 ? JSON.stringify(rentRoll) : 'Not provided'}

Write the following sections in professional CRE marketing language. Be specific, concise, and compelling. Format your response as a JSON object with these keys:
{
  "executiveSummary": "2-3 paragraph executive summary",
  "propertyDescription": "2-3 paragraph detailed property description",  
  "locationOverview": "2 paragraph location and market overview for ${prop.city}, ${prop.state}",
  "investmentHighlights": ["highlight 1", "highlight 2", "highlight 3", "highlight 4", "highlight 5"],
  "tenantSummary": "1-2 paragraph tenant overview based on rent roll",
  "valuationNarrative": "2 paragraph valuation commentary and pricing rationale"
}

Return ONLY the JSON object, no other text.`;
}

function renderDocument(prop, fin, broker, rentRoll, aiRaw, sections, disclaimer, docTypeLabel, settings) {
  let ai = {};
  if (aiRaw) {
    try {
      const clean = aiRaw.replace(/```json|```/g, '').trim();
      ai = JSON.parse(clean);
    } catch(e) { ai = {}; }
  }

  const fullAddress = `${prop.address}, ${prop.city}, ${prop.state} ${prop.zip}`;
  const firmName = settings.firm || 'Colliers International';

  let html = `<div class="doc-wrapper">`;

  // Cover
  if (sections.includes('cover')) {
    html += `
    <div class="doc-cover">
      <div class="doc-cover-eyebrow">Colliers International · ${docTypeLabel}</div>
      <div class="doc-cover-title">${prop.name || 'Subject Property'}</div>
      <div class="doc-cover-address">${fullAddress}</div>
      <div class="doc-cover-stats">
        ${fin.price ? `<div><div class="doc-cover-stat-label">Asking Price</div><div class="doc-cover-stat-value">$${formatNum(fin.price)}</div></div>` : ''}
        ${prop.sf ? `<div><div class="doc-cover-stat-label">Building SF</div><div class="doc-cover-stat-value">${formatNum(prop.sf)}</div></div>` : ''}
        ${fin.capRate ? `<div><div class="doc-cover-stat-label">Cap Rate</div><div class="doc-cover-stat-value">${fin.capRate}%</div></div>` : ''}
        ${fin.noi ? `<div><div class="doc-cover-stat-label">NOI</div><div class="doc-cover-stat-value">$${formatNum(fin.noi)}</div></div>` : ''}
        ${fin.occupancy ? `<div><div class="doc-cover-stat-label">Occupancy</div><div class="doc-cover-stat-value">${fin.occupancy}%</div></div>` : ''}
      </div>
    </div>`;
  }

  // Investment Highlights
  if (sections.includes('highlights')) {
    const hiList = ai.investmentHighlights?.length
      ? ai.investmentHighlights
      : (prop.highlights ? prop.highlights.split('\n').filter(Boolean) : ['See property details below.']);
    html += `
    <div class="doc-section">
      <div class="doc-section-title">Investment Highlights</div>
      <ul class="doc-highlights-list">
        ${hiList.map(h => `<li><div class="doc-highlight-bullet"></div><span>${h}</span></li>`).join('')}
      </ul>
    </div>`;
  }

  // Executive Summary
  html += `
  <div class="doc-section">
    <div class="doc-section-title">Executive Summary</div>
    ${ai.executiveSummary
      ? `<p class="ai-content">${ai.executiveSummary}</p>`
      : `<p>Colliers International is pleased to present the opportunity to acquire ${prop.name || 'this property'} located at ${fullAddress}. 
         This ${prop.propType} property offers ${prop.sf ? formatNum(prop.sf) + ' square feet' : 'substantial space'} of quality space 
         ${prop.yearBuilt ? 'built in ' + prop.yearBuilt : ''}.
         ${fin.price ? 'The property is offered at $' + formatNum(fin.price) + (fin.ppsf ? ' ($' + fin.ppsf + '/SF)' : '') + '.' : ''}
         ${fin.capRate ? ' The offering represents a ' + fin.capRate + '% cap rate' + (fin.noi ? ' on an NOI of $' + formatNum(fin.noi) : '') + '.' : ''}</p>`}
  </div>`;

  // Property Details
  if (sections.includes('property')) {
    const rows = [
      ['Property Name', prop.name],
      ['Address', fullAddress],
      ['County', prop.county],
      ['Property Type', prop.propType],
      ['Year Built', prop.yearBuilt],
      ['Building Size', prop.sf ? formatNum(prop.sf) + ' SF' : ''],
      ['Lot Size', prop.acres ? prop.acres + ' Acres' : ''],
      ['Number of Buildings', prop.buildings],
      ['Number of Suites', prop.units],
      ['Zoning', prop.zoning],
      ['Clear Height', prop.clearHeight ? prop.clearHeight + ' ft' : ''],
      ['Parking', prop.parking ? prop.parking + ' spaces' : ''],
    ].filter(([, v]) => v);

    html += `
    <div class="doc-section">
      <div class="doc-section-title">Property Details</div>
      ${ai.propertyDescription ? `<p class="ai-content" style="margin-bottom:20px;">${ai.propertyDescription}</p>` : ''}
      <div class="doc-prop-grid">
        ${rows.map(([l, v]) => `
          <div class="doc-prop-label">${l}</div>
          <div class="doc-prop-value">${v}</div>`).join('')}
      </div>
    </div>`;
  }

  // Financial Summary
  if (sections.includes('financials')) {
    const cards = [
      ['Asking Price', fin.price ? '$' + formatNum(fin.price) : '—'],
      ['Price / SF', fin.ppsf ? '$' + fin.ppsf : '—'],
      ['NOI', fin.noi ? '$' + formatNum(fin.noi) : '—'],
      ['Cap Rate', fin.capRate ? fin.capRate + '%' : '—'],
      ['Occupancy', fin.occupancy ? fin.occupancy + '%' : '—'],
      ['WALT', fin.walt ? fin.walt + ' Yrs' : '—'],
      ['Gross Rent', fin.gpr ? '$' + formatNum(fin.gpr) : '—'],
      ['Op. Expenses', fin.opex ? '$' + formatNum(fin.opex) : '—'],
    ].filter(([, v]) => v !== '—');

    html += `
    <div class="doc-section">
      <div class="doc-section-title">Financial Summary</div>
      <div class="doc-fin-grid" style="grid-template-columns: repeat(${Math.min(cards.length, 4)}, 1fr);">
        ${cards.map(([l, v]) => `
          <div class="doc-fin-card">
            <div class="doc-fin-label">${l}</div>
            <div class="doc-fin-value">${v}</div>
          </div>`).join('')}
      </div>
    </div>`;
  }

  // Rent Roll
  if (sections.includes('rentroll') && rentRoll.length > 0) {
    html += `
    <div class="doc-section">
      <div class="doc-section-title">Rent Roll</div>
      <table class="doc-table">
        <thead>
          <tr>
            <th>Tenant</th><th>Suite</th><th>SF</th>
            <th>Lease Start</th><th>Lease End</th>
            <th>Annual Rent</th><th>Rent/SF</th>
          </tr>
        </thead>
        <tbody>
          ${rentRoll.map(r => {
            const rpsf = r.sf && r.annualRent ? (parseFloat(r.annualRent)/parseFloat(r.sf)).toFixed(2) : '—';
            return `<tr>
              <td>${r.tenant}</td>
              <td>${r.suite}</td>
              <td>${r.sf ? formatNum(r.sf) : ''}</td>
              <td>${r.leaseStart}</td>
              <td>${r.leaseEnd}</td>
              <td>${r.annualRent ? '$' + formatNum(r.annualRent) : ''}</td>
              <td>${rpsf !== '—' ? '$' + rpsf : '—'}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;
  }

  // Tenant Summary
  if (sections.includes('tenants')) {
    html += `
    <div class="doc-section">
      <div class="doc-section-title">Tenant Summary</div>
      ${ai.tenantSummary
        ? `<p class="ai-content">${ai.tenantSummary}</p>`
        : `<p>The property ${fin.occupancy ? 'is currently ' + fin.occupancy + '% occupied' : 'is occupied'} ${fin.walt ? 'with a weighted average lease term of ' + fin.walt + ' years' : ''}. Please refer to the rent roll above for complete tenant information.</p>`}
    </div>`;
  }

  // Location Overview
  if (sections.includes('location')) {
    html += `
    <div class="doc-section">
      <div class="doc-section-title">Location & Market Overview</div>
      ${ai.locationOverview
        ? `<p class="ai-content">${ai.locationOverview}</p>`
        : `<p>${prop.city}, ${prop.state} offers a strong commercial real estate market with excellent transportation access and a growing employment base. The subject property benefits from its strategic location within the ${prop.city} submarket.</p>`}
    </div>`;
  }

  // Valuation
  if (sections.includes('valuation')) {
    html += `
    <div class="doc-section">
      <div class="doc-section-title">Valuation & Pricing Guidance</div>
      ${ai.valuationNarrative
        ? `<p class="ai-content">${ai.valuationNarrative}</p>`
        : `<p>${docTypeLabel === 'Offering Memorandum' ? 'The Seller is offering the property' : 'Based on our analysis, we estimate the value of the property'} at 
           ${fin.price ? '$' + formatNum(fin.price) : 'a price to be determined'}
           ${fin.ppsf ? ' ($' + fin.ppsf + '/SF)' : ''}
           ${fin.capRate ? ', representing a ' + fin.capRate + '% capitalization rate' : ''}.
           ${fin.noi ? ' This pricing is supported by a net operating income of $' + formatNum(fin.noi) + '.' : ''}</p>`}
    </div>`;
  }

  // Disclaimer
  if (sections.includes('disclaimer')) {
    html += `<div class="doc-disclaimer">${disclaimer}</div>`;
  }

  // Footer
  html += `
  <div class="doc-footer">
    <div class="doc-footer-firm">${firmName} · ${settings.city || 'Denver, Colorado'}</div>
    ${broker.name ? `<div class="doc-footer-broker">${broker.name}${broker.title ? ' · ' + broker.title : ''} · ${broker.phone || ''} · ${broker.email || ''}</div>` : ''}
  </div>`;

  html += `</div>`;

  document.getElementById('document-output').innerHTML = html;
}

// ============ PRINT ============
function printDocument() {
  window.print();
}

// ============ DASHBOARD ============
function refreshDashboard() {
  const projects = JSON.parse(localStorage.getItem('colliers_projects') || '[]');
  state.projects = projects;

  document.getElementById('stat-docs').textContent = projects.filter(p => p.status === 'complete').length;
  document.getElementById('stat-props').textContent = projects.length;
  const thisMonth = new Date().toLocaleDateString('en-US', { month: 'numeric', year: 'numeric' });
  document.getElementById('stat-month').textContent = projects.filter(p => {
    const d = new Date(p.createdAt);
    return d.toLocaleDateString('en-US', { month: 'numeric', year: 'numeric' }) === thisMonth;
  }).length;

  const container = document.getElementById('recent-projects-list');
  if (!projects.length) {
    container.innerHTML = `<div class="empty-state">
      <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <p>No projects yet. Create your first OM or BOV to get started.</p>
      <button class="btn-primary" onclick="navigate('om-builder')">Create Document</button>
    </div>`;
    return;
  }

  container.innerHTML = `
  <table class="projects-table">
    <thead>
      <tr>
        <th>Property</th>
        <th>Type</th>
        <th>Price</th>
        <th>Cap Rate</th>
        <th>SF</th>
        <th>Date</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>
      ${projects.map(p => `
        <tr>
          <td>
            <div style="font-weight:500;">${p.name}</div>
            <div style="font-size:11px;color:var(--text-muted);">${p.address}</div>
          </td>
          <td><span class="doc-type-badge ${p.docType}">${p.docType === 'om' ? 'OM' : 'BOV'}</span></td>
          <td>${p.price ? '$' + formatNum(p.price) : '—'}</td>
          <td>${p.capRate ? p.capRate + '%' : '—'}</td>
          <td>${p.sf ? formatNum(p.sf) + ' SF' : '—'}</td>
          <td>${p.createdAt}</td>
          <td><span style="font-size:11px;color:${p.status==='complete'?'#1a7a4a':'#b8720a'};font-weight:500;">${p.status}</span></td>
        </tr>`).join('')}
    </tbody>
  </table>`;
}

// ============ FORMAT HELPERS ============
function formatNum(n) {
  const num = parseFloat(n);
  if (isNaN(num)) return n;
  return num.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

// ============ TOAST ============
function showToast(msg) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3000);
}

// ============ INIT ============
function init() {
  if (state.apiKey) {
    const input = document.getElementById('api-key-input');
    if (input) input.value = state.apiKey;
    const status = document.getElementById('api-key-status');
    if (status) { status.textContent = '✓ Key loaded'; status.className = 'api-key-status ok'; }
  }
  refreshDashboard();
}

init();
