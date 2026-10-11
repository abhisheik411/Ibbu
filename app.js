/* =============================================
   COOL AUTO – APPLICATION JAVASCRIPT
   ============================================= */

// ========== DATA STORE ==========
let db = {
  customers: [],
  parts: [],
  services: []
};

let searchFilter = 'all';
let currentRecordsTab = 'customers';

// ========== EXCEL AUTO-SAVE ==========
let excelFileHandle = null;   // File System Access API handle
let autoSaveTimer   = null;
let lastSavedAt     = null;
const SUPPORTS_FSA  = ('showSaveFilePicker' in window);  // Chrome/Edge 86+

// ========== INIT ==========
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  document.getElementById('regDate').value = today();
  document.getElementById('serviceDate').value = today();
  renderAll();
  showExcelPathInfo();
});

function today() {
  return new Date().toISOString().split('T')[0];
}

// ========== LOCAL STORAGE ==========
function saveData() {
  // 1. Always save to localStorage immediately (never lost)
  localStorage.setItem('coolauto_db', JSON.stringify(db));
  // 2. Auto-write Excel file if user has linked one
  scheduleExcelAutoSave();
}

function loadData() {
  const raw = localStorage.getItem('coolauto_db');
  if (raw) {
    try { db = JSON.parse(raw); } catch(e) { db = { customers:[], parts:[], services:[] }; }
  }
}

function renderAll() {
  renderRecentRegistrations();
  updateCustomerCount();
  renderPartsList();
  updatePartsStats();
  renderRecords();
  renderServicePartsDropdown();
}

// ========== TABS ==========
function switchTab(name) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
  if (name === 'records') renderRecords();
  if (name === 'service') renderServicePartsDropdown();
  if (name === 'parts') { renderPartsList(); updatePartsStats(); }
}

// ========== REGISTER CUSTOMER ==========
function submitRegistration(e) {
  e.preventDefault();
  const carNum = document.getElementById('carNumber').value.trim().toUpperCase();

  // ── Returning car? Show smart prompt instead of hard block ──
  const exists = db.customers.find(c => c.carNumber === carNum);
  if (exists) {
    const services = db.services.filter(s => s.customerId === exists.id);
    document.getElementById('modalContent').innerHTML = `
      <div class="returning-car-banner">
        <div class="returning-car-icon">🔄</div>
        <div class="returning-car-title">Returning Vehicle Detected!</div>
        <div class="returning-car-sub">This car is already in the system</div>
      </div>
      <div class="modal-row"><label>Customer</label><span>${escHtml(exists.name)}</span></div>
      <div class="modal-row"><label>Car Number</label><span>${escHtml(exists.carNumber)}</span></div>
      <div class="modal-row"><label>Car Model</label><span>${escHtml(exists.carModel)}</span></div>
      <div class="modal-row"><label>Phone</label><span>${escHtml(exists.phone)}</span></div>
      <div class="modal-row"><label>Total Visits</label><span>${services.length} service${services.length !== 1 ? 's' : ''} recorded</span></div>
      <div class="returning-car-actions">
        <button class="btn btn-primary" onclick="closeModal(); goToService('${exists.id}')">
          ⚙️ Add New Service Visit
        </button>
        <button class="btn btn-outline" onclick="closeModal(); openCustomerModal('${exists.id}')">
          📋 View Full History
        </button>
      </div>
      <p class="returning-car-note">💡 Use <strong>"Add New Service Visit"</strong> to record today's visit and deduct parts.</p>
    `;
    openModal();
    return;
  }

  const customer = {
    id: genId(),
    name: document.getElementById('custName').value.trim(),
    phone: document.getElementById('custPhone').value.trim(),
    carNumber: carNum,
    carModel: document.getElementById('carModel').value.trim(),
    acType: document.getElementById('acType').value,
    date: document.getElementById('regDate').value,
    problem: document.getElementById('problem').value.trim(),
    address: document.getElementById('address').value.trim(),
    createdAt: new Date().toISOString()
  };

  db.customers.unshift(customer);
  saveData();
  renderRecentRegistrations();
  updateCustomerCount();
  document.getElementById('registerForm').reset();
  document.getElementById('regDate').value = today();
  showToast('✅ Customer registered successfully!', 'success');
}

function resetRegForm() {
  document.getElementById('registerForm').reset();
  document.getElementById('regDate').value = today();
}

function renderRecentRegistrations() {
  const el = document.getElementById('recentRegistrations');
  const recent = db.customers.slice(0, 5);
  if (recent.length === 0) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">🚗</div><p>No customers registered yet.</p></div>';
    return;
  }
  el.innerHTML = recent.map(c => customerCardHTML(c)).join('');
}

function updateCustomerCount() {
  document.getElementById('custCountBadge').textContent = db.customers.length + ' customers';
}

function customerCardHTML(c) {
  const srvCount = db.services.filter(s => s.customerId === c.id).length;
  return `
  <div class="customer-card" onclick="openCustomerModal('${c.id}')">
    <div class="cust-info">
      <div class="cust-name">👤 ${escHtml(c.name)}</div>
      <div class="cust-car">🔢 ${escHtml(c.carNumber)} &nbsp;|&nbsp; 🚘 ${escHtml(c.carModel)}</div>
      <div class="cust-meta">
        <span>📞 ${escHtml(c.phone)}</span>
        ${c.acType ? `<span>❄️ ${escHtml(c.acType)}</span>` : ''}
        ${srvCount > 0 ? `<span class="visit-count">🔧 ${srvCount} visit${srvCount>1?'s':''}</span>` : ''}
        ${c.problem ? `<span>🛠️ ${escHtml(c.problem.substring(0,40))}${c.problem.length>40?'…':''}</span>` : ''}
      </div>
    </div>
    <div>
      <div class="cust-date">${formatDate(c.date)}</div>
      <div class="cust-actions" style="justify-content:flex-end;margin-top:0.5rem;" onclick="event.stopPropagation()">
        <button class="btn btn-sm btn-service" onclick="goToService('${c.id}')" title="Add new service visit">🔧 Service</button>
        <button class="btn btn-sm btn-danger" onclick="deleteCustomer('${c.id}')">🗑️</button>
      </div>
    </div>
  </div>`;
}

function deleteCustomer(id) {
  if (!confirm('Delete this customer record?')) return;
  db.customers = db.customers.filter(c => c.id !== id);
  saveData();
  renderRecentRegistrations();
  updateCustomerCount();
  renderRecords();
  showToast('🗑️ Customer deleted', 'info');
}

// ========== GO TO SERVICE (pre-select customer) ==========
function goToService(customerId) {
  const c = db.customers.find(x => x.id === customerId);
  if (!c) return;
  switchTab('service');
  // Pre-fill the service customer search
  document.getElementById('serviceCarSearch').value = c.carNumber + ' — ' + c.name;
  document.getElementById('selectedCustomerId').value = customerId;
  document.getElementById('serviceDate').value = today();
  document.getElementById('serviceCustomerDropdown').style.display = 'none';
  const card = document.getElementById('selectedCustomerCard');
  const srvCount = db.services.filter(s => s.customerId === customerId).length;
  card.innerHTML = `
    <div style="font-size:1.5rem">🚗</div>
    <div style="flex:1">
      <div class="sel-cust-name">${escHtml(c.name)}</div>
      <div class="sel-cust-car">${escHtml(c.carNumber)} · ${escHtml(c.carModel)}</div>
      <div style="font-size:0.75rem;color:var(--text-muted);margin-top:0.15rem;">🔧 ${srvCount} previous visit${srvCount!==1?'s':''} on record</div>
    </div>
    <span class="returning-tag">🔄 Returning</span>
  `;
  card.style.display = 'flex';
  serviceCustomer = c;
  showToast(`🔄 Returning vehicle: ${c.carNumber} — ready to add service`, 'info');
}
function openCustomerModal(id) {
  const c = db.customers.find(x => x.id === id);
  if (!c) return;
  const services = db.services.filter(s => s.customerId === id);

  document.getElementById('modalContent').innerHTML = `
    <div class="modal-title">👤 ${escHtml(c.name)}</div>
    <div class="modal-row"><label>Phone</label><span>${escHtml(c.phone)}</span></div>
    <div class="modal-row"><label>Car Number</label><span>${escHtml(c.carNumber)}</span></div>
    <div class="modal-row"><label>Car Model</label><span>${escHtml(c.carModel)}</span></div>
    ${c.acType ? `<div class="modal-row"><label>AC Type</label><span>${escHtml(c.acType)}</span></div>` : ''}
    ${c.address ? `<div class="modal-row"><label>Address</label><span>${escHtml(c.address)}</span></div>` : ''}
    ${c.problem ? `<div class="modal-row"><label>Problem/Notes</label><span>${escHtml(c.problem)}</span></div>` : ''}
    <div class="modal-row"><label>Registered On</label><span>${formatDate(c.date)}</span></div>

    <div class="modal-section">
      <div class="modal-section-title">⚙️ Service History (${services.length})</div>
      ${services.length === 0 ? '<p style="color:var(--text-muted);font-size:0.85rem;">No services recorded.</p>' :
        services.map(s => {
          const partsUsed = (s.parts || []).map(p => {
            const pt = db.parts.find(x => x.id === p.partId);
            return pt ? `<span class="srv-part-tag">${escHtml(pt.partName)} x${p.qty}</span>` : '';
          }).join('');
          return `
          <div class="modal-part-item">
            <div><strong>${escHtml(s.serviceType)}</strong> — ${formatDate(s.date)}</div>
            ${partsUsed ? `<div style="margin-top:0.25rem;">${partsUsed}</div>` : ''}
            ${s.notes ? `<div style="color:var(--text-muted);font-size:0.75rem;margin-top:0.2rem;">${escHtml(s.notes)}</div>` : ''}
            ${s.laborCost ? `<div style="color:var(--success);font-size:0.75rem;">Labour: ₹${s.laborCost}</div>` : ''}
          </div>`;
        }).join('')}
    </div>
  `;
  openModal();
}

// ========== SEARCH ==========
function setFilter(f) {
  searchFilter = f;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('filter-' + f).classList.add('active');
  performSearch();
}

function performSearch() {
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const el = document.getElementById('searchResults');
  if (!q) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><p>Start typing to search customers...</p></div>';
    return;
  }
  let results = db.customers.filter(c => {
    if (searchFilter === 'name') return c.name.toLowerCase().includes(q);
    if (searchFilter === 'car') return c.carNumber.toLowerCase().includes(q);
    return c.name.toLowerCase().includes(q) || c.carNumber.toLowerCase().includes(q) || c.phone.includes(q);
  });

  if (results.length === 0) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">😕</div><p>No customers found for "<strong>' + escHtml(q) + '</strong>"</p></div>';
    return;
  }
  el.innerHTML = results.map(c => customerCardHTML(c)).join('');
}

function clearSearch() {
  document.getElementById('searchInput').value = '';
  performSearch();
}

// ========== PARTS ==========
function submitPart(e) {
  e.preventDefault();
  const pn = document.getElementById('partNumber').value.trim().toUpperCase();
  if (db.parts.find(p => p.partNumber === pn)) {
    showToast('⚠️ Part number already exists!', 'error');
    return;
  }
  const count = parseInt(document.getElementById('partCount').value);
  const part = {
    id: genId(),
    partName: document.getElementById('partName').value.trim(),
    partNumber: pn,
    totalCount: count,
    usedCount: 0,
    remaining: count,
    price: parseFloat(document.getElementById('partPrice').value) || 0,
    createdAt: new Date().toISOString()
  };
  db.parts.push(part);
  saveData();
  renderPartsList();
  updatePartsStats();
  renderServicePartsDropdown();
  document.getElementById('addPartForm').reset();
  showToast('✅ Part added to inventory!', 'success');
}

function renderPartsList() {
  const q = (document.getElementById('partsSearch')?.value || '').toLowerCase();
  const el = document.getElementById('partsList');
  let parts = db.parts.filter(p =>
    p.partName.toLowerCase().includes(q) || p.partNumber.toLowerCase().includes(q)
  );

  if (parts.length === 0) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">🔧</div><p>No parts in inventory yet.</p></div>';
    return;
  }

  el.innerHTML = parts.map(p => partCardHTML(p)).join('');
}

function partCardHTML(p) {
  const pct = p.totalCount > 0 ? Math.round((p.remaining / p.totalCount) * 100) : 0;
  const stockClass = p.remaining === 0 ? 'out-stock' : p.remaining <= 3 ? 'low-stock' : '';
  const badgeClass = p.remaining === 0 ? 'stock-out' : p.remaining <= 3 ? 'stock-low' : 'stock-ok';
  const badgeText = p.remaining === 0 ? '❌ Out of Stock' : p.remaining <= 3 ? '⚠️ Low Stock' : '✅ In Stock';
  const barColor = p.remaining === 0 ? '#ff5252' : p.remaining <= 3 ? '#ffab40' : '#00e676';

  return `
  <div class="part-card ${stockClass}">
    <div class="part-top">
      <div>
        <div class="part-name">🔩 ${escHtml(p.partName)}</div>
        <div class="part-num">#${escHtml(p.partNumber)} ${p.price > 0 ? `· ₹${p.price}` : ''}</div>
      </div>
      <div class="part-stock-badge ${badgeClass}">${badgeText}</div>
    </div>
    <div class="part-bar-wrap">
      <div class="part-bar-bg">
        <div class="part-bar-fill" style="width:${pct}%;background:${barColor};"></div>
      </div>
      <div class="part-counts">
        <span>Used: ${p.usedCount}</span>
        <span>Remaining: <strong style="color:${barColor}">${p.remaining}</strong> / ${p.totalCount}</span>
      </div>
    </div>
    <div class="part-actions">
      <input class="form-input restock-input" type="number" min="1" placeholder="Qty" id="restock_${p.id}" />
      <button class="btn btn-sm btn-success" onclick="restockPart('${p.id}')">+ Restock</button>
      <button class="btn btn-danger" onclick="deletePart('${p.id}')">🗑️</button>
    </div>
  </div>`;
}

function restockPart(id) {
  const qty = parseInt(document.getElementById('restock_' + id)?.value);
  if (!qty || qty < 1) { showToast('Enter a valid quantity', 'error'); return; }
  const p = db.parts.find(x => x.id === id);
  if (!p) return;
  p.totalCount += qty;
  p.remaining += qty;
  saveData();
  renderPartsList();
  updatePartsStats();
  showToast(`✅ Restocked ${qty} units of "${p.partName}"`, 'success');
}

function deletePart(id) {
  if (!confirm('Delete this part from inventory?')) return;
  db.parts = db.parts.filter(p => p.id !== id);
  saveData();
  renderPartsList();
  updatePartsStats();
  renderServicePartsDropdown();
  showToast('🗑️ Part deleted', 'info');
}

function updatePartsStats() {
  document.getElementById('totalPartsCount').textContent = db.parts.length;
  document.getElementById('inStockCount').textContent = db.parts.filter(p => p.remaining > 0).length;
  document.getElementById('lowStockCount').textContent = db.parts.filter(p => p.remaining > 0 && p.remaining <= 3).length;
}

// ========== SERVICE ==========
let serviceCustomer = null;

function searchServiceCustomer() {
  const q = document.getElementById('serviceCarSearch').value.trim().toLowerCase();
  const dd = document.getElementById('serviceCustomerDropdown');
  if (!q) { dd.style.display = 'none'; return; }

  const matches = db.customers.filter(c =>
    c.name.toLowerCase().includes(q) || c.carNumber.toLowerCase().includes(q)
  ).slice(0, 8);

  if (matches.length === 0) { dd.style.display = 'none'; return; }

  dd.innerHTML = matches.map(c =>
    `<div class="dropdown-item" onclick="selectServiceCustomer('${c.id}')">
      <strong>${escHtml(c.name)}</strong> — ${escHtml(c.carNumber)} (${escHtml(c.carModel)})
    </div>`
  ).join('');
  dd.style.display = 'block';
}

function selectServiceCustomer(id) {
  const c = db.customers.find(x => x.id === id);
  if (!c) return;
  serviceCustomer = c;
  document.getElementById('selectedCustomerId').value = id;
  document.getElementById('serviceCarSearch').value = c.carNumber + ' — ' + c.name;
  document.getElementById('serviceCustomerDropdown').style.display = 'none';
  const card = document.getElementById('selectedCustomerCard');
  card.innerHTML = `
    <div style="font-size:1.5rem">👤</div>
    <div>
      <div class="sel-cust-name">${escHtml(c.name)}</div>
      <div class="sel-cust-car">${escHtml(c.carNumber)} · ${escHtml(c.carModel)}</div>
    </div>`;
  card.style.display = 'flex';
}

function renderServicePartsDropdown() {
  // Called when parts list changes
  const rows = document.querySelectorAll('.part-row select');
  rows.forEach(sel => {
    const val = sel.value;
    sel.innerHTML = partsOptions();
    sel.value = val;
  });
}

function partsOptions() {
  if (db.parts.length === 0) return '<option value="">No parts in inventory</option>';
  return '<option value="">Select part...</option>' +
    db.parts.filter(p => p.remaining > 0).map(p =>
      `<option value="${p.id}">${escHtml(p.partName)} (#${escHtml(p.partNumber)}) — Stock: ${p.remaining}</option>`
    ).join('');
}

function addPartRow() {
  const container = document.getElementById('partsUsedContainer');
  const div = document.createElement('div');
  div.className = 'part-row';
  div.innerHTML = `
    <select class="form-input" onchange="checkPartStock(this)">${partsOptions()}</select>
    <input class="form-input" type="number" min="1" placeholder="Qty" value="1" />
    <button type="button" class="remove-row" onclick="this.parentElement.remove()">✕</button>
  `;
  container.appendChild(div);
}

function checkPartStock(sel) {
  const partId = sel.value;
  if (!partId) return;
  const p = db.parts.find(x => x.id === partId);
  if (!p) return;
  const qtyInput = sel.nextElementSibling;
  qtyInput.max = p.remaining;
  showToast(`Available stock: ${p.remaining} units`, 'info');
}

function submitService(e) {
  e.preventDefault();
  const custId = document.getElementById('selectedCustomerId').value;
  if (!custId) { showToast('⚠️ Please select a customer first!', 'error'); return; }

  // Collect parts used
  const rows = document.querySelectorAll('#partsUsedContainer .part-row');
  const partsUsed = [];
  let valid = true;

  rows.forEach(row => {
    const sel = row.querySelector('select');
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (!sel.value) return;
    const part = db.parts.find(p => p.id === sel.value);
    if (!part) return;
    if (qty < 1 || qty > part.remaining) {
      showToast(`⚠️ Invalid quantity for "${part.partName}". Available: ${part.remaining}`, 'error');
      valid = false; return;
    }
    partsUsed.push({ partId: part.id, qty });
  });

  if (!valid) return;

  // Deduct parts
  partsUsed.forEach(p => {
    const part = db.parts.find(x => x.id === p.partId);
    if (part) { part.usedCount += p.qty; part.remaining -= p.qty; }
  });

  const service = {
    id: genId(),
    customerId: custId,
    date: document.getElementById('serviceDate').value,
    serviceType: document.getElementById('serviceType').value,
    parts: partsUsed,
    notes: document.getElementById('serviceNotes').value.trim(),
    laborCost: parseFloat(document.getElementById('laborCost').value) || 0,
    createdAt: new Date().toISOString()
  };

  db.services.unshift(service);
  saveData();
  renderPartsList();
  updatePartsStats();
  renderRecords();

  // Reset form
  document.getElementById('serviceForm').reset();
  document.getElementById('serviceDate').value = today();
  document.getElementById('partsUsedContainer').innerHTML = '';
  document.getElementById('selectedCustomerCard').style.display = 'none';
  document.getElementById('selectedCustomerId').value = '';
  serviceCustomer = null;

  showToast('✅ Service record saved! Parts deducted from inventory.', 'success');
}

// ========== RECORDS ==========
function switchRecordsTab(tab) {
  currentRecordsTab = tab;
  document.querySelectorAll('.rec-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('rectab-' + tab).classList.add('active');
  document.getElementById('records-customers-panel').style.display = tab === 'customers' ? 'block' : 'none';
  document.getElementById('records-services-panel').style.display = tab === 'services' ? 'block' : 'none';
}

function renderRecords() {
  renderAllCustomers();
  renderAllServices();
}

function renderAllCustomers() {
  const el = document.getElementById('allCustomersList');
  if (db.customers.length === 0) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">👥</div><p>No customers registered yet.</p></div>';
    return;
  }
  el.innerHTML = db.customers.map(c => customerCardHTML(c)).join('');
}

function renderAllServices() {
  const el = document.getElementById('allServicesList');
  if (db.services.length === 0) {
    el.innerHTML = '<div class="empty-state"><div class="empty-icon">⚙️</div><p>No services recorded yet.</p></div>';
    return;
  }
  el.innerHTML = db.services.map(s => {
    const cust = db.customers.find(c => c.id === s.customerId);
    const partsHtml = (s.parts || []).map(p => {
      const pt = db.parts.find(x => x.id === p.partId);
      return pt ? `<span class="srv-part-tag">🔩 ${escHtml(pt.partName)} ×${p.qty}</span>` : '';
    }).join('');
    return `
    <div class="service-record-card">
      <div class="srv-header">
        <div>
          <div class="srv-customer">👤 ${cust ? escHtml(cust.name) : 'Unknown'} — ${cust ? escHtml(cust.carNumber) : ''}</div>
          <span class="srv-type">${escHtml(s.serviceType)}</span>
        </div>
        <div class="srv-date">📅 ${formatDate(s.date)}</div>
      </div>
      ${partsHtml ? `<div class="srv-parts">${partsHtml}</div>` : ''}
      ${s.notes ? `<div style="font-size:0.8rem;color:var(--text-muted);margin-top:0.4rem;">📝 ${escHtml(s.notes)}</div>` : ''}
      ${s.laborCost > 0 ? `<div style="font-size:0.8rem;color:var(--success);margin-top:0.25rem;">💰 Labour: ₹${s.laborCost}</div>` : ''}
    </div>`;
  }).join('');
}

// ========== EXCEL WORKBOOK BUILDER ==========
function buildWorkbook() {
  const wb = XLSX.utils.book_new();
  const ts = new Date().toLocaleString('en-IN');

  // Sheet 1: Customers
  const custData = [
    ['CoolAuto – Customer & Car Registry'],
    ['Last Updated:', ts],
    [],
    ['#', 'Customer Name', 'Phone', 'Car Number', 'Car Model', 'AC Type', 'Problem / Notes', 'Address', 'Registered Date']
  ];
  db.customers.forEach((c, i) =>
    custData.push([i+1, c.name, c.phone, c.carNumber, c.carModel, c.acType||'', c.problem||'', c.address||'', formatDate(c.date)])
  );
  const ws1 = XLSX.utils.aoa_to_sheet(custData);
  ws1['!cols'] = [5,20,14,14,16,14,30,20,14].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, 'Customer Registry');

  // Sheet 2: Parts Inventory
  const partsData = [
    ['CoolAuto – Parts Inventory'],
    ['Last Updated:', ts],
    [],
    ['#', 'Part Name', 'Part Number', 'Total Stock', 'Used', 'Remaining', 'Unit Price (₹)', 'Stock Status']
  ];
  db.parts.forEach((p, i) => {
    const status = p.remaining === 0 ? 'Out of Stock' : p.remaining <= 3 ? 'Low Stock' : 'In Stock';
    partsData.push([i+1, p.partName, p.partNumber, p.totalCount, p.usedCount, p.remaining, p.price||0, status]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(partsData);
  ws2['!cols'] = [5,20,14,12,10,12,14,12].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, 'Parts Inventory');

  // Sheet 3: Service Records
  const srvData = [
    ['CoolAuto – Service Records'],
    ['Last Updated:', ts],
    [],
    ['#', 'Date', 'Customer Name', 'Car Number', 'Car Model', 'Service Type', 'Parts Used', 'Labour Cost (₹)', 'Notes']
  ];
  db.services.forEach((s, i) => {
    const cust = db.customers.find(c => c.id === s.customerId);
    const partsStr = (s.parts||[]).map(p => {
      const pt = db.parts.find(x => x.id === p.partId);
      return pt ? `${pt.partName} x${p.qty}` : '';
    }).filter(Boolean).join(', ');
    srvData.push([
      i+1, formatDate(s.date),
      cust ? cust.name : 'Unknown',
      cust ? cust.carNumber : '',
      cust ? cust.carModel : '',
      s.serviceType, partsStr, s.laborCost||0, s.notes||''
    ]);
  });
  const ws3 = XLSX.utils.aoa_to_sheet(srvData);
  ws3['!cols'] = [5,12,20,14,16,16,30,14,25].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws3, 'Service Records');

  return wb;
}

// ========== AUTO-SAVE SCHEDULER (debounced, 2 sec) ==========
function scheduleExcelAutoSave() {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(async () => {
    if (excelFileHandle) {
      await writeToPicked();   // Linked file → silent write
    }
    // If not linked yet, do nothing – user clicks "Link Excel File" to set up
  }, 2000);
}

// ========== WRITE TO PICKED FILE (silent, no popup) ==========
async function writeToPicked() {
  if (!excelFileHandle) return;
  try {
    // Re-check permission (needed after page reload)
    let perm = await excelFileHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      perm = await excelFileHandle.requestPermission({ mode: 'readwrite' });
    }
    if (perm !== 'granted') {
      excelFileHandle = null;
      updateExcelStatus('unlinked');
      showToast('⚠️ Excel file permission lost. Please re-link.', 'error');
      return;
    }
    const wb  = buildWorkbook();
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const writable = await excelFileHandle.createWritable();
    await writable.write(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    await writable.close();
    lastSavedAt = new Date();
    updateExcelStatus('saved');
  } catch (err) {
    console.error('Auto-write failed:', err);
    showToast('❌ Excel auto-save failed. Check file permissions.', 'error');
    updateExcelStatus('error');
  }
}

// ========== LINK EXCEL FILE (one-time picker) ==========
async function linkExcelFile() {
  if (!SUPPORTS_FSA) {
    // Fallback: just download
    exportToExcel();
    showToast('ℹ️ Auto-link not supported on this browser. File downloaded instead.', 'info');
    return;
  }
  try {
    excelFileHandle = await window.showSaveFilePicker({
      suggestedName: 'CoolAuto_Data.xlsx',
      types: [{ description: 'Excel Spreadsheet', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
    });
    await writeToPicked();
    updateExcelStatus('saved');
    showToast('✅ Excel file linked! It will auto-update on every change.', 'success');
  } catch (e) {
    if (e.name !== 'AbortError') {
      showToast('❌ Could not link Excel file.', 'error');
    }
  }
}

// ========== MANUAL DOWNLOAD (always available) ==========
function exportToExcel() {
  const wb = buildWorkbook();
  const filename = `CoolAuto_Data_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast('📊 Excel downloaded! Check your Downloads folder.', 'success');
}

// ========== STATUS BADGE ==========
function updateExcelStatus(state) {
  const badge = document.getElementById('excelStatusBadge');
  if (!badge) return;
  if (state === 'saved') {
    const t = lastSavedAt ? lastSavedAt.toLocaleTimeString('en-IN') : '';
    badge.textContent = `✅ Auto-saved ${t}`;
    badge.className = 'excel-status saved';
  } else if (state === 'unlinked') {
    badge.textContent = '🔗 Not linked';
    badge.className = 'excel-status unlinked';
  } else {
    badge.textContent = '❌ Save error';
    badge.className = 'excel-status error';
  }
}

// ========== EXCEL INFO PANEL ==========
function showExcelPathInfo() {
  const regPanel = document.getElementById('panel-register');
  const info = document.createElement('div');
  info.innerHTML = `
  <div class="excel-info-card" id="excelInfoCard">
    <div class="excel-info-header">
      <span class="excel-info-icon">📊</span>
      <span class="excel-info-title">Excel Auto-Update</span>
      <span class="excel-status unlinked" id="excelStatusBadge">🔗 Not linked</span>
    </div>
    <div class="excel-info-body">
      <p>Link an Excel file <strong>once</strong> — it will <strong>update automatically</strong> every time you add a customer, part, or service record.</p>
      <div class="excel-actions">
        <button class="btn btn-primary btn-sm" onclick="linkExcelFile()" id="linkExcelBtn">
          ${SUPPORTS_FSA ? '🔗 Link Excel File (Auto-Update)' : '📥 Download Excel Now'}
        </button>
        <button class="btn btn-outline btn-sm" onclick="exportToExcel()">
          💾 Download Backup
        </button>
      </div>
      <div class="excel-path-info">
        <div class="path-row">💻 <strong>PC</strong> &nbsp; C:\\Users\\[Name]\\Downloads\\CoolAuto_Data.xlsx</div>
        <div class="path-row">📱 <strong>Mobile</strong> &nbsp; Downloads → CoolAuto_Data.xlsx</div>
        ${SUPPORTS_FSA
          ? '<div class="path-row fsa-note">✅ Your browser supports auto-update. Click "Link Excel File" to set it up.</div>'
          : '<div class="path-row fsa-note">ℹ️ Auto-update requires Chrome or Edge. Use Download Backup on other browsers.</div>'}
      </div>
    </div>
  </div>`;
  regPanel.appendChild(info);
}

// ========== MODAL ==========
function openModal() {
  document.getElementById('modalOverlay').classList.add('active');
  document.getElementById('modal').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('active');
  document.getElementById('modal').classList.remove('active');
  document.body.style.overflow = '';
}

// ========== TOAST ==========
let toastTimer;
function showToast(msg, type = 'info') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `toast ${type} show`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.classList.remove('show'); }, 3200);
}

// ========== UTILS ==========
function genId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate(d) {
  if (!d) return '—';
  try {
    const dt = new Date(d + 'T00:00:00');
    return dt.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' });
  } catch(e) { return d; }
}

// Close dropdown when clicking outside
document.addEventListener('click', (e) => {
  if (!e.target.closest('.form-group')) {
    const dd = document.getElementById('serviceCustomerDropdown');
    if (dd) dd.style.display = 'none';
  }
});
