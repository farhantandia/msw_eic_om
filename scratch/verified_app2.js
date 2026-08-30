
        let currentUnit = 1;
        let currentTab = 'wo';
        let currentViewMode = 'cards';
        let actuatorViewMode = 'unit';
        let instSubtab = 'ptx';
        let quickFilter = 'all';
        let fullData = null;
        let matrixData = null;
        let currentPage = 1;
        let pageSize = 10;
        let editModeState = {};
        let activeFinding = null;

        // Date Helpers for Calendar Pickers
        function formatDateForInput(dateVal) {
            if(!dateVal) return '';
            const str = String(dateVal).trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
                return str.substring(0, 10);
            }
            if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
                const parts = str.split('/');
                const day = parts[0].padStart(2, '0');
                const month = parts[1].padStart(2, '0');
                const year = parts[2].substring(0, 4);
                return `${year}-${month}-${day}`;
            }
            const d = new Date(str);
            if (!isNaN(d.getTime())) {
                const year = d.getFullYear();
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const day = String(d.getDate()).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            return '';
        }

        function formatDateForStorage(inputVal) {
            if(!inputVal) return '';
            const str = String(inputVal).trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
                const parts = str.split('-');
                return `${parts[2]}/${parts[1]}/${parts[0]}`;
            }
            return str;
        }

        // Toast System
        function showToast(message, type = 'success', duration = 3000) {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            const icon = type === 'success' ? '✅' : (type === 'error' ? '❌' : 'ℹ️');
            toast.innerHTML = `<span style="font-size:1.1rem;">${icon}</span><div>${message}</div>`;
            container.appendChild(toast);
            setTimeout(() => {
                toast.style.opacity = '0';
                toast.style.transform = 'translateY(20px)';
                toast.style.transition = 'all 0.3s';
                setTimeout(() => toast.remove(), 300);
            }, duration);
        }

        async function loadData() {
            try {
                const [res, compRes] = await Promise.all([
                    fetch(`/api/data?unit=${currentUnit}`),
                    fetch(`/api/master_components?unit=${currentUnit}`).catch(() => null)
                ]);
                if(!res.ok) throw new Error(`HTTP ${res.status} Server Error`);
                fullData = await res.json();
                if(compRes && compRes.ok) {
                    const compData = await compRes.json();
                    fullData.master_actuators = compData.actuators || [];
                    fullData.master_instruments = compData.instruments || [];
                }
                const titleEl = document.getElementById('outage-unit-title');
                if(titleEl) titleEl.innerText = `Monitoring Progress Outage EIC Unit ${currentUnit}`;
                
                renderStats();
                populateFilterDropdowns();
                renderTabContent();
            } catch(e) {
                console.error("Error loading data:", e);
                document.getElementById('tab-content').innerHTML = `
                <div style="color:#ef4444; padding:24px; background:var(--bg-card); border-radius:12px; border:1px solid rgba(239,68,68,0.3);">
                    <strong style="font-size:1.1rem;">⚠️ Gagal memuat data Unit ${currentUnit}</strong>
                    <div style="margin-top:8px; font-size:0.88rem; color:var(--text-main);">Detail Kendala: ${e.message}</div>
                    <div style="margin-top:12px; font-size:0.82rem; color:var(--text-muted);">Tips: Pastikan server.exe atau start_app.bat sedang berjalan dan file Excel <code>Template_Outage_EIC_Monitoring_unit ${currentUnit}.xlsx</code> berada di folder yang sama.</div>
                </div>`;
            }
        }

        function switchUnit(unit) {
            currentUnit = unit;
            currentPage = 1;
            document.getElementById('btn-unit-1').classList.toggle('active', unit === 1);
            document.getElementById('btn-unit-2').classList.toggle('active', unit === 2);
            loadData();
        }

        function switchTab(tab) {
            currentTab = tab;
            currentPage = 1;
            document.querySelectorAll('.nav-tabs .tab-btn').forEach((btn, i) => {
                btn.classList.toggle('active', ['wo', 'actuator', 'instrument', 'scope'][i] === tab);
            });
            document.getElementById('view-switcher-box').style.display = (tab === 'scope') ? 'none' : 'flex';
            renderTabContent();
        }

        function switchViewMode(mode) {
            currentViewMode = mode;
            document.getElementById('btn-view-cards').classList.toggle('active', mode === 'cards');
            document.getElementById('btn-view-table').classList.toggle('active', mode === 'table');
            renderTabContent();
        }

        function setQuickFilter(f) {
            quickFilter = f;
            ['all', 'findings', 'inprog', 'finish'].forEach(p => {
                const el = document.getElementById(`pill-${p}`);
                if(el) el.classList.toggle('active', p === f);
            });
            currentPage = 1;
            renderTabContent();
        }

        function renderStats() {
            if(!fullData || !fullData.summary) return;
            const s = fullData.summary;

            document.getElementById('grand-pct').innerText = `${s.grand_pct}%`;
            document.getElementById('grand-counts').innerText = `${s.grand_done} / ${s.grand_total} Sub-task / Item`;
            document.getElementById('grand-bar-fill').style.width = `${s.grand_pct}%`;

            document.getElementById('wo-pct').innerText = `${s.wo.pct}%`;
            document.getElementById('wo-counts').innerText = `${s.wo.subtask_done} / ${s.wo.subtask_total} Sub-task`;
            document.getElementById('wo-sub').innerText = `WO Selesai: ${s.wo.finish} / ${s.wo.total} | In-Prog: ${s.wo.in_progress}`;

            document.getElementById('act-pct').innerText = `${s.actuator.pct}%`;
            document.getElementById('act-counts').innerText = `${s.actuator.subtask_done} / ${s.actuator.subtask_total} Sub-task`;
            document.getElementById('act-sub').innerText = `Valve Selesai: ${s.actuator.finish} / ${s.actuator.total} | In-Prog: ${s.actuator.in_progress}`;

            document.getElementById('inst-pct').innerText = `${s.instrument.pct}%`;
            document.getElementById('inst-counts').innerText = `${s.instrument.done} / ${s.instrument.total}`;
            document.getElementById('inst-sub').innerText = `PTX (${(fullData.pressure_tx||[]).length}) | TTX (${(fullData.temperature_tx||[]).length}) | PSW (${(fullData.pressure_switch||[]).length})`;

            document.getElementById('findings-count').innerText = s.findings_count;

            document.getElementById('tab-cnt-wo').innerText = (fullData.work_orders || []).length;
            document.getElementById('tab-cnt-act').innerText = (fullData.actuators || []).length;
            document.getElementById('tab-cnt-inst').innerText = s.instrument.total;

            // Update Sticky Summary Bar
            const uNum = document.getElementById('sticky-unit-num');
            const gPct = document.getElementById('sticky-grand-pct');
            const wPct = document.getElementById('sticky-wo-pct');
            const aPct = document.getElementById('sticky-act-pct');
            const iPct = document.getElementById('sticky-inst-pct');
            if(uNum) uNum.innerText = currentUnit;
            if(gPct) gPct.innerText = `${s.grand_pct}%`;
            if(wPct) wPct.innerText = `${s.wo.pct}%`;
            if(aPct) aPct.innerText = `${s.actuator.pct}%`;
            if(iPct) iPct.innerText = `${s.instrument.pct}%`;
        }

        function populateFilterDropdowns() {
            const picSelect = document.getElementById('filter-pic');
            if(picSelect) {
                const current = picSelect.value;
                picSelect.innerHTML = '<option value="">Semua PIC</option>';
                (fullData.pics || []).forEach(p => {
                    if(p) picSelect.innerHTML += `<option value="${p}" ${current===p?'selected':''}>${p}</option>`;
                });
            }

            const areaSelect = document.getElementById('filter-area');
            if(areaSelect) {
                const current = areaSelect.value;
                const areas = new Set();
                (fullData.work_orders || []).forEach(w => { if(w && w.area) areas.add(w.area); });
                (fullData.actuators || []).forEach(a => { if(a && a.area) areas.add(a.area); });
                (fullData.pressure_tx || []).forEach(p => { if(p && p.area) areas.add(p.area); });
                areaSelect.innerHTML = '<option value="">Semua Area</option>';
                Array.from(areas).sort().forEach(a => {
                    areaSelect.innerHTML += `<option value="${a}" ${current===a?'selected':''}>${a}</option>`;
                });
            }
        }

        function applyFilters() {
            currentPage = 1;
            renderTabContent();
        }

        function toggleEditMode(itemId) {
            editModeState[itemId] = !editModeState[itemId];
            renderTabContent();
        }

        function renderPaginationControls(totalItems) {
            const totalPages = Math.ceil(totalItems / pageSize) || 1;
            if(currentPage > totalPages) currentPage = totalPages;

            return `
            <div class="pagination-bar">
                <div>Menampilkan <strong>${totalItems === 0 ? 0 : (currentPage - 1) * pageSize + 1} - ${Math.min(currentPage * pageSize, totalItems)}</strong> dari <strong>${totalItems}</strong> item</div>
                <div style="display:flex; align-items:center; gap:8px;">
                    <button class="page-btn" onclick="changePage(1)" ${currentPage===1?'disabled':''}>⏮️ Awal</button>
                    <button class="page-btn" onclick="changePage(${currentPage-1})" ${currentPage===1?'disabled':''}>◀️ Prev</button>
                    <span style="font-weight:700; margin:0 6px;">Halaman ${currentPage} / ${totalPages}</span>
                    <button class="page-btn" onclick="changePage(${currentPage+1})" ${currentPage===totalPages?'disabled':''}>Next ▶️</button>
                    <button class="page-btn" onclick="changePage(${totalPages})" ${currentPage===totalPages?'disabled':''}>Akhir ⏭️</button>
                    <select class="filter-input" style="padding:4px 8px; margin-left:10px;" onchange="changePageSize(this.value)">
                        <option value="10" ${pageSize===10?'selected':''}>10 / hal</option>
                        <option value="25" ${pageSize===25?'selected':''}>25 / hal</option>
                        <option value="50" ${pageSize===50?'selected':''}>50 / hal</option>
                        <option value="1000" ${pageSize===1000?'selected':''}>Semua</option>
                    </select>
                </div>
            </div>`;
        }

        function changePage(page) {
            currentPage = page;
            renderTabContent();
            window.scrollTo({top: 350, behavior: 'smooth'});
        }

        function changePageSize(size) {
            pageSize = parseInt(size);
            currentPage = 1;
            renderTabContent();
        }

        function filterItem(item, searchStr, statusFilter, picFilter, areaFilter, codeKey, descKey, statusKey, picKey, areaKey) {
            const code = (item[codeKey] || '').toLowerCase();
            const desc = (item[descKey] || '').toLowerCase();
            const remarks = (item.remarks || '').toLowerCase();
            const matchSearch = !searchStr || code.includes(searchStr) || desc.includes(searchStr) || remarks.includes(searchStr);
            
            const matchStatus = !statusFilter || item[statusKey] === statusFilter;
            const matchPic = !picFilter || item[picKey] === picFilter;
            const matchArea = !areaFilter || item[areaKey] === areaFilter;

            let matchQuick = true;
            if(quickFilter === 'findings') {
                matchQuick = !!item.temuan || (item.jumlah_foto > 0);
            } else if(quickFilter === 'inprog') {
                matchQuick = item[statusKey] !== 'FINISH' && !item.status_wdone;
            } else if(quickFilter === 'finish') {
                matchQuick = item[statusKey] === 'FINISH' || item.status_wdone === true;
            }

            return matchSearch && matchStatus && matchPic && matchArea && matchQuick;
        }

        /* ---------------- WORK ORDER RENDERING ---------------- */
        function renderWorkOrders(container) {
            const searchStr = document.getElementById('search-input').value.toLowerCase();
            const statusFilter = document.getElementById('filter-status').value;
            const picFilter = document.getElementById('filter-pic').value;
            const areaFilter = document.getElementById('filter-area').value;

            let filteredItems = (fullData.work_orders || []).filter(w => 
                filterItem(w, searchStr, statusFilter, picFilter, areaFilter, 'no_wo', 'job_description', 'status', 'pic', 'area')
            );

            const isAddWoOpen = openCardIds.has('add-wo-form');
            let html = '';
            html += `
            <div style="background: linear-gradient(135deg, rgba(6,182,212,0.12) 0%, rgba(245,158,11,0.12) 100%); border: 1px solid rgba(56,189,248,0.3); border-radius: 10px; padding: 12px 18px; margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap;">
                <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.4rem;">⚙️📟</span>
                    <div>
                        <div style="font-size:0.9rem; font-weight:700; color:var(--text-main);">Sinkronisasi Dua Arah & Dropdown Komponen Aktif!</div>
                        <div style="font-size:0.8rem; color:var(--text-muted);">Centang subtask Actuator/Instrument di WO akan otomatis mengupdate tab Actuator & Instrument, dan sebaliknya. Klik kartu WO untuk melihat checklist & dropdown selector.</div>
                    </div>
                </div>
                <div style="display:flex; gap:8px;">
                    <span class="badge-tag-comp badge-tag-act">⚙️ 92 Actuator Tersinkron</span>
                    <span class="badge-tag-comp badge-tag-inst">📟 153 Instrument Tersinkron</span>
                </div>
            </div>`;

            html += `
            <div style="margin-bottom:20px; background:var(--bg-card); border-radius:var(--radius-md); padding:16px; border:1px solid var(--border-color);">
                <div style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;" onclick="toggleAccordion('add-wo-form')">
                    <h3 style="font-size:0.95rem; color:var(--primary); font-weight:700;">➕ Tambah Work Order (WO) Baru</h3>
                    <span id="arrow-add-wo-form" style="font-weight:700; color:var(--text-muted); font-size:0.85rem;">${isAddWoOpen ? '▲ Tutup Form' : '▼ Buka Form'}</span>
                </div>
                <div id="add-wo-form" class="accordion-form ${isAddWoOpen ? 'open' : ''}" style="${isAddWoOpen ? 'display:block;' : ''}">
                    <div class="form-grid">
                        <div class="form-group">
                            <label>No Work Order (WO) <span style="color:#f43f5e;">*</span></label>
                            <input type="text" id="new-wo-code" class="filter-input" placeholder="mis. WO-100826-0099">
                        </div>
                        <div class="form-group">
                            <label>Job Description <span style="color:#f43f5e;">*</span></label>
                            <input type="text" id="new-wo-desc" class="filter-input" placeholder="Deskripsi pekerjaan WO...">
                        </div>
                        <div class="form-group">
                            <label>Area System</label>
                            <input type="text" id="new-wo-area" class="filter-input" placeholder="BOILER, ID FAN, COOLING TOWER...">
                        </div>
                        <div class="form-group">
                            <label>PIC Penanggung Jawab</label>
                            <select id="new-wo-pic" class="filter-input">
                                <option value="">Pilih PIC...</option>
                                ${(fullData.pics || []).map(p => `<option value="${p}">${p}</option>`).join('')}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Tanggal Schedule</label>
                            <input type="date" id="new-wo-sched" class="filter-input">
                        </div>
                        <div class="form-group" style="grid-column: span 2;">
                            <label>Checklist Sub-task (Pisahkan dengan koma atau baris baru)</label>
                            <textarea id="new-wo-checklist" class="textarea-full" placeholder="mis. General Inspection, Cleaning Contact, Function Test, Tightening Bolt"></textarea>
                        </div>
                    </div>
                    <div style="display:flex; gap:10px; margin-top:14px;">
                        <button class="btn-save" onclick="saveNewWO()">💾 Simpan Work Order Baru</button>
                        <button class="page-btn" onclick="toggleAccordion('add-wo-form')">Batal</button>
                    </div>
                </div>
            </div>`;

            html += renderPaginationControls(filteredItems.length);

            if(filteredItems.length === 0) {
                html += '<div style="text-align:center; padding:50px; color:var(--text-muted); background:var(--bg-card); border-radius:12px;">Tidak ada Work Order yang sesuai dengan filter pencarian.</div>';
                container.innerHTML = html;
                return;
            }

            const startIndex = (currentPage - 1) * pageSize;
            const pageItems = filteredItems.slice(startIndex, startIndex + pageSize);

            if(currentViewMode === 'cards') {
                if(openCardIds.size === 0 && pageItems.length > 0) {
                    const firstBodyId = getCardBodyId('wo', pageItems[0].no_wo || 0);
                    openCardIds.add(firstBodyId);
                }
                html += '<div class="card-list">';
                pageItems.forEach((item, idx) => {
                    const st = String(item.status || 'SCHED-OK').replace(/\s+/g, '_');
                    const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                    const doneCount = (item.checklist || []).filter(c => c.selesai).length;
                    const totalCount = item.checklist ? item.checklist.length : 0;
                    const bodyId = getCardBodyId('wo', item.no_wo || idx);
                    const isOpen = openCardIds.has(bodyId);

                    html += `
                    <div class="item-card" id="card-wo-${item.no_wo}">
                        <div class="item-header" onclick="toggleAccordion('${bodyId}')">
                            <div class="item-title-box">
                                <div class="item-code">${item.no_wo} &bull; ${item.area || 'GENERAL'} ${item.tanggal_finish ? '&bull; 🏁 Selesai: ' + item.tanggal_finish : ''}</div>
                                <div class="item-name">${item.job_description}</div>
                            </div>
                            <div class="header-actions">
                                <button class="btn-finding ${hasFindings?'active':''}" onclick="event.stopPropagation(); openFindingModal('wo', '${item.no_wo}', '${item.no_wo} - ${item.job_description.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}')">
                                    📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Temuan' : '+ Temuan')}
                                </button>
                                <span style="font-size:0.82rem; color:var(--text-muted); font-weight:600;">👤 ${item.pic || '-'}</span>
                                <span class="wo-subtask-progress" style="font-size:0.82rem; font-weight:700; color:var(--text-muted);">${doneCount} / ${totalCount} Sub-task</span>
                                <span class="status-badge badge-${st}">${item.status}</span>
                                <div class="progress-box">
                                    <div class="progress-bar-bg"><div class="progress-bar-fill progress-fill" style="width:${item.persen_progress}%;"></div></div>
                                    <span class="progress-text" style="font-size:0.85rem; font-weight:800; font-family:'JetBrains Mono';">${item.persen_progress}%</span>
                                </div>
                            </div>
                        </div>
                        <div class="item-body ${isOpen ? 'open' : ''}" id="${bodyId}" style="${isOpen ? 'display:block;' : ''}">
                            <div class="checklist-section">
                                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:8px;">
                                    <div class="section-h4" style="margin-bottom:0;">📋 Checklist Sub-Task (${doneCount} / ${totalCount} Selesai)</div>
                                    ${totalCount > 0 ? `
                                        <div style="display:flex; gap:6px;">
                                            <button class="btn-batch-check" onclick="batchToggleSubtasks('${item.no_wo}', 'mark_all_done')" title="Tandai semua sub-task selesai">✓ Selesai Semua</button>
                                            <button class="btn-batch-reset" onclick="batchToggleSubtasks('${item.no_wo}', 'reset_all')" title="Reset semua sub-task">↺ Reset</button>
                                        </div>
                                    ` : ''}
                                </div>
                                <div class="checklist-grid">
                                    ${(item.checklist || []).map((c, cIdx) => {
                                        const typeBadge = getSubtaskTypeBadge(c.sub_task);
                                        return `
                                        <div class="checklist-item ${c.selesai ? 'done' : ''}">
                                            <div class="checklist-item-header">
                                                <div class="header-left">
                                                    ${typeBadge}
                                                </div>
                                                <div class="header-right">
                                                    ${c.tanggal ? `<span class="date-badge" style="font-size:0.7rem; color:var(--status-finish); font-family:'JetBrains Mono',monospace; background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.25); padding:1px 6px; border-radius:3px;" title="Tanggal Dikerjakan">${c.tanggal}</span>` : ''}
                                                    <button class="btn-del-subtask-cross" title="Hapus Subtask" onclick="deleteSubtask('${item.no_wo}', '${(c.sub_task||'').toString().replace(/'/g, "\\'")}')" aria-label="Hapus">&times;</button>
                                                </div>
                                            </div>
                                            <label class="checklist-item-body">
                                                <input type="checkbox" id="chk-${item.no_wo}-${cIdx}" ${c.selesai ? 'checked' : ''} onchange="toggleLocalSubtask('${item.no_wo}', ${cIdx}, this.checked)">
                                                <span>${c.sub_task}</span>
                                            </label>
                                        </div>`;
                                    }).join('')}
                                </div>
                                
                                <div style="margin-top:14px; background:rgba(0,0,0,0.2); border:1px solid var(--border-color); border-radius:8px; padding:10px 12px;">
                                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:6px;">
                                        <span style="font-size:0.78rem; font-weight:700; color:var(--text-muted);">➕ Tambah Sub-task ke WO:</span>
                                        <div style="display:flex; gap:5px;">
                                            <button type="button" class="comp-mode-btn btn-mode-manual ${(subtaskAddModes[item.no_wo]||'manual')==='manual'?'active':''}" onclick="setSubtaskMode('${item.no_wo}', 'manual')">✏️ Manual</button>
                                            <button type="button" class="comp-mode-btn btn-mode-act ${subtaskAddModes[item.no_wo]==='actuator'?'active':''}" onclick="setSubtaskMode('${item.no_wo}', 'actuator')">⚙️ Pilih Actuator (${(fullData.master_actuators||fullData.actuators||[]).length})</button>
                                            <button type="button" class="comp-mode-btn btn-mode-inst ${subtaskAddModes[item.no_wo]==='instrument'?'active':''}" onclick="setSubtaskMode('${item.no_wo}', 'instrument')">📟 Pilih Instrument (${((fullData.pressure_tx||[]).length + (fullData.temperature_tx||[]).length + (fullData.pressure_switch||[]).length)})</button>
                                        </div>
                                    </div>
                                    
                                    <!-- Mode 1: Manual Input -->
                                    <div id="box-subtask-manual-${item.no_wo}" style="display:${(subtaskAddModes[item.no_wo]||'manual')==='manual'?'flex':'none'}; gap:8px;">
                                        <input type="text" id="new-subtask-${item.no_wo}" class="filter-input" placeholder="Ketik deskripsi sub-task (mis. INSPECTION FRAME MOTOR)..." style="flex-grow:1; font-size:0.82rem;">
                                        <button class="btn-save" style="padding:6px 14px; font-size:0.8rem; white-space:nowrap;" onclick="addSubtask('${item.no_wo}', 'manual')">➕ Tambah</button>
                                    </div>

                                    <!-- Mode 2: Actuator Dropdown Picker -->
                                    <div id="box-subtask-act-${item.no_wo}" style="display:${subtaskAddModes[item.no_wo]==='actuator'?'flex':'none'}; gap:8px;">
                                        <select id="new-subtask-act-${item.no_wo}" class="filter-input" style="flex-grow:1; font-size:0.82rem;">
                                            <option value="">-- Pilih Actuator Valve dari Master List --</option>
                                            ${(fullData.master_actuators || fullData.actuators || []).map(a => `
                                                <option value="${a.equipment_description} ${a.kks||''}">[${a.area}] ${a.equipment_description} ${a.kks ? '('+a.kks+')' : ''}</option>
                                            `).join('')}
                                        </select>
                                        <button class="btn-save" style="padding:6px 14px; font-size:0.8rem; white-space:nowrap; background:#f59e0b; border-color:#d97706; color:#000;" onclick="addSubtask('${item.no_wo}', 'actuator')">⚙️ Tambah Actuator</button>
                                    </div>

                                    <!-- Mode 3: Instrument Dropdown Picker -->
                                    <div id="box-subtask-inst-${item.no_wo}" style="display:${subtaskAddModes[item.no_wo]==='instrument'?'flex':'none'}; gap:8px;">
                                        <select id="new-subtask-inst-${item.no_wo}" class="filter-input" style="flex-grow:1; font-size:0.82rem;">
                                            <option value="">-- Pilih Instrument dari Master List (PT/TT/PS) --</option>
                                            <optgroup label="Pressure Transmitter (PTX)">
                                                ${(fullData.pressure_tx || []).map(p => `
                                                    <option value="${p.kks ? p.kks+': ' : ''}${p.equipment}">[PTX - ${p.area}] ${p.kks ? p.kks+' : ' : ''}${p.equipment}</option>
                                                `).join('')}
                                            </optgroup>
                                            <optgroup label="Temperature Transmitter (TTX)">
                                                ${(fullData.temperature_tx || []).map(t => `
                                                    <option value="${t.kks ? t.kks+': ' : ''}${t.equipment}">[TTX - ${t.area}] ${t.kks ? t.kks+' : ' : ''}${t.equipment}</option>
                                                `).join('')}
                                            </optgroup>
                                            <optgroup label="Pressure Switch (PSW)">
                                                ${(fullData.pressure_switch || []).map(s => `
                                                    <option value="${s.kks ? s.kks+': ' : ''}${s.equipment}">[PSW - ${s.area}] ${s.kks ? s.kks+' : ' : ''}${s.equipment}</option>
                                                `).join('')}
                                            </optgroup>
                                        </select>
                                        <button class="btn-save" style="padding:6px 14px; font-size:0.8rem; white-space:nowrap; background:#06b6d4; border-color:#0891b2; color:#000;" onclick="addSubtask('${item.no_wo}', 'instrument')">📟 Tambah Instrument</button>
                                    </div>
                                </div>
                            </div>

                            <div class="form-grid">
                                <div class="form-group">
                                    <label>PIC Penanggung Jawab</label>
                                    <select id="pic-${item.no_wo}" class="filter-input">
                                        <option value="">Pilih PIC...</option>
                                        ${fullData.pics.map(p => `<option value="${p}" ${item.pic===p?'selected':''}>${p}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Tanggal Selesai (Finish Date)</label>
                                    <input type="date" id="finish-wo-${item.no_wo}" class="filter-input" value="${formatDateForInput(item.tanggal_finish)}">
                                </div>
                                <div class="form-group" style="grid-column: span 2;">
                                    <label>Remarks / Catatan Lapangan</label>
                                    <input type="text" id="rem-${item.no_wo}" class="filter-input" value="${item.remarks || ''}" placeholder="Catatan pekerjaan...">
                                </div>
                            </div>

                            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:15px; border-top:1px solid var(--border-color); padding-top:12px;">
                                <div>
                                    <button class="btn-danger" onclick="deleteWO('${item.no_wo}')">🗑️ Hapus WO</button>
                                </div>
                                <div style="display:flex; gap:10px;">
                                    <button class="btn-save" onclick="saveWorkOrder('${item.no_wo}')">💾 Simpan Perubahan</button>
                                </div>
                            </div>
                        </div>
                    </div>`;
                });
                html += '</div>';
            } else {
                html += `
                <div class="table-wrap">
                    <table class="dense-table">
                        <thead>
                            <tr>
                                <th>No WO</th>
                                <th>Deskripsi Pekerjaan</th>
                                <th>Area</th>
                                <th>PIC</th>
                                <th>Subtask</th>
                                <th>Progress</th>
                                <th>Status</th>
                                <th>Tgl Selesai</th>
                                <th>Temuan & Foto</th>
                                <th>Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${pageItems.map(item => {
                                const st = (item.status || 'SCHED-OK').replace(/\s+/g, '_');
                                const doneCount = (item.checklist || []).filter(c => c.selesai).length;
                                const totalCount = item.checklist ? item.checklist.length : 0;
                                const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                                return `
                                <tr>
                                    <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.no_wo}</td>
                                    <td style="font-weight:600;">${item.job_description}</td>
                                    <td><span style="font-size:0.8rem; color:var(--text-muted);">${item.area || 'GENERAL'}</span></td>
                                    <td style="font-weight:600;">${item.pic || '-'}</td>
                                    <td><span style="font-size:0.82rem; font-weight:700;">${doneCount}/${totalCount}</span></td>
                                    <td>
                                        <div style="display:flex; align-items:center; gap:6px;">
                                            <div class="progress-bar-bg" style="width:60px;"><div class="progress-bar-fill" style="width:${item.persen_progress}%;"></div></div>
                                            <span style="font-size:0.8rem; font-weight:700;">${item.persen_progress}%</span>
                                        </div>
                                    </td>
                                    <td><span class="status-badge badge-${st}">${item.status}</span></td>
                                    <td style="font-family:'JetBrains Mono'; font-size:0.8rem; color:var(--text-muted);">${item.tanggal_finish || '-'}</td>
                                    <td>
                                        <button class="btn-finding ${hasFindings?'active':''}" onclick="openFindingModal('wo', '${item.no_wo}', '${item.no_wo} - ${item.job_description.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}')">
                                            📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Ada Temuan' : '+ Foto')}
                                        </button>
                                    </td>
                                    <td>
                                        <button class="page-btn" style="padding:4px 8px; font-size:0.78rem;" onclick="switchViewMode('cards'); toggleAccordion('body-wo-0');">Detail</button>
                                    </td>
                                </tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                </div>`;
            }

            html += renderPaginationControls(filteredItems.length);
            container.innerHTML = html;
        }

        /* ---------------- ACTUATOR VALVE RENDERING ---------------- */
        function renderActuators(container) {
            const searchStr = document.getElementById('search-input').value.toLowerCase();
            const statusFilter = document.getElementById('filter-status').value;
            const picFilter = document.getElementById('filter-pic').value;
            const areaFilter = document.getElementById('filter-area').value;

            let html = '';

            let filteredItems = (fullData.actuators || []).filter(a => 
                filterItem(a, searchStr, statusFilter, picFilter, areaFilter, 'equipment_id', 'equipment_description', 'status', 'pic', 'area')
            );

            const isAddActOpen = openCardIds.has('add-act-form');
            html += `
            <div style="margin-bottom:20px; background:var(--bg-card); border-radius:var(--radius-md); padding:16px; border:1px solid var(--border-color);">
                <div style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;" onclick="toggleAccordion('add-act-form')">
                    <h3 style="font-size:0.95rem; color:var(--primary); font-weight:700;">➕ Tambah Actuator Valve Baru</h3>
                    <span id="arrow-add-act-form" style="font-weight:700; color:var(--text-muted); font-size:0.85rem;">${isAddActOpen ? '▲ Tutup Form' : '▼ Buka Form'}</span>
                </div>
                <div id="add-act-form" class="accordion-form ${isAddActOpen ? 'open' : ''}" style="${isAddActOpen ? 'display:block;' : ''}">
                    <div class="form-grid">
                        <div class="form-group">
                            <label>Equipment ID <span style="color:#f43f5e;">*</span></label>
                            <input type="text" id="new-act-id" class="filter-input" placeholder="mis. AV-099">
                        </div>
                        <div class="form-group">
                            <label>Deskripsi Equipment <span style="color:#f43f5e;">*</span></label>
                            <input type="text" id="new-act-desc" class="filter-input" placeholder="ACTUATOR FEED WATER...">
                        </div>
                        <div class="form-group">
                            <label>Area System</label>
                            <input type="text" id="new-act-area" class="filter-input" placeholder="BOILER, ID FAN...">
                        </div>
                        <div class="form-group">
                            <label>Tag KKS</label>
                            <input type="text" id="new-act-kks" class="filter-input" placeholder="10LAB30AA210">
                        </div>
                        <div class="form-group">
                            <label>PIC Penanggung Jawab</label>
                            <select id="new-act-pic" class="filter-input">
                                <option value="">Pilih PIC...</option>
                                ${(fullData.pics || []).map(p => `<option value="${p}">${p}</option>`).join('')}
                            </select>
                        </div>
                    </div>
                    <div style="display:flex; gap:10px; margin-top:14px;">
                        <button class="btn-save" onclick="saveNewActuator()">💾 Simpan Actuator Baru</button>
                        <button class="page-btn" onclick="toggleAccordion('add-act-form')">Batal</button>
                    </div>
                </div>
            </div>`;

            html += renderPaginationControls(filteredItems.length);

            if(filteredItems.length === 0) {
                html += '<div style="text-align:center; padding:50px; color:var(--text-muted); background:var(--bg-card); border-radius:12px;">Tidak ada Actuator Valve yang sesuai filter.</div>';
                container.innerHTML = html;
                return;
            }

            const startIndex = (currentPage - 1) * pageSize;
            const pageItems = filteredItems.slice(startIndex, startIndex + pageSize);

            if(currentViewMode === 'cards') {
                if(openCardIds.size === 0 && pageItems.length > 0) {
                    const firstBodyId = getCardBodyId('wo', pageItems[0].no_wo || 0);
                    openCardIds.add(firstBodyId);
                }
                html += '<div class="card-list">';
                pageItems.forEach((item, idx) => {
                    const st = String(item.status || 'SCHED-OK').replace(/\s+/g, '_');
                    const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                    const bodyId = getCardBodyId('act', item.equipment_id || idx);
                    const isOpen = openCardIds.has(bodyId);

                    html += `
                    <div class="item-card">
                        <div class="item-header" onclick="toggleAccordion('${bodyId}')">
                            <div class="item-title-box">
                                <div class="item-code">${item.equipment_id} &bull; ${item.area} ${item.kks ? '&bull; KKS: ' + item.kks : ''} ${item.finish_date ? '&bull; 🏁 Selesai: ' + item.finish_date : ''}</div>
                                <div class="item-name">${item.equipment_description}</div>
                            </div>
                            <div class="header-actions">
                                <button class="btn-finding ${hasFindings?'active':''}" onclick="event.stopPropagation(); openFindingModal('actuator', '${item.equipment_id}', '${item.equipment_id} - ${item.equipment_description.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}')">
                                    📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Temuan' : '+ Temuan')}
                                </button>
                                <span style="font-size:0.82rem; color:var(--text-muted); font-weight:600;">👤 ${item.pic || '-'}</span>
                                <span class="status-badge badge-${st}">${item.status}</span>
                                <div class="progress-box">
                                    <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${item.persen_progress}%;"></div></div>
                                    <span style="font-size:0.85rem; font-weight:800; font-family:'JetBrains Mono';">${item.persen_progress}%</span>
                                </div>
                            </div>
                        </div>
                        <div class="item-body ${isOpen ? 'open' : ''}" id="${bodyId}" style="${isOpen ? 'display:block;' : ''}">
                            <div class="checklist-section">
                                <div class="section-h4">⚙️ Sub-Task Actuator Valve</div>
                                <div class="checklist-grid">
                                    <div class="checklist-item ${item.general_inspection ? 'done' : ''}">
                                        <label>
                                            <input type="checkbox" id="gen-${item.equipment_id}" ${item.general_inspection ? 'checked' : ''} onchange="quickToggleActuator('${item.equipment_id}', 'general_inspection', this.checked)">
                                            <span>General Inspection, Cleaning & Calibration (50%)</span>
                                        </label>
                                    </div>
                                    <div class="checklist-item ${item.function_test ? 'done' : ''}">
                                        <label>
                                            <input type="checkbox" id="func-${item.equipment_id}" ${item.function_test ? 'checked' : ''} onchange="quickToggleActuator('${item.equipment_id}', 'function_test', this.checked)">
                                            <span>Function Test & Stroke Check (50%)</span>
                                        </label>
                                    </div>
                                </div>
                            </div>

                            <div class="form-grid">
                                <div class="form-group">
                                    <label>PIC Penanggung Jawab</label>
                                    <select id="pic-act-${item.equipment_id}" class="filter-input">
                                        ${fullData.pics.map(p => `<option value="${p}" ${item.pic===p?'selected':''}>${p}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Tanggal Selesai</label>
                                    <input type="date" id="finish-act-${item.equipment_id}" class="filter-input" value="${formatDateForInput(item.finish_date)}">
                                </div>
                                <div class="form-group">
                                    <label>Remarks / Catatan</label>
                                    <input type="text" id="rem-act-${item.equipment_id}" class="filter-input" value="${item.remarks || ''}" placeholder="Catatan...">
                                </div>
                            </div>

                            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:15px; border-top:1px solid var(--border-color); padding-top:12px;">
                                <div>
                                    <button class="btn-danger" onclick="deleteActuator('${item.equipment_id}')">🗑️ Hapus Actuator</button>
                                </div>
                                <div style="display:flex; gap:10px;">
                                    <button class="btn-save" onclick="saveActuator('${item.equipment_id}', '${item.equipment_description}')">💾 Simpan Actuator</button>
                                </div>
                            </div>
                        </div>
                    </div>`;
                });
                html += '</div>';
            } else {
                html += `
                <div class="table-wrap">
                    <table class="dense-table">
                        <thead>
                            <tr>
                                <th>Equipment ID</th>
                                <th>Deskripsi Actuator</th>
                                <th>KKS</th>
                                <th>Area</th>
                                <th>PIC</th>
                                <th>General Insp</th>
                                <th>Function Test</th>
                                <th>Status</th>
                                <th>Temuan & Foto</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${pageItems.map(item => {
                                const st = String(item.status || 'SCHED-OK').replace(/\s+/g, '_');
                                const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                                return `
                                <tr>
                                    <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.equipment_id}</td>
                                    <td style="font-weight:600;">${item.equipment_description}</td>
                                    <td><span style="font-size:0.78rem; font-family:'JetBrains Mono'; color:var(--text-muted);">${item.kks || '-'}</span></td>
                                    <td><span style="font-size:0.8rem; color:var(--text-muted);">${item.area}</span></td>
                                    <td style="font-weight:600;">${item.pic || '-'}</td>
                                    <td>
                                        <input type="checkbox" ${item.general_inspection?'checked':''} onchange="quickToggleActuator('${item.equipment_id}', 'general_inspection', this.checked)" style="width:18px; height:18px; accent-color:var(--primary); cursor:pointer;">
                                    </td>
                                    <td>
                                        <input type="checkbox" ${item.function_test?'checked':''} onchange="quickToggleActuator('${item.equipment_id}', 'function_test', this.checked)" style="width:18px; height:18px; accent-color:var(--primary); cursor:pointer;">
                                    </td>
                                    <td><span class="status-badge badge-${st}">${item.status} (${item.persen_progress}%)</span></td>
                                    <td>
                                        <button class="btn-finding ${hasFindings?'active':''}" onclick="openFindingModal('actuator', '${item.equipment_id}', '${item.equipment_id} - ${item.equipment_description.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}')">
                                            📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Ada Temuan' : '+ Foto')}
                                        </button>
                                    </td>
                                </tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                </div>`;
            }

            html += renderPaginationControls(filteredItems.length);
            container.innerHTML = html;
        }

        /* ---------------- INSTRUMENTS RENDERING ---------------- */
        function switchInstSubtab(sub) {
            instSubtab = sub;
            currentPage = 1;
            renderTabContent();
        }

        function renderInstruments(container) {
            const searchStr = document.getElementById('search-input').value.toLowerCase();
            const statusFilter = document.getElementById('filter-status').value;
            const areaFilter = document.getElementById('filter-area').value;

            let items = [];
            if(instSubtab === 'ptx') items = fullData.pressure_tx || [];
            else if(instSubtab === 'ttx') items = fullData.temperature_tx || [];
            else if(instSubtab === 'psw') items = fullData.pressure_switch || [];

            let filteredItems = items.filter(itm => {
                const title = (itm.equipment || itm.description || '').toLowerCase();
                const kks = (itm.kks || '').toLowerCase();
                const area = (itm.area || '').toLowerCase();
                const matchSearch = !searchStr || title.includes(searchStr) || kks.includes(searchStr) || area.includes(searchStr);
                const matchArea = !areaFilter || itm.area === areaFilter;
                
                let matchStatus = true;
                if(statusFilter === 'FINISH') matchStatus = itm.verifikasi === true;
                else if(statusFilter === 'SCHED-OK' || statusFilter === 'IN PROGRESS') matchStatus = itm.verifikasi === false;

                let matchQuick = true;
                if(quickFilter === 'findings') matchQuick = !!itm.temuan || (itm.jumlah_foto > 0);
                else if(quickFilter === 'inprog') matchQuick = !itm.verifikasi;
                else if(quickFilter === 'finish') matchQuick = !!itm.verifikasi;

                return matchSearch && matchArea && matchStatus && matchQuick;
            });

            const isAddInstOpen = openCardIds.has('add-inst-form');
            let html = `
            <div style="margin-bottom:18px; background:var(--bg-card); border-radius:var(--radius-md); padding:16px; border:1px solid var(--border-color);">
                <div style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;" onclick="toggleAccordion('add-inst-form')">
                    <h3 style="font-size:0.95rem; color:var(--primary); font-weight:700;">➕ Tambah Instrument Baru</h3>
                    <span id="arrow-add-inst-form" style="font-weight:700; color:var(--text-muted); font-size:0.85rem;">${isAddInstOpen ? '▲ Tutup Form' : '▼ Buka Form'}</span>
                </div>
                <div id="add-inst-form" class="accordion-form ${isAddInstOpen ? 'open' : ''}" style="${isAddInstOpen ? 'display:block;' : ''}">
                    <div class="form-grid">
                        <div class="form-group">
                            <label>Tipe Instrument <span style="color:#f43f5e;">*</span></label>
                            <select id="new-inst-type" class="filter-input">
                                <option value="pressure_tx" ${instSubtab==='ptx'?'selected':''}>Pressure Transmitter (PTX)</option>
                                <option value="temperature_tx" ${instSubtab==='ttx'?'selected':''}>Temperature Transmitter (TTX)</option>
                                <option value="pressure_switch" ${instSubtab==='psw'?'selected':''}>Pressure Switch (PSW)</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Nama Equipment / Description <span style="color:#f43f5e;">*</span></label>
                            <input type="text" id="new-inst-desc" class="filter-input" placeholder="INLET ID FAN 1...">
                        </div>
                        <div class="form-group">
                            <label>Tag KKS</label>
                            <input type="text" id="new-inst-kks" class="filter-input" placeholder="10HNA61CP001">
                        </div>
                        <div class="form-group">
                            <label>Area System</label>
                            <input type="text" id="new-inst-area" class="filter-input" placeholder="ESP #2, Boiler#2...">
                        </div>
                        <div class="form-group">
                            <label>Range / Set Point</label>
                            <input type="text" id="new-inst-range" class="filter-input" placeholder="-70 - 70 mbar / 8 Bar">
                        </div>
                    </div>
                    <div style="display:flex; gap:10px; margin-top:14px;">
                        <button class="btn-save" onclick="saveNewInstrument()">💾 Simpan Instrument Baru</button>
                        <button class="page-btn" onclick="toggleAccordion('add-inst-form')">Batal</button>
                    </div>
                </div>
            </div>

            <!-- Instrument Sub-tabs -->
            <div class="nav-tabs" style="margin-bottom:16px;">
                <button class="tab-btn ${instSubtab==='ptx'?'active':''}" onclick="switchInstSubtab('ptx')">Pressure Transmitter (${(fullData.pressure_tx||[]).length})</button>
                <button class="tab-btn ${instSubtab==='ttx'?'active':''}" onclick="switchInstSubtab('ttx')">Temperature Transmitter (${(fullData.temperature_tx||[]).length})</button>
                <button class="tab-btn ${instSubtab==='psw'?'active':''}" onclick="switchInstSubtab('psw')">Pressure Switch (${(fullData.pressure_switch||[]).length})</button>
            </div>`;

            html += renderPaginationControls(filteredItems.length);

            if(filteredItems.length === 0) {
                html += '<div style="text-align:center; padding:50px; color:var(--text-muted); background:var(--bg-card); border-radius:12px;">Tidak ada instrumen yang sesuai filter.</div>';
                container.innerHTML = html;
                return;
            }

            const startIndex = (currentPage - 1) * pageSize;
            const pageItems = filteredItems.slice(startIndex, startIndex + pageSize);

            if(currentViewMode === 'cards') {
                if(openCardIds.size === 0 && pageItems.length > 0) {
                    const firstBodyId = getCardBodyId('wo', pageItems[0].no_wo || 0);
                    openCardIds.add(firstBodyId);
                }
                html += '<div class="card-list">';
                pageItems.forEach((item, idx) => {
                    const title = (instSubtab==='psw' ? item.description : item.equipment) || `Item #${item.no}`;
                    const isVerif = !!item.verifikasi;
                    const isCalib = !!item.kalibrasi;
                    const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                    const cardPrefix = instSubtab === 'psw' ? 'psw' : 'inst';
                    const bodyId = getCardBodyId(cardPrefix, item.kks || item.no || idx);
                    const isOpen = openCardIds.has(bodyId);

                    html += `
                    <div class="item-card">
                        <div class="item-header" onclick="toggleAccordion('${bodyId}')">
                            <div style="display:flex; align-items:center; gap:12px; flex-grow:1;">
                                <div class="item-title-box">
                                    <div class="item-code">${item.kks || 'No Tag'} &bull; ${item.area} ${item.tanggal ? '&bull; 🏁 Selesai: ' + item.tanggal : ''}</div>
                                    <div class="item-name">${title} ${instSubtab!=='psw' && item.range ? ' (Range: ' + item.range + ')' : ''}</div>
                                </div>
                            </div>
                            <div class="header-actions">
                                <button class="btn-finding ${hasFindings?'active':''}" onclick="event.stopPropagation(); openFindingModal('instrument', '${item.kks || item.no}', '${item.kks} - ${title.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}', '${instSubtab==='ptx'?'pressure_tx':'temperature_tx'}')">
                                    📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Temuan' : '+ Temuan')}
                                </button>
                                <span class="status-badge ${isVerif ? 'badge-FINISH' : (isCalib ? 'badge-IN-PROGRESS' : 'badge-SCHED-OK')}" id="badge-${cardPrefix}-${idx}">${isVerif ? 'DONE (100%)' : (isCalib ? 'IN PROGRESS (Kalibrasi OK)' : 'SCHEDULED')}</span>
                            </div>
                        </div>
                        <div class="item-body ${isOpen ? 'open' : ''}" id="${bodyId}" style="${isOpen ? 'display:block;' : ''}">
                            <div class="checklist-section" style="margin-bottom:14px;">
                                <div class="section-h4">📋 Checklist Progress Instrumen (Penentu Finish: Verifikasi)</div>
                                <div class="checklist-grid">
                                    <div class="checklist-item ${isCalib ? 'done' : ''}" id="card-calib-${cardPrefix}-${idx}">
                                        <label style="cursor:pointer; display:flex; align-items:center; gap:8px;">
                                             <input type="checkbox" id="inst-calib-${idx}" ${isCalib ? 'checked' : ''} onchange="toggleLocalInstCheck('${instSubtab}', ${idx}, 'kalibrasi', this.checked)">
                                             <span style="font-weight:700;">🛠️ 1. Kalibrasi Selesai</span>
                                        </label>
                                    </div>
                                    <div class="checklist-item ${isVerif ? 'done' : ''}" id="card-verif-${cardPrefix}-${idx}" style="${isVerif ? 'border-color:var(--status-finish);' : ''}">
                                        <label style="cursor:pointer; display:flex; align-items:center; gap:8px;">
                                             <input type="checkbox" id="inst-verif-${idx}" ${isVerif ? 'checked' : ''} onchange="toggleLocalInstCheck('${instSubtab}', ${idx}, 'verifikasi', this.checked)">
                                             <span style="font-weight:800; color:${isVerif ? 'var(--status-finish)' : 'var(--primary)'};">🔍 2. Verifikasi Selesai (Penentu Finish)</span>
                                        </label>
                                    </div>
                                </div>
                            </div>
                            ${instSubtab === 'psw' ? `
                            <div class="calib-grid">
                                <div class="calib-col">
                                    <h5>📥 AS FOUND (Kondisi Awal)</h5>
                                    <div class="calib-fields">
                                        <div class="form-group">
                                            <label>Set Point</label>
                                            <input type="text" id="af-set-${idx}" class="filter-input" value="${item.asfound_set || ''}" placeholder="mis. 7.8 Bar">
                                        </div>
                                        <div class="form-group">
                                            <label>Reset Point</label>
                                            <input type="text" id="af-reset-${idx}" class="filter-input" value="${item.asfound_reset || ''}" placeholder="mis. 7.2 Bar">
                                        </div>
                                    </div>
                                </div>
                                <div class="calib-col">
                                    <h5>📤 AS LEFT (Setelah Kalibrasi)</h5>
                                    <div class="calib-fields">
                                        <div class="form-group">
                                            <label>Set Point</label>
                                            <input type="text" id="al-set-${idx}" class="filter-input" value="${item.asleft_set || ''}" placeholder="mis. 8.0 Bar">
                                        </div>
                                        <div class="form-group">
                                            <label>Reset Point</label>
                                            <input type="text" id="al-reset-${idx}" class="filter-input" value="${item.asleft_reset || ''}" placeholder="mis. 7.5 Bar">
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Hasil Kalibrasi</label>
                                    <select id="res-psw-${idx}" class="filter-input">
                                        <option value="OK" ${item.status_ok_notok==='OK'?'selected':''}>OK / Sesuai Standar</option>
                                        <option value="NOT OK" ${item.status_ok_notok==='NOT OK'?'selected':''}>NOT OK / Deviasi</option>
                                    </select>
                                </div>` : `<div class="form-grid">`}
                                <div class="form-group" style="grid-column: span 2;">
                                    <label>Catatan / Remarks</label>
                                    <input type="text" id="inst-rem-${idx}" class="filter-input" value="${item.remarks || ''}" placeholder="Catatan kalibrasi / verifikasi...">
                                </div>
                            </div>

                            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:15px; border-top:1px solid var(--border-color); padding-top:12px;">
                                <div>
                                    <button class="btn-danger" onclick="deleteInstrument('${instSubtab}', '${item.kks || item.no}')">🗑️ Hapus Instrument</button>
                                </div>
                                <div style="display:flex; gap:10px;">
                                    <button class="btn-save" onclick="${instSubtab==='psw' ? 'savePressureSwitch' : 'saveTransmitter'}('${instSubtab}', '${item.kks || item.no}', ${idx})">💾 Simpan Kalibrasi & Verifikasi</button>
                                </div>
                            </div>
                        </div>
                    </div>`;
                });
                html += '</div>';
            } else {
                html += `
                <div class="table-wrap">
                    <table class="dense-table">
                        <thead>
                            <tr>
                                <th>Tag KKS</th>
                                <th>Equipment Description</th>
                                <th>Area</th>
                                <th>Range / Satuan</th>
                                <th style="text-align:center;">🛠️ Kalibrasi</th>
                                <th style="text-align:center;">🔍 Verifikasi</th>
                                <th>Status Finish</th>
                                <th>Temuan & Foto</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${pageItems.map((item, idx) => {
                                const title = (instSubtab==='psw' ? item.description : item.equipment) || `Item #${item.no}`;
                                const isVerif = !!item.verifikasi;
                                const isCalib = !!item.kalibrasi;
                                const hasFindings = !!item.temuan || (item.jumlah_foto > 0);
                                return `
                                <tr>
                                    <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.kks || '-'}</td>
                                    <td style="font-weight:600;">${title}</td>
                                    <td><span style="font-size:0.8rem; color:var(--text-muted);">${item.area}</span></td>
                                    <td><span style="font-size:0.82rem; font-family:'JetBrains Mono';">${item.range || '-'}</span></td>
                                    <td style="text-align:center;">
                                        <input type="checkbox" ${isCalib?'checked':''} onchange="toggleDirectInstCheck('${instSubtab}', '${item.kks || item.no}', 'kalibrasi', this.checked)" style="width:18px; height:18px; accent-color:var(--primary); cursor:pointer;">
                                    </td>
                                    <td style="text-align:center;">
                                        <input type="checkbox" ${isVerif?'checked':''} onchange="toggleDirectInstCheck('${instSubtab}', '${item.kks || item.no}', 'verifikasi', this.checked)" style="width:18px; height:18px; accent-color:var(--status-finish); cursor:pointer;">
                                    </td>
                                    <td>
                                        <span class="status-badge ${isVerif?'badge-FINISH':(isCalib?'badge-IN-PROGRESS':'badge-SCHED-OK')}">${isVerif?'DONE (100%)':(isCalib?'IN PROGRESS':'SCHEDULED')}</span>
                                    </td>
                                    <td>
                                        <button class="btn-finding ${hasFindings?'active':''}" onclick="openFindingModal('instrument', '${item.kks || item.no}', '${item.kks} - ${title.replace(/'/g, "\\'")}', '${item.area}', '${(item.temuan||'').replace(/'/g, "\\'")}', '${(item.tindak_lanjut||'').replace(/'/g, "\\'")}', '${instSubtab==='ptx'?'pressure_tx':'temperature_tx'}')">
                                            📷 ${item.jumlah_foto > 0 ? item.jumlah_foto + ' Foto' : (hasFindings ? 'Ada Temuan' : '+ Foto')}
                                        </button>
                                    </td>
                                </tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                </div>`;
            }

            html += renderPaginationControls(filteredItems.length);
            container.innerHTML = html;
        }

        /* ---------------- SCOPE & PIC MASTER RENDERING ---------------- */
        function renderScopeMaster(container) {
            let html = `
            <div style="display:flex; flex-direction:column; gap:20px;">
                <!-- Master PIC Card -->
                <div style="background:var(--bg-card); border-radius:var(--radius-lg); padding:22px; border:1px solid var(--border-color);">
                    <h3 style="margin-bottom:8px; color:var(--primary); font-size:1.1rem; font-weight:800;">👥 Master PIC Tim EIC</h3>
                    <p style="font-size:0.82rem; color:var(--text-muted); margin-bottom:16px;">Daftar nama penanggung jawab EIC. Semua pilihan dropdown PIC di Work Order, Actuator, & Instrumen tersinkron otomatis dari daftar master ini.</p>
                    
                    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:18px;">
                        ${(fullData.pics || []).map(p => `
                            <span style="padding:6px 14px; background:var(--bg-sub); border:1px solid var(--border-color); border-radius:20px; font-size:0.85rem; font-weight:600; color:var(--text-main); display:inline-flex; align-items:center; gap:8px; box-shadow:0 2px 5px rgba(0,0,0,0.05);">
                                👤 ${p}
                                <button style="background:none; border:none; color:#f43f5e; cursor:pointer; font-weight:700; font-size:0.85rem;" title="Hapus PIC" onclick="deletePic('${(p||'').toString().replace(/'/g, "\\'")}')">✖</button>
                            </span>
                        `).join('')}
                    </div>

                    <div style="display:flex; gap:10px; max-width:480px;">
                        <input type="text" id="new-pic-input" class="filter-input" placeholder="Masukkan nama personil / vendor PIC baru..." style="flex-grow:1;">
                        <button class="btn-save" onclick="addNewPic()">➕ Tambah PIC</button>
                    </div>
                </div>

                <!-- Scope Master Table -->
                <div style="background:var(--bg-card); border-radius:var(--radius-lg); padding:22px; border:1px solid var(--border-color);">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
                        <h3 style="color:var(--primary); font-size:1.1rem; font-weight:800;">📋 Master Scope Pekerjaan Outage (Vendor & MSW Scope)</h3>
                        <button class="page-btn" style="font-size:0.82rem;" onclick="toggleAccordion('add-scope-form')">${openCardIds.has('add-scope-form') ? '▲ Tutup Scope' : '➕ Tambah Scope'}</button>
                    </div>
                    
                    <div id="add-scope-form" class="accordion-form ${openCardIds.has('add-scope-form') ? 'open' : ''}" style="${openCardIds.has('add-scope-form') ? 'display:block;' : ''}; margin-bottom:18px;">
                        <div class="form-grid">
                            <div class="form-group">
                                <label>Kategori Scope</label>
                                <input type="text" id="new-scope-cat" class="filter-input" placeholder="BOILER, TURBINE...">
                            </div>
                            <div class="form-group">
                                <label>Equipment / Scope Pekerjaan <span style="color:#f43f5e;">*</span></label>
                                <input type="text" id="new-scope-eq" class="filter-input" placeholder="Inspeksi Burner...">
                            </div>
                            <div class="form-group">
                                <label>Tipe Scope</label>
                                <select id="new-scope-type" class="filter-input">
                                    <option value="MSW">MSW</option>
                                    <option value="Vendor">Vendor</option>
                                    <option value="Internal">Internal</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>PIC Penanggung Jawab</label>
                                <select id="new-scope-pic" class="filter-input">
                                    <option value="">Pilih PIC...</option>
                                    ${(fullData.pics || []).map(p => `<option value="${p}">${p}</option>`).join('')}
                                </select>
                            </div>
                        </div>
                        <div style="display:flex; gap:10px; margin-top:12px;">
                            <button class="btn-save" onclick="saveNewScope()">💾 Simpan Scope Baru</button>
                            <button class="page-btn" onclick="toggleAccordion('add-scope-form')">Batal</button>
                        </div>
                    </div>

                    <div class="table-wrap">
                        <table class="dense-table">
                            <thead>
                                <tr>
                                    <th>Kategori Scope</th>
                                    <th>Equipment / Scope Pekerjaan</th>
                                    <th>Tipe Scope</th>
                                    <th>PIC Penanggung Jawab</th>
                                    <th>Aksi</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${(fullData.scope_master || []).map((s, sIdx) => `
                                    <tr>
                                        <td style="font-size:0.8rem; color:var(--text-muted);">${s.kategori || '-'}</td>
                                        <td>
                                            <input type="text" id="scope-eq-${sIdx}" class="filter-input" value="${(s.nama_equipment || '').replace(/"/g, '&quot;')}" style="padding:5px 8px; font-size:0.85rem; width:100%; min-width:200px;" onblur="saveScopeRow(${sIdx})">
                                        </td>
                                        <td>
                                            <select id="scope-type-${sIdx}" class="filter-input" style="padding:5px 8px; font-size:0.85rem;" onchange="saveScopeRow(${sIdx})">
                                                <option value="Vendor" ${s.tipe_scope==='Vendor'?'selected':''}>Vendor</option>
                                                <option value="MSW" ${s.tipe_scope==='MSW'?'selected':''}>MSW</option>
                                                <option value="Internal" ${s.tipe_scope==='Internal'?'selected':''}>Internal</option>
                                            </select>
                                        </td>
                                        <td>
                                            <select id="scope-pic-${sIdx}" class="filter-input" style="padding:5px 8px; font-size:0.85rem;" onchange="saveScopeRow(${sIdx})">
                                                <option value="">Pilih PIC...</option>
                                                ${(fullData.pics || []).map(p => {
                                                    const isSel = (s.pic === p || (s.pic||'').toUpperCase() === (p||'').toUpperCase());
                                                    return `<option value="${p}" ${isSel ? 'selected' : ''}>${p}</option>`;
                                                }).join('')}
                                            </select>
                                        </td>
                                        <td>
                                            <div style="display:flex; gap:6px;">
                                                <button class="btn-save" style="padding:5px 12px; font-size:0.78rem;" onclick="saveScopeRow(${sIdx})">💾 Simpan</button>
                                                <button class="btn-danger" style="padding:5px 10px; font-size:0.78rem;" onclick="deleteScopeRow(${sIdx})">🗑️ Hapus</button>
                                            </div>
                                        </td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>`;
            container.innerHTML = html;
        }

        function renderTabContent() {
            const container = document.getElementById('tab-content');
            if(!fullData) return;

            if (currentTab === 'wo') renderWorkOrders(container);
            else if (currentTab === 'actuator') renderActuators(container);
            else if (currentTab === 'instrument') renderInstruments(container);
            else if (currentTab === 'scope') renderScopeMaster(container);
        }

        const openCardIds = new Set();
        function getCardBodyId(prefix, key) {
            return 'body-' + prefix + '-' + String(key).replace(/[^a-zA-Z0-9_-]/g, '_');
        }

        function toggleAccordion(id) {
            const el = document.getElementById(id);
            if(!el) return;
            const wasOpen = openCardIds.has(id) || el.classList.contains('open') || (el.style.display === 'block');
            if (wasOpen) {
                openCardIds.delete(id);
                el.classList.remove('open');
                el.style.display = 'none';
                const arrow = document.getElementById(`arrow-${id}`);
                if(arrow) arrow.innerText = id === 'add-scope-form' ? '➕ Tambah Scope' : '▼ Buka Form';
            } else {
                openCardIds.add(id);
                el.classList.add('open');
                el.style.display = 'block';
                const arrow = document.getElementById(`arrow-${id}`);
                if(arrow) arrow.innerText = id === 'add-scope-form' ? '▲ Tutup Scope' : '▲ Tutup Form';
            }
        }

        /* ---------------- FINDING & PHOTO MODAL ---------------- */
        async function openFindingModal(itemType, id, title, area, temuan, tindakLanjut, instType = '') {
            activeFinding = { itemType, id, title, area, temuan, tindakLanjut, instType };
            document.getElementById('modal-finding-title').innerText = `📷 Bukti & Temuan: ${id}`;
            document.getElementById('modal-finding-subtitle').innerText = `${title} (${area || 'GENERAL'})`;
            document.getElementById('modal-finding-text').value = temuan || '';
            document.getElementById('modal-tl-text').value = tindakLanjut || '';

            try {
                const res = await fetch(`/api/findings?id=${encodeURIComponent(id)}`);
                const data = await res.json();
                renderModalPhotos(data.photos || []);
                if(data.temuan) document.getElementById('modal-finding-text').value = data.temuan;
                if(data.tindak_lanjut) document.getElementById('modal-tl-text').value = data.tindak_lanjut;
            } catch(e) {
                renderModalPhotos([]);
            }

            document.getElementById('finding-modal').classList.add('open');
        }

        function closeFindingModal() {
            document.getElementById('finding-modal').classList.remove('open');
            activeFinding = null;
        }

        function renderModalPhotos(photos) {
            document.getElementById('modal-photo-count').innerText = photos.length;
            const grid = document.getElementById('modal-photo-grid');
            if(photos.length === 0) {
                grid.innerHTML = '<div style="font-size:0.8rem; color:var(--text-muted); grid-column:span 4; padding:8px 0;">Belum ada foto yang diunggah.</div>';
                return;
            }
            grid.innerHTML = photos.map(p => `
                <div class="photo-thumb-box">
                    <img src="${p.url}" alt="${p.filename}" onclick="openLightbox('${p.url}')" title="Klik untuk perbesar">
                    <button class="photo-delete-btn" onclick="deleteModalPhoto('${p.filename}')" title="Hapus foto ini">🗑️</button>
                </div>
            `).join('');
        }

        function openLightbox(url) {
            document.getElementById('lightbox-img').src = url;
            document.getElementById('lightbox-modal').classList.add('open');
        }

        function closeLightbox() {
            document.getElementById('lightbox-modal').classList.remove('open');
        }

        async function handleModalFileSelect(files) {
            if(!activeFinding || !files || files.length === 0) return;

            for(let file of files) {
                const reader = new FileReader();
                reader.onload = async function(e) {
                    const base64Data = e.target.result;
                    try {
                        const payload = {
                            id: activeFinding.id,
                            unit: currentUnit,
                            type: activeFinding.itemType,
                            inst_type: activeFinding.instType,
                            filename: file.name,
                            image_base64: base64Data,
                            temuan: document.getElementById('modal-finding-text').value,
                            tindak_lanjut: document.getElementById('modal-tl-text').value
                        };
                        const res = await fetch('/api/upload_finding_photo', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify(payload)
                        });
                        const result = await res.json();
                        if(result.status === 'success') {
                            showToast(`Foto '${file.name}' berhasil disimpan!`, 'success');
                            renderModalPhotos(result.photos || []);
                            loadData();
                        } else {
                            showToast(result.message || 'Gagal menyimpan foto', 'error');
                        }
                    } catch(err) {
                        showToast('Error uploading photo: ' + err.message, 'error');
                    }
                };
                reader.readAsDataURL(file);
            }
        }

        async function deleteModalPhoto(filename) {
            if(!activeFinding || !filename) return;
            if(!confirm(`Hapus foto ${filename}?`)) return;

            try {
                const payload = {
                    id: activeFinding.id,
                    unit: currentUnit,
                    filename: filename,
                    type: activeFinding.itemType,
                    inst_type: activeFinding.instType
                };
                const res = await fetch('/api/delete_finding_photo', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast('Foto berhasil dihapus', 'info');
                    renderModalPhotos(result.photos || []);
                    loadData();
                }
            } catch(e) {
                showToast('Gagal menghapus foto', 'error');
            }
        }

        async function saveFindingModalData() {
            if(!activeFinding) return;
            const temuan = document.getElementById('modal-finding-text').value;
            const tindakLanjut = document.getElementById('modal-tl-text').value;

            try {
                const payload = {
                    id: activeFinding.id,
                    unit: currentUnit,
                    type: activeFinding.itemType,
                    inst_type: activeFinding.instType,
                    temuan: temuan,
                    tindak_lanjut: tindakLanjut
                };
                const res = await fetch('/api/upload_finding_photo', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast('Temuan dan rekomendasi tindak lanjut berhasil disimpan!', 'success');
                closeFindingModal();
                loadData();
            } catch(e) {
                showToast('Gagal menyimpan temuan', 'error');
            }
        }

        /* ---------------- DEDICATED REPORT & PRINT ---------------- */
        function openReportModal() {
            const today = new Date();
            const yesterday = new Date(Date.now() - 86400000);
            
            const startInput = document.getElementById('report-start-date');
            const endInput = document.getElementById('report-end-date');
            if(startInput && !startInput.value) startInput.value = formatDateForInput(yesterday);
            if(endInput && !endInput.value) endInput.value = formatDateForInput(today);
            
            generateReportContent();
            document.getElementById('report-modal').classList.add('open');
        }

        function closeReportModal() {
            document.getElementById('report-modal').classList.remove('open');
        }

        function printReportModal() {
            window.print();
        }

        function parseDateStrToTime(str) {
            if(!str) return null;
            str = String(str).trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
                const p = str.split('-');
                return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2])).getTime();
            }
            if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
                const parts = str.split('/');
                return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0])).getTime();
            }
            const d = new Date(str);
            return isNaN(d.getTime()) ? null : d.getTime();
        }

        function isDateInRange(dateStr, startYMD, endYMD) {
            if(!startYMD && !endYMD) return true;
            if(!dateStr) return false;
            const itemTime = parseDateStrToTime(dateStr);
            if(!itemTime) return false;
            
            let startTime = startYMD ? parseDateStrToTime(startYMD) : null;
            let endTime = endYMD ? parseDateStrToTime(endYMD) : null;
            
            if(startTime && itemTime < startTime) return false;
            if(endTime && itemTime > endTime + (24 * 60 * 60 * 1000 - 1)) return false;
            return true;
        }

        let currentReportType = 'harian';

        function setReportType(type) {
            currentReportType = type;
            ['harian', 'wo_detail', 'actuator', 'instruments'].forEach(t => {
                const btn = document.getElementById(`reptab-${t}`);
                if(btn) btn.classList.toggle('active', t === type);
            });
            const dateBar = document.getElementById('report-date-bar');
            if(dateBar) {
                dateBar.style.display = (type === 'harian') ? 'flex' : 'none';
            }
            generateReportContent();
        }

        function generateReportContent() {
            const container = document.getElementById('report-printable-content');
            if(!container || !fullData || !fullData.summary) return;
            
            const s = fullData.summary;
            const printDateStr = new Date().toLocaleString('id-ID');
            
            if(currentReportType === 'harian') {
                renderHarianReport(container, s, printDateStr);
            } else if(currentReportType === 'wo_detail') {
                renderWODetailReport(container, s, printDateStr);
            } else if(currentReportType === 'actuator') {
                renderActuatorReport(container, s, printDateStr);
            } else if(currentReportType === 'instruments') {
                renderInstrumentsReport(container, s, printDateStr);
            }
        }

        /* OPSI 1: Laporan Progress Harian & Temuan (WO, Actuator, Instruments) */
        function renderHarianReport(container, s, printDateStr) {
            const startYMD = document.getElementById('report-start-date').value;
            const endYMD = document.getElementById('report-end-date').value;
            const startDisp = startYMD ? formatDateForStorage(startYMD) : 'Awal';
            const endDisp = endYMD ? formatDateForStorage(endYMD) : 'Hari Ini';

            const completedTasks = [];
            
            // 1. Work Orders Updates
            (fullData.work_orders || []).forEach(w => {
                let subtasksAdded = 0;
                (w.checklist || []).forEach(c => {
                    if(c.selesai && isDateInRange(c.tanggal, startYMD, endYMD)) {
                        completedTasks.push({
                            type: 'Work Order',
                            code: w.no_wo,
                            item_name: w.job_description,
                            subtask: `Sub-task: ${c.sub_task}`,
                            area: w.area || 'GENERAL',
                            pic: c.pic_task || w.pic || '-',
                            date: c.tanggal || w.tanggal_finish || '-'
                        });
                        subtasksAdded++;
                    }
                });
                // If WO has no subtask list or entire WO marked finish in range
                if(subtasksAdded === 0 && (w.status === 'FINISH' || (w.persen_progress && w.persen_progress > 0)) && isDateInRange(w.tanggal_finish, startYMD, endYMD)) {
                    completedTasks.push({
                        type: 'Work Order',
                        code: w.no_wo,
                        item_name: w.job_description,
                        subtask: w.status === 'FINISH' ? 'Pekerjaan WO Selesai (100%)' : `Progress WO (${w.persen_progress}%)`,
                        area: w.area || 'GENERAL',
                        pic: w.pic || '-',
                        date: w.tanggal_finish || '-'
                    });
                }
            });

            // 2. Actuator Valves Updates
            (fullData.actuators || []).forEach(a => {
                const isGen = !!a.general_inspection;
                const isFunc = !!a.function_test;
                const isDone = isGen && isFunc;
                if((isGen || isFunc) && isDateInRange(a.finish_date, startYMD, endYMD)) {
                    let desc = isDone ? 'General Inspection & Function Test Selesai (100% FINISH)' : (isGen ? 'General Inspection Selesai (50%)' : 'Function Test Selesai (50%)');
                    if(a.remarks) desc += ` [${a.remarks}]`;
                    completedTasks.push({
                        type: 'Actuator Valve',
                        code: a.equipment_id,
                        item_name: a.equipment_description,
                        subtask: desc,
                        area: a.area || 'BOILER',
                        pic: a.pic || '-',
                        date: a.finish_date || '-'
                    });
                }
            });

            // 3. Instruments Updates (Pressure TX, Temperature TX, Pressure Switch)
            const checkInst = (list, typeName) => {
                (list || []).forEach(inst => {
                    const isVerif = !!inst.verifikasi || !!inst.status_wdone;
                    const isCalib = !!inst.kalibrasi;
                    const instDate = inst.tanggal || inst.finish_date || inst.dated;
                    if((isVerif || isCalib) && isDateInRange(instDate, startYMD, endYMD)) {
                        let desc = isVerif ? 'Kalibrasi & Verifikasi Selesai (100% DONE)' : 'Kalibrasi Selesai (In Progress)';
                        if(inst.remarks) desc += ` [${inst.remarks}]`;
                        completedTasks.push({
                            type: typeName,
                            code: inst.kks || `Item #${inst.no}`,
                            item_name: inst.equipment || inst.description || typeName,
                            subtask: desc,
                            area: inst.area || 'GENERAL',
                            pic: inst.pic || 'Tim EIC',
                            date: instDate || '-'
                        });
                    }
                });
            };
            checkInst(fullData.pressure_tx, 'Pressure TX');
            checkInst(fullData.temperature_tx, 'Temperature TX');
            checkInst(fullData.pressure_switch, 'Pressure Switch');

            // 4. Findings Collection
            const findingsList = [];
            const collectFinding = (list, typeName, codeField, descField) => {
                (list || []).forEach(item => {
                    if(item.temuan || (item.jumlah_foto > 0)) {
                        findingsList.push({
                            type: typeName,
                            code: item[codeField] || 'Item',
                            desc: item[descField] || item.equipment || item.job_description || '-',
                            area: item.area || 'GENERAL',
                            temuan: item.temuan || '(Belum ada deskripsi temuan, tercatat foto lampiran)',
                            tindak_lanjut: item.tindak_lanjut || 'Menunggu verifikasi lapangan',
                            foto_count: item.jumlah_foto || 0
                        });
                    }
                });
            };
            collectFinding(fullData.work_orders, 'Work Order', 'no_wo', 'job_description');
            collectFinding(fullData.actuators, 'Actuator Valve', 'equipment_id', 'equipment_description');
            collectFinding(fullData.pressure_tx, 'Pressure TX', 'kks', 'equipment');
            collectFinding(fullData.temperature_tx, 'Temperature TX', 'kks', 'equipment');
            collectFinding(fullData.pressure_switch, 'Pressure Switch', 'kks', 'equipment');

            let html = `
            <div class="report-paper">
                <div class="report-header-box">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:2px solid #334155; padding-bottom:12px; margin-bottom:16px;">
                        <div>
                            <div style="font-size:1.1rem; font-weight:800; color:var(--primary); letter-spacing:0.5px;">⚡ PLTU MSW &bull; SECTION ELECTRIC, INSTRUMENT & CONTROL</div>
                            <h2 style="font-size:1.25rem; margin:4px 0 2px 0; color:#fff;">LAPORAN PROGRESS HARIAN OUTAGE (WO, ACTUATOR & INSTRUMENT)</h2>
                            <div style="font-size:0.85rem; color:var(--text-muted);">Monitoring Progress Harian & Rekapitulasi Temuan Pekerjaan Unit ${currentUnit}</div>
                        </div>
                        <div style="text-align:right; font-size:0.8rem; color:var(--text-muted);">
                            <div><strong>Periode Update:</strong> ${startDisp} s/d ${endDisp}</div>
                            <div><strong>Waktu Cetak:</strong> ${printDateStr}</div>
                            <div style="color:var(--primary); font-weight:700; margin-top:2px;">UNIT ${currentUnit}</div>
                        </div>
                    </div>
                </div>

                <div class="report-section-title">📊 1. RINGKASAN PROGRESS KESELURUHAN (UNIT ${currentUnit})</div>
                <div class="report-kpi-grid" style="grid-template-columns: repeat(3, 1fr);">
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">WORK ORDERS (SUB-TASK)</div>
                        <div class="report-kpi-val">${s.wo.pct}%</div>
                        <div class="report-kpi-sub">${s.wo.subtask_done} / ${s.wo.subtask_total} Sub-task (${s.wo.finish}/${s.wo.total} WO Finish)</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">ACTUATOR VALVES</div>
                        <div class="report-kpi-val">${s.actuator.pct}%</div>
                        <div class="report-kpi-sub">${s.actuator.subtask_done} / ${s.actuator.subtask_total} Test (${s.actuator.finish}/${s.actuator.total} Valve Finish)</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">INSTRUMENTS (TX & PSW)</div>
                        <div class="report-kpi-val">${s.instrument.pct}%</div>
                        <div class="report-kpi-sub">${s.instrument.done} / ${s.instrument.total} Selesai Verifikasi</div>
                    </div>
                </div>

                <div class="report-section-title" style="margin-top:22px;">
                    📋 2. DAFTAR UPDATE PEKERJAAN YANG DISELESAIKAN (WO, ACTUATOR & INSTRUMENT)
                    <span style="font-size:0.8rem; font-weight:600; color:var(--text-muted); float:right;">Total: ${completedTasks.length} Item Update</span>
                </div>
                ${completedTasks.length === 0 ? `
                    <div style="padding:14px; background:var(--bg-sub); border:1px dashed var(--border-color); border-radius:var(--radius-sm); font-size:0.85rem; color:var(--text-muted); text-align:center;">
                        ℹ️ Tidak ada update sub-task, actuator valve, atau instrumen yang tercatat selesai pada rentang tanggal <strong>${startDisp} s/d ${endDisp}</strong>.
                    </div>
                ` : `
                    <table class="report-table">
                        <thead>
                            <tr>
                                <th style="width:35px;">No</th>
                                <th style="width:115px;">Kategori</th>
                                <th style="width:135px;">No WO / Tag KKS</th>
                                <th>Uraian Pekerjaan / Sub-Task / Status Update</th>
                                <th style="width:110px;">Area</th>
                                <th style="width:110px;">PIC</th>
                                <th style="width:90px; text-align:center;">Tgl Update</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${completedTasks.map((t, idx) => `
                                <tr>
                                    <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                    <td><span class="report-badge">${t.type}</span></td>
                                    <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${t.code}</td>
                                    <td>
                                        <div style="font-weight:600;">${t.subtask}</div>
                                        <div style="font-size:0.75rem; color:var(--text-muted);">${t.item_name}</div>
                                    </td>
                                    <td>${t.area}</td>
                                    <td style="font-weight:600;">${t.pic}</td>
                                    <td style="text-align:center; font-family:'JetBrains Mono'; font-size:0.8rem;">${t.date}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                `}

                <div class="report-section-title" style="margin-top:24px;">
                    ⚠️ 3. REKAPITULASI TEMUAN LAPANGAN & TINDAK LANJUT (ACTIVE FINDINGS)
                    <span style="font-size:0.8rem; font-weight:600; color:var(--text-muted); float:right;">Total: ${findingsList.length} Temuan</span>
                </div>
                ${findingsList.length === 0 ? `
                    <div style="padding:14px; background:var(--bg-sub); border:1px dashed var(--border-color); border-radius:var(--radius-sm); font-size:0.85rem; color:#10b981; text-align:center;">
                        ✅ Nihil. Seluruh peralatan dan instrumen dalam kondisi normal tanpa catatan temuan terbuka.
                    </div>
                ` : `
                    <table class="report-table">
                        <thead>
                            <tr>
                                <th style="width:35px;">No</th>
                                <th style="width:110px;">Kategori</th>
                                <th style="width:130px;">No WO / Tag</th>
                                <th style="width:160px;">Equipment</th>
                                <th>Uraian Temuan Masalah</th>
                                <th>Rekomendasi Tindak Lanjut</th>
                                <th style="width:75px; text-align:center;">Foto</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${findingsList.map((f, idx) => `
                                <tr>
                                    <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                    <td><span class="report-badge alert">${f.type}</span></td>
                                    <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--status-alert);">${f.code}</td>
                                    <td style="font-weight:600; font-size:0.82rem;">${f.desc}<div style="font-size:0.75rem; color:var(--text-muted);">${f.area}</div></td>
                                    <td style="color:#fca5a5; font-size:0.82rem;">${f.temuan}</td>
                                    <td style="color:#fef08a; font-size:0.82rem;">${f.tindak_lanjut}</td>
                                    <td style="text-align:center; font-size:0.8rem; font-weight:700;">${f.foto_count > 0 ? f.foto_count + ' 📷' : '-'}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                `}
            </div>`;
            container.innerHTML = html;
        }

        /* OPSI 2: Laporan Work Order Keseluruhan & Detail Sub-Task */
        function renderWODetailReport(container, s, printDateStr) {
            const woList = fullData.work_orders || [];
            let html = `
            <div class="report-paper">
                <div class="report-header-box">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:2px solid #334155; padding-bottom:12px; margin-bottom:16px;">
                        <div>
                            <div style="font-size:1.1rem; font-weight:800; color:var(--primary); letter-spacing:0.5px;">⚡ PLTU MSW &bull; SECTION ELECTRIC, INSTRUMENT & CONTROL</div>
                            <h2 style="font-size:1.25rem; margin:4px 0 2px 0; color:#fff;">LAPORAN STATUS WORK ORDER & DETAIL SUB-TASK LENGKAP</h2>
                            <div style="font-size:0.85rem; color:var(--text-muted);">Rekapitulasi Seluruh Pekerjaan Work Order Unit ${currentUnit} Beserta Rincian Sub-Task</div>
                        </div>
                        <div style="text-align:right; font-size:0.8rem; color:var(--text-muted);">
                            <div><strong>Total WO:</strong> ${woList.length} Pekerjaan</div>
                            <div><strong>Waktu Cetak:</strong> ${printDateStr}</div>
                            <div style="color:var(--primary); font-weight:700; margin-top:2px;">UNIT ${currentUnit}</div>
                        </div>
                    </div>
                </div>

                <div class="report-section-title">📊 1. RINGKASAN PROGRESS WORK ORDERS</div>
                <div class="report-kpi-grid" style="grid-template-columns: repeat(4, 1fr);">
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">TOTAL WORK ORDERS</div>
                        <div class="report-kpi-val">${s.wo.total}</div>
                        <div class="report-kpi-sub">Total Item Pekerjaan</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">WO FINISH (100%)</div>
                        <div class="report-kpi-val" style="color:var(--status-finish);">${s.wo.finish}</div>
                        <div class="report-kpi-sub">Pekerjaan Selesai Penuh</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">WO IN PROGRESS</div>
                        <div class="report-kpi-val" style="color:#38bdf8;">${s.wo.in_progress}</div>
                        <div class="report-kpi-sub">Dalam Proses Pengerjaan</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">SUB-TASK PROGRESS</div>
                        <div class="report-kpi-val">${s.wo.pct}%</div>
                        <div class="report-kpi-sub">${s.wo.subtask_done} / ${s.wo.subtask_total} Sub-Task Selesai</div>
                    </div>
                </div>

                <div class="report-section-title" style="margin-top:22px;">
                    📋 2. RINCIAN SELURUH WORK ORDER & CHECKLIST SUB-TASK (${woList.length} WO)
                </div>

                <div style="display:flex; flex-direction:column; gap:16px;">
                    ${woList.map((w, idx) => {
                        const doneCount = (w.checklist || []).filter(c => c.selesai).length;
                        const totalCount = (w.checklist || []).length;
                        const stBadge = w.status === 'FINISH' ? 'var(--status-finish)' : (w.status === 'IN PROGRESS' ? '#38bdf8' : '#94a3b8');
                        
                        return `
                        <div style="background:var(--bg-sub); border:1px solid var(--border-color); border-radius:var(--radius-sm); padding:12px 14px; page-break-inside:avoid;">
                            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border-color); padding-bottom:8px; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                                <div>
                                    <span style="font-family:'JetBrains Mono'; font-weight:800; color:var(--primary); font-size:0.95rem;">${idx + 1}. [${w.no_wo}]</span>
                                    <strong style="color:var(--text-main); font-size:0.95rem; margin-left:6px;">${w.job_description}</strong>
                                    <span style="font-size:0.8rem; color:var(--text-muted); margin-left:8px;">&bull; Area: <strong>${w.area || 'GENERAL'}</strong></span>
                                </div>
                                <div style="display:flex; align-items:center; gap:10px;">
                                    <span style="font-size:0.8rem; color:var(--text-muted);">PIC: <strong style="color:var(--text-main);">${w.pic || '-'}</strong></span>
                                    <span style="display:inline-block; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:800; background:rgba(56, 189, 248, 0.1); color:${stBadge}; border:1px solid ${stBadge};">${w.status} (${w.persen_progress}%)</span>
                                </div>
                            </div>

                            ${totalCount === 0 ? `
                                <div style="font-size:0.8rem; color:var(--text-muted); padding:4px 0;">(Tidak ada rincian sub-task)</div>
                            ` : `
                                <table class="report-table" style="margin-bottom:4px;">
                                    <thead>
                                        <tr>
                                            <th style="width:30px;">No</th>
                                            <th>Uraian Checklist Sub-Task (${doneCount} / ${totalCount} Selesai)</th>
                                            <th style="width:110px; text-align:center;">Status</th>
                                            <th style="width:100px; text-align:center;">Tgl Selesai</th>
                                            <th style="width:110px;">PIC Sub-Task</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${(w.checklist || []).map((c, cIdx) => `
                                            <tr style="${c.selesai ? 'background:rgba(16, 185, 129, 0.03);' : ''}">
                                                <td style="text-align:center; font-family:'JetBrains Mono'; font-size:0.78rem;">${cIdx + 1}</td>
                                                <td style="${c.selesai ? 'color:var(--text-main); font-weight:600;' : 'color:var(--text-muted);'}">
                                                    ${c.sub_task}
                                                </td>
                                                <td style="text-align:center;">
                                                    <span style="font-size:0.75rem; font-weight:700; color:${c.selesai ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                                        ${c.selesai ? '✅ Selesai' : '⬜ Belum'}
                                                    </span>
                                                </td>
                                                <td style="text-align:center; font-family:'JetBrains Mono'; font-size:0.78rem; color:var(--text-muted);">${c.tanggal || '-'}</td>
                                                <td style="font-size:0.8rem; font-weight:600;">${c.pic_task || w.pic || '-'}</td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            `}

                            ${(w.temuan || (w.jumlah_foto > 0)) ? `
                                <div style="margin-top:8px; padding:6px 10px; background:rgba(244,63,94,0.08); border:1px solid rgba(244,63,94,0.25); border-radius:4px; font-size:0.8rem;">
                                    <strong style="color:#fda4af;">⚠️ Temuan:</strong> ${w.temuan || '(Tercatat bukti foto)'} &bull; 
                                    <strong style="color:#fef08a;">Tindak Lanjut:</strong> ${w.tindak_lanjut || 'Menunggu verifikasi'} 
                                    ${w.jumlah_foto > 0 ? `(${w.jumlah_foto} 📷 Foto)` : ''}
                                </div>
                            ` : ''}
                        </div>`;
                    }).join('')}
                </div>
            </div>`;
            container.innerHTML = html;
        }

        /* OPSI 3: Laporan Actuator Valves Keseluruhan */
        function renderActuatorReport(container, s, printDateStr) {
            const actList = fullData.actuators || [];
            let html = `
            <div class="report-paper">
                <div class="report-header-box">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:2px solid #334155; padding-bottom:12px; margin-bottom:16px;">
                        <div>
                            <div style="font-size:1.1rem; font-weight:800; color:var(--primary); letter-spacing:0.5px;">⚡ PLTU MSW &bull; SECTION ELECTRIC, INSTRUMENT & CONTROL</div>
                            <h2 style="font-size:1.25rem; margin:4px 0 2px 0; color:#fff;">LAPORAN MONITORING ACTUATOR VALVES KESELURUHAN</h2>
                            <div style="font-size:0.85rem; color:var(--text-muted);">Status General Inspection, Function Test & Temuan Actuator Unit ${currentUnit}</div>
                        </div>
                        <div style="text-align:right; font-size:0.8rem; color:var(--text-muted);">
                            <div><strong>Total Actuator:</strong> ${actList.length} Valve</div>
                            <div><strong>Waktu Cetak:</strong> ${printDateStr}</div>
                            <div style="color:var(--primary); font-weight:700; margin-top:2px;">UNIT ${currentUnit}</div>
                        </div>
                    </div>
                </div>

                <div class="report-section-title">📊 1. RINGKASAN PROGRESS ACTUATOR VALVES</div>
                <div class="report-kpi-grid" style="grid-template-columns: repeat(4, 1fr);">
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">TOTAL ACTUATOR VALVES</div>
                        <div class="report-kpi-val">${s.actuator.total}</div>
                        <div class="report-kpi-sub">Total Valve Terjadwal</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">VALVE SELESAI (100%)</div>
                        <div class="report-kpi-val" style="color:var(--status-finish);">${s.actuator.finish}</div>
                        <div class="report-kpi-sub">Insp & Func Selesai Penuh</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">VALVE IN PROGRESS</div>
                        <div class="report-kpi-val" style="color:#38bdf8;">${s.actuator.in_progress}</div>
                        <div class="report-kpi-sub">Selesai Salah Satu Tahap</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">OVERALL PROGRESS</div>
                        <div class="report-kpi-val">${s.actuator.pct}%</div>
                        <div class="report-kpi-sub">${s.actuator.subtask_done} / ${s.actuator.subtask_total} Uji Selesai</div>
                    </div>
                </div>

                <div class="report-section-title" style="margin-top:22px;">
                    📋 2. TABEL LENGKAP STATUS ACTUATOR VALVES (${actList.length} Valve)
                </div>

                <table class="report-table">
                    <thead>
                        <tr>
                            <th style="width:30px;">No</th>
                            <th style="width:120px;">Equipment ID</th>
                            <th style="width:110px;">KKS / Tag</th>
                            <th>Deskripsi Actuator Valve</th>
                            <th style="width:85px;">Area</th>
                            <th style="width:75px; text-align:center;">General Insp</th>
                            <th style="width:75px; text-align:center;">Function Test</th>
                            <th style="width:65px; text-align:center;">% Prog</th>
                            <th style="width:85px; text-align:center;">Status</th>
                            <th style="width:85px;">PIC</th>
                            <th style="width:85px; text-align:center;">Tgl Finish</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${actList.map((a, idx) => `
                            <tr>
                                <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${a.equipment_id}</td>
                                <td style="font-family:'JetBrains Mono'; font-size:0.78rem; color:var(--text-muted);">${a.kks || '-'}</td>
                                <td style="font-weight:600; font-size:0.83rem;">
                                    ${a.equipment_description}
                                    ${(a.temuan || (a.jumlah_foto > 0)) ? `
                                        <div style="font-size:0.75rem; color:#fda4af; font-weight:normal; margin-top:2px;">⚠️ ${a.temuan || 'Tercatat foto'}</div>
                                    ` : ''}
                                </td>
                                <td><span style="font-size:0.8rem; color:var(--text-muted);">${a.area || 'BOILER'}</span></td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${a.general_inspection ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                    ${a.general_inspection ? '✅ OK' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${a.function_test ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                    ${a.function_test ? '✅ OK' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center; font-family:'JetBrains Mono'; font-weight:800;">${a.persen_progress || 0}%</td>
                                <td style="text-align:center;">
                                    <span class="report-badge" style="background:${a.status === 'FINISH' ? 'rgba(16,185,129,0.15)' : (a.status === 'IN PROGRESS' ? 'rgba(56,189,248,0.15)' : 'rgba(148,163,184,0.15)')}; color:${a.status === 'FINISH' ? 'var(--status-finish)' : (a.status === 'IN PROGRESS' ? '#38bdf8' : '#94a3b8')};">
                                        ${a.status}
                                    </span>
                                </td>
                                <td style="font-weight:600; font-size:0.8rem;">${a.pic || '-'}</td>
                                <td style="text-align:center; font-family:'JetBrains Mono'; font-size:0.78rem;">${a.finish_date || '-'}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>`;
            container.innerHTML = html;
        }

        /* OPSI 4: Laporan Instruments (TX & PSW) Keseluruhan */
        function renderInstrumentsReport(container, s, printDateStr) {
            const ptxList = fullData.pressure_tx || [];
            const ttxList = fullData.temperature_tx || [];
            const pswList = fullData.pressure_switch || [];
            const totalInst = ptxList.length + ttxList.length + pswList.length;

            let html = `
            <div class="report-paper">
                <div class="report-header-box">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:2px solid #334155; padding-bottom:12px; margin-bottom:16px;">
                        <div>
                            <div style="font-size:1.1rem; font-weight:800; color:var(--primary); letter-spacing:0.5px;">⚡ PLTU MSW &bull; SECTION ELECTRIC, INSTRUMENT & CONTROL</div>
                            <h2 style="font-size:1.25rem; margin:4px 0 2px 0; color:#fff;">LAPORAN MONITORING INSTRUMENTS KESELURUHAN</h2>
                            <div style="font-size:0.85rem; color:var(--text-muted);">Status Kalibrasi & Verifikasi Transmitter (PTX, TTX) dan Pressure Switch Unit ${currentUnit}</div>
                        </div>
                        <div style="text-align:right; font-size:0.8rem; color:var(--text-muted);">
                            <div><strong>Total Instrumen:</strong> ${totalInst} Item</div>
                            <div><strong>Waktu Cetak:</strong> ${printDateStr}</div>
                            <div style="color:var(--primary); font-weight:700; margin-top:2px;">UNIT ${currentUnit}</div>
                        </div>
                    </div>
                </div>

                <div class="report-section-title">📊 1. RINGKASAN STATUS INSTRUMENTS</div>
                <div class="report-kpi-grid" style="grid-template-columns: repeat(4, 1fr);">
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">TOTAL INSTRUMENTS</div>
                        <div class="report-kpi-val">${s.instrument.total}</div>
                        <div class="report-kpi-sub">PTX, TTX, dan Pressure Switch</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">SELESAI VERIFIKASI (DONE)</div>
                        <div class="report-kpi-val" style="color:var(--status-finish);">${s.instrument.done}</div>
                        <div class="report-kpi-sub">Penentu Finish Tercapai</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">IN PROGRESS (KALIBRASI)</div>
                        <div class="report-kpi-val" style="color:#38bdf8;">${s.instrument.calib_done || 0}</div>
                        <div class="report-kpi-sub">Kalibrasi Telah Dilakukan</div>
                    </div>
                    <div class="report-kpi-card">
                        <div class="report-kpi-lbl">VERIFIKASI PROGRESS</div>
                        <div class="report-kpi-val">${s.instrument.pct}%</div>
                        <div class="report-kpi-sub">${s.instrument.done} / ${s.instrument.total} Selesai Penuh</div>
                    </div>
                </div>

                <!-- Bagian A: Pressure Transmitter -->
                <div class="report-section-title" style="margin-top:22px;">
                    🎛️ 2A. PRESSURE TRANSMITTERS (PTX) &bull; ${ptxList.length} Item
                </div>
                <table class="report-table">
                    <thead>
                        <tr>
                            <th style="width:30px;">No</th>
                            <th style="width:130px;">KKS / Tag</th>
                            <th>Equipment Description</th>
                            <th style="width:120px;">Range / Unit</th>
                            <th style="width:90px; text-align:center;">🛠️ Kalibrasi</th>
                            <th style="width:105px; text-align:center;">🔍 Verifikasi (Done)</th>
                            <th style="width:95px; text-align:center;">Status</th>
                            <th>Catatan / Temuan</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ptxList.map((item, idx) => {
                            const isDone = !!item.verifikasi;
                            const isCalib = !!item.kalibrasi;
                            return `
                            <tr>
                                <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.kks}</td>
                                <td style="font-weight:600;">${item.equipment}</td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.range || '-'} ${item.eng_unit || ''}</td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isCalib ? '#38bdf8' : 'var(--text-muted)'};">
                                    ${isCalib ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isDone ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                    ${isDone ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center;">
                                    <span class="report-badge" style="background:${isDone ? 'rgba(16,185,129,0.15)' : (isCalib ? 'rgba(56,189,248,0.15)' : 'rgba(148,163,184,0.15)')}; color:${isDone ? 'var(--status-finish)' : (isCalib ? '#38bdf8' : '#94a3b8')};">
                                        ${isDone ? 'DONE' : (isCalib ? 'CALIB OK' : 'SCHEDULED')}
                                    </span>
                                </td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.temuan ? '⚠️ ' + item.temuan : (item.remarks || '-')}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>

                <!-- Bagian B: Temperature Transmitter -->
                <div class="report-section-title" style="margin-top:22px;">
                    🌡️ 2B. TEMPERATURE TRANSMITTERS (TTX) &bull; ${ttxList.length} Item
                </div>
                <table class="report-table">
                    <thead>
                        <tr>
                            <th style="width:30px;">No</th>
                            <th style="width:130px;">KKS / Tag</th>
                            <th>Equipment Description</th>
                            <th style="width:120px;">Range / Unit</th>
                            <th style="width:90px; text-align:center;">🛠️ Kalibrasi</th>
                            <th style="width:105px; text-align:center;">🔍 Verifikasi (Done)</th>
                            <th style="width:95px; text-align:center;">Status</th>
                            <th>Catatan / Temuan</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ttxList.map((item, idx) => {
                            const isDone = !!item.verifikasi;
                            const isCalib = !!item.kalibrasi;
                            return `
                            <tr>
                                <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.kks}</td>
                                <td style="font-weight:600;">${item.equipment}</td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.range || '-'} ${item.eng_unit || ''}</td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isCalib ? '#38bdf8' : 'var(--text-muted)'};">
                                    ${isCalib ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isDone ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                    ${isDone ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center;">
                                    <span class="report-badge" style="background:${isDone ? 'rgba(16,185,129,0.15)' : (isCalib ? 'rgba(56,189,248,0.15)' : 'rgba(148,163,184,0.15)')}; color:${isDone ? 'var(--status-finish)' : (isCalib ? '#38bdf8' : '#94a3b8')};">
                                        ${isDone ? 'DONE' : (isCalib ? 'CALIB OK' : 'SCHEDULED')}
                                    </span>
                                </td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.temuan ? '⚠️ ' + item.temuan : (item.remarks || '-')}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>

                <!-- Bagian C: Pressure Switch -->
                <div class="report-section-title" style="margin-top:22px;">
                    🔘 2C. PRESSURE SWITCHES (PSW) &bull; ${pswList.length} Item
                </div>
                <table class="report-table">
                    <thead>
                        <tr>
                            <th style="width:30px;">No</th>
                            <th style="width:130px;">KKS / Tag</th>
                            <th>Description</th>
                            <th style="width:120px;">Setpoint / Range</th>
                            <th style="width:90px; text-align:center;">🛠️ Kalibrasi</th>
                            <th style="width:105px; text-align:center;">🔍 Verifikasi (Done)</th>
                            <th style="width:95px; text-align:center;">Status</th>
                            <th>Catatan / Temuan</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${pswList.map((item, idx) => {
                            const isDone = !!item.verifikasi;
                            const isCalib = !!item.kalibrasi;
                            return `
                            <tr>
                                <td style="text-align:center; font-family:'JetBrains Mono';">${idx + 1}</td>
                                <td style="font-family:'JetBrains Mono'; font-weight:700; color:var(--primary);">${item.kks}</td>
                                <td style="font-weight:600;">${item.description}</td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.setpoint || '-'} ${item.range || ''}</td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isCalib ? '#38bdf8' : 'var(--text-muted)'};">
                                    ${isCalib ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center; font-size:0.8rem; font-weight:700; color:${isDone ? 'var(--status-finish)' : 'var(--text-muted)'};">
                                    ${isDone ? '✅ Selesai' : '⬜ Belum'}
                                </td>
                                <td style="text-align:center;">
                                    <span class="report-badge" style="background:${isDone ? 'rgba(16,185,129,0.15)' : (isCalib ? 'rgba(56,189,248,0.15)' : 'rgba(148,163,184,0.15)')}; color:${isDone ? 'var(--status-finish)' : (isCalib ? '#38bdf8' : '#94a3b8')};">
                                        ${isDone ? 'DONE' : (isCalib ? 'CALIB OK' : 'SCHEDULED')}
                                    </span>
                                </td>
                                <td style="font-size:0.8rem; color:var(--text-muted);">${item.temuan ? '⚠️ ' + item.temuan : (item.remarks || '-')}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>`;
            container.innerHTML = html;
        }

        
        function getSubtaskTypeBadge(desc) {
            if(!desc) return '';
            const s = String(desc).replace(/\xa0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
            
            // 1. Actuator Detection
            if(s.includes('ACTUATOR') || s.includes(' MOV') || s.includes('MOV ') || s.includes(' AOV') || s.includes('AOV ') || s.includes('GATE ACTUATOR') || s.includes('DAMPER ACTUATOR') || s.includes('VALVE ACTUATOR') || s.includes('FEED WATER CONTROL VALVE')) {
                return '<span class="badge-tag-comp badge-tag-act">ACTUATOR</span>';
            }
            const acts = (fullData.master_actuators || fullData.actuators || []);
            for(let i = 0; i < acts.length; i++) {
                const a = acts[i];
                const aKks = (a.kks || '').replace(/\xa0/g, ' ').trim().toUpperCase();
                if(aKks.length >= 6) {
                    const kksCore = aKks.length >= 8 ? aKks.slice(2) : aKks;
                    if(s.includes(aKks) || s.includes(kksCore)) {
                        return '<span class="badge-tag-comp badge-tag-act">ACTUATOR</span>';
                    }
                }
                const aDesc = (a.equipment_description || '').replace(/\xa0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
                if(aDesc.length >= 8 && (s.includes(aDesc) || aDesc.includes(s))) {
                    return '<span class="badge-tag-comp badge-tag-act">ACTUATOR</span>';
                }
            }
            
            // 2. Instrument Detection (PTX, TTX, PSW)
            if(s.includes('TRANSMITTER') || s.includes('PRESSURE TRANSMITTER') || s.includes('TEMP TRANSMITTER') || s.includes('TEMPERATURE TRANSMITTER') || s.includes('PRESSURE SWITCH') || s.includes('TEMP SWITCH') || s.includes('TEMPERATURE SWITCH') || s.includes('KALIBRASI TRANSMITTER') || s.includes('CALIBRATION MEASUREMENT') || s.includes('MEASUREMENT DEVICE')) {
                return '<span class="badge-tag-comp badge-tag-inst">INSTRUMENT</span>';
            }
            const insts = [...(fullData.pressure_tx || []), ...(fullData.temperature_tx || []), ...(fullData.pressure_switch || []), ...(fullData.master_instruments || [])];
            for(let i = 0; i < insts.length; i++) {
                const inst = insts[i];
                const iKks = (inst.kks || '').replace(/\xa0/g, ' ').trim().toUpperCase();
                if(iKks.length >= 6) {
                    const kksCore = iKks.length >= 8 ? iKks.slice(2) : iKks;
                    if(s.includes(iKks) || s.includes(kksCore)) {
                        return '<span class="badge-tag-comp badge-tag-inst">INSTRUMENT</span>';
                    }
                }
                const iDesc = (inst.equipment || inst.description || '').replace(/\xa0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
                if(iDesc.length >= 8 && (s.includes(iDesc) || iDesc.includes(s))) {
                    return '<span class="badge-tag-comp badge-tag-inst">INSTRUMENT</span>';
                }
            }
            
            // Jika bukan Actuator dan bukan Instrument -> TANPA LABEL
            return '';
        }

        const subtaskAddModes = {};

        function setSubtaskMode(noWo, mode) {
            subtaskAddModes[noWo] = mode;
            const cardEl = document.getElementById(`card-wo-${noWo}`);
            if(!cardEl) return;
            
            const btnManual = cardEl.querySelector(`.btn-mode-manual`);
            const btnAct = cardEl.querySelector(`.btn-mode-act`);
            const btnInst = cardEl.querySelector(`.btn-mode-inst`);
            
            if(btnManual) btnManual.classList.toggle('active', mode === 'manual');
            if(btnAct) btnAct.classList.toggle('active', mode === 'actuator');
            if(btnInst) btnInst.classList.toggle('active', mode === 'instrument');
            
            const boxManual = document.getElementById(`box-subtask-manual-${noWo}`);
            const boxAct = document.getElementById(`box-subtask-act-${noWo}`);
            const boxInst = document.getElementById(`box-subtask-inst-${noWo}`);
            
            if(boxManual) boxManual.style.display = mode === 'manual' ? 'flex' : 'none';
            if(boxAct) boxAct.style.display = mode === 'actuator' ? 'flex' : 'none';
            if(boxInst) boxInst.style.display = mode === 'instrument' ? 'flex' : 'none';
        }

        /* ---------------- QUICK ACTIONS (NO AUTO REFRESH) ---------------- */
        function toggleLocalSubtask(noWo, cIdx, isChecked) {
            const item = (fullData.work_orders || []).find(w => w.no_wo === noWo);
            if(!item || !item.checklist || !item.checklist[cIdx]) return;
            
            item.checklist[cIdx].selesai = isChecked;
            const nowStr = new Date().toLocaleDateString('id-ID');
            item.checklist[cIdx].tanggal = isChecked ? nowStr : '';
            
            const chkItemEl = document.getElementById(`chk-${noWo}-${cIdx}`)?.closest('.checklist-item');
            if(chkItemEl) {
                chkItemEl.classList.toggle('done', isChecked);
                const spanEl = chkItemEl.querySelector('.checklist-item-body span');
                if(spanEl) {
                    spanEl.style.textDecoration = isChecked ? 'line-through' : 'none';
                    spanEl.style.color = isChecked ? 'var(--text-muted)' : 'var(--text-main)';
                    spanEl.style.opacity = isChecked ? '0.65' : '1';
                }
                const rightBox = chkItemEl.querySelector('.header-right');
                let dateBadge = rightBox?.querySelector('.date-badge');
                if(isChecked) {
                    if(!dateBadge && rightBox) {
                        const badge = document.createElement('span');
                        badge.className = 'date-badge';
                        badge.title = "Tanggal Dikerjakan";
                        badge.innerText = `${nowStr}`;
                        const delBtn = rightBox.querySelector('.btn-del-subtask-cross');
                        rightBox.insertBefore(badge, delBtn);
                    }
                } else {
                    if(dateBadge) dateBadge.remove();
                }
            }
            
            const total = item.checklist.length;
            const done = item.checklist.filter(c => c.selesai).length;
            const pct = total > 0 ? Math.round((done / total) * 100) : 0;
            item.persen_progress = pct;
            item.status = (done === total && total > 0) ? 'FINISH' : (done > 0 ? 'IN PROGRESS' : 'SCHED-OK');
            
            // Update UI
            const woIdx = (fullData.work_orders || []).indexOf(item);
            const bodyId = getCardBodyId('wo', noWo);
            const bodyEl = document.getElementById(bodyId) || document.getElementById(`body-wo-${woIdx}`);
            if(bodyEl) {
                const secH4 = bodyEl.querySelector('.section-h4');
                if(secH4) secH4.innerText = `📋 Checklist Sub-Task (${done} / ${total} Selesai)`;
                
                const headerEl = document.getElementById(`card-wo-${noWo}`)?.querySelector('.item-header');
                if(headerEl) {
                    const subCountEl = headerEl.querySelector('.wo-subtask-progress');
                    if(subCountEl) subCountEl.innerText = `${done} / ${total} Sub-task`;
                    const barFill = headerEl.querySelector('.progress-fill');
                    if(barFill) barFill.style.width = `${pct}%`;
                    const pctText = headerEl.querySelector('.progress-text');
                    if(pctText) pctText.innerText = `${pct}%`;
                    
                    const badgeEl = headerEl.querySelector('.status-badge');
                    if(badgeEl) {
                        const st = item.status.replace(/\s+/g, '_');
                        badgeEl.className = `status-badge badge-${st}`;
                        badgeEl.innerText = item.status;
                    }
                }
            }
        }

        function toggleLocalInstCheck(instSubtab, idx, field, isChecked) {
            let list = instSubtab === 'ptx' ? fullData.pressure_tx : (instSubtab === 'ttx' ? fullData.temperature_tx : fullData.pressure_switch);
            if(!list || !list[idx]) return;
            const item = list[idx];
            item[field] = isChecked;
            if(field === 'verifikasi') {
                item.status_wdone = isChecked;
            }
            
            const isVerif = !!item.verifikasi;
            const isCalib = !!item.kalibrasi;
            
            // Update card elements locally
            const cardPrefix = instSubtab === 'psw' ? 'psw' : 'inst';
            const calibCard = document.getElementById(`card-calib-${cardPrefix}-${idx}`);
            const verifCard = document.getElementById(`card-verif-${cardPrefix}-${idx}`);
            const badgeEl = document.getElementById(`badge-${cardPrefix}-${idx}`);
            
            if(calibCard) calibCard.classList.toggle('done', isCalib);
            if(verifCard) {
                verifCard.classList.toggle('done', isVerif);
                verifCard.style.borderColor = isVerif ? 'var(--status-finish)' : 'var(--border-color)';
                const verifSpan = verifCard.querySelector('label span');
                if(verifSpan) verifSpan.style.color = isVerif ? 'var(--status-finish)' : 'var(--primary)';
            }
            if(badgeEl) {
                badgeEl.className = `status-badge ${isVerif ? 'badge-FINISH' : (isCalib ? 'badge-IN-PROGRESS' : 'badge-SCHED-OK')}`;
                badgeEl.innerText = isVerif ? 'DONE (100%)' : (isCalib ? 'IN PROGRESS (Kalibrasi OK)' : 'SCHEDULED');
            }
        }

        async function toggleDirectInstCheck(instSubtab, key, field, isChecked) {
            const payload = {
                unit: currentUnit,
                type: instSubtab === 'ptx' ? 'pressure_tx' : (instSubtab === 'ttx' ? 'temperature_tx' : 'pressure_switch'),
                kks: key,
                no: key
            };
            payload[field] = isChecked;
            if(field === 'verifikasi') payload.status_wdone = isChecked;
            
            try {
                const res = await fetch('/api/update_instrument', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(`✓ ${field === 'verifikasi' ? 'Verifikasi' : 'Kalibrasi'} diperbarui!`, 'success', 1800);
                    loadData();
                }
            } catch(e) {
                showToast('Gagal memperbarui instrumen', 'error');
            }
        }

        /* ---------------- SAVE / EDIT SUBMISSIONS ---------------- */
        function getTodayFormatted() {
            const d = new Date();
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear();
            return `${day}/${month}/${year}`;
        }

        async function saveWorkOrder(noWo) {
            const item = (fullData.work_orders || []).find(w => w.no_wo === noWo);
            if (!item) return;

            const picSelect = document.getElementById(`pic-${noWo}`);
            const remInput = document.getElementById(`rem-${noWo}`);

            const checklistPayload = (item.checklist || []).map((c, cIdx) => {
                const chk = document.getElementById(`chk-${noWo}-${cIdx}`);
                const isChecked = chk ? chk.checked : !!c.selesai;
                return {
                    sub_task: c.sub_task,
                    selesai: isChecked,
                    tanggal: isChecked ? (c.tanggal || getTodayFormatted()) : '',
                    pic_task: c.pic_task || ''
                };
            });

            const allDone = checklistPayload.length > 0 && checklistPayload.every(c => c.selesai);
            const autoFinishDate = allDone ? (item.tanggal_finish || getTodayFormatted()) : '';

            const payload = {
                unit: currentUnit,
                no_wo: noWo,
                pic: picSelect ? picSelect.value : (item.pic || ''),
                tanggal_finish: autoFinishDate,
                remarks: remInput ? remInput.value : (item.remarks || ''),
                checklist: checklistPayload
            };

            try {
                const res = await fetch('/api/update_wo', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast(result.message || 'WO berhasil diperbarui!', 'success');
                loadData();
            } catch(e) {
                showToast('Gagal menyimpan perubahan WO', 'error');
            }
        }

        async function quickToggleActuator(eqId, field, isChecked) {
            try {
                const payload = {
                    unit: currentUnit,
                    equipment_id: eqId,
                    field: field,
                    value: isChecked
                };
                const res = await fetch('/api/quick_toggle_actuator', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message, 'success');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal update status actuator', 'error');
                }
            } catch(e) {
                showToast('Gagal update status actuator', 'error');
            }
        }

        async function saveActuator(eqId, desc) {
            const genChk = document.getElementById(`gen-${eqId}`);
            const funcChk = document.getElementById(`func-${eqId}`);
            const isGen = genChk ? genChk.checked : false;
            const isFunc = funcChk ? funcChk.checked : false;
            const isAllDone = isGen && isFunc;
            const autoFinishDate = isAllDone ? getTodayFormatted() : '';

            const payload = {
                unit: currentUnit,
                equipment_id: eqId,
                equipment_description: desc,
                pic: document.getElementById(`pic-act-${eqId}`).value,
                finish_date: autoFinishDate,
                remarks: document.getElementById(`rem-act-${eqId}`).value,
                general_inspection: isGen,
                function_test: isFunc
            };

            try {
                const res = await fetch('/api/update_actuator', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast(result.message || 'Actuator berhasil diperbarui!', 'success');
                loadData();
            } catch(e) {
                showToast('Gagal menyimpan actuator', 'error');
            }
        }

        async function saveTransmitter(type, key, idx) {
            const calibChk = document.getElementById(`inst-calib-${idx}`);
            const verifChk = document.getElementById(`inst-verif-${idx}`);
            const remInput = document.getElementById(`inst-rem-${idx}`);
            
            const payload = {
                unit: currentUnit,
                type: type === 'ptx' ? 'pressure_tx' : 'temperature_tx',
                kks: key,
                no: key,
                kalibrasi: calibChk ? calibChk.checked : false,
                verifikasi: verifChk ? verifChk.checked : false,
                remarks: remInput ? remInput.value : ''
            };

            try {
                const res = await fetch('/api/update_instrument', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast('Instrumen berhasil diperbarui!', 'success');
                loadData();
            } catch(e) {
                showToast('Gagal menyimpan instrumen', 'error');
            }
        }

        async function savePressureSwitch(key, idx) {
            const calibChk = document.getElementById(`inst-calib-${idx}`);
            const verifChk = document.getElementById(`inst-verif-${idx}`);
            const remInput = document.getElementById(`inst-rem-${idx}`);
            
            const payload = {
                unit: currentUnit,
                type: 'pressure_switch',
                kks: key,
                no: key,
                asfound_set: document.getElementById(`af-set-${idx}`) ? document.getElementById(`af-set-${idx}`).value : '',
                asfound_reset: document.getElementById(`af-reset-${idx}`) ? document.getElementById(`af-reset-${idx}`).value : '',
                asleft_set: document.getElementById(`al-set-${idx}`) ? document.getElementById(`al-set-${idx}`).value : '',
                asleft_reset: document.getElementById(`al-reset-${idx}`) ? document.getElementById(`al-reset-${idx}`).value : '',
                status_ok_notok: document.getElementById(`res-psw-${idx}`) ? document.getElementById(`res-psw-${idx}`).value : 'OK',
                kalibrasi: calibChk ? calibChk.checked : false,
                verifikasi: verifChk ? verifChk.checked : false,
                remarks: remInput ? remInput.value : ''
            };

            try {
                const res = await fetch('/api/update_instrument', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast('Kalibrasi & Verifikasi Pressure Switch berhasil disimpan!', 'success');
                loadData();
            } catch(e) {
                showToast('Gagal menyimpan kalibrasi PSW', 'error');
            }
        }

        async function saveNewWO() {
            const noWo = (document.getElementById('new-wo-code').value || '').trim();
            const desc = (document.getElementById('new-wo-desc').value || '').trim();
            const area = (document.getElementById('new-wo-area').value || '').trim();
            const pic = document.getElementById('new-wo-pic').value;
            const sched = formatDateForStorage(document.getElementById('new-wo-sched').value);
            const checklistStr = (document.getElementById('new-wo-checklist').value || '').trim();

            if(!noWo || !desc) {
                showToast('No WO dan Job Description wajib diisi!', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_wo', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, no_wo: noWo, job_description: desc, area: area, pic: pic, tanggal_schedule: sched, checklist_str: checklistStr})
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message || 'WO baru berhasil ditambahkan!', 'success');
                    document.getElementById('new-wo-code').value = '';
                    document.getElementById('new-wo-desc').value = '';
                    document.getElementById('new-wo-area').value = '';
                    document.getElementById('new-wo-sched').value = '';
                    document.getElementById('new-wo-checklist').value = '';
                    toggleAccordion('add-wo-form');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal menambah WO', 'error');
                }
            } catch(e) {
                showToast('Gagal menambah WO baru: ' + e.message, 'error');
            }
        }

        async function deleteWO(noWo) {
            if(!confirm(`Yakin ingin menghapus Work Order ${noWo} dan seluruh checklist subtask-nya?`)) return;

            try {
                const res = await fetch('/api/delete_wo', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, no_wo: noWo})
                });
                const result = await res.json();
                showToast(result.message || 'WO berhasil dihapus!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus WO', 'error');
            }
        }

        async function addSubtask(noWo, mode = 'manual') {
            let subTask = '';
            let defaultPic = '';
            
            if(mode === 'actuator') {
                const sel = document.getElementById(`new-subtask-act-${noWo}`);
                subTask = (sel?.value || '').trim();
                defaultPic = 'AMP';
                if(!subTask) {
                    showToast('Silakan pilih Actuator dari dropdown!', 'error');
                    return;
                }
            } else if(mode === 'instrument') {
                const sel = document.getElementById(`new-subtask-inst-${noWo}`);
                subTask = (sel?.value || '').trim();
                defaultPic = 'JAPA';
                if(!subTask) {
                    showToast('Silakan pilih Instrument dari dropdown!', 'error');
                    return;
                }
            } else {
                const input = document.getElementById(`new-subtask-${noWo}`);
                subTask = (input?.value || '').trim();
                if(!subTask) {
                    showToast('Silakan masukkan deskripsi sub-task!', 'error');
                    return;
                }
            }

            try {
                const res = await fetch('/api/add_subtask', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, no_wo: noWo, sub_task: subTask, pic: defaultPic})
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message || 'Sub-task berhasil ditambahkan dan disinkronkan!', 'success');
                    const input = document.getElementById(`new-subtask-${noWo}`);
                    if(input) input.value = '';
                    loadData();
                } else {
                    showToast(result.message || 'Gagal menambah sub-task', 'error');
                }
            } catch(e) {
                showToast('Gagal menambah sub-task: ' + e.message, 'error');
            }
        }

        async function deleteSubtask(noWo, subTask) {
            if(!confirm(`Hapus sub-task "${subTask}" dari ${noWo}?`)) return;

            try {
                const res = await fetch('/api/delete_subtask', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, no_wo: noWo, sub_task: subTask})
                });
                const result = await res.json();
                showToast(result.message || 'Sub-task berhasil dihapus!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus sub-task', 'error');
            }
        }

        async function saveNewActuator() {
            const eqId = (document.getElementById('new-act-id').value || '').trim();
            const desc = (document.getElementById('new-act-desc').value || '').trim();
            const area = (document.getElementById('new-act-area').value || '').trim();
            const kks = (document.getElementById('new-act-kks').value || '').trim();
            const pic = document.getElementById('new-act-pic').value;

            if(!eqId || !desc) {
                showToast('Equipment ID dan Description wajib diisi!', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_actuator', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, equipment_id: eqId, equipment_description: desc, area: area, kks: kks, pic: pic})
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message || 'Actuator baru berhasil ditambahkan!', 'success');
                    document.getElementById('new-act-id').value = '';
                    document.getElementById('new-act-desc').value = '';
                    document.getElementById('new-act-area').value = '';
                    document.getElementById('new-act-kks').value = '';
                    toggleAccordion('add-act-form');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal menambah Actuator', 'error');
                }
            } catch(e) {
                showToast('Gagal menambah Actuator baru: ' + e.message, 'error');
            }
        }

        async function deleteActuator(eqId) {
            if(!confirm(`Yakin ingin menghapus Actuator Valve ${eqId}?`)) return;

            try {
                const res = await fetch('/api/delete_actuator', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, equipment_id: eqId})
                });
                const result = await res.json();
                showToast(result.message || 'Actuator berhasil dihapus!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus Actuator', 'error');
            }
        }

        async function saveNewInstrument() {
            const type = document.getElementById('new-inst-type').value;
            const desc = (document.getElementById('new-inst-desc').value || '').trim();
            const kks = (document.getElementById('new-inst-kks').value || '').trim();
            const area = (document.getElementById('new-inst-area').value || '').trim();
            const range = (document.getElementById('new-inst-range').value || '').trim();

            if(!desc) {
                showToast('Nama Equipment wajib diisi!', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_instrument', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, type: type, equipment: desc, kks: kks, area: area, range: range})
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message || 'Instrument baru berhasil ditambahkan!', 'success');
                    document.getElementById('new-inst-desc').value = '';
                    document.getElementById('new-inst-kks').value = '';
                    document.getElementById('new-inst-area').value = '';
                    document.getElementById('new-inst-range').value = '';
                    toggleAccordion('add-inst-form');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal menambah Instrument', 'error');
                }
            } catch(e) {
                showToast('Gagal menambah Instrument baru: ' + e.message, 'error');
            }
        }

        async function saveNewScope() {
            const cat = (document.getElementById('new-scope-cat').value || '').trim();
            const eq = (document.getElementById('new-scope-eq').value || '').trim();
            const type = document.getElementById('new-scope-type').value;
            const pic = document.getElementById('new-scope-pic').value;

            if(!eq) {
                showToast('Nama Equipment / Scope wajib diisi!', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_scope', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, kategori: cat, nama_equipment: eq, tipe_scope: type, pic: pic})
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message || 'Scope Master berhasil ditambahkan!', 'success');
                    document.getElementById('new-scope-cat').value = '';
                    document.getElementById('new-scope-eq').value = '';
                    toggleAccordion('add-scope-form');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal menambah Scope Master', 'error');
                }
            } catch(e) {
                showToast('Gagal menambah Scope: ' + e.message, 'error');
            }
        }

        async function deleteInstrument(type, key) {
            if(!confirm(`Yakin ingin menghapus Instrument ini?`)) return;

            try {
                const res = await fetch('/api/delete_instrument', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, type: type, no: key, kks: key})
                });
                const result = await res.json();
                showToast(result.message || 'Instrument berhasil dihapus!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus Instrument', 'error');
            }
        }

        async function addNewPic() {
            const input = document.getElementById('new-pic-input');
            const picName = (input.value || '').trim();
            if(!picName) {
                showToast('Silakan masukkan nama PIC baru.', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_pic', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, pic_name: picName})
                });
                const result = await res.json();
                showToast(result.message || 'PIC berhasil ditambahkan!', 'success');
                input.value = '';
                loadData();
            } catch(e) {
                showToast('Gagal menambahkan PIC', 'error');
            }
        }

        async function deletePic(picName) {
            if(!confirm(`Yakin ingin menghapus PIC "${picName}" dari Master PIC?`)) return;

            try {
                const res = await fetch('/api/delete_pic', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, pic_name: picName})
                });
                const result = await res.json();
                showToast(result.message || 'PIC berhasil dihapus dari master!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus PIC', 'error');
            }
        }

        async function saveScopeRow(sIdx) {
            const item = (fullData.scope_master || [])[sIdx];
            if(!item) return;
            const eqInput = document.getElementById(`scope-eq-${sIdx}`);
            const typeSelect = document.getElementById(`scope-type-${sIdx}`);
            const picSelect = document.getElementById(`scope-pic-${sIdx}`);

            const payload = {
                unit: currentUnit,
                row_index: item.row_index !== undefined ? item.row_index : sIdx,
                nama_equipment: (eqInput ? eqInput.value : item.nama_equipment || '').trim(),
                tipe_scope: typeSelect ? typeSelect.value : (item.tipe_scope || 'Vendor'),
                pic: picSelect ? picSelect.value : (item.pic || '')
            };

            try {
                const res = await fetch('/api/update_scope', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                showToast(result.message || 'Scope Master berhasil diperbarui!', 'success');
                loadData();
            } catch(e) {
                showToast('Gagal memperbarui Scope Master', 'error');
            }
        }

        async function deleteScopeRow(sIdx) {
            const item = (fullData.scope_master || [])[sIdx];
            if(!item) return;
            if(!confirm(`Yakin ingin menghapus baris Master Scope "${item.nama_equipment || ''}"?`)) return;

            try {
                const res = await fetch('/api/delete_scope', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, row_index: item.row_index !== undefined ? item.row_index : sIdx})
                });
                const result = await res.json();
                showToast(result.message || 'Baris Master Scope berhasil dihapus!', 'info');
                loadData();
            } catch(e) {
                showToast('Gagal menghapus baris Master Scope', 'error');
            }
        }

        // Drag and drop dropzone handling
        const dropzone = document.getElementById('photo-dropzone');
        if(dropzone) {
            ['dragenter', 'dragover'].forEach(eventName => {
                dropzone.addEventListener(eventName, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.add('dragover'); }, false);
            });
            ['dragleave', 'drop'].forEach(eventName => {
                dropzone.addEventListener(eventName, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('dragover'); }, false);
            });
            dropzone.addEventListener('drop', (e) => {
                const dt = e.dataTransfer;
                const files = dt.files;
                handleModalFileSelect(files);
            }, false);
        }

        /* ---------------- S-CURVE, WA SUMMARY & EXCEL MODALS ---------------- */
        function openSCurveModal() {
            renderSCurveChart();
            document.getElementById('scurve-modal').classList.add('open');
        }
        function closeSCurveModal() {
            document.getElementById('scurve-modal').classList.remove('open');
        }

        function openWaSummaryModal() {
            generateWaText();
            document.getElementById('wa-modal').classList.add('open');
        }
        function closeWaSummaryModal() {
            document.getElementById('wa-modal').classList.remove('open');
        }

        function downloadExcel() {
            const url = `/api/export_excel?unit=${currentUnit}`;
            showToast(`📥 Mengunduh Laporan Excel Unit ${currentUnit}...`, 'info', 2000);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Laporan_Monitoring_Outage_EIC_Unit_${currentUnit}.xlsx`;
            document.body.appendChild(a);
            a.click();
            a.remove();
        }

        function scrollToTop() {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        // Scroll listener for Sticky Bar & Back to Top Button
        window.addEventListener('scroll', () => {
            const stBar = document.getElementById('sticky-summary-bar');
            const topBtn = document.getElementById('back-to-top-btn');
            const scrollPos = window.scrollY;
            if(stBar) {
                stBar.classList.toggle('visible', scrollPos > 180);
            }
            if(topBtn) {
                topBtn.classList.toggle('visible', scrollPos > 300);
            }
        });

        async function batchToggleSubtasks(noWo, action) {
            const isDone = (action === 'mark_all_done');
            if(!confirm(`Apakah Anda yakin ingin ${isDone ? 'menandai SEMUA sub-task selesai' : 'mereset SEMUA sub-task'} untuk WO ${noWo}?`)) {
                return;
            }
            try {
                const payload = { unit: currentUnit, no_wo: noWo, action: action };
                const res = await fetch('/api/batch_toggle_subtasks', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if(result.status === 'success') {
                    showToast(result.message, 'success');
                    loadData();
                } else {
                    showToast(result.message || 'Gagal mengubah batch sub-task', 'error');
                }
            } catch(e) {
                showToast('Gagal memproses batch checklist', 'error');
            }
        }

        function renderSCurveChart() {
            const chartBox = document.getElementById('scurve-chart-container');
            if(!chartBox || !fullData) return;

            // Collect all completed dates and count tasks per date
            const dateMap = {};
            let totalTasks = 0;

            // Total tasks pool
            (fullData.work_orders || []).forEach(w => {
                const chk = w.checklist || [];
                if(chk.length > 0) {
                    totalTasks += chk.length;
                    chk.forEach(c => {
                        if(c.selesai && c.tanggal) {
                            dateMap[c.tanggal] = (dateMap[c.tanggal] || 0) + 1;
                        }
                    });
                } else {
                    totalTasks += 1;
                    if(w.status === 'FINISH' && w.tanggal_finish) {
                        dateMap[w.tanggal_finish] = (dateMap[w.tanggal_finish] || 0) + 1;
                    }
                }
            });

            (fullData.actuators || []).forEach(a => {
                totalTasks += 2; // 2 checklists
                if(a.general_inspection && a.finish_date) dateMap[a.finish_date] = (dateMap[a.finish_date] || 0) + 1;
                if(a.function_test && a.finish_date) dateMap[a.finish_date] = (dateMap[a.finish_date] || 0) + 1;
            });

            const countInst = (list) => {
                (list || []).forEach(inst => {
                    totalTasks += 1;
                    const d = inst.tanggal || inst.finish_date || inst.dated;
                    if(inst.verifikasi && d) {
                        dateMap[d] = (dateMap[d] || 0) + 1;
                    }
                });
            };
            countInst(fullData.pressure_tx);
            countInst(fullData.temperature_tx);
            countInst(fullData.pressure_switch);

            // Sort unique dates
            const sortedDates = Object.keys(dateMap).sort((a, b) => {
                return (parseDateStrToTime(a) || 0) - (parseDateStrToTime(b) || 0);
            });

            if(sortedDates.length === 0) {
                chartBox.innerHTML = `
                <div style="padding:40px 20px; text-align:center; color:var(--text-muted);">
                    <div style="font-size:2rem; margin-bottom:8px;">📈</div>
                    Belum ada data tanggal selesai yang tercatat pada WO, Valve, atau Instrumen.<br>Centang sub-task atau instrumen untuk melihat grafik S-Curve.
                </div>`;
                return;
            }

            // Build cumulative progress
            let cumCount = 0;
            const dataPoints = sortedDates.map(d => {
                const cnt = dateMap[d];
                cumCount += cnt;
                const pct = totalTasks > 0 ? Math.min(100, Math.round((cumCount / totalTasks) * 1000) / 10) : 0;
                return { date: d, daily: cnt, cum: cumCount, pct: pct };
            });

            // SVG dimensions
            const svgW = 740;
            const svgH = 260;
            const padL = 50;
            const padR = 30;
            const padT = 30;
            const padB = 40;
            const graphW = svgW - padL - padR;
            const graphH = svgH - padT - padB;

            const n = dataPoints.length;
            const getX = (i) => padL + (n === 1 ? graphW / 2 : (i / (n - 1)) * graphW);
            const getY = (pct) => padT + graphH - (pct / 100) * graphH;

            let pathD = '';
            let targetPathD = '';
            dataPoints.forEach((pt, i) => {
                const x = getX(i);
                const y = getY(pt.pct);
                pathD += (i === 0 ? `M ${x} ${y}` : ` L ${x} ${y}`);
                const targetPct = Math.round(((i + 1) / n) * 100);
                const targetY = getY(targetPct);
                targetPathD += (i === 0 ? `M ${x} ${targetY}` : ` L ${x} ${targetY}`);
            });

            // Fill area
            const fillD = `${pathD} L ${getX(n - 1)} ${padT + graphH} L ${getX(0)} ${padT + graphH} Z`;

            let svgHTML = `
            <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; height:auto; background:var(--bg-sub); border-radius:var(--radius-md); border:1px solid var(--border-color);">
                <!-- Grid Lines -->
                ${[0, 25, 50, 75, 100].map(p => {
                    const y = getY(p);
                    return `
                    <line x1="${padL}" y1="${y}" x2="${svgW - padR}" y2="${y}" stroke="var(--border-color)" stroke-dasharray="4 4" stroke-width="1"/>
                    <text x="${padL - 8}" y="${y + 4}" fill="var(--text-muted)" font-size="10" text-anchor="end" font-family="'JetBrains Mono'">${p}%</text>`;
                }).join('')}

                <!-- Fill Area -->
                <path d="${fillD}" fill="url(#scurve-grad)" opacity="0.25"/>

                <!-- Target Line (Dashed) -->
                <path d="${targetPathD}" fill="none" stroke="var(--text-muted)" stroke-dasharray="5 5" stroke-width="2"/>

                <!-- Actual Curve Line -->
                <path d="${pathD}" fill="none" stroke="var(--primary)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>

                <!-- Points & Labels -->
                ${dataPoints.map((pt, i) => {
                    const x = getX(i);
                    const y = getY(pt.pct);
                    return `
                    <circle cx="${x}" cy="${y}" r="5" fill="var(--primary)" stroke="#fff" stroke-width="2"/>
                    <text x="${x}" y="${y - 10}" fill="var(--primary)" font-size="10" font-weight="bold" text-anchor="middle" font-family="'JetBrains Mono'">${pt.pct}%</text>
                    <text x="${x}" y="${svgH - 12}" fill="var(--text-muted)" font-size="9" text-anchor="middle" font-family="'JetBrains Mono'">${pt.date}</text>`;
                }).join('')}

                <defs>
                    <linearGradient id="scurve-grad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stop-color="var(--primary)" stop-opacity="0.8"/>
                        <stop offset="100%" stop-color="var(--primary)" stop-opacity="0.0"/>
                    </linearGradient>
                </defs>
            </svg>`;

            // Summary Breakdown Table
            let tableHTML = `
            <div style="margin-top:16px;">
                <div style="font-size:0.85rem; font-weight:700; color:var(--text-main); margin-bottom:8px;">📊 Rincian Capaian Per Hari:</div>
                <table class="dense-table" style="font-size:0.8rem;">
                    <thead>
                        <tr>
                            <th>Tanggal</th>
                            <th style="text-align:center;">Task Selesai Hari Ini</th>
                            <th style="text-align:center;">Kumulatif Selesai</th>
                            <th style="text-align:center;">Progress Kumulatif</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${dataPoints.map(pt => `
                            <tr>
                                <td style="font-family:'JetBrains Mono'; font-weight:700;">📅 ${pt.date}</td>
                                <td style="text-align:center; color:var(--status-finish); font-weight:700;">+${pt.daily} task</td>
                                <td style="text-align:center; font-family:'JetBrains Mono';">${pt.cum} / ${totalTasks}</td>
                                <td style="text-align:center; font-family:'JetBrains Mono'; font-weight:800; color:var(--primary);">${pt.pct}%</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>`;

            chartBox.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                <div style="font-size:0.85rem; color:var(--text-muted);">
                    Kurva-S Kumulatif Real-Time Outage Unit ${currentUnit}
                </div>
                <div style="display:flex; gap:14px; font-size:0.75rem;">
                    <span style="display:flex; align-items:center; gap:5px;"><span style="width:12px; height:3px; background:var(--primary); display:inline-block; border-radius:2px;"></span> Progress Aktual</span>
                    <span style="display:flex; align-items:center; gap:5px;"><span style="width:12px; height:2px; background:var(--text-muted); border-top:2px dashed var(--text-muted); display:inline-block;"></span> Target Rencana</span>
                </div>
            </div>
            ${svgHTML}
            ${tableHTML}`;
        }

        function generateWaText() {
            if(!fullData || !fullData.summary) return;
            const s = fullData.summary;
            const todayStr = getTodayFormatted();

            // Find completed tasks today
            const todayTasks = [];
            (fullData.work_orders || []).forEach(w => {
                (w.checklist || []).forEach(c => {
                    if(c.selesai && (c.tanggal === todayStr || !c.tanggal)) {
                        todayTasks.push(`[${w.no_wo}] Sub-task: ${c.sub_task} (PIC: ${c.pic_task || w.pic || '-'})`);
                    }
                });
            });
            (fullData.actuators || []).forEach(a => {
                if((a.general_inspection || a.function_test) && a.finish_date === todayStr) {
                    todayTasks.push(`[${a.equipment_id}] ${a.equipment_description} (${a.status})`);
                }
            });

            // Find open findings
            const openFindings = [];
            const collectF = (list, codeF, descF) => {
                (list || []).forEach(item => {
                    if(item.temuan) {
                        openFindings.push(`▪ [${item[codeF]}] ${item[descF] || ''}: ${item.temuan} (TL: ${item.tindak_lanjut || 'Proses verifikasi'})`);
                    }
                });
            };
            collectF(fullData.work_orders, 'no_wo', 'job_description');
            collectF(fullData.actuators, 'equipment_id', 'equipment_description');
            collectF(fullData.pressure_tx, 'kks', 'equipment');
            collectF(fullData.temperature_tx, 'kks', 'equipment');
            collectF(fullData.pressure_switch, 'kks', 'equipment');

            let msg = `*⚡ LAPORAN PROGRESS OUTAGE EIC - UNIT ${currentUnit}*\n`;
            msg += `🏭 *PLTU MSW &bull; SECTION EIC*\n`;
            msg += `📅 *Tanggal:* ${todayStr}\n\n`;

            msg += `📊 *RINGKASAN PROGRESS:*\n`;
            msg += `▪ *Grand Progress:* *${s.grand_pct}%* (${s.grand_done}/${s.grand_total} Sub-task Selesai)\n`;
            msg += `▪ *Work Orders:* ${s.wo.pct}% (${s.wo.finish}/${s.wo.total} WO Finish &bull; ${s.wo.subtask_done}/${s.wo.subtask_total} Sub-task)\n`;
            msg += `▪ *Actuator Valves:* ${s.actuator.pct}% (${s.actuator.finish}/${s.actuator.total} Valve Finish)\n`;
            msg += `▪ *Instruments:* ${s.instrument.pct}% (${s.instrument.done}/${s.instrument.total} Verifikasi OK)\n\n`;

            msg += `✅ *UPDATE PEKERJAAN TERKINI:* (${todayTasks.length} Task)\n`;
            if(todayTasks.length > 0) {
                todayTasks.slice(0, 10).forEach((t, i) => {
                    msg += `${i+1}. ${t}\n`;
                });
                if(todayTasks.length > 10) msg += `... dan ${todayTasks.length - 10} item lainnya.\n`;
            } else {
                msg += `_Belum ada task yang diselesaikan pada tanggal ${todayStr}_\n`;
            }
            msg += `\n`;

            msg += `⚠️ *REKAP TEMUAN / ACTIVE FINDINGS:* (${openFindings.length} Temuan)\n`;
            if(openFindings.length > 0) {
                openFindings.forEach(f => {
                    msg += `${f}\n`;
                });
            } else {
                msg += `_Nihil (Kondisi peralatan normal tanpa temuan terbuka)_\n`;
            }
            msg += `\n_Laporan diperbarui otomatis dari EIC Monitoring System PLTU MSW_`;

            const waBox = document.getElementById('wa-text-box');
            if(waBox) waBox.value = msg;
        }

        async function copyWaText() {
            const waBox = document.getElementById('wa-text-box');
            if(!waBox) return;
            try {
                if(navigator.clipboard && window.isSecureContext) {
                    await navigator.clipboard.writeText(waBox.value);
                } else {
                    waBox.select();
                    document.execCommand('copy');
                }
                showToast('✓ Format laporan WhatsApp berhasil disalin ke clipboard!', 'success', 2500);
            } catch(e) {
                waBox.select();
                document.execCommand('copy');
                showToast('✓ Format laporan disalin!', 'success');
            }
        }

        function initTheme() {
            const saved = localStorage.getItem('eic_theme') || 'dark';
            setTheme(saved);
        }

        function toggleTheme() {
            const current = document.documentElement.getAttribute('data-theme') || 'dark';
            const next = current === 'dark' ? 'light' : 'dark';
            setTheme(next);
        }

        function setTheme(theme) {
            document.documentElement.setAttribute('data-theme', theme);
            localStorage.setItem('eic_theme', theme);
            const icon = document.getElementById('theme-icon');
            const text = document.getElementById('theme-text');
            if(icon && text) {
                if(theme === 'light') {
                    icon.innerText = '☀️';
                    text.innerText = 'Light Mode';
                } else {
                    icon.innerText = '🌙';
                    text.innerText = 'Dark Mode';
                }
            }
        }

        window.onload = function() {
            initTheme();
            loadData();
        };
    