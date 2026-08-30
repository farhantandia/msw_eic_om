import re

with open("server.py", "r", encoding="utf-8") as f:
    text = f.read()

# 1. Add CSS for component badges and selector
css_patch = '''
        .badge-tag-comp {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            font-size: 0.68rem;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 4px;
            letter-spacing: 0.02em;
            font-family: 'JetBrains Mono', monospace;
            white-space: nowrap;
        }
        .badge-tag-act {
            background: rgba(245, 158, 11, 0.15);
            color: #f59e0b;
            border: 1px solid rgba(245, 158, 11, 0.35);
        }
        .badge-tag-inst {
            background: rgba(6, 182, 212, 0.15);
            color: #06b6d4;
            border: 1px solid rgba(6, 182, 212, 0.35);
        }
        .badge-tag-elec {
            background: rgba(139, 92, 246, 0.15);
            color: #a78bfa;
            border: 1px solid rgba(139, 92, 246, 0.35);
        }
        .comp-mode-btn {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.78rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .comp-mode-btn.active {
            background: var(--primary);
            color: #090d16;
            border-color: var(--primary);
            font-weight: 700;
        }
'''

if ".badge-tag-comp" not in text:
    text = text.replace("</style>", css_patch + "\n    </style>")

# 2. Update loadData() to fetch master_components
old_loaddata = '''                const res = await fetch(`/api/data?unit=${currentUnit}`);
                if(!res.ok) throw new Error(`HTTP ${res.status} Server Error`);
                fullData = await res.json();'''

new_loaddata = '''                const [res, compRes] = await Promise.all([
                    fetch(`/api/data?unit=${currentUnit}`),
                    fetch(`/api/master_components?unit=${currentUnit}`).catch(() => null)
                ]);
                if(!res.ok) throw new Error(`HTTP ${res.status} Server Error`);
                fullData = await res.json();
                if(compRes && compRes.ok) {
                    const compData = await compRes.json();
                    fullData.master_actuators = compData.actuators || [];
                    fullData.master_instruments = compData.instruments || [];
                }'''

if old_loaddata in text:
    text = text.replace(old_loaddata, new_loaddata)

# 3. Add helper to detect subtask type and component UI functions in JS
js_helpers = '''
        function getSubtaskTypeBadge(desc) {
            if(!desc) return '';
            const s = desc.toUpperCase();
            if(s.includes('ACTUATOR') || s.includes('MOV') || s.includes('AOV') || s.includes('GATE') || s.includes('DAMPER') || s.includes('IGV') || s.includes('VALVE')) {
                return '<span class="badge-tag-comp badge-tag-act" title="Tersinkron dengan Tab Actuator">⚡ Actuator</span>';
            } else if(s.includes('TRANSMITTER') || s.includes('SWITCH') || s.includes('CALIBRATION') || s.includes('MEASUREMENT') || s.includes('CP') || s.includes('CT') || s.includes('PT') || s.includes('TT') || s.includes('RTD')) {
                return '<span class="badge-tag-comp badge-tag-inst" title="Tersinkron dengan Tab Instrument">📟 Instrument</span>';
            } else if(s.includes('MOTOR') || s.includes('BEARING') || s.includes('WINDING') || s.includes('ISOLASI') || s.includes('SOLO RUN') || s.includes('HEATER') || s.includes('BREAKER')) {
                return '<span class="badge-tag-comp badge-tag-elec" title="Pekerjaan Motor / Elektrik">⚙️ Motor/Elec</span>';
            }
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
'''

if "function getSubtaskTypeBadge" not in text:
    text = text.replace("/* ---------------- QUICK ACTIONS (NO AUTO REFRESH) ---------------- */", js_helpers + "\n        /* ---------------- QUICK ACTIONS (NO AUTO REFRESH) ---------------- */")

# 4. Update Checklist Item rendering inside renderWorkOrders
old_chk_map = """                                    ${(item.checklist || []).map((c, cIdx) => `
                                        <div class="checklist-item ${c.selesai ? 'done' : ''}">
                                            <label style="cursor:pointer; display:flex; align-items:center; gap:8px;">
                                                <input type="checkbox" id="chk-${item.no_wo}-${cIdx}" ${c.selesai ? 'checked' : ''} onchange="toggleLocalSubtask('${item.no_wo}', ${cIdx}, this.checked)">
                                                <span style="${c.selesai?'text-decoration:line-through; color:var(--text-muted);':''}">${c.sub_task}</span>
                                            </label>
                                            <div style="display:flex; align-items:center; gap:8px;">
                                                ${c.tanggal ? `<span class="date-badge" style="font-size:0.75rem; color:var(--primary); font-family:'JetBrains Mono',monospace; background:var(--date-badge-bg); padding:2px 7px; border-radius:4px; border:1px solid var(--border-color);" title="Tanggal Dikerjakan">📅 ${c.tanggal}</span>` : ''}
                                                <button style="background:none; border:none; color:#f43f5e; cursor:pointer; font-size:0.85rem;" title="Hapus Subtask" onclick="deleteSubtask('${item.no_wo}', '${(c.sub_task||'').toString().replace(/'/g, "\\\\'")}')">🗑️</button>
                                            </div>
                                        </div>
                                    `).join('')}"""

new_chk_map = """                                    ${(item.checklist || []).map((c, cIdx) => `
                                        <div class="checklist-item ${c.selesai ? 'done' : ''}">
                                            <label style="cursor:pointer; display:flex; align-items:center; gap:8px; flex-grow:1;">
                                                <input type="checkbox" id="chk-${item.no_wo}-${cIdx}" ${c.selesai ? 'checked' : ''} onchange="toggleLocalSubtask('${item.no_wo}', ${cIdx}, this.checked)">
                                                <span style="${c.selesai?'text-decoration:line-through; color:var(--text-muted);':''}">${c.sub_task}</span>
                                                ${getSubtaskTypeBadge(c.sub_task)}
                                            </label>
                                            <div style="display:flex; align-items:center; gap:8px;">
                                                ${c.tanggal ? `<span class="date-badge" style="font-size:0.75rem; color:var(--primary); font-family:'JetBrains Mono',monospace; background:var(--date-badge-bg); padding:2px 7px; border-radius:4px; border:1px solid var(--border-color);" title="Tanggal Dikerjakan">📅 ${c.tanggal}</span>` : ''}
                                                <button style="background:none; border:none; color:#f43f5e; cursor:pointer; font-size:0.85rem;" title="Hapus Subtask" onclick="deleteSubtask('${item.no_wo}', '${(c.sub_task||'').toString().replace(/'/g, "\\\\'")}')">🗑️</button>
                                            </div>
                                        </div>
                                    `).join('')}"""

if old_chk_map in text:
    text = text.replace(old_chk_map, new_chk_map)

# 5. Update Add Subtask section with Component Dropdown Selector
old_add_subtask_box = """                                <div style="display:flex; gap:8px; margin-top:12px; max-width:550px;">
                                    <input type="text" id="new-subtask-${item.no_wo}" class="filter-input" placeholder="Tambah checklist sub-task baru..." style="flex-grow:1; font-size:0.82rem;">
                                    <button class="btn-save" style="padding:6px 14px; font-size:0.8rem;" onclick="addSubtask('${item.no_wo}')">➕ Tambah Sub-task</button>
                                </div>"""

new_add_subtask_box = """                                <div style="margin-top:14px; background:rgba(0,0,0,0.2); border:1px solid var(--border-color); border-radius:8px; padding:10px 12px;">
                                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:6px;">
                                        <span style="font-size:0.78rem; font-weight:700; color:var(--text-muted);">➕ Tambah Sub-task ke WO:</span>
                                        <div style="display:flex; gap:5px;">
                                            <button type="button" class="comp-mode-btn btn-mode-manual ${(subtaskAddModes[item.no_wo]||'manual')==='manual'?'active':''}" onclick="setSubtaskMode('${item.no_wo}', 'manual')">✏️ Manual</button>
                                            <button type="button" class="comp-mode-btn btn-mode-act ${subtaskAddModes[item.no_wo]==='actuator'?'active':''}" onclick="setSubtaskMode('${item.no_wo}', 'actuator')">⚡ Pilih Actuator (${(fullData.master_actuators||fullData.actuators||[]).length})</button>
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
                                        <button class="btn-save" style="padding:6px 14px; font-size:0.8rem; white-space:nowrap; background:#f59e0b; border-color:#d97706; color:#000;" onclick="addSubtask('${item.no_wo}', 'actuator')">⚡ Tambah Actuator</button>
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
                                </div>"""

if old_add_subtask_box in text:
    text = text.replace(old_add_subtask_box, new_add_subtask_box)

# 6. Update addSubtask JS function to support modes
old_add_func = """        async function addSubtask(noWo) {
            const input = document.getElementById(`new-subtask-${noWo}`);
            const subTask = (input.value || '').trim();
            if(!subTask) {
                showToast('Silakan masukkan deskripsi sub-task!', 'error');
                return;
            }

            try {
                const res = await fetch('/api/add_subtask', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({unit: currentUnit, no_wo: noWo, sub_task: subTask})
                });
                const result = await res.json();
                showToast(result.message || 'Sub-task berhasil ditambahkan!', 'success');
                input.value = '';
                loadData();
            } catch(e) {
                showToast('Gagal menambah sub-task', 'error');
            }
        }"""

new_add_func = """        async function addSubtask(noWo, mode = 'manual') {
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
        }"""

if old_add_func in text:
    text = text.replace(old_add_func, new_add_func)

with open("server.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Frontend component dropdown and badges successfully patched into server.py!")
