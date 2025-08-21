(function(){
  const fileInput = document.getElementById('file-input');
  const fileInfo = document.getElementById('file-info');
  const sheetSelect = document.getElementById('sheet-select');
  const loadMappingBtn = document.getElementById('load-mapping');
  const mappingContainer = document.getElementById('mapping-container');
  const previewBtn = document.getElementById('preview');
  const generateBtn = document.getElementById('generate');
  const previewOutput = document.getElementById('preview-output');
  const filenamePatternInput = document.getElementById('filename-pattern');
  const filenamesPreview = document.getElementById('filenames-preview');
  const filenamePrefixInput = document.getElementById('filename-prefix');
  const filenameHeaderSelect = document.getElementById('filename-header');

  // wire custom file-picker button (if present) to the hidden file input
  const filePickerBtn = document.getElementById('file-picker');
  if (filePickerBtn) {
    filePickerBtn.addEventListener('click', (e) => { e.preventDefault(); if (fileInput) fileInput.click(); });
  }

  let workbook = null;
  let currentFile = null;
  let mappingSheetName = null;
  let mapping = []; // array of row objects
  let tableHeaders = []; // header names
  let mappingTargets = {}; // { header: [ {sheet, addr}, ... ] }
  let headerKeyMap = {}; // normalized -> original header name

  // Output folder support (File System Access API)
  let outputDirHandle = null;

  async function pickOutputDir() {
    if (!window.showDirectoryPicker) return null;
    try {
      const dir = await window.showDirectoryPicker();
      outputDirHandle = dir; // store for session
      window._outputDirHandle = dir;
      renderOutputDirUI();
      return dir;
    } catch (err) {
      console.warn('Directory picker cancelled or failed', err);
      // Provide a helpful message for users who try to pick system folders (Documents/Downloads)
      try {
        const name = err && err.name ? err.name : null;
        // Many browsers/OS combinations will block access to system folders; show guidance
        alert('Could not open the selected folder. Some system-managed folders (e.g. root, Documents, Downloads) may be blocked by the browser/OS.\n\nPlease create or choose a regular subfolder (for example "SpreadsheetOutputs" inside Documents) and pick that instead.\n\nIf you need automatic saving into the template folder, consider using the desktop (Electron) build.\n\nError: ' + (name || String(err)));
      } catch (e) {
        /* ignore alert failures */
      }
      return null;
    }
  }

  async function saveFileToDir(dirHandle, filename, arrayBuffer) {
    try {
      const fh = await dirHandle.getFileHandle(filename, { create: true });
      const writable = await fh.createWritable();
      // write accepts ArrayBuffer or Blob
      await writable.write(arrayBuffer);
      await writable.close();
      return true;
    } catch (err) {
      console.warn('Failed to write file to directory', err);
      try {
        const name = err && err.name ? err.name : null;
        if (name === 'NotAllowedError' || name === 'SecurityError' || name === 'InvalidModificationError') {
          alert('Failed to save into the chosen folder. The browser or OS blocked write access to that folder.\n\nPlease choose a different folder (create a non-system subfolder) or use the desktop build to save files automatically.');
        }
      } catch (e) { /* ignore */ }
      return false;
    }
  }

  function renderOutputDirUI() {
    // Try to inject UI next to filename controls if present
    try {
      // only create once
      if (document.getElementById('choose-output-wrap')) return;
      const wrapTarget = filenameHeaderSelect ? filenameHeaderSelect.parentNode : null;
      const wrap = document.createElement('div'); wrap.id = 'choose-output-wrap'; wrap.style.marginTop = '8px';
      const label = document.createElement('div'); label.textContent = 'Output folder:'; label.className = 'small muted'; wrap.appendChild(label);
      const row = document.createElement('div'); row.style.display = 'flex'; row.style.gap = '8px'; row.style.alignItems = 'center';
      const btn = document.createElement('button'); btn.id = 'choose-output-btn'; btn.textContent = 'Choose output folder';
      const name = document.createElement('span'); name.id = 'output-dir-name'; name.style.fontSize = '0.9em'; name.style.color = '#ccc'; name.textContent = outputDirHandle ? (outputDirHandle.name || 'chosen folder') : 'none';
      const clear = document.createElement('button'); clear.id = 'clear-output-btn'; clear.textContent = 'Clear'; clear.style.display = outputDirHandle ? 'inline-block' : 'none';
      // small inline hint for guidance when folder selection fails or is blocked
      const hint = document.createElement('div'); hint.id = 'choose-output-hint'; hint.style.fontSize = '0.85em'; hint.style.color = '#bbb'; hint.style.marginTop = '6px'; hint.textContent = 'Tip: if Documents/Downloads are blocked, create a subfolder (e.g. "SpreadsheetOutputs") and choose that.';
      btn.addEventListener('click', async (e) => {
        e.preventDefault(); const d = await pickOutputDir(); if (d) name.textContent = d.name || 'selected folder'; clear.style.display = 'inline-block';
      });
      clear.addEventListener('click', (e) => { e.preventDefault(); outputDirHandle = null; window._outputDirHandle = null; document.getElementById('output-dir-name').textContent = 'none'; clear.style.display = 'none'; });
      row.appendChild(btn); row.appendChild(name); row.appendChild(clear); wrap.appendChild(row);
      wrap.appendChild(hint);
      if (wrapTarget && wrapTarget.parentNode) {
        wrapTarget.parentNode.insertBefore(wrap, wrapTarget.nextSibling);
      } else {
        // fallback: append to mappingContainer
        mappingContainer.insertBefore(wrap, mappingContainer.firstChild);
      }
    } catch (e) { console.warn('Could not render output dir UI', e); }
  }

  // attempt to render the output UI now (if DOM elements already exist)
  renderOutputDirUI();

  function reset() {
    sheetSelect.innerHTML = '';
    mappingContainer.innerHTML = '';
    previewOutput.textContent = '';
    workbook = null;
    currentFile = null;
    mappingSheetName = null;
    mapping = [];
    tableHeaders = [];
    mappingTargets = {};
    headerKeyMap = {};
    if (filenamesPreview) filenamesPreview.innerHTML = '';
  }

  fileInput.addEventListener('change', async (e) => {
    reset();
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    currentFile = f;
    fileInfo.textContent = `${f.name} — ${f.size} bytes`;
    const data = await f.arrayBuffer();
    try {
      workbook = XLSX.read(data, { type: 'array' });
    } catch (err) {
      fileInfo.textContent = 'Failed to read workbook: ' + err.message;
      return;
    }
    // populate sheets
    for (const name of workbook.SheetNames) {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sheetSelect.appendChild(opt);
    }
  });

  loadMappingBtn.addEventListener('click', () => {
    mappingContainer.innerHTML = '';
    previewOutput.textContent = '';
    if (filenamesPreview) filenamesPreview.innerHTML = '';
    if (!workbook) { alert('Please choose a master template first'); return; }
    const name = sheetSelect.value;
    if (!name) { alert('Select a replacement sheet'); return; }
    mappingSheetName = name;
    const sheet = workbook.Sheets[name];
    if (!sheet) { alert('Selected sheet not found'); return; }

    const parsed = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    mapping = [];
    tableHeaders = [];
    mappingTargets = {};
    headerKeyMap = {};
    // find header row
    let headerRowIndex = -1;
    for (let r = 0; r < parsed.length; r++) {
      const row = parsed[r];
      if (!row) continue;
      const nonEmpty = row.some(c => (c !== undefined && c !== null && String(c).trim() !== ''));
      if (nonEmpty) { headerRowIndex = r; break; }
    }
    if (headerRowIndex === -1) { alert('Mapping sheet appears empty'); return; }
    tableHeaders = (parsed[headerRowIndex] || []).map(h => h ? String(h).trim() : '');
    for (const h of tableHeaders) {
      mappingTargets[h] = [];
      const norm = String(h || '').trim().toLowerCase();
      if (norm) headerKeyMap[norm] = h;
    }
    // populate filename header select
    if (filenameHeaderSelect) {
      filenameHeaderSelect.innerHTML = '<option value="">Choose column for name...</option>' + tableHeaders.map(h => `<option value="${encodeURIComponent(h)}">${h}</option>`).join('');
    }
    for (let r = headerRowIndex + 1; r < parsed.length; r++) {
      const row = parsed[r];
      if (!row) continue;
      const allEmpty = row.every(c => c === undefined || c === null || String(c).trim() === '');
      if (allEmpty) continue;
      const obj = {};
      for (let c = 0; c < tableHeaders.length; c++) {
        const h = tableHeaders[c] || `Col${c+1}`;
        obj[h] = row[c] !== undefined && row[c] !== null ? String(row[c]) : '';
      }
      mapping.push(obj);
    }

    renderMappingEditor();
    // removed automatic filename preview here; user requested filenames only show with the Preview button
    // renderFilenamesPreview();
  });

  function renderMappingEditor() {
    mappingContainer.innerHTML = '';
    const info = document.createElement('div');
    info.textContent = `Table mode — ${mapping.length} rows loaded. Headers: ${tableHeaders.join(', ')}`;
    mappingContainer.appendChild(info);

    // table preview (first 20 rows)
    const tbl = document.createElement('table'); tbl.className = 'mapping-table';
    const thead = document.createElement('thead'); thead.innerHTML = '<tr>' + tableHeaders.map(h => `<th>${h}</th>`).join('') + '</tr>';
    tbl.appendChild(thead);
    const tbody = document.createElement('tbody');
    mapping.slice(0,3).forEach(rowObj => {
      const tr = document.createElement('tr');
      for (const h of tableHeaders) {
        const td = document.createElement('td');
        const inp = document.createElement('input'); inp.value = rowObj[h] || '';
        inp.addEventListener('input', () => { rowObj[h] = inp.value; /* do not auto-update filenames here */ }); td.appendChild(inp); tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
    mappingContainer.appendChild(tbl);

    // mapping targets editor
    const targetsTitle = document.createElement('h4'); targetsTitle.textContent = 'Map headers to target cells'; mappingContainer.appendChild(targetsTitle);
    const targetsWrap = document.createElement('div'); targetsWrap.className = 'targets-wrap';
    for (const h of tableHeaders) {
      const box = document.createElement('div'); box.className = 'target-box';
      const label = document.createElement('div'); label.innerHTML = `<strong>${h}</strong>`; box.appendChild(label);
      const list = document.createElement('div'); list.className = 'target-list';
      const renderList = () => { list.innerHTML = ''; (mappingTargets[h]||[]).forEach((tgt, idx) => {
        const row = document.createElement('div'); row.className = 'target-row';
        const sheetSel = document.createElement('select');
        // do not include the mapping (replacement) sheet as a target option
        for (const s of workbook.SheetNames) {
          if (s === mappingSheetName) continue;
          const opt = document.createElement('option'); opt.value = s; opt.textContent = s; if (tgt.sheet===s) opt.selected=true; sheetSel.appendChild(opt);
        }
        sheetSel.addEventListener('change', () => { mappingTargets[h][idx].sheet = sheetSel.value; });
        const addrInp = document.createElement('input'); addrInp.placeholder = 'e.g. B1'; addrInp.value = tgt.addr||''; addrInp.addEventListener('input', () => { mappingTargets[h][idx].addr = addrInp.value.toUpperCase(); });
        const rm = document.createElement('button'); rm.textContent = 'Remove'; rm.addEventListener('click', () => { mappingTargets[h].splice(idx,1); renderList(); });
        row.appendChild(sheetSel); row.appendChild(addrInp); row.appendChild(rm); list.appendChild(row);
      }); };
      const addT = document.createElement('button'); addT.textContent = 'Add target'; addT.addEventListener('click', () => {
        // choose a default sheet that is not the mapping sheet
        let defaultSheet = '';
        if (Array.isArray(workbook.SheetNames)) defaultSheet = workbook.SheetNames.find(s => s !== mappingSheetName) || (workbook.SheetNames[0]||'');
        mappingTargets[h].push({ sheet: defaultSheet, addr: '' }); renderList();
      });
      box.appendChild(list); box.appendChild(addT); targetsWrap.appendChild(box); renderList();
    }
    mappingContainer.appendChild(targetsWrap);
  }

  // modified: return a DOM fragment containing the filename preview instead of writing it automatically
  function renderFilenamesPreview() {
    // if neither previewOutput nor filenamesPreview exist, nothing to return
    if (!previewOutput && !filenamesPreview) return null;
    if (!Array.isArray(mapping) || mapping.length === 0) {
      const msg = document.createElement('div'); msg.textContent = 'No rows to preview.'; return msg;
    }
    const pattern = (typeof filenamePatternInput !== 'undefined' && filenamePatternInput) ? (filenamePatternInput.value || '') : '';
    const names = mapping.map((row, idx) => buildFilename(pattern, row, idx));

    const outer = document.createElement('div');
    const title = document.createElement('div');
    title.textContent = `Previewing ${names.length} filenames:`;
    outer.appendChild(title);

    const wrapper = document.createElement('div');
    wrapper.className = 'preview-list';
    const list = document.createElement('ol');
    list.style.margin = '8px 0 0 18px';
    names.forEach((n) => { const li = document.createElement('li'); li.textContent = n; list.appendChild(li); });
    wrapper.appendChild(list);
    outer.appendChild(wrapper);
    return outer;
  }

  function getNamingInputs() {
    const pattern = (typeof filenamePatternInput !== 'undefined' && filenamePatternInput) ? (filenamePatternInput.value || '') : '';
    const prefix = (filenamePrefixInput && filenamePrefixInput.value) ? String(filenamePrefixInput.value).trim() : '';
    const headerChoice = (filenameHeaderSelect && filenameHeaderSelect.value) ? decodeURIComponent(filenameHeaderSelect.value) : '';
    return { pattern, prefix, headerChoice };
  }

  function validateNamingInputs() {
    const { pattern, prefix, headerChoice } = getNamingInputs();
    if (pattern) return true; // pattern present is OK
    if (prefix && headerChoice) return true; // prefix+header OK
    // otherwise prompt user and focus the missing control
    alert('Provide an output filename pattern OR a filename prefix and a column header to use for naming.');
    if (!prefix && filenamePrefixInput) {
      filenamePrefixInput.focus();
    } else if (!headerChoice && filenameHeaderSelect) {
      filenameHeaderSelect.focus();
    }
    return false;
  }

  // helper: sanitize a string for use as a filename
  function sanitizeFilename(s) {
    return String(s || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
  }

  // Build an output filename from pattern or prefix/header choice
  function buildFilename(pattern, rowObj, index) {
    const fallbackBase = (currentFile && currentFile.name) ? String(currentFile.name).replace(/\.(xlsx?|xlsm?|xls)$/i, '') : 'output';
    const prefix = (filenamePrefixInput && filenamePrefixInput.value) ? String(filenamePrefixInput.value).trim() : '';
    const headerChoice = (filenameHeaderSelect && filenameHeaderSelect.value) ? decodeURIComponent(filenameHeaderSelect.value) : '';

    let filled = '';
    if (prefix && headerChoice && rowObj && Object.prototype.hasOwnProperty.call(rowObj, headerChoice)) {
      filled = `${prefix}_${sanitizeFilename(String(rowObj[headerChoice]))}`;
    } else if (pattern) {
      // tolerant placeholder replacement
      filled = String(pattern).replace(/\{([^}]+)\}/g, (m, p1) => {
        const raw = p1.trim();
        const norm = raw.toLowerCase();
        let val = undefined;
        if (rowObj && Object.prototype.hasOwnProperty.call(rowObj, raw)) val = rowObj[raw];
        else if (rowObj && headerKeyMap[norm] && Object.prototype.hasOwnProperty.call(rowObj, headerKeyMap[norm])) val = rowObj[headerKeyMap[norm]];
        if (val === undefined || val === null) return '';
        return sanitizeFilename(String(val));
      });
    }

    filled = String(filled || '').replace(/__+/g, '_').replace(/^[_\-.\s]+|[_\-.\s]+$/g, '');
    if (!filled) filled = `${fallbackBase}_row${index+1}`;

    // ensure extension
    let ext = '.xlsx';
    if (currentFile && currentFile.name) {
      const m = String(currentFile.name).match(/(\.[a-z0-9]+)$/i);
      if (m) ext = m[1];
    }
    if (!/\.[a-z0-9]+$/i.test(filled)) filled = filled + ext;
    return filled;
  }

  // helper: build a list that is scrollable (CSS controls max-height)
  function createScrollList(names) {
    const container = document.createElement('div');
    const title = document.createElement('div');
    title.textContent = `Previewing ${names.length} filenames:`;
    container.appendChild(title);

    const wrapper = document.createElement('div');
    wrapper.className = 'preview-list';
    const list = document.createElement('ol');
    list.style.margin = '8px 0 0 18px';
    names.forEach(n => { const li = document.createElement('li'); li.textContent = n; list.appendChild(li); });
    wrapper.appendChild(list);
    container.appendChild(wrapper);
    return container;
  }

  // helper: return the current replace mode chosen in the UI (radio inputs named "replace-mode").
  // Falls back to 'conditional' when the control is missing.
  function getReplaceMode() {
    try {
      const sel = document.querySelector('input[name="replace-mode"]:checked');
      return (sel && sel.value) ? sel.value : 'conditional';
    } catch (e) {
      return 'conditional';
    }
  }

  // helper: validate a simple A1-style cell address (e.g. A1, B12, AA100)
  function isValidCellAddress(addr) {
    if (!addr || typeof addr !== 'string') return false;
    // strip absolute markers like $A$1
    const raw = addr.replace(/\$/g, '').trim().toUpperCase();
    // basic pattern: 1-3 letters followed by row number (no leading zero)
    if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(raw)) return false;
    try {
      // also try decode_cell from SheetJS to be safe
      if (typeof XLSX !== 'undefined' && XLSX.utils && typeof XLSX.utils.decode_cell === 'function') {
        XLSX.utils.decode_cell(raw);
      }
      return true;
    } catch (e) {
      return false;
    }
  }

  // Preview button: validate state and show filenames in the small preview area and the previewOutput
  previewBtn.addEventListener('click', () => {
    // require workbook/mapping loaded
    if (!workbook || !currentFile || !mappingSheetName) { alert('Load workbook and mapping first'); return; }
    if (!Array.isArray(mapping) || mapping.length === 0) { alert('No rows found in mapping sheet'); return; }
    // require naming inputs (pattern OR prefix+header)
    if (!validateNamingInputs()) return;

    // Build filenames using current naming inputs
    const pattern = (typeof filenamePatternInput !== 'undefined' && filenamePatternInput) ? (filenamePatternInput.value || '') : '';
    const names = mapping.map((row, idx) => buildFilename(pattern, row, idx));

    // Clear right-side preview output (we show preview in the left compact panel only)
    if (previewOutput) previewOutput.innerHTML = '';

    // Render mapping summary + scrollable filename list into the left filenamesPreview panel
    if (filenamesPreview) {
      filenamesPreview.innerHTML = '';
      const summary = document.createElement('div');
      summary.textContent = `Rows: ${mapping.length}`;
      summary.className = 'muted small';
      filenamesPreview.appendChild(summary);

      const scrollCompact = createScrollList(names);
      // compact spacing for the left panel
      scrollCompact.querySelectorAll('li').forEach(li => li.style.padding = '4px 0');
      filenamesPreview.appendChild(scrollCompact);
    }
  });

  // ensure browser download helper exists
  async function browserDownloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename || 'output.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  generateBtn.addEventListener('click', async () => {
    if (!workbook || !currentFile || !mappingSheetName) { alert('Load workbook and mapping first'); return; }
    const mode = getReplaceMode();
    if (!Array.isArray(mapping) || mapping.length === 0) { alert('No rows found in mapping sheet'); return; }
    // safe access: guard against missing DOM nodes
    const pattern = (filenamePatternInput && filenamePatternInput.value) ? filenamePatternInput.value : '';
    const prefix = (filenamePrefixInput && filenamePrefixInput.value) ? String(filenamePrefixInput.value).trim() : '';
    const headerChoice = (filenameHeaderSelect && filenameHeaderSelect.value) ? decodeURIComponent(filenameHeaderSelect.value) : '';
    if (!pattern && (!prefix || !headerChoice)) { alert('Provide an output filename pattern OR a filename prefix and a column header to use for naming.'); return; }

    // validate mapping targets
    for (const h of tableHeaders) {
      const tg = mappingTargets[h] || [];
      if (!tg || tg.length === 0) {
        if (!confirm(`Header "${h}" has no mapping targets. Continue without writing this field?`)) return;
      } else {
        for (const t of tg) {
          if (!workbook.SheetNames.includes(t.sheet)) { alert(`Invalid sheet ${t.sheet} for header ${h}`); return; }
          if (t.addr && !isValidCellAddress(t.addr)) { alert(`Invalid cell address ${t.addr} for header ${h}`); return; }
        }
      }
    }

    let count = 0;
    for (let i = 0; i < mapping.length; i++) {
      const rowObj = mapping[i];

      // Deep-clone the original workbook so we preserve workbook-level props, formulas, styles and other metadata
      let newWb = JSON.parse(JSON.stringify(workbook));
      // Remove the mapping (replacement) sheet from the clone so outputs don't include it
      if (Array.isArray(newWb.SheetNames)) {
        newWb.SheetNames = newWb.SheetNames.filter(n => n !== mappingSheetName);
      }
      if (newWb.Sheets && newWb.Sheets[mappingSheetName]) delete newWb.Sheets[mappingSheetName];

      // apply mappings into the cloned sheets in-place
      for (const name of newWb.SheetNames) {
        const newSheet = newWb.Sheets[name]; if (!newSheet) continue;
        for (const h of tableHeaders) {
          const targets = mappingTargets[h] || [];
          const rawVal = rowObj[h] !== undefined ? rowObj[h] : '';
          const val = (rawVal === null || rawVal === undefined) ? '' : String(rawVal);
          if (mode === 'conditional' && (val === undefined || val === null || val === '')) continue;
          for (const t of targets) {
            const targetSheet = (t.sheet || '').trim();
            if (targetSheet !== name && targetSheet !== (name||'').trim()) { continue; }
            const addr = (t.addr || '').toUpperCase().trim();
            if (!addr) { console.warn('Skipping mapping target with empty address', { header: h, sheet: targetSheet }); continue; }
            if (!isValidCellAddress(addr)) { console.warn('Skipping mapping target with invalid address', { header: h, addr }); continue; }
            newSheet[addr] = newSheet[addr] || {};
            // set value and type; overwrite cell value but preserve other metadata where possible
            const num = Number(val);
            if (val !== '' && !isNaN(num) && isFinite(num)) {
              newSheet[addr].v = num; newSheet[addr].t = 'n';
            } else {
              newSheet[addr].v = val; newSheet[addr].t = 's';
            }
            // remove formula only when we intentionally replace a formula cell
            if (newSheet[addr].f) delete newSheet[addr].f;
          }
        }

        // Try to update sheet range (!ref) so viewers show newly written cells
        try {
          let minR = Infinity, minC = Infinity, maxR = -Infinity, maxC = -Infinity;
          Object.keys(newSheet).forEach(cell => {
            if (cell[0] === '!') return;
            try { const rc = XLSX.utils.decode_cell(cell); if (rc.r < minR) minR = rc.r; if (rc.r > maxR) maxR = rc.r; if (rc.c < minC) minC = rc.c; if (rc.c > maxC) maxC = rc.c; } catch (e) { /* ignore */ }
          });
          if (minR !== Infinity && minC !== Infinity && maxR >= 0 && maxC >= 0) {
            newSheet['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
          }
        } catch (err) { console.warn('Failed to update sheet range', err); }
      }

      // Ask SheetJS to write cell styles if possible so formatting is preserved
      const wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
       const blob = new Blob([wbout], { type: 'application/octet-stream' });
       const suggested = buildFilename(pattern, rowObj, i);
       console.log('Saving file:', suggested);
      // If running in Electron and we have the original template full path, write
      // outputs into the same folder as the template. Otherwise check outputDirHandle from picker or fallback to browser download.
      const dirHandle = outputDirHandle || window._outputDirHandle || null;
      if (dirHandle && typeof dirHandle.getFileHandle === 'function') {
        try {
          // write using File System Access API
          const arr = new Uint8Array(wbout);
          const ok = await saveFileToDir(dirHandle, suggested, arr);
          if (!ok) {
            // fallback to electron or download
            if (window.electron && currentFile && currentFile.path) {
              try {
                const fullPath = String(currentFile.path).replace(/\\/g, '/');
                const dir = fullPath.replace(/\/[^\/]*$/, '').replace(/\/$/, '');
                const outPath = dir + '/' + suggested;
                const arr2 = new Uint8Array(wbout);
                const writeRes = await window.electron.writeFile(outPath, arr2);
                if (!writeRes || !writeRes.ok) console.warn('Electron writeFile failed', writeRes);
              } catch (err) {
                console.warn('Failed to write via Electron, falling back to download', err);
                await browserDownloadBlob(blob, suggested);
              }
            } else {
              await browserDownloadBlob(blob, suggested);
            }
          }
        } catch (err) {
          console.warn('Failed to save via File System Access API', err);
          // fallback to electron or download
          if (window.electron && currentFile && currentFile.path) {
            try {
              const fullPath = String(currentFile.path).replace(/\\/g, '/');
              const dir = fullPath.replace(/\/[^\/]*$/, '').replace(/\/$/, '');
              const outPath = dir + '/' + suggested;
              const arr2 = new Uint8Array(wbout);
              const writeRes = await window.electron.writeFile(outPath, arr2);
              if (!writeRes || !writeRes.ok) console.warn('Electron writeFile failed', writeRes);
            } catch (err2) {
              console.warn('Failed to write via Electron, falling back to download', err2);
              await browserDownloadBlob(blob, suggested);
            }
          } else {
            await browserDownloadBlob(blob, suggested);
          }
        }
      } else if (window.electron && currentFile && currentFile.path) {
        try {
          // derive folder by trimming the filename from the path
          const fullPath = String(currentFile.path).replace(/\\/g, '/');
          const dir = fullPath.replace(/\/[^^\/]*$/, '').replace(/\/$/, '');
          const outPath = dir + '/' + suggested;
          const arr = new Uint8Array(wbout);
          const writeRes = await window.electron.writeFile(outPath, arr);
          if (!writeRes || !writeRes.ok) console.warn('Electron writeFile failed', writeRes);
        } catch (err) {
          console.warn('Failed to write via Electron, falling back to download', err);
          await browserDownloadBlob(blob, suggested);
        }
      } else {
        await browserDownloadBlob(blob, suggested);
      }
      count++; await new Promise(r => setTimeout(r, 80));
    }
    alert(`Generated ${count} files and triggered downloads/saves.`);
  });

})();
