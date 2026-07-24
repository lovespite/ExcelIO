(function () {
  'use strict';

  const $ = (s) => document.querySelector(s);

  const grid = $('#grid');
  const gridScroll = $('#grid-scroll');
  const sheetTabs = $('#sheet-tabs');
  const filenameEl = $('#filename');
  const errorEl = $('#error-msg');
  const btnSave = $('#btn-save');
  const fileInput = $('#file-input');

  let selRow = -1, selCol = -1;
  let editRow = -1, editCol = -1;
  let editOrigValue = '';

  // ── Bootstrap: wait for WASM ──

  function waitForExcelIO(cb) {
    if (window.excelIO) { cb(); return; }
    let ticks = 0;
    const timer = setInterval(() => {
      if (window.excelIO) { clearInterval(timer); cb(); return; }
      if (++ticks > 100) { clearInterval(timer); showError('WASM failed to load'); }
    }, 100);
  }

  // ── Toolbar events ──

  $('#btn-new').addEventListener('click', () => {
    window.excelIO.newWorkbook();
    refreshAll();
  });

  btnSave.addEventListener('click', () => {
    const bytes = window.excelIO.saveToBytes();
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = window.excelIO.getFileName();
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  fileInput.addEventListener('change', () => {
    const file = fileInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const bytes = new Uint8Array(reader.result);
      const ok = window.excelIO.loadFromBytes(bytes, file.name);
      if (ok) {
        errorEl.textContent = '';
        refreshAll();
      } else {
        showError('Failed to load file.');
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // ── Grid keyboard ──

  gridScroll.addEventListener('keydown', (e) => {
    if (editRow >= 0) return; // editing input handles its own keys

    if (selRow < 0 || selCol < 0) return;

    switch (e.key) {
      case 'ArrowUp':    moveSelection(selRow - 1, selCol); e.preventDefault(); break;
      case 'ArrowDown':  moveSelection(selRow + 1, selCol); e.preventDefault(); break;
      case 'ArrowLeft':  moveSelection(selRow, selCol - 1); e.preventDefault(); break;
      case 'ArrowRight': moveSelection(selRow, selCol + 1); e.preventDefault(); break;
      case 'Enter':      startEdit(selRow, selCol); break;
      case 'Tab':
        e.preventDefault();
        moveSelection(selRow, selCol + 1);
        break;
      case 'Delete':
        window.excelIO.setCellValue(selRow, selCol, '');
        reRenderCell(selRow, selCol);
        break;
    }
  });

  function moveSelection(row, col) {
    const rc = window.excelIO.getRowCount();
    const cc = window.excelIO.getColCount() || 26;
    if (row < 0 || col < 0) return;
    selRow = row;
    selCol = col;
    ensureVisible(selRow, selCol);
    renderGrid();
    gridScroll.focus();
  }

  function ensureVisible(row, col) {
    const th = $('#grid thead tr');
    if (!th) return;
    const cell = th.children[col + 1]; // +1 for corner cell
    if (cell) cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  // ── Inline editing ──

  function startEdit(row, col) {
    selRow = row;
    selCol = col;
    editRow = row;
    editCol = col;
    editOrigValue = window.excelIO.getCellValue(row, col);
    renderGrid();
    // Focus the input after render
    requestAnimationFrame(() => {
      const inp = document.querySelector('.data-cell.editing input');
      if (inp) {
        inp.focus();
        inp.select();
      }
    });
  }

  function commitEdit() {
    if (editRow < 0) return;
    const inp = document.querySelector('.data-cell.editing input');
    const val = inp ? inp.value : editOrigValue;
    window.excelIO.setCellValue(editRow, editCol, val);
    selRow = editRow;
    selCol = editCol;
    editRow = editCol = -1;
    reRenderCell(selRow, selCol);
  }

  function cancelEdit() {
    editRow = editCol = -1;
    reRenderCell(selRow, selCol);
  }

  function reRenderCell(row, col) {
    // Re-render just one cell in the DOM
    const tbody = $('#grid tbody');
    if (!tbody || row >= tbody.children.length) { renderGrid(); return; }
    const tr = tbody.children[row];
    if (!tr || col + 1 >= tr.children.length) { renderGrid(); return; }
    const td = tr.children[col + 1]; // +1 for row-header
    const val = window.excelIO.getCellValue(row, col);
    td.textContent = val || '';
    td.style.cssText = buildCellStyle(row, col);
    td.className = 'data-cell' + (row === selRow && col === selCol ? ' selected' : '');
  }

  function handleEditKeydown(e) {
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        commitEdit();
        moveSelection(selRow + 1, selCol);
        startEdit(selRow, selCol);
        break;
      case 'Tab':
        e.preventDefault();
        commitEdit();
        moveSelection(selRow, selCol + 1);
        startEdit(selRow, selCol);
        break;
      case 'Escape':
        e.preventDefault();
        cancelEdit();
        gridScroll.focus();
        break;
      case 'ArrowUp':
        e.preventDefault();
        commitEdit();
        moveSelection(Math.max(0, selRow - 1), selCol);
        startEdit(selRow, selCol);
        break;
      case 'ArrowDown':
        e.preventDefault();
        commitEdit();
        moveSelection(selRow + 1, selCol);
        startEdit(selRow, selCol);
        break;
      case 'ArrowLeft': {
        const inp = document.querySelector('.data-cell.editing input');
        if (inp && inp.selectionStart === 0 && inp.selectionEnd === 0) {
          e.preventDefault();
          commitEdit();
          moveSelection(selRow, Math.max(0, selCol - 1));
          startEdit(selRow, selCol);
        }
        break;
      }
      case 'ArrowRight': {
        const inp = document.querySelector('.data-cell.editing input');
        if (inp && inp.selectionStart === inp.value.length) {
          e.preventDefault();
          commitEdit();
          moveSelection(selRow, selCol + 1);
          startEdit(selRow, selCol);
        }
        break;
      }
    }
  }

  // ── Render ──

  function refreshAll() {
    selRow = selCol = -1;
    editRow = editCol = -1;
    const hasWb = window.excelIO.getSheetCount() > 0;
    btnSave.disabled = !hasWb;
    filenameEl.textContent = hasWb ? window.excelIO.getFileName() : '';
    renderSheetTabs();
    renderGrid();
    if (hasWb) gridScroll.focus();
  }

  function renderSheetTabs() {
    const count = window.excelIO.getSheetCount();
    const active = window.excelIO.getActiveSheetIndex();
    let html = '';
    for (let i = 0; i < count; i++) {
      const name = escHtml(window.excelIO.getSheetName(i));
      const cls = i === active ? 'sheet-tab active' : 'sheet-tab';
      html += `<button class="${cls}" data-index="${i}">${name}</button>`;
    }
    sheetTabs.innerHTML = html;

    sheetTabs.querySelectorAll('.sheet-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index);
        if (window.excelIO.switchSheet(idx)) refreshAll();
      });
    });
  }

  function renderGrid() {
    const rc = window.excelIO.getRowCount();
    const cc = window.excelIO.getColCount() || 26;
    const displayCols = Math.max(cc, 26);

    let html = '<thead><tr><th class="corner-cell"></th>';

    // Column headers
    for (let c = 0; c < displayCols; c++) {
      const colName = columnName(c);
      const width = window.excelIO.getColumnWidth(c);
      const hidden = window.excelIO.isColumnHidden(c);
      const style = width > 0 ? ` style="width:${Math.max(width * 7, 30)}px"` : '';
      const visStyle = hidden ? ' style="display:none"' : '';
      html += `<th class="col-header"${style}${visStyle}>${colName}</th>`;
    }
    html += '</tr></thead><tbody>';

    // Build merged cell map
    const mergedMap = buildMergeMap();

    // Rows
    for (let r = 0; r < rc; r++) {
      const rowHeight = window.excelIO.getRowHeight(r);
      const rowHidden = window.excelIO.isRowHidden(r);
      const rhStyle = rowHeight > 0 ? ` style="height:${rowHeight * 1.4}px"` : '';
      const rhVis = rowHidden ? ' style="display:none"' : '';
      html += `<tr${rhStyle}${rhVis}><td class="row-header">${r + 1}</td>`;

      for (let c = 0; c < displayCols; c++) {
        // Skip cells covered by a merge
        const key = `${r},${c}`;
        if (mergedMap.skip.has(key)) continue;

        const merge = mergedMap.cells[key];
        if (merge) {
          const val = escHtml(window.excelIO.getCellValue(r, c));
          const style = buildCellStyle(r, c);
          const cls = (r === selRow && c === selCol) ? 'data-cell selected' : 'data-cell';
          const rs = merge.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : '';
          const cs = merge.colspan > 1 ? ` colspan="${merge.colspan}"` : '';
          html += `<td class="${cls}" style="${style}"${rs}${cs}>${val}</td>`;
        } else if (r === editRow && c === editCol) {
          html += `<td class="data-cell editing"><input value="${escAttr(editOrigValue)}"></td>`;
        } else {
          const val = escHtml(window.excelIO.getCellValue(r, c));
          const style = buildCellStyle(r, c);
          const cls = (r === selRow && c === selCol) ? 'data-cell selected' : 'data-cell';
          html += `<td class="${cls}" style="${style}">${val}</td>`;
        }
      }
      html += '</tr>';
    }
    html += '</tbody>';

    grid.innerHTML = html;

    // Attach event handlers
    attachCellHandlers();
  }

  function buildMergeMap() {
    const map = { cells: {}, skip: new Set() };
    try {
      const json = window.excelIO.getMergedCells();
      const refs = JSON.parse(json);
      if (!Array.isArray(refs)) return map;

      for (const ref of refs) {
        const parsed = parseRange(ref);
        if (!parsed) continue;
        const { r1, c1, r2, c2 } = parsed;
        const rowspan = r2 - r1 + 1;
        const colspan = c2 - c1 + 1;
        map.cells[`${r1},${c1}`] = { rowspan, colspan };
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            if (r !== r1 || c !== c1) map.skip.add(`${r},${c}`);
          }
        }
      }
    } catch (_) { /* ignore */ }
    return map;
  }

  function parseRange(ref) {
    const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return null;
    return {
      c1: colIndex(m[1]),
      r1: parseInt(m[2]) - 1,
      c2: colIndex(m[3]),
      r2: parseInt(m[4]) - 1
    };
  }

  function buildCellStyle(row, col) {
    let json;
    try {
      json = JSON.parse(window.excelIO.getCellStyleJson(row, col));
    } catch (_) { return ''; }

    const parts = [];
    if (json.fontName)   parts.push(`font-family:'${json.fontName}'`);
    if (json.fontSize)   parts.push(`font-size:${json.fontSize}pt`);
    if (json.fontColor)  parts.push(`color:${json.fontColor}`);
    if (json.bold)       parts.push('font-weight:bold');
    if (json.italic)     parts.push('font-style:italic');
    if (json.fillColor)  parts.push(`background-color:${json.fillColor}`);
    if (json.hAlign)     parts.push(`text-align:${json.hAlign}`);
    if (json.vAlign)     parts.push(`vertical-align:${json.vAlign}`);
    if (json.wrapText)   parts.push('white-space:normal;word-wrap:break-word');

    // Borders
    const bd = [
      ['borderLeft', 'border-left'], ['borderRight', 'border-right'],
      ['borderTop', 'border-top'], ['borderBottom', 'border-bottom']
    ];
    for (const [key, cssProp] of bd) {
      if (json[key]) {
        parts.push(`${cssProp}:1px ${json[key]} ${json[key + 'Color'] || '#888'}`);
      }
    }

    return parts.join(';');
  }

  function attachCellHandlers() {
    const tbody = grid.querySelector('tbody');
    if (!tbody) return;

    const rows = tbody.querySelectorAll('tr');
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].querySelectorAll('td.data-cell');
      // Map TD back to (row, col) by counting preceding cells including merge spans
      let colOffset = 0;
      for (let tdIdx = 0; tdIdx < cells.length; tdIdx++) {
        const td = cells[tdIdx];
        const colspan = parseInt(td.getAttribute('colspan')) || 1;
        const rowspan = parseInt(td.getAttribute('rowspan')) || 1;

        const cellRow = r;
        const cellCol = colOffset;

        if (td.classList.contains('editing')) {
          const inp = td.querySelector('input');
          if (inp) {
            inp.addEventListener('keydown', handleEditKeydown);
            inp.addEventListener('blur', () => {
              // Delay to allow click on other cells to register first
              setTimeout(() => {
                if (editRow >= 0) commitEdit();
              }, 150);
            });
          }
        } else {
          td.addEventListener('click', (function (rr, cc) {
            return function () {
              selRow = rr; selCol = cc;
              editRow = editCol = -1;
              gridScroll.focus();
              renderGrid();
            };
          })(cellRow, cellCol));

          td.addEventListener('dblclick', (function (rr, cc) {
            return function () {
              startEdit(rr, cc);
            };
          })(cellRow, cellCol));
        }

        colOffset += colspan;
      }
    }
  }

  // ── Helpers ──

  function columnName(index) {
    let name = '';
    while (index >= 0) {
      name = String.fromCharCode(65 + (index % 26)) + name;
      index = Math.floor(index / 26) - 1;
    }
    return name;
  }

  function colIndex(name) {
    let idx = 0;
    for (let i = 0; i < name.length; i++) {
      idx = idx * 26 + (name.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  function escHtml(s) {
    if (!s) return '';
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;')
            .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function escAttr(s) {
    if (!s) return '';
    return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
            .replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function showError(msg) {
    errorEl.textContent = msg;
  }

  // ── Init ──

  waitForExcelIO(() => {
    // Show empty state or render an existing workbook
    if (window.excelIO.getSheetCount() > 0) {
      refreshAll();
    }
    gridScroll.setAttribute('tabindex', '0');
    gridScroll.addEventListener('click', () => gridScroll.focus());
  });
})();
