(function () {
    'use strict';

    const INITIAL_ROWS = 50;
    const INITIAL_COLS = 26;
    const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
    const MAX_ROWS = 10000;

    // ── Sheet state ──
    let sheets = [];
    let activeSheetIdx = 0;

    function currentSheet() { return sheets[activeSheetIdx]; }
    function data() { return currentSheet().data; }

    let activeCell = { row: 0, col: 0 };
    let selectionStart = null;
    let selectionEnd = null;
    let isSelecting = false;
    let isEditing = false;
    let formulaFocused = false;
    let saveTimer = null;

    // ── DOM refs ──
    const grid = document.getElementById('grid');
    const gridContainer = document.getElementById('grid-container');
    const formulaInput = document.getElementById('formula-input');
    const cellRef = document.getElementById('cell-ref');
    const fileInput = document.getElementById('file-input');
    const statusCells = document.getElementById('status-cells');
    const statusSelection = document.getElementById('status-selection');
    const sheetTabs = document.getElementById('sheet-tabs');
    const sheetAddBtn = document.getElementById('sheet-add-btn');
    const sheetScroll = document.getElementById('sheet-tabs-scroll');
    const toast = document.getElementById('toast');

    // ── Init ──
    function init() {
        const restored = loadFromCache();
        if (!restored) createSheet('Sheet 1');
        loadTheme();
        bindEvents();
    }

    function createSheet(name) {
        const sheet = {
            name: name || ('Sheet ' + (sheets.length + 1)),
            data: makeData(INITIAL_ROWS, INITIAL_COLS),
            undoStack: [],
            redoStack: [],
        };
        sheets.push(sheet);
        activeSheetIdx = sheets.length - 1;
        autoSave();
        renderSheetTabs();
        renderGrid();
        setActiveCell(0, 0);
    }

    function makeData(rows, cols) {
        const d = [];
        for (let r = 0; r < rows; r++) d.push(new Array(cols).fill(''));
        return d;
    }

    // ── Auto-save / restore ──
    const CACHE_KEY = 'csv-editor-sheets';
    const CACHE_ACTIVE = 'csv-editor-active';

    function autoSave() {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(() => {
            try {
                localStorage.setItem(CACHE_KEY, JSON.stringify(sheets));
                localStorage.setItem(CACHE_ACTIVE, activeSheetIdx);
            } catch (e) { /* quota exceeded — ignore */ }
        }, 500);
    }

    function loadFromCache() {
        try {
            const raw = localStorage.getItem(CACHE_KEY);
            if (!raw) return false;
            const cached = JSON.parse(raw);
            if (!Array.isArray(cached) || cached.length === 0) return false;

            sheets = cached.map(s => ({
                name: s.name || 'Sheet',
                data: s.data || makeData(INITIAL_ROWS, INITIAL_COLS),
                undoStack: [],
                redoStack: [],
            }));

            activeSheetIdx = parseInt(localStorage.getItem(CACHE_ACTIVE)) || 0;
            if (activeSheetIdx >= sheets.length) activeSheetIdx = 0;

            renderSheetTabs();
            renderGrid();
            setActiveCell(0, 0);
            showToast('Previous session restored');
            return true;
        } catch (e) {
            return false;
        }
    }

    function clearCache() {
        localStorage.removeItem(CACHE_KEY);
        localStorage.removeItem(CACHE_ACTIVE);
    }

    function showToast(msg) {
        toast.textContent = msg;
        toast.classList.add('visible');
        setTimeout(() => toast.classList.remove('visible'), 5000);
    }

    function switchSheet(idx) {
        if (idx === activeSheetIdx) return;
        if (isEditing) finishEditing(true);
        activeSheetIdx = idx;
        autoSave();
        renderSheetTabs();
        renderGrid();
        setActiveCell(0, 0);
    }

    // ── Column label ──
    function colLabel(c) {
        let s = '';
        c++;
        while (c > 0) { c--; s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26); }
        return s;
    }

    function cellId(r, c) { return colLabel(c) + (r + 1); }

    // ── Render ──
    function renderGrid() {
        const d = data();
        const rows = d.length;
        const cols = d[0].length;

        grid.style.gridTemplateColumns = `46px repeat(${cols}, minmax(90px, 1fr))`;
        grid.style.gridTemplateRows = `24px repeat(${rows}, 24px)`;
        grid.innerHTML = '';

        grid.appendChild(el('div', 'corner-cell'));

        for (let c = 0; c < cols; c++) {
            const h = el('div', 'col-header');
            h.textContent = colLabel(c);
            h.dataset.col = c;
            h.addEventListener('click', () => selectColumn(c));
            grid.appendChild(h);
        }

        for (let r = 0; r < rows; r++) {
            const rh = el('div', 'row-header');
            rh.textContent = r + 1;
            rh.dataset.row = r;
            rh.addEventListener('click', () => selectRow(r));
            grid.appendChild(rh);

            for (let c = 0; c < cols; c++) {
                const cell = el('div', 'cell');
                cell.textContent = d[r][c];
                cell.dataset.row = r;
                cell.dataset.col = c;
                cell.addEventListener('mousedown', onCellMouseDown);
                cell.addEventListener('touchstart', onCellTouchStart, { passive: false });
                cell.addEventListener('mouseover', onCellMouseOver);
                grid.appendChild(cell);
            }
        }

        updateStatus();
    }

    function refreshCells() {
        const d = data();
        grid.querySelectorAll('.cell').forEach(cell => {
            const r = +cell.dataset.row;
            const c = +cell.dataset.col;
            if (r < d.length && c < d[0].length) cell.textContent = d[r][c];
        });
        updateStatus();
    }

    function updateStatus() {
        const d = data();
        let count = 0;
        for (let r = 0; r < d.length; r++)
            for (let c = 0; c < d[0].length; c++)
                if (d[r][c] !== '') count++;
        statusCells.textContent = count + ' cell' + (count !== 1 ? 's' : '') + ' filled';
    }

    // ── Selection ──
    function setActiveCell(r, c, clearSelection) {
        if (isEditing) finishEditing(true);
        activeCell = { row: r, col: c };

        if (clearSelection !== false) {
            selectionStart = { row: r, col: c };
            selectionEnd = { row: r, col: c };
        }

        cellRef.textContent = cellId(r, c);
        if (!formulaFocused) {
            formulaInput.value = data()[r][c];
        }

        grid.querySelectorAll('.cell.active').forEach(e => e.classList.remove('active'));
        grid.querySelectorAll('.cell.in-range').forEach(e => e.classList.remove('in-range'));
        grid.querySelectorAll('.row-header.selected').forEach(e => e.classList.remove('selected'));
        grid.querySelectorAll('.col-header.selected').forEach(e => e.classList.remove('selected'));

        const activeEl = getCellEl(r, c);
        if (activeEl) activeEl.classList.add('active');

        highlightRange();
    }

    function highlightRange() {
        grid.querySelectorAll('.cell.in-range').forEach(e => e.classList.remove('in-range'));
        if (!selectionStart || !selectionEnd) return;
        const r1 = Math.min(selectionStart.row, selectionEnd.row);
        const r2 = Math.max(selectionStart.row, selectionEnd.row);
        const c1 = Math.min(selectionStart.col, selectionEnd.col);
        const c2 = Math.max(selectionStart.col, selectionEnd.col);

        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const el = getCellEl(r, c);
                if (el && !(r === activeCell.row && c === activeCell.col)) el.classList.add('in-range');
            }
        }

        if (r1 === r2 && c1 === c2) {
            statusSelection.textContent = '';
        } else {
            statusSelection.textContent = `${r2 - r1 + 1}R × ${c2 - c1 + 1}C`;
        }
    }

    function getCellEl(r, c) {
        return grid.querySelector(`.cell[data-row="${r}"][data-col="${c}"]`);
    }

    function selectRow(r) {
        if (isEditing) finishEditing(true);
        const d = data();
        selectionStart = { row: r, col: 0 };
        selectionEnd = { row: r, col: d[0].length - 1 };
        activeCell = { row: r, col: 0 };
        setActiveCell(r, 0, false);
        grid.querySelectorAll('.row-header.selected').forEach(e => e.classList.remove('selected'));
        const rh = grid.querySelector(`.row-header[data-row="${r}"]`);
        if (rh) rh.classList.add('selected');
    }

    function selectColumn(c) {
        if (isEditing) finishEditing(true);
        const d = data();
        selectionStart = { row: 0, col: c };
        selectionEnd = { row: d.length - 1, col: c };
        activeCell = { row: 0, col: c };
        setActiveCell(0, c, false);
        grid.querySelectorAll('.col-header.selected').forEach(e => e.classList.remove('selected'));
        const ch = grid.querySelector(`.col-header[data-col="${c}"]`);
        if (ch) ch.classList.add('selected');
    }

    // ── Mouse events ──
    let lastTapTime = 0;
    let lastTapTarget = null;

    function onCellMouseDown(e) {
        if (e.button !== 0) return;
        const r = +e.currentTarget.dataset.row;
        const c = +e.currentTarget.dataset.col;

        if (e.shiftKey) {
            selectionEnd = { row: r, col: c };
            setActiveCell(r, c, false);
            selectionStart = selectionStart || activeCell;
            highlightRange();
        } else {
            setActiveCell(r, c);
        }
        isSelecting = true;
    }

    function onCellTouchStart(e) {
        const cell = e.currentTarget;
        const r = +cell.dataset.row;
        const c = +cell.dataset.col;
        const now = Date.now();

        if (lastTapTarget === cell && now - lastTapTime < 400) {
            // double tap — enter edit mode
            e.preventDefault();
            setActiveCell(r, c);
            startEditing();
            lastTapTime = 0;
            lastTapTarget = null;
        } else {
            lastTapTime = now;
            lastTapTarget = cell;
            setActiveCell(r, c);
        }
    }

    function onCellMouseOver(e) {
        if (!isSelecting) return;
        const r = +e.currentTarget.dataset.row;
        const c = +e.currentTarget.dataset.col;
        selectionEnd = { row: r, col: c };
        highlightRange();
    }

    // ── Editing ──
    function startEditing(key) {
        if (isEditing) return;
        isEditing = true;
        const cellEl = getCellEl(activeCell.row, activeCell.col);
        if (!cellEl) return;

        cellEl.classList.add('editing');
        const input = document.createElement('input');
        input.type = 'text';
        input.inputMode = 'text';
        input.enterKeyHint = 'done';
        input.className = 'editor-input';
        input.value = key !== undefined ? key : data()[activeCell.row][activeCell.col];
        cellEl.textContent = '';
        cellEl.appendChild(input);
        // Use setTimeout to ensure the input is in the DOM before focusing (needed for mobile keyboard)
        setTimeout(() => {
            input.focus();
            if (key === undefined) input.select();
        }, 0);

        input.addEventListener('keydown', onEditKeyDown);
        input.addEventListener('blur', () => finishEditing(true));
    }

    function finishEditing(save) {
        if (!isEditing) return;
        isEditing = false;
        const cellEl = getCellEl(activeCell.row, activeCell.col);
        if (!cellEl) return;

        const input = cellEl.querySelector('.editor-input');
        if (input && save) {
            pushUndo();
            data()[activeCell.row][activeCell.col] = input.value;
        }

        cellEl.classList.remove('editing');
        cellEl.textContent = data()[activeCell.row][activeCell.col];
        if (!formulaFocused) {
            formulaInput.value = data()[activeCell.row][activeCell.col];
        }
    }

    function onEditKeyDown(e) {
        if (e.key === 'Enter') { e.preventDefault(); finishEditing(true); moveActive(1, 0); }
        else if (e.key === 'Tab') { e.preventDefault(); finishEditing(true); moveActive(0, e.shiftKey ? -1 : 1); }
        else if (e.key === 'Escape') { finishEditing(false); }
    }

    // ── Navigation ──
    function moveActive(dr, dc) {
        const d = data();
        const nr = activeCell.row + dr;
        const nc = activeCell.col + dc;
        if (nr >= 0 && nr < d.length && nc >= 0 && nc < d[0].length) {
            setActiveCell(nr, nc);
            scrollToCell(nr, nc);
        }
    }

    function scrollToCell(r, c) {
        const cellEl = getCellEl(r, c);
        if (cellEl) cellEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }

    // ── Keyboard ──
    function onKeyDown(e) {
        if (isEditing) return;
        if (formulaFocused) return;

        const ctrl = e.ctrlKey || e.metaKey;

        if (ctrl && e.key === 'z') { e.preventDefault(); undo(); return; }
        if (ctrl && e.key === 'y') { e.preventDefault(); redo(); return; }
        if (ctrl && e.key === 'a') { e.preventDefault(); selectAll(); return; }
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); clearSelection(); return; }

        if (e.key === 'ArrowDown') { e.preventDefault(); moveActive(1, 0); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); moveActive(-1, 0); }
        else if (e.key === 'ArrowRight') { e.preventDefault(); moveActive(0, 1); }
        else if (e.key === 'ArrowLeft') { e.preventDefault(); moveActive(0, -1); }
        else if (e.key === 'Tab') { e.preventDefault(); moveActive(0, e.shiftKey ? -1 : 1); }
        else if (e.key === 'Enter') { e.preventDefault(); moveActive(1, 0); }
        else if (e.key === 'F2') { e.preventDefault(); startEditing(); }
        else if (e.key.length === 1 && !ctrl) {
            e.preventDefault(); // prevent the key from also being typed into the new input
            startEditing(e.key);
        }
    }

    // ── Undo/Redo ──
    function pushUndo() {
        const s = currentSheet();
        s.undoStack.push(JSON.parse(JSON.stringify(s.data)));
        if (s.undoStack.length > 100) s.undoStack.shift();
        s.redoStack = [];
        autoSave();
    }

    function undo() {
        const s = currentSheet();
        if (s.undoStack.length === 0) return;
        s.redoStack.push(JSON.parse(JSON.stringify(s.data)));
        s.data = s.undoStack.pop();
        ensureDimensions(s.data);
        renderGrid();
        setActiveCell(activeCell.row, activeCell.col);
    }

    function redo() {
        const s = currentSheet();
        if (s.redoStack.length === 0) return;
        s.undoStack.push(JSON.parse(JSON.stringify(s.data)));
        s.data = s.redoStack.pop();
        ensureDimensions(s.data);
        renderGrid();
        setActiveCell(activeCell.row, activeCell.col);
    }

    function ensureDimensions(d) {
        while (d.length < INITIAL_ROWS) d.push(new Array(INITIAL_COLS).fill(''));
        d.forEach((row, i) => { while (row.length < INITIAL_COLS) d[i].push(''); });
    }

    // ── Clipboard (file:// compatible) ──
    function getSelectionRange() {
        if (!selectionStart || !selectionEnd) return null;
        return {
            r1: Math.min(selectionStart.row, selectionEnd.row),
            r2: Math.max(selectionStart.row, selectionEnd.row),
            c1: Math.min(selectionStart.col, selectionEnd.col),
            c2: Math.max(selectionStart.col, selectionEnd.col),
        };
    }

    function copySelection(cut) {
        const range = getSelectionRange();
        if (!range) return;
        const d = data();
        let text = '';
        for (let r = range.r1; r <= range.r2; r++) {
            const row = [];
            for (let c = range.c1; c <= range.c2; c++) row.push(d[r][c]);
            text += row.join('\t') + '\n';
        }

        // textarea fallback for file://
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.select();
        try { document.execCommand('copy'); } catch (e) { /* fallback: use clipboard API */ }
        document.body.removeChild(ta);

        // also try modern API
        try { navigator.clipboard.writeText(text); } catch (e) {}

        if (cut) {
            pushUndo();
            for (let r = range.r1; r <= range.r2; r++)
                for (let c = range.c1; c <= range.c2; c++)
                    d[r][c] = '';
            refreshCells();
        }
    }

    function handlePaste(text) {
        if (!text) return;
        pushUndo();
        const d = data();
        const rows = text.replace(/\r\n/g, '\n').split('\n').filter((_, i, a) => i < a.length - 1 || a[i] !== '');
        const parsed = rows.map(r => r.split('\t'));

        const pasteRows = parsed.length;
        const pasteCols = (parsed[0] || []).length;

        // If user has a multi-cell selection and clipboard is a single value, fill entire selection
        const range = getSelectionRange();
        if (range && pasteRows === 1 && pasteCols === 1) {
            const selRows = range.r2 - range.r1 + 1;
            const selCols = range.c2 - range.c1 + 1;
            if (selRows > 1 || selCols > 1) {
                const val = parsed[0][0];
                ensureCapacity(d, range.r2 + 1, range.c2 + 1);
                for (let r = range.r1; r <= range.r2; r++)
                    for (let c = range.c1; c <= range.c2; c++)
                        d[r][c] = val;
                refreshCells();
                return;
            }
        }

        ensureCapacity(d, activeCell.row + pasteRows, activeCell.col + pasteCols);

        for (let r = 0; r < pasteRows; r++) {
            for (let c = 0; c < pasteCols; c++) {
                // Tile: repeat clipboard contents to fill selection if range is larger
                const sr = r % parsed.length;
                const sc = c % (parsed[sr] || []).length;
                const dr = activeCell.row + r;
                const dc = activeCell.col + c;
                if (dr < d.length && dc < d[0].length) d[dr][dc] = (parsed[sr] || [])[sc] || '';
            }
        }
        refreshCells();
    }

    function onNativePaste(e) {
        if (isEditing) return; // let the input handle its own paste
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (text) {
            e.preventDefault();
            handlePaste(text);
        }
    }

    function onNativeCopy(e) {
        if (isEditing) return;
        const range = getSelectionRange();
        if (!range) return;
        const d = data();
        let text = '';
        for (let r = range.r1; r <= range.r2; r++) {
            const row = [];
            for (let c = range.c1; c <= range.c2; c++) row.push(d[r][c]);
            text += row.join('\t') + '\n';
        }
        e.preventDefault();
        e.clipboardData.setData('text/plain', text);
    }

    function selectAll() {
        const d = data();
        selectionStart = { row: 0, col: 0 };
        selectionEnd = { row: d.length - 1, col: d[0].length - 1 };
        highlightRange();
    }

    function clearSelection() {
        const range = getSelectionRange();
        if (!range) return;
        pushUndo();
        const d = data();
        for (let r = range.r1; r <= range.r2; r++)
            for (let c = range.c1; c <= range.c2; c++)
                d[r][c] = '';
        refreshCells();
        formulaInput.value = '';
    }

    function ensureCapacity(d, neededRows, neededCols) {
        let changed = false;
        while (d.length < neededRows) { d.push(new Array(d[0].length).fill('')); changed = true; }
        if (d[0].length < neededCols) {
            d.forEach((row, i) => { while (row.length < neededCols) { d[i].push(''); changed = true; } });
        }
        if (changed) renderGrid();
    }

    // ── File ops ──
    function openFile() { fileInput.click(); }

    function handleFileOpen(e) {
        const file = e.target.files[0];
        if (!file) return;
        if (file.size > MAX_FILE_SIZE) {
            alert('File is too large (' + (file.size / 1024 / 1024).toFixed(1) + 'MB). Maximum is 5MB.');
            fileInput.value = '';
            return;
        }
        const reader = new FileReader();
        reader.onload = function (evt) {
            pushUndo();
            parseCSV(evt.target.result);
            renderGrid();
            setActiveCell(0, 0);
        };
        reader.readAsText(file);
        fileInput.value = '';
    }

    function parseCSV(text) {
        const rows = [];
        let current = '';
        let inQuotes = false;
        let row = [];

        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (inQuotes) {
                if (ch === '"' && text[i + 1] === '"') { current += '"'; i++; }
                else if (ch === '"') { inQuotes = false; }
                else { current += ch; }
            } else {
                if (ch === '"') { inQuotes = true; }
                else if (ch === ',') { row.push(current); current = ''; }
                else if (ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) {
                    row.push(current); current = ''; rows.push(row); row = [];
                    if (ch === '\r') i++;
                } else { current += ch; }
            }
        }
        if (current || row.length > 0) { row.push(current); rows.push(row); }

        if (rows.length > MAX_ROWS) {
            alert('File has ' + rows.length + ' rows. Only the first ' + MAX_ROWS + ' will be loaded.');
            rows.length = MAX_ROWS;
        }

        const maxCols = Math.max(INITIAL_COLS, ...rows.map(r => r.length));
        const maxRows = Math.max(INITIAL_ROWS, rows.length);

        const d = [];
        for (let r = 0; r < maxRows; r++) {
            const source = rows[r] || [];
            const padded = [];
            for (let c = 0; c < maxCols; c++) padded.push(source[c] !== undefined ? source[c] : '');
            d.push(padded);
        }
        currentSheet().data = d;
    }

    function dataToCSV(d) {
        let text = '';
        let maxR = d.length - 1;
        while (maxR > 0 && d[maxR].every(c => c === '')) maxR--;
        let maxC = d[0].length - 1;
        while (maxC > 0 && d.every(r => r[maxC] === '')) maxC--;

        for (let r = 0; r <= maxR; r++) {
            const cells = [];
            for (let c = 0; c <= maxC; c++) {
                let val = d[r][c];
                if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                    val = '"' + val.replace(/"/g, '""') + '"';
                }
                cells.push(val);
            }
            text += cells.join(',') + '\r\n';
        }
        return text;
    }

    function downloadCSV(filename, text) {
        const blob = new Blob([text], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = (filename || 'spreadsheet') + '.csv';
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function saveFile() {
        const sheet = currentSheet();
        downloadCSV(sheet.name, dataToCSV(sheet.data));
    }

    function saveAllFiles() {
        sheets.forEach(sheet => {
            downloadCSV(sheet.name, dataToCSV(sheet.data));
        });
    }

    function newSheet() {
        createSheet();
    }

    // ── Add row/col ──
    function addRow() {
        pushUndo();
        const d = data();
        d.splice(activeCell.row, 0, new Array(d[0].length).fill(''));
        renderGrid();
        setActiveCell(activeCell.row, activeCell.col);
    }

    function addCol() {
        pushUndo();
        data().forEach(row => row.splice(activeCell.col, 0, ''));
        renderGrid();
        setActiveCell(activeCell.row, activeCell.col);
    }

    // ── Context menu (grid) ──
    function showContextMenu(e) {
        e.preventDefault();
        removeContextMenu();

        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = e.clientX + 'px';
        menu.style.top = e.clientY + 'px';

        const items = [
            { label: 'Cut', action: () => copySelection(true) },
            { label: 'Copy', action: () => copySelection() },
            { label: 'Paste', action: () => { navigator.clipboard.readText().then(handlePaste).catch(() => {}); } },
            { sep: true },
            { label: 'Insert Row Above', action: () => insertRow(activeCell.row) },
            { label: 'Insert Row Below', action: () => insertRow(activeCell.row + 1) },
            { label: 'Insert Column Left', action: () => insertCol(activeCell.col) },
            { label: 'Insert Column Right', action: () => insertCol(activeCell.col + 1) },
            { sep: true },
            { label: 'Delete Row', action: () => deleteRow(activeCell.row) },
            { label: 'Delete Column', action: () => deleteCol(activeCell.col) },
            { sep: true },
            { label: 'Clear Selection', action: () => clearSelection() },
        ];

        items.forEach(item => {
            if (item.sep) {
                menu.appendChild(el('div', 'context-menu-sep'));
            } else {
                const mi = el('div', 'context-menu-item');
                mi.textContent = item.label;
                mi.addEventListener('click', () => { removeContextMenu(); item.action(); });
                menu.appendChild(mi);
            }
        });

        document.body.appendChild(menu);
        const rect = menu.getBoundingClientRect();
        if (rect.right > window.innerWidth) menu.style.left = (window.innerWidth - rect.width - 8) + 'px';
        if (rect.bottom > window.innerHeight) menu.style.top = (window.innerHeight - rect.height - 8) + 'px';
    }

    function removeContextMenu() {
        document.querySelectorAll('.context-menu').forEach(m => m.remove());
    }

    function insertRow(at) {
        pushUndo();
        const d = data();
        d.splice(at, 0, new Array(d[0].length).fill(''));
        renderGrid();
        setActiveCell(at, activeCell.col);
    }

    function insertCol(at) {
        pushUndo();
        data().forEach(row => row.splice(at, 0, ''));
        renderGrid();
        setActiveCell(activeCell.row, at);
    }

    function deleteRow(r) {
        const d = data();
        if (d.length <= 1) return;
        pushUndo();
        d.splice(r, 1);
        renderGrid();
        setActiveCell(Math.min(r, d.length - 1), activeCell.col);
    }

    function deleteCol(c) {
        const d = data();
        if (d[0].length <= 1) return;
        pushUndo();
        d.forEach(row => row.splice(c, 1));
        renderGrid();
        setActiveCell(activeCell.row, Math.min(c, d[0].length - 1));
    }

    // ── Sheet tabs ──
    let selectedSheets = new Set(); // multi-select state

    function renderSheetTabs() {
        sheetTabs.innerHTML = '';

        sheets.forEach((sheet, idx) => {
            const tab = el('button', 'sheet-tab');
            if (idx === activeSheetIdx) tab.classList.add('active');
            if (selectedSheets.has(idx)) tab.classList.add('multi-selected');
            tab.textContent = sheet.name;
            tab.dataset.idx = idx;

            // PC events
            tab.addEventListener('click', (e) => onTabClick(e, idx));
            tab.addEventListener('dblclick', () => renameSheetTab(idx, tab));
            tab.addEventListener('contextmenu', (e) => showSheetContextMenu(e, idx));
            tab.addEventListener('mousedown', (e) => onTabMouseDown(e, idx));

            // Mobile touch events
            tab.addEventListener('touchstart', (e) => onTabTouchStart(e, idx), { passive: false });
            tab.addEventListener('touchmove', (e) => onTabTouchMove(e), { passive: false });
            tab.addEventListener('touchend', (e) => onTabTouchEnd(e, idx));

            sheetTabs.appendChild(tab);
        });
    }

    // ── PC tab click ──
    function onTabClick(e, idx) {
        if (draggedTabIdx !== null) return;

        if (e.ctrlKey || e.metaKey) {
            // Ctrl+click: toggle single tab in/out of multi-select
            if (selectedSheets.has(idx)) {
                selectedSheets.delete(idx);
            } else {
                selectedSheets.add(idx);
            }
            renderSheetTabs();
        } else if (e.shiftKey) {
            // Shift+click: select range from active to clicked
            const from = Math.min(activeSheetIdx, idx);
            const to = Math.max(activeSheetIdx, idx);
            selectedSheets.clear();
            for (let i = from; i <= to; i++) selectedSheets.add(i);
            renderSheetTabs();
        } else if (selectedSheets.size > 0) {
            selectedSheets.clear();
            switchSheet(idx);
        } else {
            switchSheet(idx);
        }
    }

    function renameSheetTab(idx, tabEl) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'sheet-tab-rename';
        input.value = sheets[idx].name;

        tabEl.textContent = '';
        tabEl.appendChild(input);
        input.focus();
        input.select();

        function finish() {
            const val = input.value.trim() || sheets[idx].name;
            sheets[idx].name = val;
            renderSheetTabs();
        }

        input.addEventListener('blur', finish);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
            if (e.key === 'Escape') { input.value = sheets[idx].name; input.blur(); }
        });
    }

    function showSheetContextMenu(e, idx) {
        if (e.preventDefault) e.preventDefault();
        if (e.stopPropagation) e.stopPropagation();
        removeContextMenu();

        // If right-clicking a tab that isn't in the multi-selection, select only it
        if (!selectedSheets.has(idx)) {
            selectedSheets.clear();
            selectedSheets.add(idx);
            renderSheetTabs();
        }

        const multiCount = selectedSheets.size;
        const canDelete = sheets.length > multiCount;

        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = e.clientX + 'px';
        menu.style.top = e.clientY + 'px';

        const items = [];

        if (multiCount > 1) {
            items.push({ label: `Delete ${multiCount} Sheets`, action: () => deleteSelectedSheets(), disabled: !canDelete });
            items.push({ sep: true });
        }

        items.push({ label: 'Rename', action: () => renameSheetTab(idx, sheetTabs.children[idx]) });
        items.push({ label: 'Duplicate', action: () => duplicateSheet(idx) });
        items.push({ sep: true });
        items.push({ label: 'Move Left', action: () => moveSheet(idx, -1), disabled: idx === 0 });
        items.push({ label: 'Move Right', action: () => moveSheet(idx, 1), disabled: idx === sheets.length - 1 });
        items.push({ sep: true });
        items.push({ label: 'Delete', action: () => {
            if (selectedSheets.size > 1) {
                deleteSelectedSheets();
            } else {
                deleteSheet(idx);
            }
        }, disabled: sheets.length <= (selectedSheets.size > 1 ? selectedSheets.size : 1) });

        items.forEach(item => {
            if (item.sep) {
                menu.appendChild(el('div', 'context-menu-sep'));
            } else {
                const mi = el('div', 'context-menu-item');
                mi.textContent = item.label;
                if (item.disabled) {
                    mi.style.opacity = '0.4';
                    mi.style.cursor = 'default';
                } else {
                    mi.addEventListener('click', () => { removeContextMenu(); item.action(); });
                }
                menu.appendChild(mi);
            }
        });

        document.body.appendChild(menu);
        const rect = menu.getBoundingClientRect();
        if (rect.right > window.innerWidth) menu.style.left = (window.innerWidth - rect.width - 8) + 'px';
        if (rect.bottom > window.innerHeight) menu.style.top = (window.innerHeight - rect.height - 8) + 'px';
    }

    function deleteSelectedSheets() {
        if (sheets.length <= selectedSheets.size) return;
        const toDelete = [...selectedSheets].sort((a, b) => b - a); // delete from highest index first
        toDelete.forEach(idx => sheets.splice(idx, 1));
        selectedSheets.clear();
        if (activeSheetIdx >= sheets.length) activeSheetIdx = sheets.length - 1;
        autoSave();
        renderSheetTabs();
        renderGrid();
        setActiveCell(0, 0);
    }

    function duplicateSheet(idx) {
        const src = sheets[idx];
        const dup = {
            name: src.name + ' (copy)',
            data: JSON.parse(JSON.stringify(src.data)),
            undoStack: [],
            redoStack: [],
        };
        sheets.splice(idx + 1, 0, dup);
        switchSheet(idx + 1);
    }

    function moveSheet(idx, dir) {
        const newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= sheets.length) return;
        [sheets[idx], sheets[newIdx]] = [sheets[newIdx], sheets[idx]];
        activeSheetIdx = newIdx;
        renderSheetTabs();
    }

    function deleteSheet(idx) {
        if (sheets.length <= 1) return;
        sheets.splice(idx, 1);
        if (activeSheetIdx >= sheets.length) activeSheetIdx = sheets.length - 1;
        else if (activeSheetIdx > idx) activeSheetIdx--;
        renderSheetTabs();
        renderGrid();
        setActiveCell(0, 0);
    }

    // ── Tab drag (PC: mouse-based with 200ms hold threshold) ──
    let draggedTabIdx = null;
    let mouseTabIdx = null;
    let mouseTimer = null;
    let mouseStartX = 0;
    let mouseStartY = 0;
    let mouseLastX = 0;
    let mouseDragging = false;
    let mouseScrolling = false;

    function onTabMouseDown(e, idx) {
        if (e.button !== 0) return;
        mouseStartX = e.clientX;
        mouseLastX = e.clientX;
        mouseStartY = e.clientY;
        mouseTabIdx = idx;
        mouseDragging = false;
        mouseScrolling = false;

        mouseTimer = setTimeout(() => {
            mouseDragging = true;
            draggedTabIdx = idx;
            const tabEl = sheetTabs.children[idx];
            if (tabEl) tabEl.classList.add('dragging');
        }, 200);
    }

    function onTabMouseMove(e) {
        if (mouseTabIdx === null) return;
        const dx = e.clientX - mouseStartX;
        const dy = e.clientY - mouseStartY;
        const dist = Math.sqrt(dx * dx + dy * dy);

        if (!mouseDragging && !mouseScrolling && dist > 5) {
            clearTimeout(mouseTimer);
            mouseTimer = null;
            mouseScrolling = true;
        }

        if (mouseScrolling) {
            const delta = e.clientX - mouseLastX;
            mouseLastX = e.clientX;
            sheetScroll.scrollLeft -= delta;
            return;
        }

        if (mouseDragging) {
            clearTimeout(mouseTimer);
            const tabs = sheetTabs.querySelectorAll('.sheet-tab');
            tabs.forEach(t => t.classList.remove('drag-over'));
            const overEl = document.elementFromPoint(e.clientX, e.clientY);
            if (overEl) {
                const targetTab = overEl.closest('.sheet-tab');
                if (targetTab && +targetTab.dataset.idx !== draggedTabIdx) {
                    targetTab.classList.add('drag-over');
                }
            }
        }
    }

    function onTabMouseUp(e) {
        clearTimeout(mouseTimer);
        mouseTimer = null;

        if (mouseDragging && draggedTabIdx !== null) {
            const overTab = sheetTabs.querySelector('.sheet-tab.drag-over');
            if (overTab) {
                const targetIdx = +overTab.dataset.idx;
                if (targetIdx !== draggedTabIdx) {
                    const moved = sheets.splice(draggedTabIdx, 1)[0];
                    sheets.splice(targetIdx, 0, moved);
                    activeSheetIdx = targetIdx;
                    autoSave();
                }
            }
            sheetTabs.querySelectorAll('.sheet-tab').forEach(t => {
                t.classList.remove('dragging');
                t.classList.remove('drag-over');
            });
            renderSheetTabs();
            draggedTabIdx = '__just_dragged__';
            setTimeout(() => { draggedTabIdx = null; }, 50);
        }

        mouseTabIdx = null;
        mouseDragging = false;
        mouseScrolling = false;
    }

    // ── Tab touch (mobile) ──
    // <200ms move = manual scroll
    // 200-500ms hold + move = drag reorder
    // 500ms hold still = show context menu
    // move after 500ms = multi-select
    let touchTabIdx = null;
    let touchTimer = null;
    let touchMenuTimer = null;
    let touchStartX = 0;
    let touchStartY = 0;
    let touchLastX = 0;
    let touchStartTime = 0;
    let touchLongPressFired = false;
    let touchDragging = false;
    let touchDragIdx = null;
    let touchMultiSelecting = false;
    let touchDragReady = false;
    let touchClaimed = false;

    function onTabTouchStart(e, idx) {
        const touch = e.touches[0];
        touchStartX = touch.clientX;
        touchLastX = touch.clientX;
        touchStartY = touch.clientY;
        touchStartTime = Date.now();
        touchTabIdx = idx;
        touchLongPressFired = false;
        touchDragging = false;
        touchDragReady = false;
        touchMultiSelecting = false;
        touchClaimed = false;

        touchTimer = setTimeout(() => {
            touchDragReady = true;
            // Show drag-ready lift animation on the held tab
            const tabEl = sheetTabs.children[idx];
            if (tabEl) tabEl.classList.add('dragging');
        }, 200);

        touchMenuTimer = setTimeout(() => {
            // Cancel drag-ready animation — menu takes over
            const tabEl = sheetTabs.children[idx];
            if (tabEl) tabEl.classList.remove('dragging');
            touchDragReady = false;
            touchLongPressFired = true;
            if (!selectedSheets.has(idx)) {
                selectedSheets.clear();
                selectedSheets.add(idx);
            }
            renderSheetTabs();
            const fakeEvent = {
                preventDefault: () => {},
                stopPropagation: () => {},
                clientX: touchStartX,
                clientY: touchStartY,
            };
            showSheetContextMenu(fakeEvent, idx);
        }, 500);
    }

    function onTabTouchMove(e) {
        if (touchTabIdx === null) return;
        const touch = e.touches[0];
        const dx = touch.clientX - touchStartX;
        const dy = touch.clientY - touchStartY;
        const dist = Math.sqrt(dx * dx + dy * dy);

        if (dist > 10) {
            // Claim gesture on first significant move to prevent browser scroll steal
            if (!touchClaimed) {
                touchClaimed = true;
                e.preventDefault();
            }

            if (touchLongPressFired) {
                // Moved after menu — multi-select mode
                removeContextMenu();
                if (!touchMultiSelecting) {
                    touchMultiSelecting = true;
                    if (!selectedSheets.has(touchTabIdx)) {
                        selectedSheets.add(touchTabIdx);
                    }
                }
                const overEl = document.elementFromPoint(touch.clientX, touch.clientY);
                if (overEl) {
                    const targetTab = overEl.closest('.sheet-tab');
                    if (targetTab) {
                        const targetIdx = +targetTab.dataset.idx;
                        if (!selectedSheets.has(targetIdx)) {
                            selectedSheets.add(targetIdx);
                            renderSheetTabs();
                        }
                    }
                }
                touchLastX = touch.clientX;
                return;
            }

            if (touchDragReady) {
                // 200-500ms window: drag reorder
                clearTimeout(touchMenuTimer);
                touchMenuTimer = null;

                if (!touchDragging) {
                    touchDragging = true;
                    touchDragIdx = touchTabIdx;
                    const tabEl = sheetTabs.children[touchDragIdx];
                    if (tabEl) tabEl.classList.add('dragging');
                }

                const tabs = sheetTabs.querySelectorAll('.sheet-tab');
                tabs.forEach(t => t.classList.remove('drag-over'));
                const overTab = document.elementFromPoint(touch.clientX, touch.clientY);
                if (overTab) {
                    const targetTab = overTab.closest('.sheet-tab');
                    if (targetTab && +targetTab.dataset.idx !== touchDragIdx) {
                        targetTab.classList.add('drag-over');
                    }
                }
            } else {
                // <200ms: manual scroll
                const delta = touch.clientX - touchLastX;
                sheetScroll.scrollLeft -= delta;
            }

            touchLastX = touch.clientX;
        }
    }

    function onTabTouchEnd(e, idx) {
        clearTimeout(touchTimer);
        clearTimeout(touchMenuTimer);
        touchTimer = null;
        touchMenuTimer = null;

        // Remove drag-ready lift if still showing
        if (touchDragReady && !touchDragging) {
            const tabEl = sheetTabs.children[idx];
            if (tabEl) tabEl.classList.remove('dragging');
        }

        if (touchDragging && touchDragIdx !== null) {
            const overTab = sheetTabs.querySelector('.sheet-tab.drag-over');
            if (overTab) {
                const targetIdx = +overTab.dataset.idx;
                if (targetIdx !== touchDragIdx) {
                    const moved = sheets.splice(touchDragIdx, 1)[0];
                    sheets.splice(targetIdx, 0, moved);
                    activeSheetIdx = targetIdx;
                    autoSave();
                }
            }
            sheetTabs.querySelectorAll('.sheet-tab').forEach(t => {
                t.classList.remove('dragging');
                t.classList.remove('drag-over');
            });
            renderSheetTabs();
        } else if (touchMultiSelecting) {
            renderSheetTabs();
        } else if (!touchLongPressFired) {
            if (selectedSheets.size > 1) {
                selectedSheets.clear();
            }
            switchSheet(idx);
        }

        touchTabIdx = null;
        touchDragIdx = null;
        touchDragging = false;
        touchLongPressFired = false;
        touchMultiSelecting = false;
        touchDragReady = false;
        touchClaimed = false;
    }

    // ── Dark mode ──
    const darkToggle = document.getElementById('dark-mode-toggle');

    function toggleDark() {
        const isDark = darkToggle.checked;
        document.documentElement.dataset.theme = isDark ? 'dark' : '';
        localStorage.setItem('csv-editor-theme', isDark ? 'dark' : 'light');
    }

    function loadTheme() {
        const saved = localStorage.getItem('csv-editor-theme');
        if (saved === 'dark') {
            document.documentElement.dataset.theme = 'dark';
            darkToggle.checked = true;
        }
    }

    // ── Formula bar ──
    function onFormulaInput(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            pushUndo();
            data()[activeCell.row][activeCell.col] = formulaInput.value;
            refreshCells();
            formulaInput.blur();
            moveActive(1, 0);
        } else if (e.key === 'Escape') {
            formulaInput.value = data()[activeCell.row][activeCell.col];
            formulaInput.blur();
        }
    }

    function onFormulaInputLive() {
        if (!formulaFocused) return;
        pushUndo();
        data()[activeCell.row][activeCell.col] = formulaInput.value;
        const cellEl = getCellEl(activeCell.row, activeCell.col);
        if (cellEl) cellEl.textContent = formulaInput.value;
    }

    function onFormulaFocus() {
        formulaFocused = true;
        formulaInput.select();
    }

    function onFormulaBlur() {
        formulaFocused = false;
        formulaInput.value = data()[activeCell.row][activeCell.col];
    }

    // ── Bind events ──
    function bindEvents() {
        document.addEventListener('mouseup', (e) => { isSelecting = false; onTabMouseUp(e); });
        document.addEventListener('mousemove', onTabMouseMove);
        document.addEventListener('click', removeContextMenu);
        gridContainer.addEventListener('contextmenu', showContextMenu);
        document.addEventListener('keydown', onKeyDown);

        // Clipboard — use native events (works on file://)
        document.addEventListener('copy', onNativeCopy);
        document.addEventListener('paste', onNativePaste);

        formulaInput.addEventListener('keydown', onFormulaInput);
        formulaInput.addEventListener('input', onFormulaInputLive);
        formulaInput.addEventListener('focus', onFormulaFocus);
        formulaInput.addEventListener('blur', onFormulaBlur);

        document.getElementById('btn-new').addEventListener('click', newSheet);
        document.getElementById('btn-open').addEventListener('click', openFile);
        sheetAddBtn.addEventListener('click', newSheet);
        document.getElementById('btn-save').addEventListener('click', saveFile);
        document.getElementById('btn-save-all').addEventListener('click', saveAllFiles);
        document.getElementById('btn-undo').addEventListener('click', undo);
        document.getElementById('btn-redo').addEventListener('click', redo);
        document.getElementById('btn-add-row').addEventListener('click', addRow);
        document.getElementById('btn-add-col').addEventListener('click', addCol);
        darkToggle.addEventListener('change', toggleDark);
        fileInput.addEventListener('change', handleFileOpen);

        // Drag & drop CSV
        gridContainer.addEventListener('dragover', e => {
            if (e.dataTransfer.types.includes('Files')) {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'copy';
            }
        });
        gridContainer.addEventListener('drop', e => {
            e.preventDefault();
            const file = e.dataTransfer.files[0];
            if (!file) return;
            if (file.size > MAX_FILE_SIZE) {
                alert('File is too large (' + (file.size / 1024 / 1024).toFixed(1) + 'MB). Maximum is 5MB.');
                return;
            }
            const reader = new FileReader();
            reader.onload = evt => {
                pushUndo();
                parseCSV(evt.target.result);
                renderGrid();
                setActiveCell(0, 0);
            };
            reader.readAsText(file);
        });
    }

    // ── Helpers ──
    function el(tag, className) {
        const e = document.createElement(tag);
        if (className) e.className = className;
        return e;
    }

    // ── Go ──
    init();

    // ── Support toast ──
    (function () {
        var el = document.getElementById('support-toast');
        var closeBtn = document.getElementById('support-close');
        if (!el) return;

        function show() {
            el.classList.add('visible');
            setTimeout(hide, 5000);
        }

        function hide() {
            el.classList.remove('visible');
        }

        closeBtn.addEventListener('click', hide);
        setTimeout(show, 3000);
    })();
})();
