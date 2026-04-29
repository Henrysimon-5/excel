/**
 * spreadsheet.js – core logic for the Excel-like spreadsheet
 *
 * Features:
 *  - 50 rows × 26 columns (A–Z)
 *  - Cell editing with inline input
 *  - Formula evaluation: =SUM, =AVERAGE, =MIN, =MAX, =COUNT, =IF,
 *    =ROUND, =ABS, +, -, *, /, arithmetic expressions, cell refs
 *  - Keyboard navigation (arrows, Tab, Enter, Esc, Delete, F2)
 *  - Toolbar: bold, italic, underline, strikethrough, font-size,
 *    text-colour, fill-colour, align (L/C/R), clear
 *  - Copy / Cut / Paste (Ctrl+C/X/V)
 *  - Undo / Redo (Ctrl+Z / Ctrl+Y)
 *  - Context menu
 *  - Status-bar aggregates (SUM / AVERAGE / COUNT for selection)
 *  - Export to CSV download
 *  - Column-resize via drag
 */

(function () {
  "use strict";

  /* ── constants ──────────────────────────────────────── */
  const ROWS = 50;
  const COLS = 26;                      // A–Z
  const COL_LETTERS = Array.from({ length: COLS }, (_, i) =>
    String.fromCharCode(65 + i)
  );

  /* ── state ──────────────────────────────────────────── */
  // cells[r][c] = { raw, value, style }
  let cells = createEmptyCells();
  let selRow = 0, selCol = 0;
  let editing = false;
  let clipboard = null;  // { mode: 'copy'|'cut', data: [{r,c,raw,style},...] }
  let history = [];      // undo stack  [{cells snapshot}]
  let future  = [];      // redo stack

  /* ── DOM refs ───────────────────────────────────────── */
  const table      = document.getElementById("spreadsheet");
  const formulaIn  = document.getElementById("formula-input");
  const cellAddr   = document.getElementById("cell-address");
  const statusMode = document.getElementById("status-mode");
  const statusSum  = document.getElementById("status-sum");
  const statusAvg  = document.getElementById("status-avg");
  const statusCnt  = document.getElementById("status-count");
  const ctxMenu    = document.getElementById("context-menu");
  const toastCont  = document.getElementById("toast-container");

  /* ── build table ────────────────────────────────────── */
  function buildTable() {
    table.innerHTML = "";

    // header row
    const thead = table.createTHead();
    const hr = thead.insertRow();
    const corner = document.createElement("th");
    corner.className = "corner";
    hr.appendChild(corner);
    COL_LETTERS.forEach((letter, c) => {
      const th = document.createElement("th");
      th.className = "col-header";
      th.dataset.col = c;
      th.textContent = letter;
      // resize handle
      const handle = document.createElement("div");
      handle.style.cssText =
        "position:absolute;right:0;top:0;width:5px;height:100%;cursor:col-resize;";
      th.style.position = "relative";
      th.appendChild(handle);
      attachColResize(handle, th, c);
      hr.appendChild(th);
    });

    // body
    const tbody = table.createTBody();
    for (let r = 0; r < ROWS; r++) {
      const tr = tbody.insertRow();
      const rh = tr.insertCell();
      rh.className = "row-header";
      rh.dataset.row = r;
      rh.textContent = r + 1;

      for (let c = 0; c < COLS; c++) {
        const td = tr.insertCell();
        td.className = "cell";
        td.dataset.row = r;
        td.dataset.col = c;
        td.tabIndex = -1;
        renderCell(td, r, c);

        td.addEventListener("mousedown", onCellMouseDown);
        td.addEventListener("dblclick", () => startEdit(r, c));
      }
    }
  }

  function getCell(r, c) { return cells[r] && cells[r][c]; }

  function getTD(r, c) {
    return table.querySelector(`td.cell[data-row="${r}"][data-col="${c}"]`);
  }

  function renderCell(td, r, c) {
    const cell = getCell(r, c);
    if (!cell) return;
    td.textContent = cell.value !== undefined ? String(cell.value) : "";
    applyStyleToTD(td, cell.style);
  }

  function applyStyleToTD(td, style) {
    td.style.fontWeight     = style.bold          ? "bold"   : "";
    td.style.fontStyle      = style.italic        ? "italic" : "";
    td.style.textDecoration =
      (style.underline ? "underline " : "") +
      (style.strikethrough ? "line-through" : "");
    td.style.color       = style.color      || "";
    td.style.background  = style.fillColor  || "";
    td.style.textAlign   = style.align      || "";
    td.style.fontSize    = style.fontSize   ? style.fontSize + "px" : "";
  }

  /* ── cell model ─────────────────────────────────────── */
  function createEmptyCells() {
    return Array.from({ length: ROWS }, () =>
      Array.from({ length: COLS }, () => ({
        raw: "",
        value: "",
        style: {}
      }))
    );
  }

  function deepCloneCells(src) {
    return src.map(row =>
      row.map(cell => ({
        raw:   cell.raw,
        value: cell.value,
        style: { ...cell.style }
      }))
    );
  }

  function saveHistory() {
    history.push(deepCloneCells(cells));
    if (history.length > 100) history.shift();
    future = [];
  }

  function undo() {
    if (!history.length) return;
    future.push(deepCloneCells(cells));
    cells = history.pop();
    rebuildValues();
    renderAll();
    toast("Undo");
  }

  function redo() {
    if (!future.length) return;
    history.push(deepCloneCells(cells));
    cells = future.pop();
    rebuildValues();
    renderAll();
    toast("Redo");
  }

  function rebuildValues() {
    // re-evaluate all formulas after load / undo
    for (let r = 0; r < ROWS; r++)
      for (let c = 0; c < COLS; c++)
        cells[r][c].value = computeValue(cells[r][c].raw, r, c);
  }

  function renderAll() {
    for (let r = 0; r < ROWS; r++)
      for (let c = 0; c < COLS; c++)
        renderCell(getTD(r, c), r, c);
    updateFormulaBar();
    updateStatus();
  }

  /* ── formula engine ─────────────────────────────────── */
  function computeValue(raw, _r, _c) {
    if (!raw || raw === "") return "";
    if (!String(raw).startsWith("=")) {
      const n = Number(raw);
      return isNaN(n) || raw.trim() === "" ? raw : n;
    }
    try {
      return evalFormula(String(raw).slice(1).trim());
    } catch (e) {
      return "#ERR";
    }
  }

  function evalFormula(expr) {
    // Replace cell references like A1, B2 with their values
    expr = expr.replace(/([A-Z]+)(\d+)/g, (_, col, row) => {
      const c = colLetterToIndex(col);
      const r = parseInt(row, 10) - 1;
      if (r < 0 || r >= ROWS || c < 0 || c >= COLS) return "0";
      const v = cells[r][c].value;
      return (v === "" || v === undefined) ? "0" : Number(v) || 0;
    });

    // Built-in functions
    expr = expr.replace(/SUM\(([^)]+)\)/gi, (_, args) =>
      rangeValues(args).reduce((a, b) => a + b, 0)
    );
    expr = expr.replace(/AVERAGE\(([^)]+)\)/gi, (_, args) => {
      const vals = rangeValues(args).filter(v => !isNaN(v));
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    });
    expr = expr.replace(/MIN\(([^)]+)\)/gi, (_, args) =>
      Math.min(...rangeValues(args))
    );
    expr = expr.replace(/MAX\(([^)]+)\)/gi, (_, args) =>
      Math.max(...rangeValues(args))
    );
    expr = expr.replace(/COUNT\(([^)]+)\)/gi, (_, args) =>
      rangeValues(args).filter(v => !isNaN(v) && v !== "").length
    );
    expr = expr.replace(/ROUND\(([^,]+),([^)]+)\)/gi, (_, num, dec) =>
      Number(Number(evalFormula(num.trim())).toFixed(Number(evalFormula(dec.trim()))))
    );
    expr = expr.replace(/ABS\(([^)]+)\)/gi, (_, num) =>
      Math.abs(Number(evalFormula(num.trim())))
    );
    // IF(condition, true_val, false_val)
    expr = expr.replace(/IF\((.+)\)/gi, (_, args) => {
      const parts = splitArgs(args);
      if (parts.length < 3) return "#ERR";
      const cond = evalFormula(parts[0]);
      return cond ? evalFormula(parts[1]) : evalFormula(parts[2]);
    });

    // By this point all cell refs and functions have been replaced with numbers.
    // Only allow digits, basic arithmetic operators, parentheses, dot (decimal), and whitespace.
    if (/[^0-9+\-*/.()\s]/.test(expr)) return "#ERR";
    // eslint-disable-next-line no-new-func
    const result = Function('"use strict"; return (' + expr + ")")();
    return isNaN(result) ? result : result;
  }

  function splitArgs(str) {
    const parts = [];
    let depth = 0, current = "";
    for (const ch of str) {
      if (ch === "(" ) depth++;
      if (ch === ")" ) depth--;
      if (ch === "," && depth === 0) { parts.push(current.trim()); current = ""; }
      else current += ch;
    }
    if (current.trim()) parts.push(current.trim());
    return parts;
  }

  function rangeValues(args) {
    // args can be: "A1:B3" or "A1,B2,3"
    const vals = [];
    const parts = args.split(",").map(s => s.trim());
    for (const part of parts) {
      if (/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.test(part)) {
        const m = part.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        const c1 = colLetterToIndex(m[1]), r1 = parseInt(m[2], 10) - 1;
        const c2 = colLetterToIndex(m[3]), r2 = parseInt(m[4], 10) - 1;
        for (let r = r1; r <= r2; r++)
          for (let c = c1; c <= c2; c++) {
            const v = cells[r] && cells[r][c] ? cells[r][c].value : 0;
            vals.push(Number(v) || 0);
          }
      } else if (/^([A-Z]+)(\d+)$/.test(part)) {
        const m = part.match(/^([A-Z]+)(\d+)$/);
        const r = parseInt(m[2], 10) - 1, c = colLetterToIndex(m[1]);
        vals.push(Number((cells[r] && cells[r][c] && cells[r][c].value) || 0));
      } else {
        const n = Number(part);
        if (!isNaN(n)) vals.push(n);
      }
    }
    return vals;
  }

  function colLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++)
      index = index * 26 + letter.charCodeAt(i) - 64;
    return index - 1;
  }

  /* ── selection ──────────────────────────────────────── */
  function select(r, c, updateFormula = true) {
    if (r < 0 || r >= ROWS || c < 0 || c >= COLS) return;
    if (editing) commitEdit();

    const prev = getTD(selRow, selCol);
    if (prev) prev.classList.remove("selected");

    selRow = r; selCol = c;

    const td = getTD(r, c);
    if (td) {
      td.classList.add("selected");
      td.focus({ preventScroll: false });
      scrollIntoView(td);
    }

    // highlight headers
    updateHeaderHighlights(r, c);

    if (updateFormula) updateFormulaBar();
    updateStatus();
  }

  function updateHeaderHighlights(r, c) {
    table.querySelectorAll(".col-header").forEach(th => {
      th.classList.toggle("selected-col", parseInt(th.dataset.col) === c);
    });
    table.querySelectorAll(".row-header").forEach(rh => {
      rh.classList.toggle("selected-row", parseInt(rh.dataset.row) === r);
    });
  }

  function scrollIntoView(td) {
    const container = document.getElementById("grid-container");
    const rect = td.getBoundingClientRect();
    const cRect = container.getBoundingClientRect();
    if (rect.bottom > cRect.bottom) container.scrollTop += rect.bottom - cRect.bottom + 4;
    if (rect.top < cRect.top + 24)  container.scrollTop -= cRect.top + 24 - rect.top + 4;
    if (rect.right > cRect.right)   container.scrollLeft += rect.right - cRect.right + 4;
    if (rect.left < cRect.left + 50) container.scrollLeft -= cRect.left + 50 - rect.left + 4;
  }

  /* ── editing ────────────────────────────────────────── */
  let editorEl = null;

  function startEdit(r, c, initialChar) {
    if (editing) commitEdit();
    editing = true;
    selRow = r; selCol = c;

    const td = getTD(r, c);
    td.classList.add("editing");
    td.classList.add("selected");

    editorEl = document.createElement("input");
    editorEl.className = "cell-editor";
    editorEl.value = initialChar !== undefined
      ? initialChar
      : (cells[r][c].raw || "");
    td.appendChild(editorEl);
    editorEl.focus();
    editorEl.select();

    statusMode.textContent = "EDIT";

    editorEl.addEventListener("keydown", onEditorKeyDown);
    editorEl.addEventListener("input", () => {
      formulaIn.value = editorEl.value;
    });

    formulaIn.value = editorEl.value;
  }

  function commitEdit(newVal) {
    if (!editing) return;
    editing = false;
    statusMode.textContent = "READY";

    const val = newVal !== undefined
      ? newVal
      : (editorEl ? editorEl.value : "");

    const td = getTD(selRow, selCol);
    if (td && editorEl) {
      td.classList.remove("editing");
      td.removeChild(editorEl);
      editorEl = null;
    }

    saveHistory();
    cells[selRow][selCol].raw   = val;
    cells[selRow][selCol].value = computeValue(val, selRow, selCol);

    // re-compute dependent formulas
    recomputeFormulas();
    renderAll();
  }

  function cancelEdit() {
    if (!editing) return;
    editing = false;
    statusMode.textContent = "READY";
    const td = getTD(selRow, selCol);
    if (td && editorEl) {
      td.classList.remove("editing");
      td.removeChild(editorEl);
      editorEl = null;
    }
    updateFormulaBar();
  }

  function recomputeFormulas() {
    for (let r = 0; r < ROWS; r++)
      for (let c = 0; c < COLS; c++)
        if (String(cells[r][c].raw).startsWith("="))
          cells[r][c].value = computeValue(cells[r][c].raw, r, c);
  }

  /* ── formula bar ────────────────────────────────────── */
  function updateFormulaBar() {
    const cell = getCell(selRow, selCol);
    formulaIn.value = cell ? cell.raw : "";
    cellAddr.textContent =
      COL_LETTERS[selCol] + (selRow + 1);
  }

  formulaIn.addEventListener("keydown", e => {
    if (e.key === "Enter") {
      const val = formulaIn.value;
      saveHistory();
      cells[selRow][selCol].raw   = val;
      cells[selRow][selCol].value = computeValue(val, selRow, selCol);
      recomputeFormulas();
      renderAll();
      getTD(selRow, selCol)?.focus();
    }
    if (e.key === "Escape") {
      formulaIn.value = cells[selRow][selCol].raw || "";
      getTD(selRow, selCol)?.focus();
    }
  });

  /* ── keyboard navigation ────────────────────────────── */
  function onEditorKeyDown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      commitEdit();
      select(selRow + 1, selCol);
    } else if (e.key === "Tab") {
      e.preventDefault();
      commitEdit();
      select(selRow, selCol + (e.shiftKey ? -1 : 1));
    } else if (e.key === "Escape") {
      cancelEdit();
    } else if (e.key === "ArrowUp") {
      commitEdit(); select(selRow - 1, selCol);
    } else if (e.key === "ArrowDown") {
      commitEdit(); select(selRow + 1, selCol);
    }
  }

  document.addEventListener("keydown", onGlobalKeyDown);

  function onGlobalKeyDown(e) {
    if (editing) return;
    if (document.activeElement === formulaIn) return;

    // Ctrl shortcuts
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case "z": e.preventDefault(); undo(); return;
        case "y": e.preventDefault(); redo(); return;
        case "c": e.preventDefault(); copySelection(); return;
        case "x": e.preventDefault(); cutSelection(); return;
        case "v": e.preventDefault(); pasteSelection(); return;
        case "b": e.preventDefault(); toggleStyle("bold"); return;
        case "i": e.preventDefault(); toggleStyle("italic"); return;
        case "u": e.preventDefault(); toggleStyle("underline"); return;
      }
    }

    const td = getTD(selRow, selCol);
    if (!td || document.activeElement !== td) return;

    switch (e.key) {
      case "ArrowUp":    e.preventDefault(); select(selRow - 1, selCol); break;
      case "ArrowDown":  e.preventDefault(); select(selRow + 1, selCol); break;
      case "ArrowLeft":  e.preventDefault(); select(selRow, selCol - 1); break;
      case "ArrowRight": e.preventDefault(); select(selRow, selCol + 1); break;
      case "Tab":
        e.preventDefault();
        select(selRow, selCol + (e.shiftKey ? -1 : 1));
        break;
      case "Enter":
        e.preventDefault();
        startEdit(selRow, selCol);
        break;
      case "F2":
        e.preventDefault();
        startEdit(selRow, selCol);
        break;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        clearCell(selRow, selCol);
        break;
      case "Escape":
        break;
      default:
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          startEdit(selRow, selCol, e.key);
        }
    }
  }

  /* ── mouse ──────────────────────────────────────────── */
  function onCellMouseDown(e) {
    const td = e.currentTarget;
    const r = parseInt(td.dataset.row);
    const c = parseInt(td.dataset.col);
    if (editing && (r !== selRow || c !== selCol)) commitEdit();
    select(r, c);
  }

  /* ── clear ──────────────────────────────────────────── */
  function clearCell(r, c) {
    saveHistory();
    cells[r][c].raw   = "";
    cells[r][c].value = "";
    recomputeFormulas();
    renderAll();
  }

  /* ── clipboard ──────────────────────────────────────── */
  function copySelection() {
    clipboard = {
      mode: "copy",
      data: [{ r: selRow, c: selCol,
               raw: cells[selRow][selCol].raw,
               style: { ...cells[selRow][selCol].style } }]
    };
    toast("Copied");
  }

  function cutSelection() {
    clipboard = {
      mode: "cut",
      data: [{ r: selRow, c: selCol,
               raw: cells[selRow][selCol].raw,
               style: { ...cells[selRow][selCol].style } }]
    };
    toast("Cut");
  }

  function pasteSelection() {
    if (!clipboard) return;
    saveHistory();
    for (const item of clipboard.data) {
      const dr = selRow - clipboard.data[0].r;
      const dc = selCol - clipboard.data[0].c;
      const tr = item.r + dr, tc = item.c + dc;
      if (tr >= 0 && tr < ROWS && tc >= 0 && tc < COLS) {
        cells[tr][tc].raw   = item.raw;
        cells[tr][tc].value = computeValue(item.raw, tr, tc);
        cells[tr][tc].style = { ...item.style };
      }
    }
    if (clipboard.mode === "cut") {
      for (const item of clipboard.data) {
        cells[item.r][item.c].raw   = "";
        cells[item.r][item.c].value = "";
      }
      clipboard = null;
    }
    recomputeFormulas();
    renderAll();
    toast("Pasted");
  }

  /* ── styling ────────────────────────────────────────── */
  function toggleStyle(prop) {
    saveHistory();
    cells[selRow][selCol].style[prop] = !cells[selRow][selCol].style[prop];
    renderCell(getTD(selRow, selCol), selRow, selCol);
    updateToolbarState();
  }

  function setStyle(prop, val) {
    saveHistory();
    cells[selRow][selCol].style[prop] = val;
    renderCell(getTD(selRow, selCol), selRow, selCol);
    updateToolbarState();
  }

  function updateToolbarState() {
    const style = cells[selRow][selCol].style;
    document.getElementById("btn-bold").classList.toggle("active", !!style.bold);
    document.getElementById("btn-italic").classList.toggle("active", !!style.italic);
    document.getElementById("btn-underline").classList.toggle("active", !!style.underline);
    document.getElementById("btn-strike").classList.toggle("active", !!style.strikethrough);
    document.getElementById("btn-align-left").classList.toggle("active", style.align === "left" || !style.align);
    document.getElementById("btn-align-center").classList.toggle("active", style.align === "center");
    document.getElementById("btn-align-right").classList.toggle("active", style.align === "right");
    document.getElementById("color-indicator").style.background =
      style.color || "#000";
    document.getElementById("fill-indicator").style.background =
      style.fillColor || "transparent";
    const fsSel = document.getElementById("font-size-select");
    if (style.fontSize) fsSel.value = String(style.fontSize);
  }

  /* ── status bar ─────────────────────────────────────── */
  function updateStatus() {
    const v = cells[selRow][selCol].value;
    const num = Number(v);
    if (!isNaN(num) && v !== "") {
      statusSum.textContent = "Sum: " + num;
      statusAvg.textContent = "Avg: " + num;
      statusCnt.textContent = "Count: 1";
    } else {
      statusSum.textContent = "";
      statusAvg.textContent = "";
      statusCnt.textContent = "";
    }
  }

  /* ── toast ──────────────────────────────────────────── */
  function toast(msg) {
    const el = document.createElement("div");
    el.className = "toast";
    el.textContent = msg;
    toastCont.appendChild(el);
    setTimeout(() => el.remove(), 3000);
  }

  /* ── context menu ───────────────────────────────────── */
  document.addEventListener("contextmenu", e => {
    const td = e.target.closest("td.cell");
    if (!td) return;
    e.preventDefault();
    const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
    select(r, c);
    showContextMenu(e.clientX, e.clientY);
  });

  function showContextMenu(x, y) {
    ctxMenu.style.left = x + "px";
    ctxMenu.style.top  = y + "px";
    ctxMenu.classList.add("visible");
  }

  document.addEventListener("click", () => ctxMenu.classList.remove("visible"));

  ctxMenu.addEventListener("click", e => {
    const item = e.target.closest(".ctx-item");
    if (!item) return;
    const action = item.dataset.action;
    if (action === "copy")       copySelection();
    if (action === "cut")        cutSelection();
    if (action === "paste")      pasteSelection();
    if (action === "clear")      clearCell(selRow, selCol);
    if (action === "insert-row") insertRow(selRow);
    if (action === "delete-row") deleteRow(selRow);
    ctxMenu.classList.remove("visible");
  });

  /* ── insert / delete row ────────────────────────────── */
  function insertRow(r) {
    saveHistory();
    cells.splice(r, 0, Array.from({ length: COLS }, () => ({
      raw: "", value: "", style: {}
    })));
    if (cells.length > ROWS) cells.length = ROWS;
    rebuildValues();
    buildTable();
    select(r, selCol);
    toast("Row inserted");
  }

  function deleteRow(r) {
    if (cells.length <= 1) return;
    saveHistory();
    cells.splice(r, 1);
    cells.push(Array.from({ length: COLS }, () => ({ raw: "", value: "", style: {} })));
    rebuildValues();
    buildTable();
    select(Math.min(r, ROWS - 1), selCol);
    toast("Row deleted");
  }

  /* ── column resize ──────────────────────────────────── */
  function attachColResize(handle, th, colIdx) {
    let startX, startW;
    handle.addEventListener("mousedown", e => {
      e.stopPropagation();
      startX = e.clientX;
      startW = th.offsetWidth;
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
    function onMove(e) {
      const w = Math.max(40, startW + e.clientX - startX);
      th.style.width    = w + "px";
      th.style.minWidth = w + "px";
      // resize all cells in this column
      table.querySelectorAll(`td.cell[data-col="${colIdx}"]`).forEach(td => {
        td.style.width    = w + "px";
        td.style.minWidth = w + "px";
        td.style.maxWidth = w + "px";
      });
    }
    function onUp() {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    }
  }

  /* ── export CSV ─────────────────────────────────────── */
  function exportCSV() {
    const rows = [];
    for (let r = 0; r < ROWS; r++) {
      const row = cells[r].map(cell => {
        const v = String(cell.value ?? "");
        return v.includes(",") || v.includes('"') || v.includes("\n")
          ? `"${v.replace(/"/g, '""')}"` : v;
      });
      rows.push(row.join(","));
    }
    // trim trailing empty rows
    while (rows.length > 1 && rows[rows.length - 1].replace(/,/g, "").trim() === "")
      rows.pop();

    const blob = new Blob([rows.join("\n")], { type: "text/csv" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = (document.getElementById("sheet-name").value || "sheet") + ".csv";
    a.click();
    URL.revokeObjectURL(url);
    toast("CSV exported");
  }

  /* ── toolbar wiring ─────────────────────────────────── */
  function wireToolbar() {
    document.getElementById("btn-bold")
      .addEventListener("click", () => toggleStyle("bold"));
    document.getElementById("btn-italic")
      .addEventListener("click", () => toggleStyle("italic"));
    document.getElementById("btn-underline")
      .addEventListener("click", () => toggleStyle("underline"));
    document.getElementById("btn-strike")
      .addEventListener("click", () => toggleStyle("strikethrough"));

    document.getElementById("btn-align-left")
      .addEventListener("click", () => setStyle("align", "left"));
    document.getElementById("btn-align-center")
      .addEventListener("click", () => setStyle("align", "center"));
    document.getElementById("btn-align-right")
      .addEventListener("click", () => setStyle("align", "right"));

    document.getElementById("btn-clear")
      .addEventListener("click", () => clearCell(selRow, selCol));

    document.getElementById("btn-undo")
      .addEventListener("click", undo);
    document.getElementById("btn-redo")
      .addEventListener("click", redo);

    document.getElementById("btn-export")
      .addEventListener("click", exportCSV);

    // font size
    document.getElementById("font-size-select")
      .addEventListener("change", function () {
        setStyle("fontSize", Number(this.value));
      });

    // text colour
    document.getElementById("color-picker")
      .addEventListener("input", function () {
        document.getElementById("color-indicator").style.background = this.value;
        setStyle("color", this.value);
      });

    // fill colour
    document.getElementById("fill-picker")
      .addEventListener("input", function () {
        document.getElementById("fill-indicator").style.background = this.value;
        setStyle("fillColor", this.value);
      });
  }

  /* ── init ───────────────────────────────────────────── */
  function init() {
    buildTable();
    wireToolbar();
    select(0, 0);
    statusMode.textContent = "READY";
  }

  document.addEventListener("DOMContentLoaded", init);
})();
