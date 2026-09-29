/*
 * Fake Excel host for tests. A plain browser script (no modules, no build step) that defines
 * window.Office and window.Excel with the subset of the Excel JavaScript API that
 * src/office/adapter.ts uses, backed by a workbook JSON:
 *
 *   window.__NYMFORM_FAKE_WORKBOOK__ = {
 *     sheets: [{ name, visibility, cells: { "A1": { value, text, type, format, formula? } }, autoFilter? }],
 *     tables: [{ name, sheet, address, showTotals? }],
 *     selection: { sheet, address },
 *     calculationMode: "Automatic" | "AutomaticExceptTables" | "Manual"
 *   }
 *
 * A sheet's `autoFilter` is the address of its AutoFilter range ("A1:F501"). A table's `address`
 * includes its totals row when `showTotals` is true, as Table.getRange() does in Excel.
 * Range.getSurroundingRegion() follows Excel's current region (see currentRegion below).
 * Range.getUsedRangeOrNullObject() gives the box around the used cells inside the range; Excel
 * documents it only as "the used range of the given range", which may instead be the range cut to
 * the sheet's used range. Set fake.usedRange = "sheet" for that reading; the adapter must work
 * with both.
 *
 * It follows Office's proxy rules: load() and property writes are queued and run in order on
 * context.sync(); a property can be read only after it was loaded and synced (reading it earlier
 * throws PropertyNotLoaded, like Office). Writes change the fake state and are logged in
 * window.__NYMFORM_FAKE__.writes; loads are logged in .loads.
 *
 * There is no formula engine. A written formula cell gets type "Double" and value 0, unless a test
 * registers window.__NYMFORM_FAKE__.evaluate = (formula, sheet, addr) => result, where result is
 * { value, type, text? } for one cell or { spill: [[{ value, type, text? }, ...], ...] } for a
 * dynamic array that spills from the cell (blocked spills become #SPILL!).
 *
 * In Node (tests/adapter.test.ts) it attaches the same globals to globalThis.
 */
/* eslint-disable @typescript-eslint/no-this-alias -- closures over proxy objects, in plain script style */
(function (root) {
  "use strict";

  var MAX_ROWS = 1048576;
  var MAX_COLS = 16384;
  var MAX_LOAD_CELLS = 1000000;

  // ------------------------------------------------------------------------------------------
  // A1 helpers

  function colLetter(index) {
    var n = index + 1;
    var out = "";
    while (n > 0) {
      var rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  function colIndex(letters) {
    if (!/^[A-Za-z]{1,3}$/.test(letters)) return -1;
    var n = 0;
    var up = letters.toUpperCase();
    for (var i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64);
    return n - 1 < MAX_COLS ? n - 1 : -1;
  }

  function parseCellRef(ref) {
    var m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref).trim());
    if (!m) return null;
    var c = colIndex(m[1]);
    var r = Number(m[2]) - 1;
    if (c < 0 || r < 0 || r >= MAX_ROWS) return null;
    return { r: r, c: c };
  }

  function keyOf(r, c) {
    return colLetter(c) + (r + 1);
  }

  /** "", "A1", "A1:B2", "A:B", "1:3" (with or without $) to a 0-based inclusive rect, or null. */
  function parseAddress(address) {
    var a = String(address === undefined || address === null ? "" : address).trim();
    if (a === "") return { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 };
    if (a.indexOf("!") >= 0) return null;
    var parts = a.split(":");
    if (parts.length > 2) return null;
    if (parts.length === 1) {
      var one = parseCellRef(parts[0]);
      return one ? { r1: one.r, c1: one.c, r2: one.r, c2: one.c } : null;
    }
    var x = parseCellRef(parts[0]);
    var y = parseCellRef(parts[1]);
    if (x && y) {
      return { r1: Math.min(x.r, y.r), c1: Math.min(x.c, y.c), r2: Math.max(x.r, y.r), c2: Math.max(x.c, y.c) };
    }
    var cols = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(a);
    if (cols) {
      var ca = colIndex(cols[1]);
      var cb = colIndex(cols[2]);
      if (ca < 0 || cb < 0) return null;
      return { r1: 0, c1: Math.min(ca, cb), r2: MAX_ROWS - 1, c2: Math.max(ca, cb) };
    }
    var rows = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(a);
    if (rows) {
      var ra = Number(rows[1]) - 1;
      var rb = Number(rows[2]) - 1;
      if (ra < 0 || rb < 0 || ra >= MAX_ROWS || rb >= MAX_ROWS) return null;
      return { r1: Math.min(ra, rb), c1: 0, r2: Math.max(ra, rb), c2: MAX_COLS - 1 };
    }
    return null;
  }

  function formatRect(rect) {
    var fullCols = rect.r1 === 0 && rect.r2 === MAX_ROWS - 1;
    var fullRows = rect.c1 === 0 && rect.c2 === MAX_COLS - 1;
    // Excel names the whole sheet by its rows: 1:1048576.
    if (fullRows) return rect.r1 + 1 + ":" + (rect.r2 + 1);
    if (fullCols) return colLetter(rect.c1) + ":" + colLetter(rect.c2);
    var a = keyOf(rect.r1, rect.c1);
    var b = keyOf(rect.r2, rect.c2);
    return a === b ? a : a + ":" + b;
  }

  function quoteSheet(name) {
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseCellRef(name)) return name;
    return "'" + name.replace(/'/g, "''") + "'";
  }

  function intersects(a, b) {
    return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
  }

  function contains(outer, inner) {
    return outer.r1 <= inner.r1 && outer.c1 <= inner.c1 && outer.r2 >= inner.r2 && outer.c2 >= inner.c2;
  }

  function cellCount(rect) {
    return (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
  }

  function clone(v) {
    return v === undefined ? undefined : JSON.parse(JSON.stringify(v));
  }

  // ------------------------------------------------------------------------------------------
  // Display text for values the fake computes itself (written constants, evaluate results)

  function grouped(n, decimals) {
    var fixed = Math.abs(n).toFixed(decimals);
    var parts = fixed.split(".");
    var int = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (n < 0 ? "-" : "") + int + (parts.length > 1 ? "." + parts[1] : "");
  }

  function generalText(n) {
    if (!isFinite(n)) return "#NUM!";
    if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
    if (Math.abs(n) >= 1e11 || (Math.abs(n) < 1e-9 && n !== 0)) return n.toExponential(5).replace("e", "E");
    return String(parseFloat(n.toPrecision(10)));
  }

  function dateText(serial) {
    var d = new Date(Math.round((Math.floor(serial) - 25569) * 86400000));
    return d.getUTCMonth() + 1 + "/" + d.getUTCDate() + "/" + d.getUTCFullYear();
  }

  function formatText(value, format) {
    if (value === null || value === undefined) return "";
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (typeof value !== "number") return String(value);
    var f = format || "General";
    if (f === "General") return generalText(value);
    if (/^m\/d\/yyyy$/i.test(f)) return dateText(value);
    if (f === "#,##0.00") return grouped(value, 2);
    if (f === "#,##0") return grouped(value, 0);
    if (f === "0.00") return value.toFixed(2);
    if (f === "0") return value.toFixed(0);
    if (f === "0%") return (value * 100).toFixed(0) + "%";
    if (f === "0.00%") return (value * 100).toFixed(2) + "%";
    return generalText(value);
  }

  function inferType(value) {
    if (value === null || value === undefined || value === "") return "Empty";
    if (typeof value === "number") return "Double";
    if (typeof value === "boolean") return "Boolean";
    if (typeof value === "string" && /^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)$/.test(value)) {
      return "Error";
    }
    return "String";
  }

  // ------------------------------------------------------------------------------------------
  // State

  var state = null;

  function normCell(cell) {
    var c = cell || {};
    var hasFormula = typeof c.formula === "string" && c.formula !== "";
    var value = c.value === undefined ? null : c.value;
    // A formula cell without a stored result shows 0, as the fake has no formula engine.
    if (hasFormula && value === null) value = 0;
    var out = { value: value, type: c.type || inferType(value), format: c.format || "General" };
    out.text = c.text !== undefined ? String(c.text) : formatText(value, out.format);
    if (hasFormula) out.formula = c.formula;
    return out;
  }

  function normalizeWorkbook(input) {
    var wb = clone(input || {}) || {};
    var sheets = (wb.sheets || []).map(function (s) {
      if (!s || typeof s.name !== "string" || s.name === "") throw new Error("Fake workbook: every sheet needs a name.");
      var cells = {};
      Object.keys(s.cells || {}).forEach(function (k) {
        var p = parseCellRef(k);
        if (!p) throw new Error("Fake workbook: bad cell address " + k + " on " + s.name);
        cells[keyOf(p.r, p.c)] = normCell(s.cells[k]);
      });
      var sheet = { name: s.name, visibility: s.visibility || "Visible", cells: cells };
      if (s.autoFilter) {
        var filter = parseAddress(s.autoFilter);
        if (!filter) throw new Error("Fake workbook: bad autoFilter address on " + s.name);
        sheet.autoFilter = formatRect(filter);
      }
      return sheet;
    });
    if (sheets.length === 0) sheets.push({ name: "Sheet1", visibility: "Visible", cells: {} });
    var tables = (wb.tables || []).map(function (t) {
      if (!parseAddress(t.address)) throw new Error("Fake workbook: bad table address for " + t.name);
      var table = { name: t.name, sheet: t.sheet, address: t.address };
      if (t.showTotals) table.showTotals = true;
      return table;
    });
    // Defined names: { name, sheet? }. With a sheet the name is sheet-scoped (worksheet.names).
    var names = (wb.names || []).map(function (n) {
      return { name: String(n.name), sheet: n.sheet || null };
    });
    var firstVisible = sheets.filter(function (s) { return s.visibility === "Visible"; })[0] || sheets[0];
    var sel = wb.selection || { sheet: firstVisible.name, address: "A1" };
    return {
      sheets: sheets,
      tables: tables,
      names: names,
      selection: { sheet: sel.sheet, address: sel.address },
      activeSheet: sel.sheet,
      calculationMode: wb.calculationMode || "Automatic",
    };
  }

  function st() {
    if (!state) state = normalizeWorkbook(root.__NYMFORM_FAKE_WORKBOOK__);
    return state;
  }

  function findSheet(name) {
    var key = String(name).toUpperCase();
    var list = st().sheets;
    for (var i = 0; i < list.length; i++) if (list[i].name.toUpperCase() === key) return list[i];
    return null;
  }

  function isEmptyCell(cell) {
    return !cell || (cell.formula === undefined && (cell.value === null || cell.value === undefined || cell.value === ""));
  }

  function filledAt(sheet, r, c) {
    return r >= 0 && r < MAX_ROWS && c >= 0 && c < MAX_COLS && !isEmptyCell(sheet.cells[keyOf(r, c)]);
  }

  /**
   * Excel's current region (Ctrl+* / getSurroundingRegion) around one cell: the rectangle grows while
   * any cell of the one-cell ring around it, corners included, is filled; the sheet's edges clip it.
   */
  function currentRegion(sheet, r, c) {
    var rect = { r1: r, c1: c, r2: r, c2: c };
    var rowHas = function (row, from, to) {
      for (var x = from; x <= to; x++) if (filledAt(sheet, row, x)) return true;
      return false;
    };
    var colHas = function (col, from, to) {
      for (var y = from; y <= to; y++) if (filledAt(sheet, y, col)) return true;
      return false;
    };
    var grew = true;
    while (grew) {
      grew = false;
      var top = rect.r1 - 1;
      var bottom = rect.r2 + 1;
      var left = rect.c1 - 1;
      var right = rect.c2 + 1;
      if (top >= 0 && rowHas(top, left, right)) {
        rect.r1 = top;
        grew = true;
      }
      if (bottom < MAX_ROWS && rowHas(bottom, left, right)) {
        rect.r2 = bottom;
        grew = true;
      }
      if (left >= 0 && colHas(left, top, bottom)) {
        rect.c1 = left;
        grew = true;
      }
      if (right < MAX_COLS && colHas(right, top, bottom)) {
        rect.c2 = right;
        grew = true;
      }
    }
    return rect;
  }

  /** Bounding box of the cells in `within` (the whole sheet when null) that are used: filled, or with valuesOnly off, formatted too. */
  function usedRect(sheet, within, valuesOnly) {
    var rect = null;
    Object.keys(sheet.cells).forEach(function (k) {
      var cell = sheet.cells[k];
      if (valuesOnly && isEmptyCell(cell)) return;
      var p = parseCellRef(k);
      if (!p || (within && !contains(within, { r1: p.r, c1: p.c, r2: p.r, c2: p.c }))) return;
      if (!rect) rect = { r1: p.r, c1: p.c, r2: p.r, c2: p.c };
      else {
        rect.r1 = Math.min(rect.r1, p.r);
        rect.c1 = Math.min(rect.c1, p.c);
        rect.r2 = Math.max(rect.r2, p.r);
        rect.c2 = Math.max(rect.c2, p.c);
      }
    });
    return rect;
  }

  // ------------------------------------------------------------------------------------------
  // Errors

  function officeError(code, message) {
    var e = new Error(message);
    e.name = "RichApi.Error";
    e.code = code;
    e.debugInfo = { code: code, message: message };
    return e;
  }

  function notLoaded(prop) {
    return officeError(
      "PropertyNotLoaded",
      "The property '" + prop + "' is not available. Before reading the property's value, call the load method " +
        'on the containing object and call "context.sync()" on the associated request context.',
    );
  }

  // ------------------------------------------------------------------------------------------
  // Cell writes

  function clearSpill(sheet, key) {
    var prev = sheet.cells[key];
    if (!prev || !prev.spill) return;
    var rect = parseAddress(prev.spill);
    Object.keys(sheet.cells).forEach(function (k) {
      var p = parseCellRef(k);
      if (p && rect && contains(rect, { r1: p.r, c1: p.c, r2: p.r, c2: p.c }) && sheet.cells[k].spilledFrom === key) {
        delete sheet.cells[k];
      }
    });
    delete prev.spill;
  }

  function blockSpill(sheet, childKey) {
    var child = sheet.cells[childKey];
    if (!child || !child.spilledFrom) return;
    var anchorKey = child.spilledFrom;
    var anchor = sheet.cells[anchorKey];
    clearSpill(sheet, anchorKey);
    if (anchor) {
      anchor.value = "#SPILL!";
      anchor.type = "Error";
      anchor.text = "#SPILL!";
    }
  }

  function applyResult(cell, res) {
    var value = res.value === undefined ? null : res.value;
    cell.value = value;
    cell.type = res.type || inferType(value);
    if (cell.type === "Empty") cell.type = "Double";
    cell.text = res.text !== undefined ? String(res.text) : formatText(value, cell.format);
  }

  function applySpill(sheet, r, c, cell, grid) {
    var key = keyOf(r, c);
    var rows = grid.length;
    var cols = 0;
    grid.forEach(function (row) { cols = Math.max(cols, (row || []).length); });
    if (rows === 0 || cols === 0) {
      applyResult(cell, { value: "#CALC!", type: "Error" });
      return;
    }
    var first = (grid[0] || [])[0] || { value: "#N/A", type: "Error" };
    if (rows === 1 && cols === 1) {
      applyResult(cell, first);
      return;
    }
    var rect = { r1: r, c1: c, r2: r + rows - 1, c2: c + cols - 1 };
    var blocked = rect.r2 >= MAX_ROWS || rect.c2 >= MAX_COLS;
    for (var i = 0; i < rows && !blocked; i++) {
      for (var j = 0; j < cols && !blocked; j++) {
        if (i === 0 && j === 0) continue;
        var existing = sheet.cells[keyOf(r + i, c + j)];
        if (existing && !isEmptyCell(existing) && existing.spilledFrom !== key) blocked = true;
      }
    }
    if (blocked) {
      applyResult(cell, { value: "#SPILL!", type: "Error", text: "#SPILL!" });
      return;
    }
    for (var y = 0; y < rows; y++) {
      for (var x = 0; x < cols; x++) {
        var item = (grid[y] || [])[x] || { value: "#N/A", type: "Error" };
        if (y === 0 && x === 0) {
          applyResult(cell, item);
          continue;
        }
        var k2 = keyOf(r + y, c + x);
        var prev = sheet.cells[k2];
        var child = { value: null, type: "Double", text: "", format: prev ? prev.format : "General", spilledFrom: key };
        applyResult(child, item);
        sheet.cells[k2] = child;
      }
    }
    cell.spill = formatRect(rect);
  }

  function writeFormulaCell(sheet, r, c, formula, format) {
    var key = keyOf(r, c);
    var cell = { formula: formula, value: 0, type: "Double", text: "0", format: format || "General" };
    sheet.cells[key] = cell;
    if (typeof fake.evaluate === "function") {
      var res = fake.evaluate(formula, sheet.name, key);
      if (res && Array.isArray(res.spill)) applySpill(sheet, r, c, cell, res.spill);
      else if (res) applyResult(cell, res);
    }
  }

  /** Writes one input the way Excel's `formulas` setter reads it. */
  function setCellInput(sheet, r, c, input) {
    var key = keyOf(r, c);
    blockSpill(sheet, key);
    clearSpill(sheet, key);
    var prev = sheet.cells[key];
    var format = prev ? prev.format : "General";
    if (typeof input === "string" && input.length > 1 && input.charAt(0) === "=") {
      writeFormulaCell(sheet, r, c, input, format);
      return;
    }
    if (input === "" || input === null || input === undefined) {
      if (format !== "General") sheet.cells[key] = { value: null, type: "Empty", text: "", format: format };
      else delete sheet.cells[key];
      return;
    }
    var value = input;
    var type;
    if (typeof input === "number") type = "Double";
    else if (typeof input === "boolean") type = "Boolean";
    else {
      var s = String(input);
      if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s.trim())) {
        value = Number(s);
        type = "Double";
      } else if (/^(true|false)$/i.test(s.trim())) {
        value = s.trim().toUpperCase() === "TRUE";
        type = "Boolean";
      } else {
        value = s;
        type = inferType(s);
      }
    }
    sheet.cells[key] = { value: value, type: type, text: formatText(value, format), format: format };
  }

  /** Shifts relative A1 references in a formula by dr rows and dc columns (autofill). */
  function shiftFormula(formula, dr, dc) {
    var out = "";
    var i = 0;
    var f = formula;
    while (i < f.length) {
      var ch = f.charAt(i);
      var j;
      if (ch === '"' || ch === "'") {
        j = i + 1;
        while (j < f.length) {
          if (f.charAt(j) === ch) {
            if (f.charAt(j + 1) === ch) {
              j += 2;
              continue;
            }
            break;
          }
          j++;
        }
        out += f.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      if (ch === "[") {
        var depth = 0;
        j = i;
        while (j < f.length) {
          if (f.charAt(j) === "[") depth++;
          else if (f.charAt(j) === "]") {
            depth--;
            if (depth === 0) break;
          }
          j++;
        }
        out += f.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      j = i;
      while (j < f.length && "\"'[".indexOf(f.charAt(j)) < 0) j++;
      out += shiftRefs(f.slice(i, j), dr, dc);
      i = j;
    }
    return out;
  }

  function shiftRefs(segment, dr, dc) {
    return segment.replace(
      /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/g,
      function (m, pre, colAbs, col, rowAbs, row) {
        var ci = colIndex(col);
        var ri = Number(row) - 1;
        if (ci < 0 || ri < 0 || ri >= MAX_ROWS) return m;
        var nc = colAbs ? ci : ci + dc;
        var nr = rowAbs ? ri : ri + dr;
        if (nc < 0 || nc >= MAX_COLS || nr < 0 || nr >= MAX_ROWS) return pre + "#REF!";
        return pre + colAbs + colLetter(nc) + rowAbs + (nr + 1);
      },
    );
  }

  function copyCell(sheet, from, r, c, dr, dc) {
    var key = keyOf(r, c);
    blockSpill(sheet, key);
    clearSpill(sheet, key);
    if (!from) {
      delete sheet.cells[key];
      return;
    }
    if (from.formula !== undefined) {
      writeFormulaCell(sheet, r, c, shiftFormula(from.formula, dr, dc), from.format);
      return;
    }
    sheet.cells[key] = { value: from.value, type: from.type, text: from.text, format: from.format };
  }

  function autoFill(src, dst, type) {
    var fillType = type || "FillDefault";
    if (fillType !== "FillDefault" && fillType !== "FillCopy") {
      throw officeError("InvalidArgument", "The fake Excel host supports only FillDefault and FillCopy.");
    }
    if (src.sheet !== dst.sheet || !contains(dst.rect, src.rect)) {
      throw officeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
    }
    var s = src.rect;
    var d = dst.rect;
    var sheet = src.sheet;
    var snapshot = {};
    for (var r0 = s.r1; r0 <= s.r2; r0++) {
      for (var c0 = s.c1; c0 <= s.c2; c0++) snapshot[keyOf(r0, c0)] = clone(sheet.cells[keyOf(r0, c0)]);
    }
    var h = s.r2 - s.r1 + 1;
    var w = s.c2 - s.c1 + 1;
    var r;
    var c;
    if (s.c1 === d.c1 && s.c2 === d.c2) {
      for (r = d.r1; r <= d.r2; r++) {
        if (r >= s.r1 && r <= s.r2) continue;
        var sr = s.r1 + ((((r - s.r1) % h) + h) % h);
        for (c = s.c1; c <= s.c2; c++) copyCell(sheet, snapshot[keyOf(sr, c)], r, c, r - sr, 0);
      }
    } else if (s.r1 === d.r1 && s.r2 === d.r2) {
      for (c = d.c1; c <= d.c2; c++) {
        if (c >= s.c1 && c <= s.c2) continue;
        var sc = s.c1 + ((((c - s.c1) % w) + w) % w);
        for (r = s.r1; r <= s.r2; r++) copyCell(sheet, snapshot[keyOf(r, sc)], r, c, 0, c - sc);
      }
    } else {
      throw officeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
    }
    fake.writes.push({
      kind: "autoFill",
      sheet: sheet.name,
      source: formatRect(s),
      destination: formatRect(d),
      type: fillType,
    });
  }

  function clearRect(sheet, rect, applyTo) {
    var mode = applyTo || "All";
    Object.keys(sheet.cells).forEach(function (k) {
      var p = parseCellRef(k);
      if (!p || !contains(rect, { r1: p.r, c1: p.c, r2: p.r, c2: p.c })) return;
      var cell = sheet.cells[k];
      if (mode === "Formats") {
        cell.format = "General";
        cell.text = formatText(cell.value, "General");
      } else if (mode === "Contents") {
        if (cell.format !== "General") sheet.cells[k] = { value: null, type: "Empty", text: "", format: cell.format };
        else delete sheet.cells[k];
      } else delete sheet.cells[k];
    });
    fake.writes.push({ kind: "clear", sheet: sheet.name, address: formatRect(rect), applyTo: mode });
  }

  // ------------------------------------------------------------------------------------------
  // Proxy objects

  function parseProps(props) {
    if (props === undefined || props === null) return [];
    if (typeof props === "string") return props.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
    if (Array.isArray(props)) return props.map(function (s) { return String(s).trim(); }).filter(Boolean);
    if (typeof props === "object" && props.select) return parseProps(props.select);
    return [];
  }

  function ClientObject(ctx, type, resolve, nullable) {
    this._ctx = ctx;
    this._type = type;
    this._resolveFn = resolve;
    this._nullable = !!nullable;
    this._resolved = false;
    this._target = undefined;
    this._isNull = undefined;
    this._synced = false;
    this._loaded = {};
    ctx._objects.push(this);
  }

  ClientObject.prototype._t = function () {
    if (!this._resolved) {
      var t = this._resolveFn();
      if (t === null || t === undefined) {
        if (!this._nullable) throw officeError("ItemNotFound", "The requested resource doesn't exist.");
        t = null;
      }
      this._target = t;
      this._isNull = t === null;
      this._resolved = true;
    }
    return this._target;
  };

  ClientObject.prototype._read = function (prop) {
    if (!Object.prototype.hasOwnProperty.call(this._loaded, prop)) throw notLoaded(prop);
    return this._loaded[prop];
  };

  ClientObject.prototype._enqueue = function (fn) {
    this._ctx._queue.push(fn);
  };

  ClientObject.prototype.load = function (props) {
    var self = this;
    var list = parseProps(props);
    this._enqueue(function () {
      self._doLoad(list);
    });
    return this;
  };

  ClientObject.prototype._doLoad = function (list) {
    var t = this._t();
    if (t === null) return;
    var names = list.length ? list : this._scalars;
    for (var i = 0; i < names.length; i++) {
      if (this._scalars.indexOf(names[i]) < 0) {
        throw officeError("InvalidArgument", "The property '" + names[i] + "' is not available on " + this._type + ".");
      }
    }
    var self = this;
    names.forEach(function (n) {
      self._loaded[n] = clone(self._get(t, n));
    });
    fake.loads.push(this._describe(t, names));
  };

  ClientObject.prototype._describe = function (_t, names) {
    return { type: this._type, props: names.slice() };
  };

  Object.defineProperty(ClientObject.prototype, "isNullObject", {
    get: function () {
      if (!this._synced) throw notLoaded("isNullObject");
      return this._isNull === true;
    },
  });

  function defineScalars(Ctor, names) {
    Ctor.prototype._scalars = names;
    names.forEach(function (n) {
      Object.defineProperty(Ctor.prototype, n, {
        get: function () {
          return this._read(n);
        },
        configurable: true,
      });
    });
  }

  function inherit(Ctor) {
    Ctor.prototype = Object.create(ClientObject.prototype);
    Ctor.prototype.constructor = Ctor;
  }

  // --- Collections

  function Collection(ctx, type, resolve, makeChild, childScalars) {
    ClientObject.call(this, ctx, type, resolve, false);
    this._makeChild = makeChild;
    this._childScalars = childScalars;
  }
  inherit(Collection);
  Collection.prototype._scalars = [];

  Collection.prototype._doLoad = function (list) {
    var targets = this._t();
    var props = [];
    var self = this;
    list.forEach(function (p) {
      var name = p.indexOf("items/") === 0 ? p.slice(6) : p;
      if (name === "items" || name === "") return;
      if (self._childScalars.indexOf(name) < 0) {
        throw officeError("InvalidArgument", "The property '" + name + "' is not available on " + self._type + " items.");
      }
      props.push(name);
    });
    if (props.length === 0) props = this._childScalars.slice();
    this._loaded.items = targets.map(function (t) {
      var child = self._makeChild(t);
      child._target = t;
      child._resolved = true;
      child._isNull = false;
      props.forEach(function (n) {
        child._loaded[n] = clone(child._get(t, n));
      });
      return child;
    });
    fake.loads.push({ type: this._type, props: props.map(function (p) { return "items/" + p; }) });
  };

  Object.defineProperty(Collection.prototype, "items", {
    get: function () {
      return this._read("items");
    },
  });

  // --- Workbook

  function Workbook(ctx) {
    this._ctx = ctx;
    this._worksheets = null;
    this._tables = null;
    this._application = null;
  }

  Object.defineProperty(Workbook.prototype, "worksheets", {
    get: function () {
      if (!this._worksheets) this._worksheets = new WorksheetCollection(this._ctx);
      return this._worksheets;
    },
  });

  Object.defineProperty(Workbook.prototype, "tables", {
    get: function () {
      if (!this._tables) {
        this._tables = new Collection(
          this._ctx,
          "TableCollection",
          function () { return st().tables.slice(); },
          makeTableChild(this._ctx),
          Table.prototype._scalars,
        );
      }
      return this._tables;
    },
  });

  Object.defineProperty(Workbook.prototype, "names", {
    get: function () {
      if (!this._names) {
        this._names = new Collection(
          this._ctx,
          "NamedItemCollection",
          function () { return st().names.filter(function (n) { return !n.sheet; }); },
          makeNamedItemChild(this._ctx),
          ["name"],
        );
      }
      return this._names;
    },
  });

  Object.defineProperty(Workbook.prototype, "application", {
    get: function () {
      if (!this._application) this._application = new Application(this._ctx);
      return this._application;
    },
  });

  Workbook.prototype.getSelectedRange = function () {
    return new Range(this._ctx, function () {
      var sel = st().selection;
      var sheet = findSheet(sel.sheet);
      var rect = parseAddress(sel.address);
      if (!sheet || !rect) throw officeError("InvalidSelection", "The current selection is invalid for this operation.");
      return { sheet: sheet, rect: rect };
    });
  };

  // --- Application

  function Application(ctx) {
    ClientObject.call(this, ctx, "Application", function () { return st(); }, false);
  }
  inherit(Application);
  defineScalars(Application, ["calculationMode"]);
  Application.prototype._get = function (t, prop) {
    if (prop === "calculationMode") return t.calculationMode;
    return undefined;
  };

  // --- Worksheets

  function WorksheetCollection(ctx) {
    Collection.call(
      this,
      ctx,
      "WorksheetCollection",
      function () { return st().sheets.slice(); },
      function (t) { return new Worksheet(ctx, function () { return t; }); },
      ["name", "visibility", "position"],
    );
  }
  WorksheetCollection.prototype = Object.create(Collection.prototype);
  WorksheetCollection.prototype.constructor = WorksheetCollection;

  WorksheetCollection.prototype.getItem = function (name) {
    return new Worksheet(this._ctx, function () { return findSheet(name); });
  };

  WorksheetCollection.prototype.getItemOrNullObject = function (name) {
    return new Worksheet(this._ctx, function () { return findSheet(name); }, true);
  };

  WorksheetCollection.prototype.getActiveWorksheet = function () {
    return new Worksheet(this._ctx, function () { return findSheet(st().activeSheet) || st().sheets[0]; });
  };

  WorksheetCollection.prototype.add = function (name) {
    var ctx = this._ctx;
    var created = null;
    this._enqueue(function () {
      var n = name;
      if (n === undefined || n === null || n === "") {
        var i = st().sheets.length + 1;
        while (findSheet("Sheet" + i)) i++;
        n = "Sheet" + i;
      }
      if (findSheet(n)) throw officeError("ItemAlreadyExists", "A resource with the same name or identifier already exists.");
      created = { name: String(n), visibility: "Visible", cells: {} };
      st().sheets.push(created);
      fake.writes.push({ kind: "addSheet", sheet: created.name });
    });
    return new Worksheet(ctx, function () {
      if (!created) throw officeError("InvalidOperation", "The sheet hasn't been added yet.");
      return created;
    });
  };

  function Worksheet(ctx, resolve, nullable) {
    ClientObject.call(this, ctx, "Worksheet", resolve, nullable);
  }
  inherit(Worksheet);
  defineScalars(Worksheet, ["name", "visibility", "position"]);
  Worksheet.prototype._get = function (t, prop) {
    if (prop === "name") return t.name;
    if (prop === "visibility") return t.visibility;
    if (prop === "position") return st().sheets.indexOf(t);
    return undefined;
  };

  Object.defineProperty(Worksheet.prototype, "names", {
    get: function () {
      var self = this;
      return new Collection(
        this._ctx,
        "NamedItemCollection",
        function () {
          var sheet = self._t();
          return st().names.filter(function (n) { return sheet && n.sheet === sheet.name; });
        },
        makeNamedItemChild(this._ctx),
        ["name"],
      );
    },
  });

  Worksheet.prototype.getRange = function (address) {
    var self = this;
    return new Range(this._ctx, function () {
      var sheet = self._t();
      if (sheet === null) throw officeError("InvalidObjectPath", "The object path isn't working for what you're trying to do.");
      var rect = parseAddress(address);
      if (!rect) throw officeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
      return { sheet: sheet, rect: rect };
    });
  };

  // RangeAreas, as far as the pane uses it: select() (ExcelApi 1.18). Selecting several areas makes
  // the selection a list, which getSelectedRange then refuses, as Excel does.
  Worksheet.prototype.getRanges = function (address) {
    var self = this;
    var areas = String(address).split(",").map(function (a) { return a.trim(); });
    var ctx = this._ctx;
    return {
      select: function () {
        if (!isSetSupported("ExcelApi", "1.18")) throw officeError("ApiNotFound", "The API you are trying to use could not be found.");
        ctx._queue.push(function () {
          var sheet = self._t();
          if (sheet === null) throw officeError("InvalidObjectPath", "The object path isn't working for what you're trying to do.");
          for (var i = 0; i < areas.length; i++) {
            if (!parseAddress(areas[i])) throw officeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
          }
          st().selection = { sheet: sheet.name, address: areas.length === 1 ? formatRect(parseAddress(areas[0])) : areas.join(", ") };
          st().activeSheet = sheet.name;
          fake.selections.push({ sheet: sheet.name, address: areas.join(", ") });
        });
      },
    };
  };

  Worksheet.prototype.getUsedRangeOrNullObject = function (valuesOnly) {
    var self = this;
    return new Range(
      this._ctx,
      function () {
        var sheet = self._t();
        if (sheet === null) return null;
        var rect = usedRect(sheet, null, valuesOnly);
        return rect ? { sheet: sheet, rect: rect } : null;
      },
      true,
    );
  };

  Object.defineProperty(Worksheet.prototype, "autoFilter", {
    get: function () {
      if (!this._autoFilter) this._autoFilter = new AutoFilter(this._ctx, this);
      return this._autoFilter;
    },
  });

  // --- AutoFilter (the sheet's own; a table's filter belongs to the table)

  function AutoFilter(ctx, worksheet) {
    ClientObject.call(this, ctx, "AutoFilter", function () { return worksheet._t(); }, false);
  }
  inherit(AutoFilter);
  defineScalars(AutoFilter, ["enabled", "isDataFiltered"]);
  AutoFilter.prototype._get = function (sheet, prop) {
    if (prop === "enabled") return !!sheet.autoFilter;
    if (prop === "isDataFiltered") return false;
    return undefined;
  };

  AutoFilter.prototype.getRangeOrNullObject = function () {
    var self = this;
    return new Range(
      this._ctx,
      function () {
        var sheet = self._t();
        return sheet && sheet.autoFilter ? { sheet: sheet, rect: parseAddress(sheet.autoFilter) } : null;
      },
      true,
    );
  };

  Worksheet.prototype.activate = function () {
    var self = this;
    this._enqueue(function () {
      st().activeSheet = self._t().name;
    });
  };

  Worksheet.prototype.calculate = function () {
    var self = this;
    this._enqueue(function () {
      self._t();
    });
  };

  // --- Range

  function Range(ctx, resolve, nullable) {
    ClientObject.call(this, ctx, "Range", resolve, nullable);
    this._worksheet = null;
  }
  inherit(Range);
  defineScalars(Range, [
    "address",
    "addressLocal",
    "cellCount",
    "isEntireRow",
    "isEntireColumn",
    "rowCount",
    "columnCount",
    "rowIndex",
    "columnIndex",
    "values",
    "text",
    "formulas",
    "valueTypes",
    "numberFormat",
  ]);

  function grid(t, fn) {
    if (cellCount(t.rect) > MAX_LOAD_CELLS) {
      throw officeError("ResponsePayloadSizeLimitExceeded", "The response payload size has exceeded the limit.");
    }
    var out = [];
    for (var r = t.rect.r1; r <= t.rect.r2; r++) {
      var row = [];
      for (var c = t.rect.c1; c <= t.rect.c2; c++) row.push(fn(t.sheet.cells[keyOf(r, c)]));
      out.push(row);
    }
    return out;
  }

  Range.prototype._get = function (t, prop) {
    var rect = t.rect;
    switch (prop) {
      case "address":
      case "addressLocal":
        return quoteSheet(t.sheet.name) + "!" + formatRect(rect);
      case "cellCount":
        // Office reports -1 past 2^31-1 cells (a whole sheet has 17 billion).
        return cellCount(rect) > 2147483647 ? -1 : cellCount(rect);
      case "isEntireRow":
        return rect.c1 === 0 && rect.c2 === MAX_COLS - 1;
      case "isEntireColumn":
        return rect.r1 === 0 && rect.r2 === MAX_ROWS - 1;
      case "rowCount":
        return rect.r2 - rect.r1 + 1;
      case "columnCount":
        return rect.c2 - rect.c1 + 1;
      case "rowIndex":
        return rect.r1;
      case "columnIndex":
        return rect.c1;
      case "values":
        return grid(t, function (cell) { return isEmptyCell(cell) ? "" : cell.value === null ? "" : cell.value; });
      case "text":
        return grid(t, function (cell) { return isEmptyCell(cell) ? "" : cell.text; });
      case "formulas":
        return grid(t, function (cell) {
          if (isEmptyCell(cell)) return "";
          return cell.formula !== undefined ? cell.formula : cell.value;
        });
      case "valueTypes":
        return grid(t, function (cell) { return isEmptyCell(cell) ? "Empty" : cell.type; });
      case "numberFormat":
        return grid(t, function (cell) { return cell ? cell.format : "General"; });
    }
    return undefined;
  };

  Range.prototype._describe = function (t, names) {
    return { type: "Range", sheet: t.sheet.name, address: formatRect(t.rect), props: names.slice() };
  };

  Object.defineProperty(Range.prototype, "formulas", {
    get: function () {
      return this._read("formulas");
    },
    set: function (data) {
      var self = this;
      var input = clone(data);
      this._enqueue(function () {
        var t = self._t();
        var rows = t.rect.r2 - t.rect.r1 + 1;
        var cols = t.rect.c2 - t.rect.c1 + 1;
        if (Array.isArray(input)) {
          var ok = input.length === rows && input.every(function (row) { return Array.isArray(row) && row.length === cols; });
          if (!ok) {
            throw officeError(
              "InvalidArgument",
              "The number of rows or columns in the input array doesn't match the size or dimensions of the range.",
            );
          }
        } else if (cellCount(t.rect) > MAX_LOAD_CELLS) {
          throw officeError("RequestPayloadSizeLimitExceeded", "The request payload size has exceeded the limit.");
        }
        for (var r = 0; r < rows; r++) {
          for (var c = 0; c < cols; c++) {
            setCellInput(t.sheet, t.rect.r1 + r, t.rect.c1 + c, Array.isArray(input) ? input[r][c] : input);
          }
        }
        fake.writes.push({ kind: "formulas", sheet: t.sheet.name, address: formatRect(t.rect), formulas: input });
      });
    },
    configurable: true,
  });

  Object.defineProperty(Range.prototype, "worksheet", {
    get: function () {
      var self = this;
      if (!this._worksheet) {
        this._worksheet = new Worksheet(this._ctx, function () {
          var t = self._t();
          return t === null ? null : t.sheet;
        });
      }
      return this._worksheet;
    },
  });

  Range.prototype.getTables = function (fullyContained) {
    var self = this;
    return new Collection(
      this._ctx,
      "TableScopedCollection",
      function () {
        var t = self._t();
        return st().tables.filter(function (tb) {
          if (tb.sheet.toUpperCase() !== t.sheet.name.toUpperCase()) return false;
          var rect = parseAddress(tb.address);
          return fullyContained ? contains(t.rect, rect) : intersects(t.rect, rect);
        });
      },
      makeTableChild(this._ctx),
      Table.prototype._scalars,
    );
  };

  Range.prototype.autoFill = function (destination, type) {
    var self = this;
    this._enqueue(function () {
      var src = self._t();
      var dst;
      if (typeof destination === "string") {
        var rect = parseAddress(destination);
        if (!rect) throw officeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
        dst = { sheet: src.sheet, rect: rect };
      } else if (destination instanceof Range) {
        dst = destination._t();
      } else {
        throw officeError("InvalidArgument", "The fake Excel host needs a destination range for autoFill.");
      }
      autoFill(src, dst, type);
    });
  };

  Range.prototype.getIntersectionOrNullObject = function (other) {
    var self = this;
    return new Range(
      this._ctx,
      function () {
        var a = self._t();
        var b = other._t();
        if (!a || !b || a.sheet !== b.sheet) return null;
        var rect = {
          r1: Math.max(a.rect.r1, b.rect.r1),
          c1: Math.max(a.rect.c1, b.rect.c1),
          r2: Math.min(a.rect.r2, b.rect.r2),
          c2: Math.min(a.rect.c2, b.rect.c2),
        };
        if (rect.r1 > rect.r2 || rect.c1 > rect.c2) return null;
        return { sheet: a.sheet, rect: rect };
      },
      true,
    );
  };

  Range.prototype.getSurroundingRegion = function () {
    var self = this;
    return new Range(this._ctx, function () {
      var t = self._t();
      return { sheet: t.sheet, rect: currentRegion(t.sheet, t.rect.r1, t.rect.c1) };
    });
  };

  Range.prototype.getUsedRangeOrNullObject = function (valuesOnly) {
    var self = this;
    return new Range(
      this._ctx,
      function () {
        var t = self._t();
        if (!t) return null;
        var rect = null;
        if (fake.usedRange === "sheet") {
          // The other reading: the range cut to the sheet's used range.
          var sheetUsed = usedRect(t.sheet, null, valuesOnly);
          if (sheetUsed) {
            rect = {
              r1: Math.max(t.rect.r1, sheetUsed.r1),
              c1: Math.max(t.rect.c1, sheetUsed.c1),
              r2: Math.min(t.rect.r2, sheetUsed.r2),
              c2: Math.min(t.rect.c2, sheetUsed.c2),
            };
            if (rect.r1 > rect.r2 || rect.c1 > rect.c2) rect = null;
          }
        } else rect = usedRect(t.sheet, t.rect, valuesOnly);
        return rect ? { sheet: t.sheet, rect: rect } : null;
      },
      true,
    );
  };

  Range.prototype.getSpillingToRangeOrNullObject = function () {
    if (!isSetSupported("ExcelApi", "1.12")) {
      throw officeError("ApiNotFound", "The API you are trying to use could not be found. It may be available in a newer version of Excel.");
    }
    var self = this;
    return new Range(
      this._ctx,
      function () {
        var t = self._t();
        var anchor = t.sheet.cells[keyOf(t.rect.r1, t.rect.c1)];
        if (!anchor || !anchor.spill) return null;
        return { sheet: t.sheet, rect: parseAddress(anchor.spill) };
      },
      true,
    );
  };

  Range.prototype.select = function () {
    var self = this;
    this._enqueue(function () {
      var t = self._t();
      st().selection = { sheet: t.sheet.name, address: formatRect(t.rect) };
      st().activeSheet = t.sheet.name;
      fake.selections.push({ sheet: t.sheet.name, address: formatRect(t.rect) });
    });
  };

  Range.prototype.clear = function (applyTo) {
    var self = this;
    this._enqueue(function () {
      var t = self._t();
      clearRect(t.sheet, t.rect, applyTo);
    });
  };

  // --- Tables

  function Table(ctx, resolve) {
    ClientObject.call(this, ctx, "Table", resolve, false);
    this._worksheet = null;
    this._columns = null;
  }
  inherit(Table);
  defineScalars(Table, ["name", "showHeaders", "showTotals"]);
  Table.prototype._get = function (t, prop) {
    if (prop === "name") return t.name;
    if (prop === "showHeaders") return true;
    if (prop === "showTotals") return t.showTotals === true;
    return undefined;
  };

  function NamedItem(ctx, resolve) {
    ClientObject.call(this, ctx, "NamedItem", resolve, false);
  }
  inherit(NamedItem);
  defineScalars(NamedItem, ["name"]);
  NamedItem.prototype._get = function (t, prop) {
    if (prop === "name") return t.name;
    return undefined;
  };

  function makeNamedItemChild(ctx) {
    return function (t) {
      return new NamedItem(ctx, function () { return t; });
    };
  }

  function makeTableChild(ctx) {
    return function (t) {
      return new Table(ctx, function () { return t; });
    };
  }

  Object.defineProperty(Table.prototype, "worksheet", {
    get: function () {
      var self = this;
      if (!this._worksheet) {
        this._worksheet = new Worksheet(this._ctx, function () {
          var sheet = findSheet(self._t().sheet);
          if (!sheet) throw officeError("ItemNotFound", "The requested resource doesn't exist.");
          return sheet;
        });
      }
      return this._worksheet;
    },
  });

  Object.defineProperty(Table.prototype, "columns", {
    get: function () {
      var self = this;
      if (!this._columns) {
        var ctx = this._ctx;
        this._columns = new Collection(
          ctx,
          "TableColumnCollection",
          function () {
            var tb = self._t();
            var sheet = findSheet(tb.sheet);
            var rect = parseAddress(tb.address);
            var out = [];
            for (var c = rect.c1; c <= rect.c2; c++) {
              var cell = sheet ? sheet.cells[keyOf(rect.r1, c)] : null;
              var name = cell && !isEmptyCell(cell) ? cell.text : "Column" + (c - rect.c1 + 1);
              out.push({ name: name, index: c - rect.c1 });
            }
            return out;
          },
          function (t) {
            return new TableColumn(ctx, function () { return t; });
          },
          ["name", "index"],
        );
      }
      return this._columns;
    },
  });

  Table.prototype.getRange = function () {
    var self = this;
    return new Range(this._ctx, function () {
      var tb = self._t();
      var sheet = findSheet(tb.sheet);
      if (!sheet) throw officeError("ItemNotFound", "The requested resource doesn't exist.");
      return { sheet: sheet, rect: parseAddress(tb.address) };
    });
  };

  function TableColumn(ctx, resolve) {
    ClientObject.call(this, ctx, "TableColumn", resolve, false);
  }
  inherit(TableColumn);
  defineScalars(TableColumn, ["name", "index"]);
  TableColumn.prototype._get = function (t, prop) {
    return t[prop];
  };

  // ------------------------------------------------------------------------------------------
  // Request context and Excel.run

  function RequestContext() {
    this._queue = [];
    this._objects = [];
    this.workbook = new Workbook(this);
  }

  RequestContext.prototype._flush = function () {
    var queue = this._queue;
    this._queue = [];
    if (fake._failNext) {
      var f = fake._failNext;
      fake._failNext = null;
      throw officeError(f.code, f.message);
    }
    if (fake._failOn) {
      fake._failOn.remaining--;
      if (fake._failOn.remaining <= 0) {
        var g = fake._failOn;
        fake._failOn = null;
        throw officeError(g.code, g.message);
      }
    }
    for (var i = 0; i < queue.length; i++) queue[i]();
    var objects = this._objects.slice();
    objects.forEach(function (o) {
      if (o._nullable && !o._resolved) {
        try {
          o._t();
        } catch {
          o._isNull = true;
          o._resolved = true;
          o._target = null;
        }
      }
      o._synced = true;
    });
  };

  RequestContext.prototype.sync = function () {
    var ctx = this;
    return Promise.resolve().then(function () {
      ctx._flush();
    });
  };

  function versionAtLeast(have, want) {
    var a = String(have).split(".").map(Number);
    var b = String(want).split(".").map(Number);
    for (var i = 0; i < Math.max(a.length, b.length); i++) {
      var x = a[i] || 0;
      var y = b[i] || 0;
      if (x !== y) return x > y;
    }
    return true;
  }

  function isSetSupported(name, version) {
    if (fake.maxApi === null || fake.maxApi === undefined) return true;
    if (String(name).toLowerCase() !== "excelapi") return true;
    return versionAtLeast(fake.maxApi, version === undefined ? "1.1" : version);
  }

  // ------------------------------------------------------------------------------------------
  // Test helpers: window.__NYMFORM_FAKE__

  var fake = {
    writes: [],
    /** Every selection the pane made (Show in sheet, restoring the selection), in order. */
    selections: [],
    loads: [],
    evaluate: null,
    host: "Excel",
    platform: "OfficeOnline",
    /** Highest ExcelApi version to report, e.g. "1.9"; null supports everything. */
    maxApi: null,
    /** Office.context.officeTheme, e.g. { isDarkTheme: true, bodyBackgroundColor: "#1f1f1f" }. */
    theme: null,
    /** "cells" (default): Range.getUsedRangeOrNullObject is the box around used cells in the range; "sheet": the range cut to the sheet's used range. */
    usedRange: "cells",
    _failNext: null,
    _failOn: null,

    /** Loads a workbook (default: window.__NYMFORM_FAKE_WORKBOOK__) and resets every helper setting. */
    reset: function (workbook) {
      state = normalizeWorkbook(workbook !== undefined ? workbook : root.__NYMFORM_FAKE_WORKBOOK__);
      fake.writes.length = 0;
      fake.loads.length = 0;
      fake.selections.length = 0;
      fake.evaluate = null;
      fake.host = "Excel";
      fake.platform = "OfficeOnline";
      fake.maxApi = null;
      fake.theme = null;
      fake.usedRange = "cells";
      fake._failNext = null;
      fake._failOn = null;
    },

    /** A copy of one cell, or null when it is empty and unformatted. */
    getCell: function (sheet, addr) {
      var s = findSheet(sheet);
      var p = parseCellRef(addr);
      if (!s || !p) return null;
      var cell = s.cells[keyOf(p.r, p.c)];
      return cell ? clone(cell) : null;
    },

    /** Sets one cell directly (no write is logged). */
    setCell: function (sheet, addr, cell) {
      var s = findSheet(sheet);
      var p = parseCellRef(addr);
      if (!s || !p) throw new Error("Fake: no such sheet or cell");
      if (cell === null || cell === undefined) delete s.cells[keyOf(p.r, p.c)];
      else s.cells[keyOf(p.r, p.c)] = normCell(cell);
    },

    setSelection: function (sheet, address) {
      var s = findSheet(sheet);
      if (!s || !parseAddress(address)) throw new Error("Fake: no such sheet or bad address");
      st().selection = { sheet: s.name, address: formatRect(parseAddress(address)) };
      st().activeSheet = s.name;
    },

    getSelection: function () {
      return clone(st().selection);
    },

    setCalculationMode: function (mode) {
      st().calculationMode = mode;
    },

    /** The next context.sync() rejects with an Office error of this code. */
    failNext: function (code, message) {
      fake._failNext = { code: code, message: message || "The operation failed." };
    },

    /**
     * The n-th context.sync() from now (1 = the next one) rejects with an Office error of this code,
     * before any of its queued operations run; the syncs before it succeed. Lets a test fail a read
     * that comes after a write has already gone through.
     */
    failOnSync: function (n, code, message) {
      fake._failOn = { remaining: n, code: code, message: message || "The operation failed." };
    },

    /** A copy of the current state in the input JSON shape. */
    workbook: function () {
      var s = st();
      return clone({
        sheets: s.sheets.map(function (sh) {
          var out = { name: sh.name, visibility: sh.visibility, cells: sh.cells };
          if (sh.autoFilter) out.autoFilter = sh.autoFilter;
          return out;
        }),
        tables: s.tables,
        selection: s.selection,
        calculationMode: s.calculationMode,
      });
    },

    sheetNames: function () {
      return st().sheets.map(function (s) { return s.name; });
    },

    shiftFormula: shiftFormula,
  };

  // ------------------------------------------------------------------------------------------
  // Globals

  var Office = {
    HostType: {
      Word: "Word",
      Excel: "Excel",
      PowerPoint: "PowerPoint",
      Outlook: "Outlook",
      OneNote: "OneNote",
      Project: "Project",
      Access: "Access",
    },
    PlatformType: {
      PC: "PC",
      OfficeOnline: "OfficeOnline",
      Mac: "Mac",
      iOS: "iOS",
      Android: "Android",
      Universal: "Universal",
    },
    context: {
      requirements: { isSetSupported: isSetSupported },
      get officeTheme() {
        return fake.theme ? clone(fake.theme) : undefined;
      },
    },
    onReady: function (callback) {
      var info = { host: fake.host || null, platform: fake.host ? fake.platform : null };
      if (typeof callback === "function") {
        try {
          callback(info);
        } catch {
          // Office ignores callback errors here too.
        }
      }
      return Promise.resolve(info);
    },
  };

  var Excel = {
    run: function (a, b) {
      var batch = typeof a === "function" ? a : b;
      if (typeof batch !== "function") return Promise.reject(officeError("InvalidArgument", "Excel.run needs a function."));
      var ctx = new RequestContext();
      return Promise.resolve()
        .then(function () {
          return batch(ctx);
        })
        .then(function (result) {
          return ctx.sync().then(function () {
            return result;
          });
        });
    },
    RequestContext: RequestContext,
    AutoFillType: { fillDefault: "FillDefault", fillCopy: "FillCopy" },
    CalculationMode: { automatic: "Automatic", automaticExceptTables: "AutomaticExceptTables", manual: "Manual" },
    ClearApplyTo: { all: "All", formats: "Formats", contents: "Contents" },
    SheetVisibility: { visible: "Visible", hidden: "Hidden", veryHidden: "VeryHidden" },
    RangeValueType: {
      unknown: "Unknown",
      empty: "Empty",
      string: "String",
      integer: "Integer",
      double: "Double",
      boolean: "Boolean",
      error: "Error",
      richValue: "RichValue",
    },
  };

  root.Office = Office;
  root.Excel = Excel;
  root.__NYMFORM_FAKE__ = fake;
  if (typeof module !== "undefined" && module && module.exports) module.exports = fake;
})(typeof window !== "undefined" ? window : globalThis);
