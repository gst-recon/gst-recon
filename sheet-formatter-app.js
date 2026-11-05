(function () {
  "use strict";
  var core = window.GSTReconSheetFormatter;
  if (!core) return;

  var el = {
    file: document.getElementById("sourceFile"),
    fileName: document.getElementById("fileName"),
    sample: document.getElementById("sampleBtn"),
    clear: document.getElementById("clearBtn"),
    sheetWrap: document.getElementById("sheetWrap"),
    sheetSelect: document.getElementById("sheetSelect"),
    headerWrap: document.getElementById("headerWrap"),
    headerSelect: document.getElementById("headerSelect"),
    headerNote: document.getElementById("headerNote"),
    mappingWrap: document.getElementById("mappingWrap"),
    mappingGrid: document.getElementById("mappingGrid"),
    message: document.getElementById("message"),
    results: document.getElementById("results"),
    rowCount: document.getElementById("rowCount"),
    readyCount: document.getElementById("readyCount"),
    reviewCount: document.getElementById("reviewCount"),
    skipNote: document.getElementById("skipNote"),
    previewHead: document.getElementById("previewHead"),
    previewBody: document.getElementById("previewBody"),
    downloadXlsx: document.getElementById("downloadXlsx"),
    downloadCsv: document.getElementById("downloadCsv"),
    csvNote: document.getElementById("csvNote")
  };

  var state = { fileName: "", sheets: [], sheetIndex: 0, grid: [], headerIndex: 0, mapping: {}, result: null };
  var MAX_BYTES = 10 * 1024 * 1024;

  function setText(node, value) { node.textContent = value === null || value === undefined ? "" : String(value); }
  function empty(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  function showMessage(text, kind) {
    setText(el.message, text);
    el.message.className = "message" + (text ? " show " + (kind || "info") : "");
  }
  function element(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = String(text);
    return node;
  }
  function hasContent(grid) {
    return Array.isArray(grid) && grid.some(function (row) {
      return Array.isArray(row) && row.some(function (cell) { return String(cell === null || cell === undefined ? "" : cell).trim() !== ""; });
    });
  }
  function clearResults() {
    state.result = null;
    el.results.hidden = true;
    el.downloadXlsx.disabled = true;
    el.downloadCsv.disabled = true;
    el.csvNote.hidden = true;
    empty(el.previewHead);
    empty(el.previewBody);
  }
  function cleanState() {
    state.fileName = "";
    state.sheets = [];
    state.sheetIndex = 0;
    state.grid = [];
    state.headerIndex = 0;
    state.mapping = {};
    state.result = null;
    el.file.value = "";
    setText(el.fileName, "");
    empty(el.sheetSelect);
    empty(el.headerSelect);
    empty(el.mappingGrid);
    el.sheetWrap.hidden = true;
    el.headerWrap.hidden = true;
    el.mappingWrap.hidden = true;
    clearResults();
    showMessage("", "info");
  }

  async function readSheets(file) {
    if (!file) throw new Error("Pehle ek spreadsheet select karein.");
    if (file.size === 0) throw new Error("File khaali hai.");
    if (file.size > MAX_BYTES) throw new Error("File 10 MB se badi hai. Chhoti file export karke dobara try karein.");
    var name = String(file.name || "").toLowerCase();
    var extension = name.indexOf(".") >= 0 ? name.slice(name.lastIndexOf(".")) : "";
    if (extension === ".json") {
      var json = await file.text();
      var grid;
      try { grid = core.jsonToGrid(json); }
      catch (_) { throw new Error("JSON file me invoice rows nahi mile. Array of objects ya site ke accepted invoice columns use karein."); }
      return [{ name: "JSON data", grid: grid }];
    }
    if (extension === ".csv" || extension === ".tsv" || extension === ".txt") {
      var text = await file.text();
      var parsed = core.parseDelimited(text, extension === ".tsv" ? "\t" : undefined);
      if (!parsed.length) throw new Error("File me padhne layak rows nahi mili.");
      return [{ name: "File data", grid: parsed }];
    }
    if (extension === ".xlsx" || extension === ".xls") {
      if (!window.XLSX) throw new Error("Excel reader load nahi hua. CSV / JSON try karein, ya internet connection check karein.");
      try {
        var workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
        if (!workbook.SheetNames || !workbook.SheetNames.length) throw new Error("No worksheet");
        return workbook.SheetNames.map(function (sheetName) {
          return {
            name: sheetName,
            grid: window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: "" })
          };
        });
      } catch (_) {
        throw new Error("Excel file read nahi ho payi. File ko Excel se dobara save karke ya CSV me export karke try karein.");
      }
    }
    throw new Error("Sirf .xlsx, .xls, .csv, .tsv, .txt ya .json file upload karein.");
  }

  function chooseSheet(index) {
    state.sheetIndex = Number(index) || 0;
    var sheet = state.sheets[state.sheetIndex];
    state.grid = sheet ? sheet.grid : [];
    if (!hasContent(state.grid)) {
      showMessage("Is worksheet me koi data nahi hai. Dusri sheet select karein.", "warn");
      el.headerWrap.hidden = true;
      el.mappingWrap.hidden = true;
      clearResults();
      return;
    }
    prepareHeader();
  }

  function rowCaption(row, index) {
    var cells = (row || []).map(function (cell) { return String(cell === null || cell === undefined ? "" : cell).trim(); })
      .filter(Boolean).slice(0, 3).map(function (cell) { return cell.length > 34 ? cell.slice(0, 31) + "…" : cell; });
    return "Row " + (index + 1) + (cells.length ? " · " + cells.join(" / ") : " · blank");
  }

  function prepareHeader() {
    var detection = core.detectHeaderRow(state.grid, 15);
    var maxRows = Math.min(state.grid.length, 100);
    empty(el.headerSelect);
    for (var i = 0; i < maxRows; i++) {
      var option = element("option", "", rowCaption(state.grid[i], i));
      option.value = String(i);
      el.headerSelect.appendChild(option);
    }
    state.headerIndex = detection.index;
    el.headerSelect.value = String(state.headerIndex);
    el.headerWrap.hidden = false;
    if (detection.autoDetected) setText(el.headerNote, "Header row auto-detected using the same invoice/GSTIN names the site accepts. You can change it.");
    else setText(el.headerNote, "Header row auto-detect nahi hui. Sahi headings wali row choose karein; columns neeche manually map kiye ja sakte hain.");
    setMappingFromHeader();
  }

  function setMappingFromHeader() {
    var headers = state.grid[state.headerIndex] || [];
    state.mapping = core.suggestMappings(headers);
    renderMapping(headers);
    renderPreview();
  }

  function renderMapping(headers) {
    empty(el.mappingGrid);
    core.FIELD_DEFS.forEach(function (field) {
      var wrapper = element("div", "map-field");
      var label = element("label", "map-label", field.label + (field.required ? " *" : ""));
      label.htmlFor = "map-" + field.key;
      var select = element("select", "map-select");
      select.id = "map-" + field.key;
      select.setAttribute("aria-label", field.label + " source column");
      var emptyOption = element("option", "", "— Not mapped —");
      emptyOption.value = "";
      select.appendChild(emptyOption);
      headers.forEach(function (header, index) {
        var text = String(header === null || header === undefined ? "" : header).trim();
        var option = element("option", "", core.columnLetter(index) + " · " + (text || "(blank header)"));
        option.value = String(index);
        select.appendChild(option);
      });
      var suggested = state.mapping[field.key];
      select.value = suggested === null || suggested === undefined ? "" : String(suggested);
      select.addEventListener("change", function () {
        state.mapping[field.key] = select.value === "" ? null : Number(select.value);
        renderPreview();
      });
      wrapper.appendChild(label);
      wrapper.appendChild(select);
      el.mappingGrid.appendChild(wrapper);
    });
    el.mappingWrap.hidden = false;
  }

  function addCell(row, value, className, title) {
    var cell = element("td", className || "", value === null || value === undefined ? "" : String(value));
    if (title) cell.title = String(title);
    row.appendChild(cell);
  }

  function renderPreview() {
    clearResults();
    var result = core.transformGrid(state.grid, state.headerIndex, state.mapping);
    state.result = result;
    if (result.errors.length) {
      showMessage(result.errors.join(" "), "err");
      return;
    }
    showMessage("Preview ready. Files are processed in this browser tab only; this tool does not upload or save them.", "ok");
    var ready = result.rows.filter(function (row) { return row.ready; }).length;
    var review = result.rows.length - ready;
    setText(el.rowCount, result.rows.length);
    setText(el.readyCount, ready);
    setText(el.reviewCount, review);
    var notes = [];
    if (result.blankSkipped) notes.push(result.blankSkipped + " blank row(s) skipped");
    if (result.totalSkipped) notes.push(result.totalSkipped + " repeated-header / total row(s) skipped");
    if (result.optionalDefaults.length) notes.push("Not mapped; output defaults: " + result.optionalDefaults.join(", "));
    if (review) notes.push("Rows needing review remain in the export; the site may flag missing or invalid GSTIN / invoice numbers.");
    setText(el.skipNote, notes.join(". "));

    empty(el.previewHead);
    empty(el.previewBody);
    var headerRow = element("tr");
    ["Source row"].concat(core.SITE_HEADERS).concat(["Check"]).forEach(function (header) { headerRow.appendChild(element("th", "", header)); });
    el.previewHead.appendChild(headerRow);
    result.rows.slice(0, 20).forEach(function (record) {
      var tr = element("tr", record.issues.length ? "needs-review" : "");
      addCell(tr, record.sourceRow, "mono");
      core.FIELD_DEFS.forEach(function (field) { addCell(tr, record.values[field.key], field.numeric ? "number" : ""); });
      addCell(tr, record.issues.length ? record.issues.join(" ") : "Ready", record.issues.length ? "check bad" : "check good");
      el.previewBody.appendChild(tr);
    });
    if (result.rows.length > 20) {
      var more = element("tr", "more-row");
      var moreCell = element("td", "", "Showing first 20 of " + result.rows.length + " rows; download contains all rows.");
      moreCell.colSpan = core.SITE_HEADERS.length + 2;
      more.appendChild(moreCell);
      el.previewBody.appendChild(more);
    }
    var hasRows = result.rows.length > 0;
    el.results.hidden = !hasRows;
    el.downloadXlsx.disabled = !hasRows || !window.XLSX;
    var formulaRisk = core.hasFormulaLikeText(result.rows);
    el.downloadCsv.disabled = !hasRows || formulaRisk;
    el.csvNote.hidden = !formulaRisk;
    if (formulaRisk) setText(el.csvNote, "CSV download disabled: a text value begins with a spreadsheet formula character. Use the XLSX download to keep text cells safe without changing the invoice number.");
  }

  function downloadBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var anchor = element("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function formattedFileName(extension) {
    return core.outputFileName(state.fileName, extension);
  }

  function exportXlsx() {
    if (!state.result || !state.result.rows.length || !window.XLSX) return;
    var matrix = core.recordsToMatrix(state.result.rows);
    var sheet = window.XLSX.utils.aoa_to_sheet([matrix[0]]);
    for (var r = 1; r < matrix.length; r++) {
      for (var c = 0; c < matrix[r].length; c++) {
        var address = window.XLSX.utils.encode_cell({ r: r, c: c });
        var value = matrix[r][c];
        sheet[address] = typeof value === "number" ? { t: "n", v: value } : { t: "s", v: String(value === null || value === undefined ? "" : value) };
      }
    }
    sheet["!ref"] = window.XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: matrix.length - 1, c: matrix[0].length - 1 } });
    var book = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(book, sheet, "GST Recon");
    window.XLSX.writeFile(book, formattedFileName("xlsx"));
  }

  function exportCsv() {
    if (!state.result || !state.result.rows.length || core.hasFormulaLikeText(state.result.rows)) return;
    var csv = core.toCsv(state.result.rows);
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), formattedFileName("csv"));
  }

  function selectWorkbook(sheets, fileName) {
    sheets = (sheets || []).filter(function (sheet) { return hasContent(sheet && sheet.grid); });
    if (!sheets.length) throw new Error("File me koi non-empty worksheet nahi mili.");
    state.sheets = sheets;
    state.fileName = fileName;
    setText(el.fileName, fileName + " · " + sheets.length + " worksheet(s)");
    if (sheets.length > 1) {
      empty(el.sheetSelect);
      sheets.forEach(function (sheet, index) {
        var option = element("option", "", sheet.name);
        option.value = String(index);
        el.sheetSelect.appendChild(option);
      });
      el.sheetWrap.hidden = false;
    } else el.sheetWrap.hidden = true;
    chooseSheet(0);
  }

  el.file.addEventListener("change", async function () {
    state.fileName = "";
    state.sheets = [];
    state.sheetIndex = 0;
    state.grid = [];
    state.headerIndex = 0;
    state.mapping = {};
    clearResults();
    empty(el.sheetSelect);
    empty(el.headerSelect);
    empty(el.mappingGrid);
    setText(el.fileName, "");
    el.sheetWrap.hidden = true;
    el.mappingWrap.hidden = true;
    el.headerWrap.hidden = true;
    if (!el.file.files || !el.file.files.length) return;
    showMessage("File is being read locally in this browser…", "info");
    try {
      var file = el.file.files[0];
      var sheets = await readSheets(file);
      if (!sheets.length) throw new Error("File me worksheet nahi mili.");
      selectWorkbook(sheets, file.name);
    } catch (error) {
      setText(el.fileName, "");
      showMessage(error && error.message ? error.message : "File read nahi ho payi. CSV, Excel ya JSON format check karein.", "err");
    }
  });

  el.sheetSelect.addEventListener("change", function () { chooseSheet(el.sheetSelect.value); });
  el.headerSelect.addEventListener("change", function () {
    state.headerIndex = Number(el.headerSelect.value) || 0;
    setMappingFromHeader();
  });
  el.downloadXlsx.addEventListener("click", exportXlsx);
  el.downloadCsv.addEventListener("click", exportCsv);
  el.clear.addEventListener("click", cleanState);

  el.sample.addEventListener("click", function () {
    cleanState();
    var grid = [
      ["Synthetic example workbook — fictional data only"],
      ["Party GSTIN", "Bill No", "Bill Date", "Assessable Value", "Central Tax", "State Tax", "Integrated Tax", "Description"],
      ["27ZZZZZ0000Z1ZQ", "DEMO-001", "05/10/2026", "₹10,000.00", "900", "900", "0", "Fictional local test invoice"],
      ["29YYYYY1111Y1ZP", "DEMO-002", "2026-10-31", "5,000", "0", "0", "900", "Fictional local interstate test invoice"]
    ];
    selectWorkbook([{ name: "Fictional sample", grid: grid }], "Fictional sample (no real data)");
  });
})();
