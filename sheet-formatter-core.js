(function (root, factory) {
  "use strict";
  var api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (root) root.GSTReconSheetFormatter = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var FIELD_DEFS = [
    { key: "supplier_gstin", header: "Supplier GSTIN", label: "GSTIN (supplier / customer)", required: true, aliases: ["supplier gstin", "gstin", "seller gstin", "ctin", "suppliergstin", "gstin no", "gstin number", "gst no", "vendor gstin", "customer gstin", "party gstin", "buyer gstin", "gstin of supplier", "supplier gstin no"] },
    { key: "invoice_number", header: "Invoice No", label: "Invoice No", required: true, aliases: ["invoice no", "inv no", "invoice number", "invoice", "inum", "invoice no.", "bill no", "bill number", "inv number", "invoice id", "document no", "document number", "voucher no", "invoiceno", "invoicenumber"] },
    { key: "invoice_date", header: "Invoice Date", label: "Invoice Date", required: false, aliases: ["invoice date", "inv date", "date", "idt", "bill date", "inv dt", "invoice dt", "voucher date", "transaction date", "invoicedate"] },
    { key: "taxable_value", header: "Taxable Value", label: "Taxable Value", required: false, numeric: true, aliases: ["taxable value", "taxable amt", "taxable amount", "taxable", "txval", "taxable val", "assessable value", "net amount", "taxablevalue"] },
    { key: "cgst", header: "CGST", label: "CGST", required: false, numeric: true, aliases: ["cgst", "cgst amt", "cgst amount", "camt", "cgst tax", "central tax"] },
    { key: "sgst", header: "SGST", label: "SGST", required: false, numeric: true, aliases: ["sgst", "sgst amt", "sgst amount", "samt", "sgst tax", "state tax", "sgst utgst", "sgst utgst amt"] },
    { key: "igst", header: "IGST", label: "IGST", required: false, numeric: true, aliases: ["igst", "igst amt", "igst amount", "iamt", "igst tax", "integrated tax"] },
    { key: "narration", header: "Narration", label: "Narration (optional)", required: false, aliases: ["narration", "particulars", "particular", "description", "narration particulars", "voucher narration", "bill narration"] }
  ];
  var SITE_HEADERS = FIELD_DEFS.map(function (f) { return f.header; });
  var GSTIN_ALPHA = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  function normHeader(value) {
    return String(value === null || value === undefined ? "" : value)
      .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  }

  function cellString(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "object") {
      try { return JSON.stringify(value); } catch (_) { return String(value); }
    }
    return String(value);
  }

  function jsonToGrid(input) {
    var data = typeof input === "string" ? JSON.parse(input.replace(/^\uFEFF/, "")) : input;
    var rows = null;
    function firstRows(value, depth) {
      if (depth > 4 || value === null || value === undefined) return null;
      if (Array.isArray(value)) {
        if (!value.length) return null;
        if (Array.isArray(value[0])) return value;
        if (value[0] && typeof value[0] === "object") return value;
        return null;
      }
      if (typeof value === "object") {
        var keys = Object.keys(value);
        for (var i = 0; i < keys.length; i++) {
          var found = firstRows(value[keys[i]], depth + 1);
          if (found) return found;
        }
      }
      return null;
    }
    if (Array.isArray(data)) rows = data;
    else if (data && typeof data === "object") {
      rows = firstRows(data, 0);
      if (!rows) {
        var keys = Object.keys(data).map(normHeader);
        var isRow = FIELD_DEFS.some(function (field) {
          return field.aliases.some(function (alias) { return keys.indexOf(alias) >= 0; });
        });
        if (isRow) rows = [data];
      }
    }
    if (!Array.isArray(rows) || !rows.length) throw new Error("JSON me invoice rows nahi mile. Array of objects ya header ke saath row arrays upload karein.");
    if (Array.isArray(rows[0])) {
      return rows.map(function (row) { return row.map(cellString); })
        .filter(function (row) { return row.some(function (cell) { return String(cell).trim() !== ""; }); });
    }
    var headers = [];
    var seen = Object.create(null);
    rows.forEach(function (obj) {
      if (obj && typeof obj === "object" && !Array.isArray(obj)) {
        Object.keys(obj).forEach(function (key) {
          if (!seen[key]) { seen[key] = true; headers.push(key); }
        });
      }
    });
    if (!headers.length) throw new Error("JSON rows me column names nahi mile.");
    return [headers].concat(rows.map(function (obj) {
      return headers.map(function (header) { return obj && Object.prototype.hasOwnProperty.call(obj, header) ? cellString(obj[header]) : ""; });
    }));
  }

  function countDelimiter(line, delimiter) {
    var quoted = false, count = 0;
    for (var i = 0; i < line.length; i++) {
      var ch = line.charAt(i);
      if (ch === '"') {
        if (quoted && line.charAt(i + 1) === '"') i++;
        else quoted = !quoted;
      } else if (!quoted && ch === delimiter) count++;
    }
    return count;
  }

  function detectDelimiter(text) {
    var lines = String(text || "").replace(/^\uFEFF/, "").split(/\r?\n/).filter(function (line) { return line.trim() !== ""; }).slice(0, 8);
    var candidates = [",", ";", "\t", "|"];
    var best = ",", bestScore = 0;
    candidates.forEach(function (delimiter) {
      var score = lines.reduce(function (sum, line) { return sum + countDelimiter(line, delimiter); }, 0);
      if (score > bestScore) { bestScore = score; best = delimiter; }
    });
    return best;
  }

  function parseDelimited(text, requestedDelimiter) {
    text = String(text === null || text === undefined ? "" : text).replace(/^\uFEFF/, "");
    if (!text.trim()) return [];
    var delimiter = requestedDelimiter || detectDelimiter(text);
    var rows = [], row = [], cell = "", quoted = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (quoted) {
        if (ch === '"') {
          if (text.charAt(i + 1) === '"') { cell += '"'; i++; }
          else quoted = false;
        } else cell += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === delimiter) { row.push(cell); cell = ""; }
      else if (ch === "\n") { row.push(cell); cell = ""; rows.push(row); row = []; }
      else if (ch !== "\r") cell += ch;
    }
    if (quoted) throw new Error("CSV me ek quoted cell band nahi hua — file ko dobara export karke try karein.");
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (cells) { return cells.some(function (value) { return String(value).trim() !== ""; }); });
  }

  function detectHeaderRow(grid, limit) {
    var max = Math.min((grid || []).length, limit || 15);
    var bestIndex = -1, bestScore = -1;
    for (var r = 0; r < max; r++) {
      var headers = (grid[r] || []).map(normHeader);
      var score = FIELD_DEFS.reduce(function (sum, field) {
        return sum + (field.aliases.some(function (alias) { return headers.indexOf(alias) >= 0; }) ? 1 : 0);
      }, 0);
      if (score > bestScore) { bestScore = score; bestIndex = r; }
    }
    var autoDetected = bestScore > 0;
    if (!autoDetected) {
      bestIndex = (grid || []).findIndex(function (row) { return (row || []).some(function (cell) { return String(cellString(cell)).trim() !== ""; }); });
    }
    return { index: bestIndex < 0 ? 0 : bestIndex, score: Math.max(0, bestScore), autoDetected: autoDetected };
  }

  function suggestMappings(headers) {
    var normalized = (headers || []).map(normHeader);
    var result = {};
    FIELD_DEFS.forEach(function (field) {
      var match = null;
      for (var a = 0; a < field.aliases.length && match === null; a++) {
        var index = normalized.indexOf(field.aliases[a]);
        if (index >= 0) match = index;
      }
      result[field.key] = match;
    });
    return result;
  }

  function parseAmount(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === "number") return isNaN(value) ? null : value;
    var text = String(value).trim();
    if (text === "" || text === "-" || text === "--") return null;
    var negative = false;
    if (/^\(.*\)$/.test(text)) { negative = true; text = text.slice(1, -1); }
    text = text.replace(/,/g, "").replace(/₹/g, "").replace(/rs\.?/gi, "").replace(/\s+/g, "");
    if (text.charAt(text.length - 1) === "-") { negative = true; text = text.slice(0, -1); }
    if (text === "") return null;
    var number = parseFloat(text);
    if (isNaN(number) || !isFinite(number)) return null;
    return negative ? -Math.abs(number) : number;
  }

  function normalizeDate(value) {
    if (value === null || value === undefined) return "";
    var text = String(value).trim();
    if (!text) return "";
    text = text.replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\s*[AP]M)?$/i, "");
    text = text.replace(/[.]/g, "/");
    var months = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
    var monthDate = text.match(/^(\d{1,2})[\/\- ]([A-Za-z]{3,9})[\/\- ](\d{2,4})$/);
    if (monthDate && months[monthDate[2].slice(0, 3).toLowerCase()]) {
      var monthYear = monthDate[3].length === 2 ? "20" + monthDate[3] : monthDate[3];
      return monthYear + "-" + ("0" + months[monthDate[2].slice(0, 3).toLowerCase()]).slice(-2) + "-" + ("0" + monthDate[1]).slice(-2);
    }
    var match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/) || text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$/);
    if (match) {
      var day = +match[1], month = +match[2];
      if (month > 12 && day <= 12) { var temp = day; day = month; month = temp; }
      var year = match[3].length === 2 ? "20" + match[3] : match[3];
      return year + "-" + ("0" + month).slice(-2) + "-" + ("0" + day).slice(-2);
    }
    match = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (match) return match[1] + "-" + ("0" + match[2]).slice(-2) + "-" + ("0" + match[3]).slice(-2);
    return text;
  }

  function validateGSTIN(raw) {
    if (raw === null || raw === undefined || !String(raw).trim()) return { ok: false, reason: "Missing GSTIN" };
    var gstin = String(raw).trim().toUpperCase();
    if (gstin.length !== 15) return { ok: false, reason: "GSTIN must be 15 characters" };
    var state = gstin.slice(0, 2);
    if (!/^\d{2}$/.test(state) || +state < 1 || +state > 38) return { ok: false, reason: "Invalid state code" };
    if (!/^[A-Z]{5}[0-9]{4}[A-Z]$/.test(gstin.slice(2, 12))) return { ok: false, reason: "Invalid PAN pattern" };
    if (!/^[0-9A-Z]$/.test(gstin.charAt(12))) return { ok: false, reason: "Invalid entity code" };
    if (gstin.charAt(13) !== "Z") return { ok: false, reason: "Character 14 must be Z" };
    var total = 0;
    for (var i = 0; i < 14; i++) {
      var value = GSTIN_ALPHA.indexOf(gstin.charAt(i));
      if (value < 0) return { ok: false, reason: "Invalid GSTIN character" };
      var product = value * (i % 2 === 1 ? 2 : 1);
      total += Math.floor(product / 36) + (product % 36);
    }
    var checkDigit = GSTIN_ALPHA.charAt((36 - (total % 36)) % 36);
    if (gstin.charAt(14) !== checkDigit) return { ok: false, reason: "GSTIN check digit mismatch" };
    return { ok: true, gstin: gstin };
  }

  function transformGrid(grid, headerRowIndex, mapping) {
    grid = (grid || []).filter(function (row) { return Array.isArray(row); });
    mapping = mapping || {};
    var errors = [];
    if (!grid.length) return { rows: [], errors: ["File me koi rows nahi mili."], blankSkipped: 0, totalSkipped: 0 };
    var headerIndex = Number(headerRowIndex);
    if (!Number.isInteger(headerIndex) || headerIndex < 0 || headerIndex >= grid.length) return { rows: [], errors: ["Header row select karein."], blankSkipped: 0, totalSkipped: 0 };
    var headers = grid[headerIndex] || [];
    FIELD_DEFS.filter(function (field) { return field.required; }).forEach(function (field) {
      var column = mapping[field.key];
      if (column === null || column === undefined || column === "" || !Number.isInteger(Number(column)) || Number(column) < 0 || Number(column) >= headers.length) errors.push(field.header + " column map karna zaroori hai.");
    });
    if (errors.length) return { rows: [], errors: errors, blankSkipped: 0, totalSkipped: 0 };

    var normalizedHeader = headers.map(normHeader).join("|");
    var rows = [], blankSkipped = 0, totalSkipped = 0;
    for (var r = headerIndex + 1; r < grid.length; r++) {
      var source = grid[r] || [];
      if (!source.some(function (cell) { return String(cellString(cell)).trim() !== ""; })) { blankSkipped++; continue; }
      if (source.map(normHeader).join("|") === normalizedHeader) { totalSkipped++; continue; }
      function rawFor(key) {
        var col = mapping[key];
        if (col === null || col === undefined || col === "" || !Number.isInteger(Number(col)) || Number(col) < 0) return "";
        return cellString(source[Number(col)]).trim();
      }
      var gstin = rawFor("supplier_gstin").toUpperCase().replace(/\s+/g, "");
      var invoiceNo = rawFor("invoice_number");
      var gstinStatus = validateGSTIN(gstin);
      var issues = [];
      if (!gstin) issues.push("GSTIN is missing.");
      else if (!gstinStatus.ok) issues.push(gstinStatus.reason + ".");
      if (!invoiceNo) issues.push("Invoice No is missing.");

      var sourceDate = rawFor("invoice_date");
      var invoiceDate = normalizeDate(sourceDate);
      var numericValues = {};
      ["taxable_value", "cgst", "sgst", "igst"].forEach(function (key) {
        var raw = rawFor(key);
        var parsed = parseAmount(raw);
        if (parsed === null) {
          numericValues[key] = 0;
          if (raw) issues.push(FIELD_DEFS.find(function (field) { return field.key === key; }).header + " is not a readable number; exported as 0.");
        } else numericValues[key] = parsed;
      });
      var values = {
        supplier_gstin: gstin,
        invoice_number: invoiceNo,
        invoice_date: invoiceDate,
        taxable_value: numericValues.taxable_value,
        cgst: numericValues.cgst,
        sgst: numericValues.sgst,
        igst: numericValues.igst,
        narration: rawFor("narration")
      };
      var invoiceLow = invoiceNo.toLowerCase(), gstinLow = gstin.toLowerCase();
      if (invoiceLow === "total" || invoiceLow === "grand total" || gstinLow === "total" || gstinLow === "grand total") { totalSkipped++; continue; }
      rows.push({ sourceRow: r + 1, values: values, issues: issues, ready: !!gstin && !!invoiceNo && gstinStatus.ok });
    }
    if (!rows.length) errors.push("Header ke neeche koi invoice row nahi mili.");
    var optionalDefaults = FIELD_DEFS.filter(function (field) {
      if (field.required) return false;
      var col = mapping[field.key];
      return col === null || col === undefined || col === "" || !Number.isInteger(Number(col)) || Number(col) < 0;
    }).map(function (field) { return field.header; });
    return { rows: rows, errors: errors, blankSkipped: blankSkipped, totalSkipped: totalSkipped, optionalDefaults: optionalDefaults };
  }

  function recordsToMatrix(records) {
    return [SITE_HEADERS].concat((records || []).map(function (record) {
      return FIELD_DEFS.map(function (field) { return record.values[field.key]; });
    }));
  }

  function outputFileName(sourceName, extension) {
    var ext = String(extension || "xlsx").replace(/^\./, "").replace(/[^a-z0-9]/gi, "").toLowerCase() || "xlsx";
    var base = String(sourceName || "spreadsheet").replace(/\.[^.]*$/, "")
      .replace(/[^a-z0-9_-]+/gi, "-").replace(/-{2,}/g, "-").replace(/^-+|-+$/g, "");
    if (!base) base = "spreadsheet";
    return base + "-gst-recon-formatted." + ext;
  }

  function csvCell(value) {
    var text = value === null || value === undefined ? "" : String(value);
    return '"' + text.replace(/"/g, '""') + '"';
  }

  function toCsv(records) {
    return "\uFEFF" + recordsToMatrix(records).map(function (row) { return row.map(csvCell).join(","); }).join("\r\n") + "\r\n";
  }

  function hasFormulaLikeText(records) {
    var textFields = ["supplier_gstin", "invoice_number", "invoice_date", "narration"];
    return (records || []).some(function (record) {
      return textFields.some(function (key) { return /^\s*[=+\-@]/.test(String(record.values[key] || "")); });
    });
  }

  function columnLetter(index) {
    var n = Number(index) + 1, result = "";
    while (n > 0) {
      var rem = (n - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  }

  return {
    FIELD_DEFS: FIELD_DEFS,
    SITE_HEADERS: SITE_HEADERS,
    normHeader: normHeader,
    jsonToGrid: jsonToGrid,
    parseDelimited: parseDelimited,
    detectDelimiter: detectDelimiter,
    detectHeaderRow: detectHeaderRow,
    suggestMappings: suggestMappings,
    parseAmount: parseAmount,
    normalizeDate: normalizeDate,
    validateGSTIN: validateGSTIN,
    transformGrid: transformGrid,
    recordsToMatrix: recordsToMatrix,
    outputFileName: outputFileName,
    toCsv: toCsv,
    hasFormulaLikeText: hasFormulaLikeText,
    columnLetter: columnLetter
  };
});
