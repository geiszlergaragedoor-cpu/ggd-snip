/* GGD Snip — core text/number logic (no Office or DOM dependencies). */
(function (root) {
  "use strict";

  // Dates like 01/05, 1/5/2026, 2026-01-05 are removed before looking for amounts,
  // so "01/05 DEPOSIT 1,234.56" yields only 1234.56.
  var DATE_RE = /\b\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?\b|\b\d{4}-\d{2}-\d{2}\b/g;
  var AMOUNT_RE = /(\()?\s*([-−])?\s*\$?\s*(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s*(\))?(-(?![\d]))?(\s*(?:CR|DR)\b)?/gi;

  function findAmounts(text) {
    if (!text) return [];
    var clean = String(text).replace(DATE_RE, " ");
    var out = [];
    var m;
    AMOUNT_RE.lastIndex = 0;
    while ((m = AMOUNT_RE.exec(clean)) !== null) {
      var whole = m[3].replace(/,/g, "");
      var n = parseFloat(whole + (m[4] || ""));
      if (isNaN(n)) continue;
      var negative = (m[1] && m[5]) || m[2] || m[6] || (m[7] && /DR/i.test(m[7]));
      out.push({
        value: negative ? -n : n,
        hasDecimal: !!m[4],
        raw: m[0].trim()
      });
    }
    // Prefer real money amounts (with cents) when the box also caught counts or reference numbers.
    var withCents = out.filter(function (a) { return a.hasDecimal; });
    return withCents.length ? withCents : out;
  }

  function round2(n) { return Math.round(n * 100) / 100; }

  /*
   * items: [{str, x, y, w, h}] in normalized page coords (0..1, origin top-left).
   * rect:  {x0, y0, x1, y1} normalized.
   * Returns {text, lines[]}. Partially covered items are cut by character position.
   */
  function extractTextInRect(items, rect) {
    var picked = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it.str || !it.w || !it.h) continue;
      var top = it.y, bottom = it.y + it.h;
      var vOverlap = Math.min(bottom, rect.y1) - Math.max(top, rect.y0);
      if (vOverlap < it.h * 0.5) continue;
      if (it.x > rect.x1 || it.x + it.w < rect.x0) continue;

      // Keep whole words whose middle falls inside the box (character widths are estimated evenly).
      var n = it.str.length;
      var cw = it.w / n;
      var first = -1, last = -1;
      var wordRe = /\S+/g, wm;
      while ((wm = wordRe.exec(it.str)) !== null) {
        var ws = wm.index, we = wm.index + wm[0].length;
        var mid = it.x + ((ws + we) / 2) * cw;
        if (mid >= rect.x0 && mid <= rect.x1) {
          if (first < 0) first = ws;
          last = we - 1;
        }
      }
      if (first < 0) continue;
      var s = it.str.substring(first, last + 1);
      if (!s.trim()) continue;
      picked.push({
        str: s,
        x0: it.x + first * cw,
        x1: it.x + (last + 1) * cw,
        y: it.y,
        h: it.h
      });
    }

    picked.sort(function (a, b) { return a.y - b.y || a.x0 - b.x0; });
    var lines = [];
    picked.forEach(function (p) {
      var line = lines.length ? lines[lines.length - 1] : null;
      if (line && Math.abs(p.y - line.y) < Math.max(p.h, line.h) * 0.5) {
        line.parts.push(p);
      } else {
        lines.push({ y: p.y, h: p.h, parts: [p] });
      }
    });

    var textLines = lines.map(function (line) {
      line.parts.sort(function (a, b) { return a.x0 - b.x0; });
      var s = "";
      var prevEnd = null;
      line.parts.forEach(function (p) {
        if (prevEnd !== null && p.x0 - prevEnd > p.h * 0.15 && !/\s$/.test(s) && !/^\s/.test(p.str)) s += " ";
        s += p.str;
        prevEnd = p.x1;
      });
      return s.replace(/\s+/g, " ").trim();
    }).filter(Boolean);

    return { text: textLines.join("\n"), lines: textLines };
  }

  /* Decide what goes in the cell for a given mode. */
  function valueForMode(mode, text) {
    var t = (text || "").trim();
    if (mode === "text") {
      if (!t) return { ok: false, error: "No text found in that box." };
      var oneLine = t.replace(/\s*\n\s*/g, " ");
      return { ok: true, value: oneLine, display: oneLine, kind: "text" };
    }
    var amounts = findAmounts(t);
    if (!amounts.length) return { ok: false, error: "No number found in that box." };
    if (mode === "sum") {
      var total = round2(amounts.reduce(function (s, a) { return s + a.value; }, 0));
      return { ok: true, value: total, kind: "number", count: amounts.length,
               display: total + "  (sum of " + amounts.length + ")" };
    }
    var warn = amounts.length > 1 ? "Box had " + amounts.length + " numbers — used the first. Use Sum to add them." : null;
    return { ok: true, value: amounts[0].value, kind: "number", display: String(amounts[0].value), warning: warn };
  }

  /* Excel turns strings starting with = + - @ into formulas/numbers; force text. */
  function safeCellText(s) {
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  }

  function normalizeAddress(addr) {
    if (!addr) return "";
    var i = addr.lastIndexOf("!");
    var sheet = i >= 0 ? addr.substring(0, i) : "";
    var cell = i >= 0 ? addr.substring(i + 1) : addr;
    sheet = sheet.replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
    cell = cell.replace(/\$/g, "").split(":")[0].toUpperCase();
    return sheet + "!" + cell;
  }

  function splitAddress(addr) {
    var n = normalizeAddress(addr);
    var i = n.lastIndexOf("!");
    return { sheet: n.substring(0, i), cell: n.substring(i + 1) };
  }

  var api = {
    findAmounts: findAmounts,
    extractTextInRect: extractTextInRect,
    valueForMode: valueForMode,
    safeCellText: safeCellText,
    normalizeAddress: normalizeAddress,
    splitAddress: splitAddress
  };
  root.SnipCore = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
