/* GGD Snip — Excel task pane: snip values from PDFs and link cells back to the source. */
(function () {
  "use strict";
  var S = window.SnipCore;
  var $ = function (id) { return document.getElementById(id); };

  var PDFJS_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  var TESSERACT_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
  var MARK_COLOR = "#DDEBF7";
  var SHEET = "GGD_Snips";
  var NAME_PREFIX = "GGDSNIP_";
  var HEAD = ["ID", "Doc Key", "Document", "Page", "X0", "Y0", "X1", "Y1", "Snipped Text", "Value", "Snipped At", "Cell (when snipped)"];
  var NUM_FORMAT = "#,##0.00_);(#,##0.00)";

  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;

  var state = {
    mode: "number",
    doc: null,          // {key, name, pdf, numPages}
    pageNum: 1,
    zoom: 1,
    textCache: {},
    snips: [],
    byAddr: new Map(),
    hitId: null,
    renderSeq: 0,
    renderTask: null,
    busy: false
  };
  var XL = null;

  /* ---------------- status ---------------- */
  function setStatus(msg, kind) {
    var el = $("status");
    el.textContent = msg;
    el.className = "status" + (kind ? " " + kind : "");
  }
  function errMsg(e) { return (e && (e.message || e.toString())) || "Unknown error"; }

  /* ---------------- local document store (IndexedDB) ---------------- */
  var DocStore = (function () {
    var memory = new Map();
    var dbp = null;
    function open() {
      if (dbp) return dbp;
      dbp = new Promise(function (resolve) {
        try {
          var req = indexedDB.open("ggd-snip", 1);
          req.onupgradeneeded = function () {
            var db = req.result;
            db.createObjectStore("meta", { keyPath: "key" });
            db.createObjectStore("bytes");
          };
          req.onsuccess = function () { resolve(req.result); };
          req.onerror = function () { resolve(null); };
        } catch (e) { resolve(null); }
      });
      return dbp;
    }
    function tx(db, stores, mode, fn) {
      return new Promise(function (resolve, reject) {
        var t = db.transaction(stores, mode);
        var out = fn(t);
        t.oncomplete = function () { resolve(out && out.result !== undefined ? out.result : out); };
        t.onerror = function () { reject(t.error); };
      });
    }
    return {
      put: function (meta, bytes) {
        memory.set(meta.key, { meta: meta, bytes: bytes });
        return open().then(function (db) {
          if (!db) return;
          return tx(db, ["meta", "bytes"], "readwrite", function (t) {
            t.objectStore("meta").put(meta);
            t.objectStore("bytes").put(bytes, meta.key);
          }).catch(function () {});
        });
      },
      list: function () {
        return open().then(function (db) {
          if (!db) return Array.from(memory.values()).map(function (v) { return v.meta; });
          return tx(db, ["meta"], "readonly", function (t) { return t.objectStore("meta").getAll(); });
        }).then(function (rows) {
          return (rows || []).sort(function (a, b) { return b.added - a.added; });
        });
      },
      get: function (key) {
        if (memory.has(key)) return Promise.resolve(memory.get(key).bytes);
        return open().then(function (db) {
          if (!db) return null;
          return tx(db, ["bytes"], "readonly", function (t) { return t.objectStore("bytes").get(key); });
        });
      },
      del: function (key) {
        memory.delete(key);
        return open().then(function (db) {
          if (!db) return;
          return tx(db, ["meta", "bytes"], "readwrite", function (t) {
            t.objectStore("meta").delete(key);
            t.objectStore("bytes").delete(key);
          });
        });
      }
    };
  })();

  function hashBytes(buf) {
    if (window.crypto && crypto.subtle) {
      return crypto.subtle.digest("SHA-256", buf).then(function (h) {
        return Array.from(new Uint8Array(h)).slice(0, 10).map(function (b) { return b.toString(16).padStart(2, "0"); }).join("");
      });
    }
    // Fallback: simple FNV-1a over the bytes plus length.
    var a = new Uint8Array(buf), h = 2166136261;
    for (var i = 0; i < a.length; i++) { h ^= a[i]; h = Math.imul(h, 16777619); }
    return Promise.resolve((h >>> 0).toString(16) + "-" + a.length.toString(16));
  }

  /* ---------------- Excel adapter ---------------- */
  function rowToRec(r) {
    return {
      id: String(r[0]), docKey: String(r[1]), docName: String(r[2]), page: Number(r[3]),
      x0: Number(r[4]), y0: Number(r[5]), x1: Number(r[6]), y1: Number(r[7]),
      text: String(r[8]), value: r[9], at: String(r[10]), address: null
    };
  }
  function recToRow(rec, addr) {
    return [rec.id, rec.docKey, rec.docName, rec.page,
      +rec.x0.toFixed(5), +rec.y0.toFixed(5), +rec.x1.toFixed(5), +rec.y1.toFixed(5),
      S.safeCellText(rec.text.slice(0, 2000)), typeof rec.value === "number" ? rec.value : S.safeCellText(String(rec.value)),
      rec.at, addr];
  }

  var ExcelAdapter = {
    loadSnips: function () {
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
        var recs = [], nis = [], rgs = [];
        return ctx.sync().then(function () {
          if (ws.isNullObject) return null;
          var used = ws.getUsedRangeOrNullObject(true);
          used.load("values");
          return ctx.sync().then(function () { return used; });
        }).then(function (used) {
          if (!used || used.isNullObject) return;
          recs = used.values.slice(1).filter(function (r) { return r[0]; }).map(rowToRec);
          nis = recs.map(function (r) { return ctx.workbook.names.getItemOrNullObject(NAME_PREFIX + r.id); });
          return ctx.sync().then(function () {
            rgs = nis.map(function (n) {
              if (n.isNullObject) return null;
              var rg = n.getRangeOrNullObject();
              rg.load("address");
              return rg;
            });
            return ctx.sync();
          }).then(function () {
            recs.forEach(function (r, i) {
              r.address = rgs[i] && !rgs[i].isNullObject ? S.normalizeAddress(rgs[i].address) : null;
            });
          });
        }).then(function () { return recs; });
      });
    },

    getSelection: function () {
      return Excel.run(function (ctx) {
        var r = ctx.workbook.getSelectedRange();
        r.load("address");
        return ctx.sync().then(function () { return S.normalizeAddress(r.address); });
      });
    },

    selectCell: function (address) {
      var p = S.splitAddress(address);
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItem(p.sheet);
        ws.activate();
        ws.getRange(p.cell).select();
        return ctx.sync();
      });
    },

    addSnip: function (rec, result, opts, replaceId) {
      return Excel.run(function (ctx) {
        var wb = ctx.workbook;
        var target = wb.getSelectedRange().getCell(0, 0);
        target.load("address,numberFormat");
        var addr;
        return ctx.sync().then(function () {
          addr = S.normalizeAddress(target.address);
          if (result.kind === "number") {
            target.values = [[result.value]];
            if (target.numberFormat[0][0] === "General") target.numberFormat = [[NUM_FORMAT]];
          } else {
            target.values = [[S.safeCellText(result.value)]];
          }
          if (opts.mark) target.format.fill.color = MARK_COLOR;
          if (replaceId) return removeInContext(ctx, replaceId, false);
        }).then(function () {
          wb.names.add(NAME_PREFIX + rec.id, target, "GGD Snip link");
          return ensureSheet(ctx);
        }).then(function (ws) {
          var used = ws.getUsedRange(true);
          used.load("rowCount");
          return ctx.sync().then(function () {
            ws.getRangeByIndexes(used.rowCount, 0, 1, HEAD.length).values = [recToRow(rec, addr)];
            if (opts.moveDown) target.getOffsetRange(1, 0).select();
            return ctx.sync();
          });
        }).then(function () { return addr; });
      });
    },

    removeSnip: function (id) {
      return Excel.run(function (ctx) { return removeInContext(ctx, id, true); });
    }
  };

  function ensureSheet(ctx) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
    return ctx.sync().then(function () {
      if (!ws.isNullObject) return ws;
      var nws = ctx.workbook.worksheets.add(SHEET);
      var head = nws.getRangeByIndexes(0, 0, 1, HEAD.length);
      head.values = [HEAD];
      head.format.font.bold = true;
      nws.visibility = Excel.SheetVisibility.hidden;
      return ctx.sync().then(function () { return nws; });
    });
  }

  function removeInContext(ctx, id, clearFill) {
    var ni = ctx.workbook.names.getItemOrNullObject(NAME_PREFIX + id);
    var ws = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
    return ctx.sync().then(function () {
      if (!ni.isNullObject) {
        if (clearFill) {
          var rg = ni.getRangeOrNullObject();
          rg.load("address");
          rg.format.fill.load("color");
          return ctx.sync().then(function () {
            if (!rg.isNullObject && String(rg.format.fill.color).toUpperCase() === MARK_COLOR) rg.format.fill.clear();
            ni.delete();
          });
        }
        ni.delete();
      }
    }).then(function () {
      if (ws.isNullObject) return ctx.sync();
      var used = ws.getUsedRange(true);
      used.load("values");
      return ctx.sync().then(function () {
        for (var i = used.values.length - 1; i >= 1; i--) {
          if (String(used.values[i][0]) === String(id)) ws.getRange((i + 1) + ":" + (i + 1)).delete(Excel.DeleteShiftDirection.up);
        }
        return ctx.sync();
      });
    });
  }

  /* In-browser stand-in so the pane can be tried outside Excel. */
  var MockAdapter = (function () {
    var cells = {}, links = {}, sel = "Sheet1!B2";
    function down(a) {
      var p = S.splitAddress(a), m = /^([A-Z]+)(\d+)$/.exec(p.cell);
      return p.sheet + "!" + m[1] + (Number(m[2]) + 1);
    }
    var api = {
      cells: cells,
      loadSnips: function () {
        return Promise.resolve(Object.keys(links).map(function (k) { return links[k]; }));
      },
      getSelection: function () { return Promise.resolve(sel); },
      selectCell: function (a) { sel = a; setTimeout(onSelectionChanged, 0); return Promise.resolve(); },
      addSnip: function (rec, result, opts, replaceId) {
        if (replaceId) delete links[replaceId];
        cells[sel] = result.value;
        var copy = Object.assign({}, rec, { address: sel });
        links[rec.id] = copy;
        var addr = sel;
        if (opts.moveDown) sel = down(sel);
        return Promise.resolve(addr);
      },
      removeSnip: function (id) { delete links[id]; return Promise.resolve(); },
      setSelection: function (a) { sel = a; onSelectionChanged(); }
    };
    return api;
  })();

  /* ---------------- links ---------------- */
  function rebuildIndex() {
    state.byAddr = new Map();
    state.snips.forEach(function (r) { if (r.address) state.byAddr.set(r.address, r); });
  }

  function refreshSnips() {
    return XL.loadSnips().then(function (recs) {
      state.snips = recs;
      rebuildIndex();
      drawBoxes();
    });
  }

  /* ---------------- PDF ---------------- */
  function openFiles(files) {
    var list = Array.from(files || []);
    if (!list.length) return;
    setStatus("Opening " + list.length + " file" + (list.length > 1 ? "s" : "") + "…");
    var lastKey = null;
    return list.reduce(function (p, f) {
      return p.then(function () {
        return f.arrayBuffer().then(function (buf) {
          return hashBytes(buf).then(function (key) {
            lastKey = key;
            return DocStore.put({ key: key, name: f.name, size: f.size, added: Date.now() }, buf);
          });
        });
      });
    }, Promise.resolve()).then(function () {
      return refreshDocList();
    }).then(function () {
      return openDoc(lastKey, 1);
    }).catch(function (e) { setStatus("Couldn't open that file: " + errMsg(e), "err"); });
  }

  function refreshDocList() {
    return DocStore.list().then(function (docs) {
      var sel = $("docSelect");
      sel.innerHTML = "";
      if (!docs.length) {
        sel.innerHTML = '<option value="">No documents yet</option>';
        return;
      }
      docs.forEach(function (d) {
        var o = document.createElement("option");
        o.value = d.key;
        o.textContent = d.name;
        sel.appendChild(o);
      });
      if (state.doc) sel.value = state.doc.key;
    });
  }

  function openDoc(key, page) {
    if (!key) return Promise.resolve(false);
    if (state.doc && state.doc.key === key) {
      return goToPage(page || state.pageNum).then(function () { return true; });
    }
    return Promise.all([DocStore.get(key), DocStore.list()]).then(function (res) {
      var bytes = res[0];
      var meta = res[1].filter(function (m) { return m.key === key; })[0];
      if (!bytes || !meta) return false;
      return pdfjsLib.getDocument({ data: new Uint8Array(bytes.slice(0)) }).promise.then(function (pdf) {
        if (state.doc && state.doc.pdf) state.doc.pdf.destroy();
        state.doc = { key: key, name: meta.name, pdf: pdf, numPages: pdf.numPages };
        state.textCache = {};
        state.pageNum = Math.min(Math.max(1, page || 1), pdf.numPages);
        $("docSelect").value = key;
        $("empty").hidden = true;
        $("pageWrap").hidden = false;
        setStatus("Opened " + meta.name + ". Click a cell in Excel, then drag a box on the PDF.");
        return renderPage().then(function () { return true; });
      });
    });
  }

  function goToPage(n) {
    if (!state.doc) return Promise.resolve();
    n = Math.min(Math.max(1, n), state.doc.numPages);
    if (n === state.pageNum && $("canvas").width) { drawBoxes(); return Promise.resolve(); }
    state.pageNum = n;
    return renderPage();
  }

  function renderPage() {
    if (!state.doc) return Promise.resolve();
    var seq = ++state.renderSeq;
    var n = state.pageNum;
    $("pageLabel").textContent = n + " / " + state.doc.numPages;
    return state.doc.pdf.getPage(n).then(function (page) {
      if (seq !== state.renderSeq) return;
      var base = page.getViewport({ scale: 1 });
      var avail = Math.max(200, $("viewer").clientWidth - 18);
      var scale = (avail / base.width) * state.zoom;
      var vp = page.getViewport({ scale: scale });
      var dpr = window.devicePixelRatio || 1;
      var canvas = $("canvas");
      canvas.width = Math.floor(vp.width * dpr);
      canvas.height = Math.floor(vp.height * dpr);
      canvas.style.width = Math.floor(vp.width) + "px";
      canvas.style.height = Math.floor(vp.height) + "px";
      $("pageWrap").style.width = Math.floor(vp.width) + "px";
      $("pageWrap").style.height = Math.floor(vp.height) + "px";
      if (state.renderTask) { try { state.renderTask.cancel(); } catch (e) {} }
      var task = page.render({
        canvasContext: canvas.getContext("2d"),
        viewport: vp,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null
      });
      state.renderTask = task;
      drawBoxes();
      return task.promise.catch(function (e) {
        if (e && e.name === "RenderingCancelledException") return;
        throw e;
      });
    });
  }

  function getTextItems(n) {
    if (state.textCache[n]) return Promise.resolve(state.textCache[n]);
    return state.doc.pdf.getPage(n).then(function (page) {
      var vp = page.getViewport({ scale: 1 });
      return page.getTextContent().then(function (tc) {
        var items = tc.items.filter(function (i) { return i.str && i.str.trim(); }).map(function (i) {
          var t = pdfjsLib.Util.transform(vp.transform, i.transform);
          var fh = Math.hypot(t[2], t[3]) || i.height || 10;
          var w = i.width * vp.scale;
          return {
            str: i.str,
            x: t[4] / vp.width,
            y: (t[5] - fh * 0.85) / vp.height,
            w: w / vp.width,
            h: (fh * 1.05) / vp.height
          };
        });
        state.textCache[n] = items;
        return items;
      });
    });
  }

  /* ---------------- OCR (scanned pages) ---------------- */
  var tessWorkerP = null;
  function getOcrWorker() {
    if (tessWorkerP) return tessWorkerP;
    tessWorkerP = new Promise(function (resolve, reject) {
      if (window.Tesseract) return resolve(window.Tesseract);
      var s = document.createElement("script");
      s.src = TESSERACT_URL;
      s.onload = function () { resolve(window.Tesseract); };
      s.onerror = function () { reject(new Error("Couldn't load the OCR reader (check your internet).")); };
      document.head.appendChild(s);
    }).then(function (T) { return T.createWorker("eng"); });
    tessWorkerP.catch(function () { tessWorkerP = null; });
    return tessWorkerP;
  }

  function ocrRect(r) {
    return Promise.all([getOcrWorker(), state.doc.pdf.getPage(state.pageNum)]).then(function (res) {
      var worker = res[0], page = res[1];
      var vp = page.getViewport({ scale: 3 });
      var pad = 12;
      var cx = r.x0 * vp.width, cy = r.y0 * vp.height;
      var cw = (r.x1 - r.x0) * vp.width, ch = (r.y1 - r.y0) * vp.height;
      var c = document.createElement("canvas");
      c.width = Math.ceil(cw + pad * 2);
      c.height = Math.ceil(ch + pad * 2);
      var g = c.getContext("2d");
      g.fillStyle = "#fff";
      g.fillRect(0, 0, c.width, c.height);
      return page.render({ canvasContext: g, viewport: vp, transform: [1, 0, 0, 1, -cx + pad, -cy + pad] }).promise
        .then(function () { return worker.recognize(c); })
        .then(function (out) { return (out && out.data && out.data.text) || ""; });
    });
  }

  /* ---------------- snipping ---------------- */
  function newId() {
    return "S" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();
  }

  function shortAddr(a) {
    var p = S.splitAddress(a);
    return p.cell;
  }

  function doSnip(r) {
    if (!state.doc || state.busy) return;
    state.busy = true;
    $("overlay").classList.add("busy");
    var viaOcr = false;
    var targetAddr;
    getTextItems(state.pageNum).then(function (items) {
      var text = S.extractTextInRect(items, r).text;
      if (text.trim() || !$("optOcr").checked) return text;
      viaOcr = true;
      setStatus("No text layer here — reading the scan (first time takes a few seconds)…");
      return ocrRect(r);
    }).then(function (text) {
      var res = S.valueForMode(state.mode, text);
      if (!res.ok) {
        setStatus(res.error + (viaOcr ? "" : " Try a slightly bigger box."), "err");
        return;
      }
      return XL.getSelection().then(function (sel) {
        targetAddr = sel;
        var existing = state.byAddr.get(sel);
        var rec = {
          id: newId(), docKey: state.doc.key, docName: state.doc.name, page: state.pageNum,
          x0: r.x0, y0: r.y0, x1: r.x1, y1: r.y1,
          text: text.trim(), value: res.value, at: new Date().toISOString()
        };
        var opts = { moveDown: $("optMoveDown").checked, mark: $("optMark").checked };
        return XL.addSnip(rec, res, opts, existing && existing.id).then(function (addr) {
          rec.address = addr;
          state.snips = state.snips.filter(function (s) { return s.address !== addr && (!existing || s.id !== existing.id); });
          state.snips.push(rec);
          rebuildIndex();
          state.hitId = rec.id;
          drawBoxes();
          var msg = shortAddr(addr) + " ← " + res.display;
          if (res.warning) msg += "\n" + res.warning;
          if (viaOcr) msg += "\nRead from a scan — double-check the value.";
          setStatus(msg, res.warning || viaOcr ? "" : "ok");
        });
      });
    }).catch(function (e) {
      setStatus("Snip failed" + (targetAddr ? " for " + shortAddr(targetAddr) : "") + ": " + errMsg(e), "err");
    }).then(function () {
      state.busy = false;
      $("overlay").classList.remove("busy");
    });
  }

  /* ---------------- overlay ---------------- */
  function drawBoxes() {
    var ov = $("overlay");
    Array.from(ov.querySelectorAll(".box")).forEach(function (b) { b.remove(); });
    if (!state.doc) return;
    var showAll = $("optShowBoxes").checked;
    state.snips.forEach(function (s) {
      if (s.docKey !== state.doc.key || s.page !== state.pageNum || !s.address) return;
      var isHit = s.id === state.hitId;
      if (!showAll && !isHit) return;
      var b = document.createElement("div");
      b.className = "box" + (isHit ? " hit" : "");
      b.dataset.id = s.id;
      b.style.left = (s.x0 * 100) + "%";
      b.style.top = (s.y0 * 100) + "%";
      b.style.width = ((s.x1 - s.x0) * 100) + "%";
      b.style.height = ((s.y1 - s.y0) * 100) + "%";
      b.title = s.address + "  ←  " + (s.text || "").replace(/\s+/g, " ").slice(0, 80);
      var tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = shortAddr(s.address);
      b.appendChild(tag);
      ov.appendChild(b);
    });
  }

  function flashHit() {
    var el = $("overlay").querySelector(".box.hit");
    if (!el) return;
    el.classList.remove("flash");
    void el.offsetWidth;
    el.classList.add("flash");
    var viewer = $("viewer");
    var wrap = $("pageWrap");
    var top = wrap.offsetTop + el.offsetTop - viewer.clientHeight / 3;
    var left = wrap.offsetLeft + el.offsetLeft - viewer.clientWidth / 3;
    viewer.scrollTo({ top: Math.max(0, top), left: Math.max(0, left), behavior: "smooth" });
  }

  function setupDrag() {
    var ov = $("overlay");
    var start = null, dragEl = null, startBox = null;
    function rel(e) {
      var b = ov.getBoundingClientRect();
      return {
        x: Math.min(Math.max(e.clientX - b.left, 0), b.width),
        y: Math.min(Math.max(e.clientY - b.top, 0), b.height),
        w: b.width, h: b.height
      };
    }
    ov.addEventListener("pointerdown", function (e) {
      if (e.button !== 0 || !state.doc) return;
      start = rel(e);
      var box = e.target.closest && e.target.closest(".box");
      startBox = box ? box.dataset.id : null;
      dragEl = document.createElement("div");
      dragEl.className = "drag";
      dragEl.hidden = true;
      ov.appendChild(dragEl);
      ov.setPointerCapture(e.pointerId);
      e.preventDefault();
    });
    ov.addEventListener("pointermove", function (e) {
      if (!start) return;
      var p = rel(e);
      var x = Math.min(start.x, p.x), y = Math.min(start.y, p.y);
      var w = Math.abs(p.x - start.x), h = Math.abs(p.y - start.y);
      if (w > 4 || h > 4) dragEl.hidden = false;
      dragEl.style.left = x + "px";
      dragEl.style.top = y + "px";
      dragEl.style.width = w + "px";
      dragEl.style.height = h + "px";
    });
    function end(e, cancelled) {
      if (!start) return;
      var p = rel(e);
      var w = Math.abs(p.x - start.x), h = Math.abs(p.y - start.y);
      var r = {
        x0: Math.min(start.x, p.x) / p.w, y0: Math.min(start.y, p.y) / p.h,
        x1: Math.max(start.x, p.x) / p.w, y1: Math.max(start.y, p.y) / p.h
      };
      if (dragEl) dragEl.remove();
      var clickedBox = startBox;
      start = null; dragEl = null; startBox = null;
      if (cancelled) return;
      if (w < 5 && h < 5) {
        if (clickedBox) {
          var rec = state.snips.filter(function (s) { return s.id === clickedBox; })[0];
          if (rec) XL.selectCell(rec.address).catch(function (er) { setStatus("Couldn't select " + rec.address + ": " + errMsg(er), "err"); });
        }
        return;
      }
      doSnip(r);
    }
    ov.addEventListener("pointerup", function (e) { end(e, false); });
    ov.addEventListener("pointercancel", function (e) { end(e, true); });
  }

  /* ---------------- selected cell ---------------- */
  var selTimer = null;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(handleSelection, 150);
  }

  function handleSelection() {
    var addr;
    XL.getSelection().then(function (a) {
      addr = a;
      return refreshSnips();
    }).then(function () {
      var rec = state.byAddr.get(addr);
      showCellInfo(addr, rec);
      if (!rec) {
        if (state.hitId) { state.hitId = null; drawBoxes(); }
        return;
      }
      state.hitId = rec.id;
      return openDoc(rec.docKey, rec.page).then(function (ok) {
        if (ok) { drawBoxes(); setTimeout(flashHit, 60); }
      });
    }).catch(function (e) { setStatus(errMsg(e), "err"); });
  }

  function showCellInfo(addr, rec) {
    var el = $("cellInfo");
    if (!rec) { el.hidden = true; el.innerHTML = ""; return; }
    el.hidden = false;
    el.innerHTML = "";
    var t = document.createElement("span");
    t.textContent = shortAddr(addr) + " is linked to " + rec.docName + ", page " + rec.page + ". ";
    el.appendChild(t);
    DocStore.get(rec.docKey).then(function (bytes) {
      if (!bytes) {
        var miss = document.createElement("span");
        miss.textContent = "That PDF isn't on this computer yet — open it to restore the link. ";
        el.appendChild(miss);
      }
      var rm = document.createElement("button");
      rm.className = "btn link";
      rm.textContent = "Remove link";
      rm.onclick = function () {
        XL.removeSnip(rec.id).then(function () {
          state.hitId = null;
          return refreshSnips();
        }).then(function () {
          showCellInfo(addr, null);
          setStatus("Link removed from " + shortAddr(addr) + ". The value stays in the cell.");
        }).catch(function (e) { setStatus(errMsg(e), "err"); });
      };
      el.appendChild(rm);
    });
  }

  /* ---------------- wiring ---------------- */
  function wireUi() {
    $("btnOpen").onclick = function () { $("fileInput").click(); };
    $("fileInput").onchange = function (e) {
      openFiles(e.target.files);
      e.target.value = "";
    };
    $("docSelect").onchange = function (e) { if (e.target.value) openDoc(e.target.value, 1); };
    $("btnForget").onclick = function () {
      var key = $("docSelect").value;
      if (!key) return;
      // Office add-ins can't show confirm() pop-ups, so ask for a second click instead.
      var btn = $("btnForget");
      if (btn.dataset.armed !== key) {
        btn.dataset.armed = key;
        var linked = state.snips.filter(function (s) { return s.docKey === key; }).length;
        setStatus("Click ✕ again to remove this PDF from the list." +
          (linked ? " Its " + linked + " linked cell(s) keep their links; open the same PDF again to restore them." : ""));
        setTimeout(function () { if (btn.dataset.armed === key) btn.dataset.armed = ""; }, 4000);
        return;
      }
      btn.dataset.armed = "";
      DocStore.del(key).then(function () {
        setStatus("Removed from the list.");
        if (state.doc && state.doc.key === key) {
          state.doc.pdf.destroy();
          state.doc = null;
          $("pageWrap").hidden = true;
          $("empty").hidden = false;
          $("pageLabel").textContent = "–";
        }
        return refreshDocList();
      });
    };
    Array.from(document.querySelectorAll(".seg button")).forEach(function (b) {
      b.onclick = function () {
        state.mode = b.dataset.mode;
        Array.from(document.querySelectorAll(".seg button")).forEach(function (x) { x.classList.toggle("on", x === b); });
        setStatus({ number: "Number: snips one amount.", sum: "Sum: adds every amount in the box.", text: "Text: copies the words." }[state.mode]);
      };
    });
    $("btnPrev").onclick = function () { goToPage(state.pageNum - 1); };
    $("btnNext").onclick = function () { goToPage(state.pageNum + 1); };
    $("btnZoomIn").onclick = function () { state.zoom = Math.min(4, state.zoom * 1.25); renderPage(); };
    $("btnZoomOut").onclick = function () { state.zoom = Math.max(0.5, state.zoom / 1.25); renderPage(); };
    // Remember settings between sessions.
    ["optMoveDown", "optMark", "optOcr", "optShowBoxes"].forEach(function (id) {
      var el = $(id);
      try {
        var saved = localStorage.getItem("ggdsnip." + id);
        if (saved !== null) el.checked = saved === "1";
      } catch (e) {}
      el.addEventListener("change", function () {
        try { localStorage.setItem("ggdsnip." + id, el.checked ? "1" : "0"); } catch (e) {}
        if (id === "optShowBoxes") drawBoxes();
      });
    });
    var rt = null;
    window.addEventListener("resize", function () {
      clearTimeout(rt);
      rt = setTimeout(function () { if (state.doc) renderPage(); }, 200);
    });
    // Drop PDFs straight onto the pane.
    var viewer = $("viewer");
    viewer.addEventListener("dragover", function (e) { e.preventDefault(); });
    viewer.addEventListener("drop", function (e) {
      e.preventDefault();
      if (e.dataTransfer && e.dataTransfer.files.length) openFiles(e.dataTransfer.files);
    });
    setupDrag();
  }

  function start(inExcel) {
    XL = inExcel ? ExcelAdapter : MockAdapter;
    window.__ggdsnip = { state: state, mock: MockAdapter, core: S };
    wireUi();
    if (!window.pdfjsLib) {
      setStatus("Couldn't load the PDF viewer. Check your internet connection and reopen the pane.", "err");
      return;
    }
    if (inExcel && !Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      setStatus("This Excel version is too old for GGD Snip. Update Microsoft 365 and try again.", "err");
      return;
    }
    refreshDocList();
    refreshSnips().catch(function () {});
    if (inExcel) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    } else {
      setStatus("Preview mode (not inside Excel). Snips go to a pretend cell starting at B2.");
    }
  }

  var started = false;
  function startOnce(inExcel) { if (!started) { started = true; start(inExcel); } }
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) { startOnce(info && info.host === Office.HostType.Excel); });
  } else {
    document.addEventListener("DOMContentLoaded", function () { startOnce(false); });
    if (document.readyState !== "loading") startOnce(false);
  }
})();
