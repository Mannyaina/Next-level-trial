/* Next Level · read a flyer, PDF, slides or document and suggest the details for a post.
   Everything runs on the phone: nothing is sent anywhere. Loaded only when a leader adds a file. */
(function () {
  "use strict";
  var BASE = (function () { var s = document.currentScript && document.currentScript.src; return s ? s.replace(/[^\/]*$/, "") : "vendor/nlread/"; })();
  var loaded = {};
  function load(src) {
    if (loaded[src]) return loaded[src];
    loaded[src] = new Promise(function (res, rej) { var s = document.createElement("script"); s.src = BASE + src; s.onload = res; s.onerror = function () { delete loaded[src]; rej(new Error("load " + src)); }; document.head.appendChild(s); });
    return loaded[src];
  }

  /* ---------- 1. get the words out of the file, with how big each line is ---------- */
  async function linesFromPdf(file) {
    await load("pdf.min.js");
    var lib = window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc = BASE + "pdf.worker.min.js";
    var pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    var out = [], pages = Math.min(pdf.numPages, 12);
    for (var p = 1; p <= pages; p++) {
      var page = await pdf.getPage(p), tc = await page.getTextContent(), rows = {};
      tc.items.forEach(function (it) {
        if (!it.str || !it.str.trim()) return;
        var y = Math.round(it.transform[5]), h = Math.abs(it.transform[3]) || it.height || 10, k = p + ":" + Math.round(y / 3);
        (rows[k] = rows[k] || { y: y, size: 0, parts: [] }).parts.push({ x: it.transform[4], s: it.str }); rows[k].size = Math.max(rows[k].size, h);
      });
      Object.keys(rows).map(function (k) { return rows[k]; }).sort(function (a, b) { return b.y - a.y; }).forEach(function (r) {
        out.push({ text: r.parts.sort(function (a, b) { return a.x - b.x; }).map(function (x) { return x.s; }).join(" ").replace(/\s+/g, " ").trim(), size: r.size, page: p });
      });
    }
    if (out.join("").length < 20 && pdf.numPages) { // a scanned PDF: read the first page as a picture
      var pg = await pdf.getPage(1), vp = pg.getViewport({ scale: 2 }), c = document.createElement("canvas"); c.width = vp.width; c.height = vp.height;
      await pg.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
      return linesFromImage(c);
    }
    return out;
  }
  async function linesFromOffice(file) {
    await load("jszip.min.js");
    var zip = await window.JSZip.loadAsync(await file.arrayBuffer()), out = [];
    var slides = Object.keys(zip.files).filter(function (n) { return /^ppt\/slides\/slide\d+\.xml$/.test(n); })
      .sort(function (a, b) { return parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10); });
    if (slides.length) {
      for (var i = 0; i < slides.length && i < 40; i++) {
        var doc = new DOMParser().parseFromString(await zip.file(slides[i]).async("string"), "application/xml");
        Array.prototype.forEach.call(doc.getElementsByTagName("p:sp"), function (sp) {
          var ph = sp.getElementsByTagName("p:ph")[0], kind = ph ? (ph.getAttribute("type") || "body") : "";
          Array.prototype.forEach.call(sp.getElementsByTagName("a:p"), function (para) {
            var t = Array.prototype.map.call(para.getElementsByTagName("a:t"), function (x) { return x.textContent; }).join("").replace(/\s+/g, " ").trim();
            if (!t) return; var sz = 0; Array.prototype.forEach.call(para.getElementsByTagName("a:rPr"), function (r) { sz = Math.max(sz, +(r.getAttribute("sz") || 0) / 100); });
            out.push({ text: t, size: sz || (/title|ctrTitle/.test(kind) ? 40 : 20), page: i + 1, title: /title|ctrTitle/.test(kind) });
          });
        });
      }
      return out;
    }
    var w = zip.file("word/document.xml");
    if (w) {
      var d = new DOMParser().parseFromString(await w.async("string"), "application/xml");
      Array.prototype.forEach.call(d.getElementsByTagName("w:p"), function (para) {
        var t = Array.prototype.map.call(para.getElementsByTagName("w:t"), function (x) { return x.textContent; }).join("").trim(); if (!t) return;
        var sz = 0; Array.prototype.forEach.call(para.getElementsByTagName("w:sz"), function (r) { sz = Math.max(sz, +(r.getAttribute("w:val") || 0) / 2); });
        var st = para.getElementsByTagName("w:pStyle")[0], isT = st && /Title|Heading1/.test(st.getAttribute("w:val") || "");
        out.push({ text: t, size: sz || (isT ? 28 : 12), page: 1, title: isT });
      });
    }
    return out;
  }
  async function toCanvas(src, scale, maxW) {
    var bmp = src instanceof HTMLCanvasElement ? src : await createImageBitmap(src);
    var w = bmp.width, h = bmp.height, k = Math.min(scale, maxW / w), c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
    var g = c.getContext("2d"); g.imageSmoothingQuality = "high"; g.drawImage(bmp, 0, 0, c.width, c.height); return { c: c, k: k };
  }
  async function linesFromImage(img) {
    await load("tesseract.min.js");
    var worker = await window.Tesseract.createWorker("eng", 1, { workerPath: BASE + "worker.min.js", corePath: BASE + "core", langPath: BASE + "lang", gzip: true, logger: function () {} });
    try {
      // read twice: at normal size for small print, and shrunk so very large headline letters are read too
      var big = await toCanvas(img, 1, 1800), small = await toCanvas(img, 0.38, 700), out = [], seen = {};
      var norm = function (t) { return t.toLowerCase().replace(/[^a-z0-9]/g, ""); };
      for (var pass of [big, small]) {
        var r = await worker.recognize(pass.c);
        (r.data.lines || []).forEach(function (l) {
          var t = (l.text || "").replace(/\s+/g, " ").trim(), n = norm(t);
          if (!t || n.length < 2 || l.confidence < 45 || seen[n]) return;
          if (Object.keys(seen).some(function (x) { return x.length > 4 && (x.indexOf(n) >= 0 || n.indexOf(x) >= 0); })) return;
          seen[n] = 1; out.push({ text: t, size: ((l.bbox.y1 - l.bbox.y0) || 10) / pass.k, y: l.bbox.y0 / pass.k, page: 1 });
        });
      }
      return out.sort(function (a, b) { return a.y - b.y; });
    } finally { worker.terminate(); }
  }
  async function lines(file) {
    var n = (file.name || "").toLowerCase(), t = file.type || "";
    if (/pdf/.test(t) || /\.pdf$/.test(n)) return linesFromPdf(file);
    if (/\.(pptx|docx)$/.test(n) || /presentationml|wordprocessingml/.test(t)) return linesFromOffice(file);
    if (/^image\//.test(t) || /\.(jpe?g|png|webp|heic)$/.test(n)) return linesFromImage(file);
    if (/^text\//.test(t)) return (await file.text()).split(/\n/).map(function (s) { return { text: s.trim(), size: 12, page: 1 }; }).filter(function (x) { return x.text; });
    return [];
  }

  /* ---------- 2. find dates, times, places, scripture and the title ---------- */
  var MON = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  var DOW = "(?:mon|tue|tues|wed|weds|thu|thur|thurs|fri|sat|sun)(?:day|nesday|rsday|urday)?";
  var MONR = "(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)";
  var pad = function (n) { return String(n).padStart(2, "0"); };
  function iso(y, m, d) { return y + "-" + pad(m + 1) + "-" + pad(d); }
  function pickYear(m, d, y) {
    if (y) { y = +y; if (y < 100) y += 2000; return y; }
    var now = new Date(), t = new Date(now.getFullYear(), m, d);
    return t < new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1) ? now.getFullYear() + 1 : now.getFullYear();
  }
  function findDates(text) {
    var out = [], m, re;
    re = new RegExp("(?:" + DOW + ",?\\s+)?(\\d{1,2})(?:st|nd|rd|th)?\\s*(?:of\\s+)?" + MONR + "\\.?,?(?:\\s+(\\d{4}))?", "gi");
    while ((m = re.exec(text))) { var mo = MON.indexOf(m[2].slice(0, 3).toLowerCase()), d = +m[1]; if (mo >= 0 && d >= 1 && d <= 31) out.push({ i: m.index, v: iso(pickYear(mo, d, m[3]), mo, d), raw: m[0] }); }
    re = new RegExp("(?:" + DOW + ",?\\s+)?" + MONR + "\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?(?:\\s+(\\d{4}))?", "gi");
    while ((m = re.exec(text))) { var mo2 = MON.indexOf(m[1].slice(0, 3).toLowerCase()), d2 = +m[2]; if (mo2 >= 0 && d2 >= 1 && d2 <= 31) out.push({ i: m.index, v: iso(pickYear(mo2, d2, m[3]), mo2, d2), raw: m[0] }); }
    re = /\b(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})\b/g; // UK order: day/month/year
    while ((m = re.exec(text))) { var dd = +m[1], mm = +m[2] - 1; if (dd >= 1 && dd <= 31 && mm >= 0 && mm < 12) out.push({ i: m.index, v: iso(pickYear(mm, dd, m[3]), mm, dd), raw: m[0] }); }
    return out.sort(function (a, b) { return a.i - b.i; });
  }
  function to24(h, mi, ap) { h = +h; mi = +(mi || 0); if (ap) { ap = ap.toLowerCase(); if (ap[0] === "p" && h < 12) h += 12; if (ap[0] === "a" && h === 12) h = 0; } return pad(h) + ":" + pad(mi); }
  function findTimes(text) {
    var m, re = /\b(\d{1,2})(?:[:.](\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)?\s*(?:-|–|—|to|till|until)\s*(\d{1,2})(?:[:.](\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)?/i;
    if ((m = re.exec(text)) && (m[3] || m[6] || m[2] || m[5])) { var ap2 = m[6] || m[3], ap1 = m[3] || (m[6] && +m[1] <= +m[4] ? m[6] : null); return { start: to24(m[1], m[2], ap1), end: to24(m[4], m[5], ap2) }; }
    re = /\b(\d{1,2})(?:[:.](\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)/i; if ((m = re.exec(text))) return { start: to24(m[1], m[2], m[3]) };
    re = /\b([01]?\d|2[0-3])[:.]([0-5]\d)\b/; if ((m = re.exec(text))) return { start: to24(m[1], m[2]) };
    if (/\bnoon\b/i.test(text)) return { start: "12:00" };
    return null;
  }
  var BOOKS = "(?:[1-3]\\s?)?(?:Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|Samuel|Kings|Chronicles|Ezra|Nehemiah|Esther|Job|Psalms?|Proverbs|Ecclesiastes|Song of Solomon|Song of Songs|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi|Matthew|Mark|Luke|John|Acts|Romans|Corinthians|Galatians|Ephesians|Philippians|Colossians|Thessalonians|Timothy|Titus|Philemon|Hebrews|James|Peter|Jude|Revelation)";
  function findScripture(text) { var m = new RegExp("\\b" + BOOKS + "\\s+\\d{1,3}(?::\\d{1,3}(?:\\s*[-–]\\s*\\d{1,3})?)?", "i").exec(text); return m ? m[0].replace(/\s+/g, " ") : ""; }
  var PLACE = /\b(venue|location|where|address|at\s+the|@|hall|centre|center|church|house of praise|road|rd\b|street|st\b|avenue|lane|park|hotel|restaurant|cafe|room|auditorium|online|zoom)\b/i;
  var NOISE = /^(free|all welcome|everyone welcome|join us|you'?re invited|save the date|rsvp|book now|register|tickets?|www\.|https?:|@\w+|#\w+)/i;

  function suggest(ls, col) {
    ls = ls.filter(function (l) { return l.text && l.text.length > 1; });
    if (!ls.length) return null;
    var all = ls.map(function (l) { return l.text; }).join("\n"), flat = all.replace(/\n/g, " ");
    var dates = findDates(flat), times = findTimes(flat), scr = findScripture(flat);
    var isDetail = function (t) { return findDates(t).length || findTimes(t) || NOISE.test(t) || /^\d+$/.test(t); };
    // title: the biggest words on the first page that aren't a date, time or "join us"
    var first = ls.filter(function (l) { return (l.page || 1) === 1; }), cand = first.filter(function (l) { return !isDetail(l.text) && l.text.length >= 3 && l.text.length <= 70; });
    var clean = cand.filter(function (l) { return !/:/.test(l.text) && !PLACE.test(l.text) && !/presents$/i.test(l.text); }); if (clean.length) cand = clean;
    var flagged = cand.filter(function (l) { return l.title; });
    var biggest = (flagged.length ? flagged : cand).slice().sort(function (a, b) { return b.size - a.size; });
    var title = "";
    if (biggest.length) {
      var top = biggest[0]; title = top.text;
      var idx = first.indexOf(top), nx = first[idx + 1]; // big titles often wrap onto two lines
      if (nx && !isDetail(nx.text) && Math.abs(nx.size - top.size) / top.size < 0.15 && (title + " " + nx.text).length <= 60) title += " " + nx.text;
    }
    title = title.replace(/\s+/g, " ").replace(/^[^\w"']+|[^\w!?)"']+$/g, "");
    if (title === title.toUpperCase() && /[A-Z]/.test(title)) title = title.toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }).replace(/\b(And|Of|The|In|At|On|For|With|To|A)\b/g, function (w, x, i) { return i ? w.toLowerCase() : w; });
    var where = "";
    ls.some(function (l) { if (l.text !== title && PLACE.test(l.text) && !findTimes(l.text) && l.text.length < 90) { where = l.text.replace(/^(venue|location|where|address)\s*[:\-–]\s*/i, "").replace(/^@\s*/, ""); return true; } });
    var used = [title, where].filter(Boolean), placeRaw = ls.filter(function (l) { return where && l.text.indexOf(where) >= 0; }).map(function (l) { return l.text; });
    used = used.concat(placeRaw, title.split(' ').length > 1 ? ls.filter(function (l) { return title.toLowerCase().indexOf(l.text.toLowerCase()) >= 0; }).map(function (l) { return l.text; }) : []);
    var rest = ls.map(function (l) { return l.text; }).filter(function (t) { return used.indexOf(t) < 0 && !(findDates(t).length && t.length < 40) && !(findTimes(t) && t.length < 30); });
    var out = {};
    if (title) { out.title = title; out.name = title; }
    if (dates.length) out.date = dates[0].v;
    if (times) { out.time = times.start; if (times.end) out.end = times.end; }
    if (where) out.where = where;
    if (scr) out.scripture = scr;
    if (/house of praise|\bhop\b/i.test(flat) && /service|sunday|main church/i.test(flat)) out.church = true;
    if (col === "notes") {
      var sub = first.filter(function (l) { return !l.title && l.text !== title && findDates(l.text).length; })[0];
      if (sub) { var ser = sub.text.replace(findDates(sub.text)[0].raw, '').replace(/[\s·\-–|,]+$/,'').replace(/^[\s·\-–|,]+/,'').trim(); if (ser.length > 2 && ser.length < 50) out.series = ser; }
      // slides become the outline: each slide's title as a heading, its points underneath
      var outline = [], lastPage = 0;
      ls.forEach(function (l) { if (l.page === 1 && (l.text === title || (sub && l.text === sub.text))) return; if (l.page !== lastPage) { lastPage = l.page; if (l.title) { outline.push("## " + l.text); return; } } outline.push(new RegExp("^" + BOOKS + "\\s+\\d", "i").test(l.text) ? "> " + l.text : l.text); });
      if (outline.length) out.body = outline.join("\n").slice(0, 6000);
    } else if (rest.length) {
      out.desc = rest.join("\n").slice(0, 1200); out.body = out.desc; out.text = (title ? title + "\n" : "") + out.desc;
    }
    return out;
  }

  window.NLREAD = {
    read: async function (file, col) { var ls = await lines(file); return suggest(ls, col); },
    _suggest: suggest, _lines: lines
  };
})();
