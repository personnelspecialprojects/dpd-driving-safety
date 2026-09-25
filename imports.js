/* ============================================================
   imports.js — Excel import (parsers + write engine + screen)
   Loaded after data.js. Replaces the Power Automate + Office Script
   pipeline: files are parsed in the browser (SheetJS) and written
   straight to SharePoint.

   Parsers use dynamic header detection (find the row containing a
   signature column) rather than hardcoded row numbers, so leading
   title/blank rows don't break them. Roster and Physicals parsers
   were verified against real sample files; Courses and Accidents are
   ports of the Office Scripts already validated against real exports.

   SCHEMA NOTE — Physicals import writes ExpirationDate (and Result).
   Add these columns to DrivingSafety_Physicals before using it:
     • ExpirationDate  (Date and Time, Date Only)   — required
     • Result          (Single line of text)         — optional
   ============================================================ */
(function () {
  const el = DS.el;

  /* ---- lazy-load SheetJS: local repo copy first (best on a locked-down
         network), then public CDNs as fallback ---- */
  let xlsxPromise = null;
  function ensureXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxPromise) return xlsxPromise;
    xlsxPromise = new Promise((resolve, reject) => {
      const sources = [
        "./xlsx.full.min.js",
        "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
        "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
      ];
      (function tryAt(i) {
        if (i >= sources.length) { reject(new Error("Couldn't load the spreadsheet library. If the network blocks CDNs, add xlsx.full.min.js to the repo.")); return; }
        const s = document.createElement("script");
        s.src = sources[i];
        s.onload = () => resolve(window.XLSX);
        s.onerror = () => tryAt(i + 1);
        document.head.appendChild(s);
      })(0);
    });
    return xlsxPromise;
  }

  /* ---- parse helpers ---- */
  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }); }
  function locate(rows, signature) {
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const idx = (rows[i] || []).findIndex(c => String(c).trim() === signature);
      if (idx >= 0) return { headerIdx: i, header: rows[i] };
    }
    return null;
  }
  function colFinder(header) { return name => header.findIndex(c => String(c).trim() === name); }
  function extractPoints(text) {
    if (!text) return 0;
    const m = String(text).match(/(\d+)\s*(pts?|points?)/i);
    return m ? Number(m[1]) : 0;
  }
  function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
  // Spreadsheet row number (as shown in Excel) for index i of rowsOf(ws)
  function rowNo(ws, i) { try { return XLSX.utils.decode_range(ws["!ref"]).s.r + i + 1; } catch (_) { return i + 1; } }
  // Fields starting with "__" are working data for matching/reporting — never sent to SharePoint
  function stripMeta(rec) { const o = {}; Object.keys(rec).forEach(k => { if (!k.startsWith("__")) o[k] = rec[k]; }); return o; }

  /* ============================================================
     PARSERS — each returns { records:[...], warnings:[...] }
     ============================================================ */

  function parseCourses(wb, requiredTitles) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    const loc = locate(r, "Course Title");
    if (!loc) throw new Error("Couldn't find the 'Course Title' header — is this the PoliceOne/Lexipol courses export?");
    const ci = colFinder(loc.header);
    const iBadge = ci("Badge/ID #"), iTitle = ci("Course Title"), iStatus = ci("Completion Status"),
          iDate = ci("Date Completed"), iScore = ci("Score"), iAttempts = ci("Attempts"), iName = ci("Full Name");
    const req = requiredTitles || [];
    const records = []; let skippedStatus = 0, skippedTitle = 0;
    for (let i = loc.headerIdx + 1; i < r.length; i++) {
      const row = r[i]; if (!row || !row[iBadge]) continue;
      const status = String(row[iStatus] || "").trim();
      const title = String(row[iTitle] || "").trim();
      if (status !== "Passed") { skippedStatus++; continue; }
      if (req.length && !req.includes(title)) { skippedTitle++; continue; }
      records.push({
        EmployeeId: String(row[iBadge]).trim(),
        CourseTitle: title,
        CompletionStatus: status,
        DateCompleted: DS.isoDate(row[iDate]),
        Score: num(row[iScore]),
        Attempts: num(row[iAttempts]),
        Source: "Bulk Upload",
        __name: iName >= 0 ? String(row[iName] || "").trim() : "",
        __row: rowNo(ws, i),
      });
    }
    const warnings = [];
    if (skippedStatus) warnings.push(skippedStatus + " row(s) skipped (not marked Passed).");
    if (skippedTitle) warnings.push(skippedTitle + " row(s) skipped (not a required course).");
    return { records, warnings };
  }

  function parseAccidents(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    const loc = locate(r, "Incident Number");
    if (!loc) throw new Error("Couldn't find the 'Incident Number' header — is this the Origami Risk equipment-incidents export?");
    const ci = colFinder(loc.header);
    const iNum = ci("Incident Number"), iLoss = ci("Loss Date"), iEmp = ci("Employee Number"),
          iIRC = ci("IRC Decision"), iFinal = ci("IRC/IAB Final Decision"),
          iStreet = ci("Accident Street1"), iVeh = ci("Vehicle"), iMake = ci("Vehicle Make"), iModel = ci("Vehicle Model");
    const iName = [ci("Employee Name"), ci("Employee"), ci("Driver Name"), ci("Driver")].find(x => x >= 0);
    const records = [];
    for (let i = loc.headerIdx + 1; i < r.length; i++) {
      const row = r[i]; if (!row || !row[iNum]) continue;
      const initial = String(row[iIRC] || "");
      const final = String(row[iFinal] || "");
      records.push({
        IncidentNumber: String(row[iNum]).trim(),
        EmployeeId: String(row[iEmp] || "").trim(),
        AccidentDate: DS.isoDate(row[iLoss]),
        InitialDecision: initial,
        InitialPoints: extractPoints(initial),
        FinalDecision: final,
        FinalPoints: extractPoints(final),
        Vehicle: String(row[iVeh] || ""),
        VehicleMakeModel: [row[iMake], row[iModel]].filter(Boolean).join(" ").trim(),
        Location: String(row[iStreet] || ""),
        Source: "Bulk Upload",
        CountsAgainstStreak: "Auto",
        __name: iName != null ? String(row[iName] || "").trim() : "",
        __row: rowNo(ws, i),
      });
    }
    return { records, warnings: [] };
  }

  /* ---- flexible column matching for Roster. The real SQL report headers are
     Emp#, FirstName, LastName, Rank, WorkingOrg, Workgroup, Supervisor,
     AdjSvcDate — column ORDER doesn't matter (matched by name, not position),
     and common spacing/wording variants are accepted too (e.g. "Emp #",
     "Employee Number", "Working Org", "1st Line Supervisor", "Hire Date"). */
  function normHeader(s) { return String(s).trim().toLowerCase(); }
  function findFlexCol(header, pattern) {
    for (let i = 0; i < header.length; i++) if (pattern.test(normHeader(header[i]))) return i;
    return -1;
  }
  const ROSTER_PATTERNS = {
    empId: /^emp(loyee)?\s*(#|number|num|no\.?|id)?$/,
    firstName: /^f(irst)?\s*name$/,
    lastName: /^l(ast)?\s*name$/,
    rank: /^rank$/,
    workingOrg: /^(working\s*)?org(anization)?$/,
    workgroup: /^(working\s*)?work\s*group$/,
    supervisor: /^(1st\s*line\s*)?supervisor$/,
    // AdjSvcDate (Adjusted Service Date) is what the real export provides in
    // place of a true hire date — used as-is for course grace-period and
    // award-clock timing. Note this can differ from actual hire date where
    // service credit (e.g. prior military time) shifts it earlier.
    hireDate: /^(adj(usted)?\s*(svc|service)\s*date|hire\s*date)$/,
    badge: /^badge\s*(#|number|num|no\.?)?$/,
  };

  function parseRoster(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    let headerIdx = -1, header = null;
    for (let i = 0; i < Math.min(r.length, 40); i++) {
      if (findFlexCol(r[i] || [], ROSTER_PATTERNS.empId) >= 0) { headerIdx = i; header = r[i]; break; }
    }
    if (headerIdx < 0) throw new Error("Couldn't find an employee ID column (Emp#, Emp #, Employee Number, ...) — is this the roster export?");
    const iEmp = findFlexCol(header, ROSTER_PATTERNS.empId);
    const iFirst = findFlexCol(header, ROSTER_PATTERNS.firstName);
    const iLast = findFlexCol(header, ROSTER_PATTERNS.lastName);
    const iRank = findFlexCol(header, ROSTER_PATTERNS.rank);
    const iOrg = findFlexCol(header, ROSTER_PATTERNS.workingOrg);
    const iWg = findFlexCol(header, ROSTER_PATTERNS.workgroup);
    const iSup = findFlexCol(header, ROSTER_PATTERNS.supervisor);
    const iHire = findFlexCol(header, ROSTER_PATTERNS.hireDate);
    const iBadge = findFlexCol(header, ROSTER_PATTERNS.badge);
    const records = []; let noHire = 0; const noEmp = [];
    for (let i = headerIdx + 1; i < r.length; i++) {
      const row = r[i]; if (!row) continue;
      if (!row[iEmp]) {                                   // a named row with no Employee # can't be imported
        const nm = [iFirst >= 0 ? row[iFirst] : "", iLast >= 0 ? row[iLast] : ""].join(" ").trim();
        if (nm) noEmp.push({ __row: rowNo(ws, i), __name: nm, EmployeeId: "", Badge: iBadge >= 0 ? String(row[iBadge] || "").trim() : "" });
        continue;
      }
      const first = iFirst >= 0 ? String(row[iFirst] || "").trim() : "";
      const last = iLast >= 0 ? String(row[iLast] || "").trim() : "";
      const hire = iHire >= 0 && row[iHire] ? DS.isoDate(row[iHire]) : null;
      if (!hire) noHire++;
      records.push({
        Title: (first + " " + last).trim(),
        LastName: last,
        EmployeeId: String(row[iEmp]).trim(),
        Division: iOrg >= 0 ? String(row[iOrg] || "") : "",
        Assignment: iWg >= 0 ? String(row[iWg] || "") : "",
        Supervisor: iSup >= 0 ? String(row[iSup] || "") : "",
        HireDate: hire,
        Rank: iRank >= 0 ? String(row[iRank] || "") : "",
        __row: rowNo(ws, i),
      });
      if (iBadge >= 0) records[records.length - 1].Badge = String(row[iBadge] == null ? "" : row[iBadge]).trim();
    }
    const warnings = [];
    if (iBadge < 0) warnings.push("No 'Badge' column found — badges on the roster won't be updated. Uploads that use badge numbers need them.");
    if (noEmp.length) warnings.push(noEmp.length + " row(s) have a name but no employee number and were skipped — listed in Upload results.");
    if (iHire < 0) warnings.push("No hire/service date column found — course due-dates for new hires need it.");
    else if (noHire) warnings.push(noHire + " employee(s) have no hire/service date.");
    return { records, warnings, noEmp, hasBadge: iBadge >= 0 };
  }

  function parsePhysicals(wb) {
    const ws = wb.Sheets["Data"] || wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    const loc = locate(r, "Employee Number");
    if (!loc) throw new Error("Couldn't find the 'Employee Number' header — is this the driver-physicals export?");
    const ci = colFinder(loc.header);
    const iDate = ci("Date Tested"), iEmp = ci("Employee Number"), iRes = ci("Driver Physical Test Results"),
          iExp = ci("Phy Exp Date"), iClinic = ci("Clinic Location"), iName = ci("Employee");
    const records = []; let noExp = 0, fails = 0;
    for (let i = loc.headerIdx + 1; i < r.length; i++) {
      const row = r[i]; if (!row || !row[iEmp]) continue;
      const exp = DS.isoDate(row[iExp]);
      const result = String(row[iRes] || "").trim();
      if (!exp) noExp++;
      if (result.toLowerCase() === "fail") fails++;
      records.push({
        EmployeeId: String(row[iEmp]).trim(),
        PhysicalDate: DS.isoDate(row[iDate]),
        ExpirationDate: exp,
        Result: result,
        Provider: String(row[iClinic] || ""),
        Source: "Bulk Upload",
        __name: iName >= 0 ? String(row[iName] || "").trim() : "",
        __row: rowNo(ws, i),
      });
    }
    const warnings = [];
    if (fails) warnings.push(fails + " row(s) are marked Fail — imported as-is; review whether they should count.");
    if (noExp) warnings.push(noExp + " row(s) have no expiration date — those won't drive a due date.");
    return { records, warnings };
  }

  /* ---- what each upload's real column headers look like, shown on hover ---- */
  const FORMAT_SPECS = {
    courses: {
      source: "PoliceOne / Lexipol course completion export",
      columns: ["Badge/ID #", "Course Title", "Completion Status", "Date Completed", "Score", "Attempts"],
      sample: ["123456", "Defensive Driving Basics", "Passed", "3/12/2026", "94", "1"],
      note: 'Only rows marked "Passed" for a required course are imported.',
    },
    accidents: {
      source: "Origami Risk — DPD Equipment Incidents export",
      columns: ["Incident Number", "Loss Date", "Employee Number", "IRC Decision", "IRC/IAB Final Decision", "Accident Street1", "Vehicle", "Vehicle Make", "Vehicle Model"],
      sample: ["INC-48213", "3/2/2026", "123456", "2 pts", "2 pts", "Main St", "1", "Ford", "Explorer"],
      note: "Deduped by Incident Number — safe to re-upload the same export.",
    },
    roster: {
      source: "DPD personnel SQL report export",
      columns: ["Emp#", "Badge", "FirstName", "LastName", "Rank", "WorkingOrg", "Workgroup", "Supervisor", "AdjSvcDate"],
      sample: ["123456", "8812", "Jordan", "Smith", "Police Officer", "1498", "Patrol", "Garcia, M.", "6/1/2022"],
      note: 'Column order doesn\'t matter, and common variants (e.g. "Emp #", "Employee Number") are accepted. Badge may be blank or start with R/T. Driver designation is never touched by this import.',
    },
    physicals: {
      source: 'Driver physicals report (sheet "Data")',
      columns: ["Date Tested", "Employee Number", "Driver Physical Test Results", "Phy Exp Date", "Clinic Location"],
      sample: ["7/1/2026", "123456", "Pass", "6/30/2028", "Wheatland"],
      note: "Needs ExpirationDate and Result columns added to DrivingSafety_Physicals.",
    },
  };

  /* ============================================================
     IMPORT TYPES — parser + write strategy per type
     ============================================================ */
  const TYPES = {
    courses: {
      label: "Courses", list: () => DS.LISTS.courses, mode: "append-dedup",
      parse: (wb, ctx) => parseCourses(wb, ctx.requiredTitles),
      cols: [["Employee", "EmployeeId"], ["Course", "CourseTitle"], ["Completed", "DateCompleted"], ["Score", "Score"]],
      describe: r => (r.CourseTitle || "?") + " \u00b7 " + DS.fmtDate(r.DateCompleted),
      dedupKey: r => [DS.util.empKey(r.EmployeeId), String(r.CourseTitle || "").trim(), DS.isoDate(r.DateCompleted)].join("|"),
      existingKeys: async () => {
        const rows = await DS.spGet(DS.LISTS.courses, { select: ["EmployeeId", "CourseTitle", "DateCompleted"] });
        return new Set(rows.map(x => [DS.util.empKey(x.EmployeeId), String(x.CourseTitle || "").trim(), DS.isoDate(x.DateCompleted)].join("|")));
      },
    },
    accidents: {
      label: "Accidents", list: () => DS.LISTS.accidents, mode: "append-dedup",
      parse: (wb) => parseAccidents(wb),
      cols: [["Incident #", "IncidentNumber"], ["Employee", "EmployeeId"], ["Date", "AccidentDate"], ["Final pts", "FinalPoints"]],
      describe: r => "Incident " + (r.IncidentNumber || "?") + " \u00b7 " + DS.fmtDate(r.AccidentDate),
      dedupKey: r => String(r.IncidentNumber).trim(),
      existingKeys: async () => {
        const rows = await DS.spGet(DS.LISTS.accidents, { select: ["IncidentNumber"] });
        return new Set(rows.map(x => String(x.IncidentNumber || "").trim()));
      },
    },
    roster: {
      label: "Roster", list: () => DS.LISTS.roster, mode: "upsert",
      parse: (wb) => parseRoster(wb),
      describe: r => r.Title || "",
      cols: [["Name", "Title"], ["ID", "EmployeeId"], ["Badge", "Badge"], ["Division", "Division"], ["Rank", "Rank"]],
    },
    physicals: {
      label: "Physicals", list: () => DS.LISTS.physicals, mode: "append-dedup",
      parse: (wb) => parsePhysicals(wb),
      cols: [["Employee", "EmployeeId"], ["Tested", "PhysicalDate"], ["Expires", "ExpirationDate"], ["Result", "Result"]],
      // one exam per employee per test date — the same exam from any source isn't logged twice
      describe: r => "Tested " + DS.fmtDate(r.PhysicalDate) + (r.Result ? " \u00b7 " + r.Result : ""),
      dedupKey: r => [DS.util.empKey(r.EmployeeId), DS.isoDate(r.PhysicalDate)].join("|"),
      existingKeys: async () => {
        const rows = await DS.spGet(DS.LISTS.physicals, { select: ["EmployeeId", "PhysicalDate"] });
        return new Set(rows.filter(x => x.PhysicalDate).map(x => [DS.util.empKey(x.EmployeeId), DS.isoDate(x.PhysicalDate)].join("|")));
      },
    },
  };

  /* ---- concurrency-limited runner with retry-with-backoff for transient failures.
     A "transient" failure (429/503/504, or the request never reaching the server
     at all) is retried with exponential backoff — up to 5 attempts, honoring a
     Retry-After header when SharePoint sends one. A permanent failure (400, 403,
     a genuinely malformed record) is NOT retried — retrying it would just fail
     the same way every time. Large batches also start at lower concurrency,
     since a big burst of simultaneous requests is what triggers throttling in
     the first place. ---- */
  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
  async function runBatched(items, worker, onProgress, concurrency, jobLabel) {
    concurrency = concurrency || 4;
    const ownsJob = !DS.job.active;                     // outermost batch owns the indicator
    if (ownsJob) DS.job.start(jobLabel, items.length);
    const report = (d, total, label) => { DS.job.update(d, total, label); onProgress(d, total, label); };
    if (items.length > 500) concurrency = Math.min(concurrency, 2);   // ease off for big batches
    const MAX_ATTEMPTS = 5;
    let done = 0, retried = 0;
    const errors = [];
    const queue = items.map(item => ({ item, attempt: 0 }));

    async function lane() {
      while (queue.length) {
        const entry = queue.shift();
        try {
          await worker(entry.item);
        } catch (e) {
          if (e && e.transient && entry.attempt < MAX_ATTEMPTS - 1) {
            entry.attempt++;
            retried++;
            const wait = e.retryAfterMs || Math.min(20000, 400 * Math.pow(2, entry.attempt));
            report(done, items.length, "Waiting to retry after a busy response (" + retried + " so far)…");
            await sleep(wait);
            queue.push(entry);   // back of the line, not counted as done yet
            continue;
          }
          if (e && typeof e === "object") e.item = entry.item;
          errors.push(e);
        }
        done++; report(done, items.length);
      }
    }
    try { await Promise.all(Array.from({ length: Math.min(concurrency, items.length || 1) }, lane)); }
    finally { if (ownsJob) DS.job.end(); }
    return errors;
  }

  /* ============================================================
     SCREEN
     ============================================================ */
  const state = { type: "courses", parsed: null, existing: null };

  async function renderImports(container) {
    const cache = await DS.data.load();               // for requiredTitles + roster count
    state.ctx = { requiredTitles: cache.idx.requiredTitles };
    container.innerHTML = "";

    // type selector
    const seg = el("div", { class: "seg" });
    Object.keys(TYPES).forEach(key => {
      const b = el("button", { text: TYPES[key].label });
      if (key === state.type) b.classList.add("active");
      b.addEventListener("click", () => { state.type = key; state.parsed = null; renderImports(container); });
      seg.appendChild(b);
    });
    container.appendChild(seg);

    const t = TYPES[state.type];

    // mode note
    const note = el("div", { class: "import-note" });
    if (t.mode === "upsert") note.textContent = "Updates existing employees and adds new ones. Driver designations (Primary / Secondary / Non-Driver) are preserved — never overwritten by an import.";
    else if (t.mode === "append-dedup") note.textContent = "New records are added; rows that already exist are skipped.";
    else note.textContent = "New records are added to the list.";
    container.appendChild(note);

    if (state.type === "physicals") {
      container.appendChild(el("div", { class: "import-note", html:
        "Physicals import needs two columns on <b>DrivingSafety_Physicals</b>: <b>ExpirationDate</b> (Date) and <b>Result</b> (text). Add them first if they aren't there." }));
    }

    // dropzone
    const dzTitle = el("h3", { text: "Drop the " + t.label.toLowerCase() + " Excel file here" });
    dzTitle.appendChild(DS.formatHint(FORMAT_SPECS[state.type]));
    const dz = el("div", { class: "dropzone" }, [
      el("div", { class: "dz-ico", text: "⬆" }),
      dzTitle,
      el("p", { text: "or click to choose a file (.xlsx)" }),
    ]);
    const fileInput = el("input", { type: "file", accept: ".xlsx,.xls", style: "display:none" });
    dz.appendChild(fileInput);
    dz.addEventListener("click", () => fileInput.click());
    dz.addEventListener("dragover", e => { e.preventDefault(); dz.classList.add("drag"); });
    dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
    dz.addEventListener("drop", e => {
      e.preventDefault(); dz.classList.remove("drag");
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0], container);
    });
    fileInput.addEventListener("change", () => { if (fileInput.files[0]) handleFile(fileInput.files[0], container); });
    container.appendChild(dz);

    const result = el("div", { id: "importResult", style: "margin-top:20px" });
    container.appendChild(result);
  }

  async function handleFile(file, container) {
    const result = document.getElementById("importResult");
    DS.showLoading(result, "Reading " + file.name + "…");
    const t = TYPES[state.type];
    try {
      const XLSXlib = await ensureXlsx();
      const buf = await file.arrayBuffer();
      const wb = XLSXlib.read(buf, { type: "array", cellDates: true });
      const parsed = t.parse(wb, state.ctx);
      const records = parsed.records, warnings = parsed.warnings || [];
      if (!records.length) { renderMsg(result, "No records found in that file. Check that it's the right export.", "warn"); return; }
      const key = v => DS.util.empKey(v);

      if (t.mode === "upsert") {
        DS.showLoading(result, "Comparing with the current roster…");
        const existing = await DS.spGet(DS.LISTS.roster, { select: ["Id", "EmployeeId"] });
        if (parsed.hasBadge && !(await DS.spFieldExists(DS.LISTS.roster, "Badge"))) {
          records.forEach(r => { delete r.Badge; });
          warnings.push("This file has badge numbers, but DrivingSafety_Roster has no 'Badge' column yet. Add it (Single line of text) and upload again so badges are stored.");
        }
        const existingMap = {};
        existing.forEach(e => { const k = key(e.EmployeeId); if (k) existingMap[k] = e; });
        const seen = new Map(), dupRows = [];
        const unique = records.filter(r => {
          const k = key(r.EmployeeId);
          if (seen.has(k)) { dupRows.push({ row: r, first: seen.get(k) }); return false; }
          seen.set(k, r); return true;
        });
        if (dupRows.length) warnings.push(dupRows.length + " employee number(s) appear more than once in the file \u2014 the first row for each is used.");
        const fileIds = new Set(unique.map(r => key(r.EmployeeId)));
        const toUpdate = [], toCreate = [];
        unique.forEach(r => { const ex = existingMap[key(r.EmployeeId)]; if (ex) toUpdate.push({ id: ex.Id, fields: r }); else toCreate.push(r); });
        const toInactivate = existing.filter(e => { const k = key(e.EmployeeId); return k && !fileIds.has(k); });
        state.parsed = { mode: "upsert", records, toUpdate, toCreate, toInactivate, warnings, fileName: file.name, noEmp: parsed.noEmp || [], dupRows };
        renderPreview(result, container);
        return;
      }

      // Match every row to a person: Employee # first, then badge.
      DS.showLoading(result, "Matching rows to the roster…");
      const cache = await DS.data.load();
      const ix = cache.idx.ids;
      const match = { employee: 0, badge: 0 }, unmatched = [], conflicts = [], matched = [];
      records.forEach(r => {
        const res = DS.ids.resolve(ix, r.EmployeeId, r.__name);
        r.__sourceId = String(r.EmployeeId == null ? "" : r.EmployeeId).trim();
        if (!res.rec) {
          const sug = res.suggestions || [];
          const full = sug.filter(x => DS.ids.nameScore(x, r.__name) === 2);
          unmatched.push({ row: r, suggestions: sug, nameCandidate: full.length === 1 ? full[0] : null });
          return;
        }
        match[res.method]++;
        r.__method = res.method;
        if (res.conflict) conflicts.push({ row: r, chosen: res.rec, other: res.other, byName: res.byName });
        r.EmployeeId = String(res.rec.EmployeeId).trim();   // always saved under the roster's Employee #
        r.Title = res.rec.Title || "";
        matched.push(r);
      });

      let toWrite = matched, dupCount = 0;
      const onFileDups = [], inFileDups = [];
      if (t.mode === "append-dedup" && t.existingKeys) {
        DS.showLoading(result, "Checking for existing records…");
        const existing = await t.existingKeys();
        const seen = new Map();
        toWrite = matched.filter(r => {
          const k = t.dedupKey(r);
          if (existing.has(k)) { onFileDups.push(r); return false; }
          if (seen.has(k)) { inFileDups.push({ row: r, first: seen.get(k) }); return false; }
          seen.set(k, r); return true;
        });
        dupCount = onFileDups.length + inFileDups.length;
        // Rows that could be matched by name (opt-in): prepare them now, with the same duplicate checks
        var nameWrite = [];
        unmatched.filter(u => u.nameCandidate).forEach(u => {
          const r2 = Object.assign({}, u.row, { EmployeeId: String(u.nameCandidate.EmployeeId).trim(), Title: u.nameCandidate.Title || "", __method: "name" });
          const k = t.dedupKey(r2);
          if (existing.has(k) || seen.has(k)) return;
          seen.set(k, r2); nameWrite.push(r2);
        });
      }

      // Courses: people this file has only some of the required courses for
      let partial = [];
      if (state.type === "courses") {
        const req = state.ctx.requiredTitles || [];
        const byEmp = {};
        matched.forEach(r => {
          const k = key(r.EmployeeId);
          (byEmp[k] = byEmp[k] || { row: r, titles: new Set() }).titles.add(String(r.CourseTitle || "").trim());
        });
        partial = Object.values(byEmp).filter(x => x.titles.size < req.length)
          .map(x => ({ row: x.row, has: req.filter(tt => x.titles.has(tt)), missing: req.filter(tt => !x.titles.has(tt)) }));
      }

      state.parsed = { records, toWrite, dupCount, onFileDups, inFileDups, warnings, fileName: file.name, match, unmatched, conflicts, partial,
        nameWrite: typeof nameWrite !== "undefined" ? nameWrite : [] };
      state.useNameMatch = false;
      renderPreview(result, container);
    } catch (e) {
      renderMsg(result, e.message, "warn");
    }
  }

  function renderPreview(result, container) {
    const t = TYPES[state.type];
    const p = state.parsed;
    result.innerHTML = "";

    const head = el("div", { class: "card__head" }, [
      el("h3", { text: p.fileName }),
      el("span", { class: "count-pill", text: p.records.length + " row(s) parsed" }),
    ]);
    const body = el("div", { class: "card__body" });

    let sample = p.records.slice(0, 8);
    let confirmLabel, doRun, inactivateCb = null, updateConfirm = null;

    if (p.mode === "upsert") {
      body.appendChild(el("div", { class: "import-note", text:
        p.toUpdate.length + " existing updated \u00b7 " + p.toCreate.length + " new \u00b7 designations preserved." }));
      if (p.toInactivate.length) {
        const wrap = el("label", { class: "check", style: "margin:2px 0 12px" });
        inactivateCb = el("input", { type: "checkbox" }); inactivateCb.checked = true;
        wrap.appendChild(inactivateCb);
        wrap.appendChild(el("span", { text: "Mark " + p.toInactivate.length + " employee(s) not in this file as inactive" }));
        body.appendChild(wrap);
        body.appendChild(el("div", { class: "import-note warn", text:
          "That count should roughly match recent separations. If it's unexpectedly large, this may be only a partial roster \u2014 uncheck the box." }));
      }
      confirmLabel = "Apply roster update";
      doRun = () => runUpsert(result, container, inactivateCb ? inactivateCb.checked : false);
    } else {
      let summary;
      if (t.mode === "append-dedup") summary = p.toWrite.length + " new \u00b7 " +
        (p.onFileDups || []).length + " already in SharePoint \u00b7 " +
        (p.inFileDups || []).length + " repeated within this file (skipped).";
      else summary = "Will add " + p.toWrite.length + " record(s).";
      body.appendChild(el("div", { class: "import-note", text: summary }));
      if (p.match) {
        const um = p.unmatched.length, cf = p.conflicts.filter(c => !c.byName).length;
        body.appendChild(el("div", { class: "import-note" + (um || cf ? " warn" : ""), text:
          "Matched to the roster: " + p.match.employee + " by employee number \u00b7 " + p.match.badge + " by badge" +
          (um ? " \u00b7 " + um + " row(s) match no one and will be skipped" : "") +
          (cf ? " \u00b7 " + cf + " number(s) belong to two people and the name didn't settle it \u2014 check below" : "") + "." }));
        if ((p.nameWrite || []).length) {
          const cb = el("input", { type: "checkbox" }); cb.checked = !!state.useNameMatch;
          cb.addEventListener("change", () => { state.useNameMatch = cb.checked; if (updateConfirm) updateConfirm(); });
          body.appendChild(el("label", { class: "check", style: "margin:2px 0 10px; align-items:flex-start" }, [cb,
            el("span", { text: "Also match " + p.nameWrite.length + " of those row(s) by name \u2014 the ID matched no one, but exactly one person on the roster has the same first and last name (likely a mistyped ID). They'll be listed in Upload results so you can spot-check them." })]));
        }
      }
      sample = p.toWrite.slice(0, 8);
      confirmLabel = "Import (" + p.toWrite.length + ")";
      doRun = () => runImport(result, container);
    }

    (p.warnings || []).forEach(w => body.appendChild(el("div", { class: "import-note warn", text: w })));
    const probs = problemRows(t, p);
    if (probs.length) body.appendChild(problemsDetails(probs, p.fileName));

    // preview table
    const cols = t.cols;
    if (sample.length) {
      body.appendChild(el("table", { class: "tbl", style: "margin-top:6px" }, [
        el("thead", null, el("tr", null, cols.map(c => el("th", { text: c[0] })))),
        el("tbody", null, sample.map(r => el("tr", null, cols.map(c => {
          let v = r[c[1]];
          if (/date/i.test(c[1])) v = DS.fmtDate(v);
          return el("td", { text: v == null || v === "" ? "—" : String(v) });
        })))),
      ]));
    }

    const confirmBtn = el("button", { class: "btn", text: confirmLabel });
    confirmBtn.disabled = (p.mode !== "upsert" && p.toWrite.length === 0);
    if (p.mode !== "upsert") {
      updateConfirm = () => {
        const n = p.toWrite.length + (state.useNameMatch ? (p.nameWrite || []).length : 0);
        confirmBtn.textContent = "Import (" + n + ")"; confirmBtn.disabled = n === 0;
      };
      updateConfirm();
    }
    const cancelBtn = el("button", { class: "btn btn--ghost", text: "Cancel" });
    cancelBtn.addEventListener("click", () => { state.parsed = null; document.getElementById("importResult").innerHTML = ""; });
    confirmBtn.addEventListener("click", doRun);
    body.appendChild(el("div", { style: "display:flex; gap:10px; margin-top:18px" }, [confirmBtn, cancelBtn]));

    result.appendChild(el("div", { class: "card" }, [head, body]));
  }

  /* ---------------- Problem rows (preview, result screen, Upload results log) ---------------- */
  const REVIEW_REASONS = ["No match on roster", "Two possible people", "Failed to save", "No employee number", "Employee # repeated in file"];
  function problemRows(t, p, failed) {
    const out = [];
    const base = r => ({ row: r.__row || "", id: r.__sourceId != null ? r.__sourceId : String(r.EmployeeId || ""),
      name: r.__name || r.Title || "", info: t && t.describe ? t.describe(r) : "" });
    (failed || []).forEach(e => { const it = e.item && (e.item.fields || e.item); out.push(Object.assign(it ? base(it) : { row: "", id: "", name: "", info: "" }, { why: "Failed to save", note: e.message || String(e) })); });
    (p.unmatched || []).forEach(u => {
      if (p.usedNameMatch && u.nameCandidate) {
        out.push(Object.assign(base(u.row), { why: "Matched by name (ID didn't match)",
          note: "Saved to " + (u.nameCandidate.Title || "?") + " (#" + u.nameCandidate.EmployeeId + (u.nameCandidate.Badge ? ", badge " + u.nameCandidate.Badge : "") + ")" }));
        return;
      }
      out.push(Object.assign(base(u.row), { why: "No match on roster",
        note: u.suggestions && u.suggestions.length ? "Possible: " + u.suggestions.map(x => (x.Title || "?") + " (#" + x.EmployeeId + (x.Badge ? ", badge " + x.Badge : "") + ")").join("; ")
          : "Likely a former employee or a mistyped ID" }));
    });
    (p.conflicts || []).forEach(c => out.push(Object.assign(base(c.row), { why: c.byName ? "Two possible people (name confirmed)" : "Two possible people",
      note: "Saved to " + (c.chosen.Title || "?") + " (#" + c.chosen.EmployeeId + ", by " + (c.row.__method === "badge" ? "badge" : "employee #") + ")" +
        (c.byName ? ", which the name in the file agrees with" : ", but the name in the file didn't settle it") +
        ". The same number is " + (c.other.Title || "?") + "'s " + (c.row.__method === "badge" ? "employee #" : "badge") + "." })));
    (p.noEmp || []).forEach(r => out.push(Object.assign(base(r), { why: "No employee number", note: r.Badge ? "Badge " + r.Badge : "" })));
    (p.dupRows || []).forEach(d => out.push(Object.assign(base(d.row), { why: "Employee # repeated in file", note: "Row " + (d.first.__row || "?") + " used instead" })));
    (p.partial || []).forEach(x => out.push(Object.assign(base(x.row), { why: "Only some required courses", info: "In file: " + (x.has.join(", ") || "none"), note: "Not in file: " + x.missing.join(", ") })));
    (p.inFileDups || []).forEach(d => out.push(Object.assign(base(d.row), { why: "Repeated in file", note: "Same as row " + (d.first.__row || "?") })));
    (p.onFileDups || []).forEach(r => out.push(Object.assign(base(r), { why: "Already in SharePoint", note: "" })));
    return out;
  }

  // Expandable table of problem rows + Excel download
  function problemsDetails(probs, fileName, open) {
    const counts = {};
    probs.forEach(x => { counts[x.why] = (counts[x.why] || 0) + 1; });
    const review = probs.filter(x => REVIEW_REASONS.includes(x.why)).length;
    const summary = Object.keys(counts).map(k => counts[k] + " " + k.toLowerCase()).join(" \u00b7 ");
    const dl = el("button", { class: "btn btn--ghost btn--sm", type: "button", text: "Download as Excel" });
    dl.addEventListener("click", e => { e.preventDefault(); downloadProblems(probs, fileName); });
    const shown = probs.slice(0, 500);
    const det = el("details", { style: "margin:6px 0 10px" }, [
      el("summary", { style: "cursor:pointer; font-size:13px; color:" + (review ? "var(--overdue)" : "var(--navy-500)"),
        text: (review ? review + " row(s) need review \u2014 " : "") + "show details (" + summary + ")" }),
      el("div", { style: "margin:8px 0" }, dl),
      el("div", { style: "overflow-x:auto; max-height:380px; overflow-y:auto" }, el("table", { class: "tbl" }, [
        el("thead", null, el("tr", null, ["Why", "Row", "ID in file", "Name in file", "Details", "Note"].map(h => el("th", { text: h })))),
        el("tbody", null, shown.map(x => el("tr", null, [x.why, String(x.row || ""), x.id, x.name, x.info, x.note].map(v => el("td", { text: v || "\u2014" }))))),
      ])),
      probs.length > shown.length ? el("div", { class: "help", text: "Showing 500 of " + probs.length + " \u2014 download for the full list." }) : null,
    ]);
    if (open) det.open = true;
    return det;
  }

  async function downloadProblems(probs, fileName) {
    try {
      const X = await ensureXlsx();
      const ws = X.utils.json_to_sheet(probs.map(x => ({ "Why": x.why, "Row in file": x.row, "ID in file": x.id, "Name in file": x.name, "Details": x.info, "Note": x.note })));
      const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Upload problems");
      X.writeFile(wb, "Upload problems - " + String(fileName || "upload").replace(/\.[^.]+$/, "") + ".xlsx");
    } catch (e) { DS.toast("Couldn't create the download: " + e.message, "error"); }
  }

  /* Save one entry to the Upload results log (DrivingSafety_UploadLog).
     Problem rows are stored as JSON in the Details column (large lists trimmed). */
  async function logUpload(entry) {
    const doc = { v: 1, at: new Date().toISOString(),
      by: (DS.me && (DS.me.mail || DS.me.userPrincipalName)) || "", type: entry.type, file: entry.file,
      counts: entry.counts || {}, warnings: entry.warnings || [],
      review: (entry.problems || []).filter(x => REVIEW_REASONS.includes(x.why)).length,
      problems: (entry.problems || []).slice(0, 800), truncated: (entry.problems || []).length > 800 };
    let json = JSON.stringify(doc);
    while (json.length > 60000 && doc.problems.length > 20) {
      doc.problems = doc.problems.slice(0, Math.floor(doc.problems.length / 2)); doc.truncated = true; json = JSON.stringify(doc);
    }
    try { await DS.spCreate(DS.LISTS.uploads, { Title: String(entry.file || entry.type), UploadType: entry.type, Details: json }); return true; }
    catch (e) { console.warn("Upload results log not saved:", e.message); return false; }
  }
  const logNote = ok => ok ? "Saved to Audit log \u2192 Upload results." :
    "Couldn't save to Upload results \u2014 create the DrivingSafety_UploadLog list (see setup notes).";

  async function runUpsert(result, container, doInactivate) {
    if (DS.job.busy()) return;
    const p = state.parsed;
    const t = TYPES.roster;
    const listName = DS.LISTS.roster;
    const ops = [];
    // update existing: org fields (+ badge) + reactivate; NEVER DriverStatus → designation preserved
    p.toUpdate.forEach(u => ops.push({ kind: "update", id: u.id, fields: Object.assign(stripMeta(u.fields), { ActiveEmployee: true }), rec: u.fields }));
    // new hires: designation left blank → "Needs a designation" on the Dashboard
    p.toCreate.forEach(r => ops.push({ kind: "create", fields: Object.assign(stripMeta(r), { ActiveEmployee: true }), rec: r }));
    // departures: mark inactive (keeps the record + its designation for a possible return)
    if (doInactivate) p.toInactivate.forEach(e => ops.push({ kind: "update", id: e.Id, fields: { ActiveEmployee: false }, rec: { EmployeeId: e.EmployeeId } }));

    result.innerHTML = "";
    const bar = el("div", { class: "progress" }, el("div", { class: "progress__bar" }));
    const barFill = bar.firstChild;
    const status = el("div", { style: "font-size:13px; color:var(--slate)" });
    result.appendChild(el("div", { class: "card" }, el("div", { class: "card__body" }, [
      el("h3", { text: "Updating roster…", style: "font-size:15px; margin-bottom:10px" }), bar, status,
    ])));

    const errors = await runBatched(ops, async op => {
      if (op.kind === "create") await DS.spCreate(listName, op.fields);
      else await DS.spUpdate(listName, op.id, op.fields);
    }, (done, total, label) => { barFill.style.width = (total ? done / total * 100 : 100) + "%"; status.textContent = label || ("Processing " + done + " of " + total + "…"); }, null, "Roster upload");
    errors.forEach(e => { if (e.item && e.item.rec) e.item = e.item.rec; });

    const updated = p.toUpdate.length, created = p.toCreate.length, inactivated = doInactivate ? p.toInactivate.length : 0;
    await DS.audit("Roster import", listName, null,
      updated + " updated, " + created + " new, " + inactivated + " inactivated; designations preserved");
    const problems = problemRows(t, p, errors);
    const logged = await logUpload({ type: "Roster", file: p.fileName, warnings: p.warnings, problems, counts: {
      rows: p.records.length + (p.noEmp || []).length, updated, added: created, inactivated,
      noEmployeeNumber: (p.noEmp || []).length, repeatedInFile: (p.dupRows || []).length, failed: errors.length } });
    DS.data.clear();
    const errCount = errors.length;
    renderMsg(result,
      "Roster processed \u2014 " + updated + " existing updated, " + created + " added, " + inactivated + " marked inactive. " +
      (ops.length - errCount) + " of " + ops.length + " operations succeeded." +
      (errCount ? " Re-running the same file is safe and will retry only what's needed." : " Designations were preserved.") + " " + logNote(logged),
      errCount ? "warn" : "ok");
    const cb = result.querySelector(".card__body");
    if (errCount) cb.appendChild(errorSummaryEl(errors));
    if (problems.length) cb.appendChild(problemsDetails(problems, p.fileName));
    DS.toast("Roster update complete.", errCount ? "error" : "success");
    state.parsed = null;
  }

  async function runImport(result, container) {
    if (DS.job.busy()) return;
    const t = TYPES[state.type];
    const p = state.parsed;
    const listName = t.list();
    result.innerHTML = "";
    const bar = el("div", { class: "progress" }, el("div", { class: "progress__bar" }));
    const barFill = bar.firstChild;
    const status = el("div", { style: "font-size:13px; color:var(--slate)" });
    result.appendChild(el("div", { class: "card" }, el("div", { class: "card__body" }, [
      el("h3", { text: "Importing…", style: "font-size:15px; margin-bottom:10px" }), bar, status,
    ])));

    function progress(done, total, label) {
      barFill.style.width = (total ? (done / total * 100) : 100) + "%";
      status.textContent = label || ("Writing " + done + " of " + total + "…");
    }

    try {
      p.usedNameMatch = !!state.useNameMatch && (p.nameWrite || []).length > 0;
      const writes = p.usedNameMatch ? p.toWrite.concat(p.nameWrite) : p.toWrite;
      const createErrors = await runBatched(writes, rec => DS.spCreate(listName, stripMeta(rec)),
        (d, tot, label) => progress(d, tot, label || "Adding records"), null, t.label + " upload");
      const created = writes.length - createErrors.length;
      const byName = p.usedNameMatch ? p.nameWrite.length : 0;

      await DS.audit("Bulk import \u2014 " + t.label, listName, null,
        "Imported: " + created + " added" + (p.dupCount ? ", " + p.dupCount + " duplicates skipped" : "") +
        (p.unmatched && p.unmatched.length ? ", " + p.unmatched.length + " matched no one" : ""));

      const problems = problemRows(t, p, createErrors);
      const m = p.match || { employee: 0, badge: 0 };
      const logged = await logUpload({ type: t.label, file: p.fileName, warnings: p.warnings, problems, counts: {
        rows: p.records.length, added: created, byEmployeeNumber: m.employee, byBadge: m.badge, byName,
        noMatch: (p.unmatched || []).length - byName, matchedTwoPeople: (p.conflicts || []).filter(c => !c.byName).length,
        resolvedByName: (p.conflicts || []).filter(c => c.byName).length,
        alreadyInSharePoint: (p.onFileDups || []).length, repeatedInFile: (p.inFileDups || []).length,
        onlySomeCourses: (p.partial || []).length, failed: createErrors.length } });

      DS.data.clear();  // recompute dashboard/roster off fresh data next visit

      const errCount = createErrors.length;
      const um = (p.unmatched || []).length - byName;
      renderMsg(result,
        created + " " + t.label.toLowerCase() + " record(s) imported" +
        (p.dupCount ? ", " + p.dupCount + " skipped as duplicates" : "") +
        (um ? ", " + um + " skipped because they match no one on the roster" : "") +
        (errCount ? " \u2014 " + errCount + " failed." : ".") + " " + logNote(logged),
        errCount || um ? "warn" : "ok");
      const cb = result.querySelector(".card__body");
      if (errCount) cb.appendChild(errorSummaryEl(createErrors));
      if (problems.length) cb.appendChild(problemsDetails(problems, p.fileName, !!(errCount || um)));
      DS.toast(created + " " + t.label.toLowerCase() + " record(s) imported.", errCount ? "error" : "success");
      state.parsed = null;
    } catch (e) {
      renderMsg(result, "Import failed: " + e.message, "warn");
      DS.toast("Import failed: " + e.message, "error");
    }
  }

  function renderMsg(container, msg, kind) {
    container.innerHTML = "";
    container.appendChild(el("div", { class: "card" }, el("div", { class: "card__body" },
      el("div", { class: "import-note" + (kind === "warn" ? " warn" : ""), text: msg }))));
  }

  /* ---- on-screen error summary: groups failures by message so 777 errors
     read as "3 distinct problems" instead of requiring the dev console ---- */
  function errorSummaryEl(errors) {
    const counts = {};
    errors.forEach(e => { const m = (e && e.message) || String(e); counts[m] = (counts[m] || 0) + 1; });
    const top = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8);
    console.warn("Write errors:", errors.map(e => e && e.message));   // full list still in the console
    return el("div", { class: "import-note warn", style: "margin-top:10px" }, [
      el("div", { text: "What went wrong (" + errors.length + " total, grouped):", style: "font-weight:600; margin-bottom:6px" }),
      el("ul", { style: "margin:0; padding-left:18px" }, top.map(([m, n]) => el("li", { text: n + " \u00d7 " + m }))),
    ]);
  }

  DS.registerScreen("imports", { title: "Imports", icon: "▾", render: renderImports });
  DS.ensureXlsx = ensureXlsx;   // shared with the reports export
  DS.runBatched = runBatched;   // shared with the legacy migration tool
  DS.errorSummaryEl = errorSummaryEl;
  DS.logUpload = logUpload;            // migration tool records its runs too
  DS.problemsDetails = problemsDetails; // Audit log → Upload results
  DS.downloadProblems = downloadProblems;
  DS.REVIEW_REASONS = REVIEW_REASONS;

})();
