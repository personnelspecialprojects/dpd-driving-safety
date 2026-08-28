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
          iDate = ci("Date Completed"), iScore = ci("Score"), iAttempts = ci("Attempts");
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
      });
    }
    return { records, warnings: [] };
  }

  function parseRoster(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    const loc = locate(r, "Emp #");
    if (!loc) throw new Error("Couldn't find the 'Emp #' header — is this the Employees-with-Supervisors export?");
    const ci = colFinder(loc.header);
    const iEmp = ci("Emp #"), iLast = ci("LastName"), iFirst = ci("FirstName"), iRank = ci("Rank"),
          iOrg = ci("Working Org"), iWg = ci("Working Workgroup"), iSup = ci("1st Line Supervisor"), iHire = ci("Hire Date");
    const records = []; let noHire = 0;
    for (let i = loc.headerIdx + 1; i < r.length; i++) {
      const row = r[i]; if (!row || !row[iEmp]) continue;
      const first = String(row[iFirst] || "").trim(), last = String(row[iLast] || "").trim();
      const hire = iHire >= 0 && row[iHire] ? DS.isoDate(row[iHire]) : null;
      if (!hire) noHire++;
      records.push({
        Title: (first + " " + last).trim(),
        LastName: last,
        EmployeeId: String(row[iEmp]).trim(),
        Division: String(row[iOrg] || ""),
        Assignment: String(row[iWg] || ""),
        Supervisor: String(row[iSup] || ""),
        HireDate: hire,
        Rank: String(row[iRank] || ""),
        DriverStatus: "Primary",
        ActiveEmployee: true,
      });
    }
    const warnings = [];
    if (iHire < 0) warnings.push("No 'Hire Date' column found — course due-dates for new hires need it.");
    else if (noHire) warnings.push(noHire + " employee(s) have no hire date.");
    return { records, warnings };
  }

  function parsePhysicals(wb) {
    const ws = wb.Sheets["Data"] || wb.Sheets[wb.SheetNames[0]];
    const r = rowsOf(ws);
    const loc = locate(r, "Employee Number");
    if (!loc) throw new Error("Couldn't find the 'Employee Number' header — is this the driver-physicals export?");
    const ci = colFinder(loc.header);
    const iDate = ci("Date Tested"), iEmp = ci("Employee Number"), iRes = ci("Driver Physical Test Results"),
          iExp = ci("Phy Exp Date"), iClinic = ci("Clinic Location");
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
      });
    }
    const warnings = [];
    if (fails) warnings.push(fails + " row(s) are marked Fail — imported as-is; review whether they should count.");
    if (noExp) warnings.push(noExp + " row(s) have no expiration date — those won't drive a due date.");
    return { records, warnings };
  }

  /* ============================================================
     IMPORT TYPES — parser + write strategy per type
     ============================================================ */
  const TYPES = {
    courses: {
      label: "Courses", list: () => DS.LISTS.courses, mode: "append-dedup",
      parse: (wb, ctx) => parseCourses(wb, ctx.requiredTitles),
      cols: [["Employee", "EmployeeId"], ["Course", "CourseTitle"], ["Completed", "DateCompleted"], ["Score", "Score"]],
      dedupKey: r => [r.EmployeeId, r.CourseTitle, r.DateCompleted].join("|"),
      existingKeys: async () => {
        const rows = await DS.spGet(DS.LISTS.courses, { select: ["EmployeeId", "CourseTitle", "DateCompleted"] });
        return new Set(rows.map(x => [String(x.EmployeeId || "").trim(), String(x.CourseTitle || "").trim(), DS.isoDate(x.DateCompleted)].join("|")));
      },
    },
    accidents: {
      label: "Accidents", list: () => DS.LISTS.accidents, mode: "append-dedup",
      parse: (wb) => parseAccidents(wb),
      cols: [["Incident #", "IncidentNumber"], ["Employee", "EmployeeId"], ["Date", "AccidentDate"], ["Final pts", "FinalPoints"]],
      dedupKey: r => String(r.IncidentNumber).trim(),
      existingKeys: async () => {
        const rows = await DS.spGet(DS.LISTS.accidents, { select: ["IncidentNumber"] });
        return new Set(rows.map(x => String(x.IncidentNumber || "").trim()));
      },
    },
    roster: {
      label: "Roster", list: () => DS.LISTS.roster, mode: "replace",
      parse: (wb) => parseRoster(wb),
      cols: [["Name", "Title"], ["ID", "EmployeeId"], ["Division", "Division"], ["Rank", "Rank"]],
    },
    physicals: {
      label: "Physicals", list: () => DS.LISTS.physicals, mode: "append",
      parse: (wb) => parsePhysicals(wb),
      cols: [["Employee", "EmployeeId"], ["Tested", "PhysicalDate"], ["Expires", "ExpirationDate"], ["Result", "Result"]],
    },
  };

  /* ---- concurrency-limited runner with progress ---- */
  async function runBatched(items, worker, onProgress, concurrency) {
    concurrency = concurrency || 4;
    let done = 0; const errors = [];
    const queue = items.slice();
    async function lane() {
      while (queue.length) {
        const item = queue.shift();
        try { await worker(item); } catch (e) { errors.push(e); }
        done++; onProgress(done, items.length);
      }
    }
    await Promise.all(Array.from({ length: Math.min(concurrency, items.length || 1) }, lane));
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
    const note = el("div", { class: "import-note" + (t.mode === "replace" ? " warn" : "") });
    if (t.mode === "replace") note.textContent = "Roster import replaces every current roster record with the uploaded file.";
    else if (t.mode === "append-dedup") note.textContent = "New records are added; rows that already exist are skipped.";
    else note.textContent = "New records are added to the list.";
    container.appendChild(note);

    if (state.type === "physicals") {
      container.appendChild(el("div", { class: "import-note", html:
        "Physicals import needs two columns on <b>DrivingSafety_Physicals</b>: <b>ExpirationDate</b> (Date) and <b>Result</b> (text). Add them first if they aren't there." }));
    }

    // dropzone
    const dz = el("div", { class: "dropzone" }, [
      el("div", { class: "dz-ico", text: "⬆" }),
      el("h3", { text: "Drop the " + t.label.toLowerCase() + " Excel file here" }),
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
      const { records, warnings } = t.parse(wb, state.ctx);
      if (!records.length) { renderMsg(result, "No records found in that file. Check that it's the right export.", "warn"); return; }

      // dedup preview
      let toWrite = records, dupCount = 0;
      if (t.mode === "append-dedup" && t.existingKeys) {
        DS.showLoading(result, "Checking for existing records…");
        const existing = await t.existingKeys();
        toWrite = records.filter(r => !existing.has(t.dedupKey(r)));
        dupCount = records.length - toWrite.length;
      }
      state.parsed = { records, toWrite, dupCount, warnings, fileName: file.name };
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

    // summary line
    let summary;
    if (t.mode === "replace") summary = "Will replace all roster records with these " + p.toWrite.length + ".";
    else if (t.mode === "append-dedup") summary = p.toWrite.length + " new · " + p.dupCount + " already on file (skipped).";
    else summary = "Will add " + p.toWrite.length + " record(s).";

    const body = el("div", { class: "card__body" });
    body.appendChild(el("div", { class: "import-note", text: summary }));
    (p.warnings || []).forEach(w => body.appendChild(el("div", { class: "import-note warn", text: w })));

    // preview table (first 8)
    const cols = t.cols;
    const sample = p.toWrite.slice(0, 8);
    if (sample.length) {
      body.appendChild(el("table", { class: "tbl", style: "margin-top:6px" }, [
        el("thead", null, el("tr", null, cols.map(c => el("th", { text: c[0] })))),
        el("tbody", null, sample.map(r => el("tr", null, cols.map(c => {
          let v = r[c[1]];
          if (/date/i.test(c[1]) || c[1] === "PhysicalDate" || c[1] === "ExpirationDate" || c[1] === "AccidentDate" || c[1] === "DateCompleted")
            v = DS.fmtDate(v);
          return el("td", { text: v == null || v === "" ? "—" : String(v) });
        })))),
      ]));
    }

    // action row
    const confirmBtn = el("button", { class: "btn", text: (t.mode === "replace" ? "Replace roster" : "Import") + " (" + p.toWrite.length + ")" });
    confirmBtn.disabled = p.toWrite.length === 0;
    const cancelBtn = el("button", { class: "btn btn--ghost", text: "Cancel" });
    cancelBtn.addEventListener("click", () => { state.parsed = null; document.getElementById("importResult").innerHTML = ""; });
    confirmBtn.addEventListener("click", () => runImport(result, container));
    body.appendChild(el("div", { style: "display:flex; gap:10px; margin-top:18px" }, [confirmBtn, cancelBtn]));

    result.appendChild(el("div", { class: "card" }, [head, body]));
  }

  async function runImport(result, container) {
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
      status.textContent = (label || "Writing") + " " + done + " of " + total + "…";
    }

    let deleted = 0, deleteErrors = [];
    try {
      // REPLACE: delete existing first
      if (t.mode === "replace") {
        const existing = await DS.spGet(listName, { select: ["Id"] });
        deleteErrors = await runBatched(existing, item => DS.spDelete(listName, item.Id),
          (d, tot) => progress(d, tot, "Clearing old records"));
        deleted = existing.length - deleteErrors.length;
      }

      // CREATE
      const createErrors = await runBatched(p.toWrite, rec => DS.spCreate(listName, rec),
        (d, tot) => progress(d, tot, "Adding records"));
      const created = p.toWrite.length - createErrors.length;

      await DS.audit("Bulk import — " + t.label, listName, null,
        (t.mode === "replace" ? "Replaced roster: " : "Imported: ") +
        created + " added" + (t.mode === "replace" ? ", " + deleted + " cleared" : "") +
        (p.dupCount ? ", " + p.dupCount + " duplicates skipped" : ""));

      DS.data.clear();  // recompute dashboard/roster off fresh data next visit

      const errCount = createErrors.length + deleteErrors.length;
      renderMsg(result,
        created + " " + t.label.toLowerCase() + " record(s) imported" +
        (t.mode === "replace" ? " (" + deleted + " old cleared)" : "") +
        (p.dupCount ? ", " + p.dupCount + " skipped as duplicates" : "") +
        (errCount ? " — " + errCount + " failed (see console)." : "."),
        errCount ? "warn" : "ok");
      if (errCount) console.warn("Import errors:", createErrors.concat(deleteErrors).map(e => e.message));
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

  DS.registerScreen("imports", { title: "Imports", icon: "▾", render: renderImports });
  DS.ensureXlsx = ensureXlsx;   // shared with the reports export

})();
