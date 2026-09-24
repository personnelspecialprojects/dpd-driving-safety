/* ============================================================
   migrate-legacy.js — ONE-TIME legacy data migration tool
   ------------------------------------------------------------
   Not linked in the sidebar. Reach it at #/migrate. Safe to delete
   this file (and its <script> tag in index.html) once Julie's team
   is fully on the standard Import screen.

   Inputs (both parsed entirely in the browser — files never leave it):
     1. "Safety Team Main" — the current-status roster. Has real headers:
        Last Name, First Name, Emp#, Badge, Rank, Driver Physical Date,
        Defensive Driving Date, Future verified.
        HIGHEST PRECEDENCE for designation and for physical/course due dates.
     2. "Reference physicals" workbook — the exam history. No header row.
        Sheets matched by fuzzy name:
          - Master List / Add to Main Spreadsheet: TestDate, "Last, First",
            EmpID, Result, ExpirationDate
          - New Hires: TestDate, LastName, FirstName, Result, ExpirationDate
            (NO Employee ID column — matched by name against the roster,
            unlike every other sheet, which is why it's the one source
            that can land in the review queue just for being unmatched)
          - Pending: Date, "First Last", EmpID, "Pending", note

   Policy encoded here (confirmed with Jared):
     - "Non-Primary" (any spelling) → Secondary (needs courses, not physical)
     - "Non-Driver"/"ND License" (any spelling) → Non-Driver (needs neither)
     - "Helicopter-Exempt"/"Air One" → flagged for individual review, suggested
       Secondary — they have their own, separate aviation physical
     - "Non-Driving-Medical" → flagged (may be temporary, not a designation)
     - "Expired <old date>" → flagged as stale, per Jared: treat as unreliable
     - The date in Driver Physical Date / Defensive Driving Date IS the due
       date (not last-completed). Physicals has a real ExpirationDate field
       for this. Courses doesn't, so a completion date is back-computed
       (due date minus the renewal cycle) and written to all required
       course titles as one bundled event, per Jared's call.
     - "Not needed"/"Needs"/"Academy" in the DD column → no record needed;
       that's the app's normal "not started" state already.
     - A designation is only ever SET by this tool when the employee is
       still at the untouched default ("Primary"). Anything already
       Secondary/Non-Driver — from a prior run or a manual edit — is left
       alone. This makes the tool safe to re-run.
     - Dates are ADDED as history, never overwritten. Every exam in the
       reference workbook becomes its own "Legacy History" record. Master-sheet
       due dates become separate "Master Sheet" records that only count when no
       real record (upload, manual entry, exam history) exists — Julie's sheet
       has stale dates, so real data always takes precedence.
     - Re-running only adds what's missing; existing records are never edited.
   ============================================================ */
(function () {
  const el = DS.el;
  const L = DS.LISTS;
  const PLAUS_MIN = new Date(2015, 0, 1);

  function plausibleDate(v) {
    const d = DS.parseDate(v);
    if (!d) return null;
    const max = DS.util.addYears(new Date(), 10);
    return (d >= PLAUS_MIN && d <= max) ? d : null;
  }
  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }); }
  function findSheet(wb, patterns) {
    const name = wb.SheetNames.find(n => patterns.some(p => n.toLowerCase().includes(p)));
    return name ? wb.Sheets[name] : null;
  }

  /* ---------------- classification of free-text status values ---------------- */
  function classifyPhysicalText(raw) {
    const low = String(raw).trim().toLowerCase();
    if (/non[\s-]*primary/.test(low)) return { kind: "secondary" };
    if (/helicopter/.test(low) || /^air\s*one/.test(low))
      return { kind: "review", reason: "Aviation physical exemption (their own, separate cycle)", suggested: "Secondary" };
    if (/non[\s-]*driving[\s-]*medical/.test(low))
      return { kind: "review", reason: "Possibly a temporary medical restriction, not a permanent designation" };
    if (/non[\s-]*driv(er|ing)\b/.test(low) || /\bnd\b.*licen[sc]e/.test(low)) return { kind: "nondriver" };
    if (/^expired\b/.test(low)) return { kind: "review", reason: "Stale entry — treat as unreliable" };
    return { kind: "review", reason: 'Unrecognized text: "' + String(raw).trim() + '"' };
  }
  function classifyDDText(raw) {
    const low = String(raw).trim().toLowerCase();
    if (/not\s*needed/.test(low)) return { kind: "notneeded" };
    if (/^needs?\b/.test(low) || /academy/.test(low)) return { kind: "skip" };
    return { kind: "review", reason: 'Unrecognized note: "' + String(raw).trim() + '"' };
  }
  function noteWorthShowing(v) {
    if (v == null) return false;
    const s = String(v).trim();
    return s !== "" && s !== "." && s.length > 1;
  }

  /* ---------------- parsers ---------------- */
  function parseStatusRoster(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = rowsOf(ws);
    const header = (rows[0] || []).map(h => String(h).trim().toLowerCase());
    const ci = name => header.indexOf(name);
    const iLast = ci("last name"), iFirst = ci("first name"), iEmp = ci("emp#"),
          iPhys = ci("driver physical date"), iDD = ci("defensive driving date"), iFuture = ci("future verified");
    if (iEmp < 0) throw new Error("Couldn't find an 'Emp#' column — is this the Safety Team Main spreadsheet?");
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i]; if (!row || !row[iEmp]) continue;
      out.push({
        employeeId: String(row[iEmp]).trim(),
        name: [row[iLast], row[iFirst]].filter(Boolean).join(", "),
        physicalRaw: row[iPhys], ddRaw: row[iDD],
        futureNote: iFuture >= 0 ? row[iFuture] : null,
      });
    }
    return out;
  }

  function parseHistoryWorkbook(wb) {
    const master = findSheet(wb, ["master"]);
    const addToMain = findSheet(wb, ["add to main", "add-to-main"]);
    const newHires = findSheet(wb, ["new hire"]);
    const pending = findSheet(wb, ["pending"]);
    const examLog = [];

    function pullLog(ws, source) {
      if (!ws) return;
      rowsOf(ws).forEach(row => {
        const emp = row[2]; if (!emp) return;
        examLog.push({ employeeId: String(emp).trim(), testDate: row[0], result: String(row[3] || "").trim(), expirationDate: row[4], source });
      });
    }
    pullLog(master, "Master List");
    pullLog(addToMain, "Add to Main Spreadsheet");

    if (newHires) {
      rowsOf(newHires).forEach(row => {
        const last = String(row[1] || "").trim(), first = String(row[2] || "").trim();
        if (!last && !first) return;
        examLog.push({ employeeId: null, nameHint: { last, first }, testDate: row[0], result: String(row[3] || "").trim(), expirationDate: row[4], source: "New Hires" });
      });
    }

    const pendingRows = [];
    if (pending) {
      rowsOf(pending).forEach(row => {
        const emp = row[2]; if (!emp) return;
        pendingRows.push({ employeeId: String(emp).trim(), date: row[0], note: String(row[4] || "").trim() });
      });
    }
    const foundSheets = [["Master List", master], ["Add to Main Spreadsheet", addToMain], ["New Hires", newHires], ["Pending", pending]]
      .filter(([, ws]) => ws).map(([n]) => n);
    return { examLog, pendingRows, foundSheets };
  }

  /* ---------------- name matching (New Hires has no Employee ID) ---------------- */
  function buildNameIndex(cache) {
    const idx = {};
    cache.roster.forEach(r => {
      const last = String(r.LastName || "").trim().toLowerCase();
      if (!last) return;
      (idx[last] = idx[last] || []).push(r);
    });
    return idx;
  }
  function matchByName(nameIdx, last, first) {
    const cands = nameIdx[String(last || "").trim().toLowerCase()] || [];
    if (cands.length === 1) return cands[0];
    if (cands.length > 1 && first) {
      const fl = String(first).trim().toLowerCase();
      const narrowed = cands.filter(c => String(c.Title || "").toLowerCase().includes(fl));
      if (narrowed.length === 1) return narrowed[0];
    }
    return null;
  }

  /* ---------------- reconciliation ---------------- */
  function reconcile(cache, statusRows, history, pendingRows) {
    const nameIdx = buildNameIndex(cache);
    const historyByEmp = {}; let unmatchedHistory = 0;
    history.forEach(h => {
      let emp = h.employeeId;
      if (!emp && h.nameHint) {
        const m = matchByName(nameIdx, h.nameHint.last, h.nameHint.first);
        if (m) emp = String(m.EmployeeId).trim();
      }
      if (!emp) { unmatchedHistory++; return; }
      (historyByEmp[emp] = historyByEmp[emp] || []).push(h);
    });
    const pendingByEmp = {};
    pendingRows.forEach(p => (pendingByEmp[p.employeeId] = pendingByEmp[p.employeeId] || []).push(p));

    const plan = [];
    statusRows.forEach(sr => {
      const roster = cache.idx.rosterByEmp[sr.employeeId];
      if (!roster) { plan.push({ employeeId: sr.employeeId, name: sr.name, notOnRoster: true }); return; }

      const entry = { employeeId: sr.employeeId, name: roster.Title || sr.name, flags: [] };
      entry.currentDesignation = DS.util.designation(roster);
      entry.designationLocked = entry.currentDesignation !== "Primary";
      entry.rawDesignation = String(roster.DriverStatus || "").trim();

      // designation + physical due date
      const pDate = plausibleDate(sr.physicalRaw);
      let suggested = null;
      if (pDate) {
        suggested = "Primary";
        entry.physicalDueDate = DS.isoDate(pDate);
      } else if (sr.physicalRaw != null && String(sr.physicalRaw).trim() !== "") {
        const c = classifyPhysicalText(sr.physicalRaw);
        if (c.kind === "secondary") suggested = "Secondary";
        else if (c.kind === "nondriver") suggested = "Non-Driver";
        else entry.flags.push({ type: "designation", raw: sr.physicalRaw, reason: c.reason, suggested: c.suggested || null });
      }
      entry.suggestedDesignation = suggested;

      // full exam history — every exam becomes its own record (employee log)
      entry.exams = examsFor(historyByEmp[sr.employeeId], pendingByEmp[sr.employeeId]);

      // defensive driving due date
      const dDate = plausibleDate(sr.ddRaw);
      if (dDate) {
        entry.courseDueDate = DS.isoDate(dDate);
      } else if (sr.ddRaw != null && String(sr.ddRaw).trim() !== "") {
        const c = classifyDDText(sr.ddRaw);
        if (c.kind === "review") entry.flags.push({ type: "courses", raw: sr.ddRaw, reason: c.reason });
        if (c.kind === "notneeded" && suggested === "Primary")
          entry.flags.push({ type: "conflict", raw: sr.ddRaw, reason: 'Physical column suggests Primary, but this says "Not needed" — please confirm' });
      }

      if (noteWorthShowing(sr.futureNote)) entry.futureNote = String(sr.futureNote).trim();

      plan.push(entry);
    });

    // Employees with exam history who aren't in Safety Team Main: log their
    // history too (no designation change for them).
    const inStatus = new Set(statusRows.map(sr => sr.employeeId));
    Object.keys(historyByEmp).concat(Object.keys(pendingByEmp)).forEach(emp => {
      if (inStatus.has(emp)) return;
      inStatus.add(emp);
      const roster = cache.idx.rosterByEmp[emp];
      if (!roster) return;
      plan.push({ employeeId: emp, name: roster.Title || emp, flags: [], historyOnly: true,
        currentDesignation: DS.util.designation(roster), rawDesignation: String(roster.DriverStatus || "").trim(),
        designationLocked: true, suggestedDesignation: null, exams: examsFor(historyByEmp[emp], pendingByEmp[emp]) });
    });

    return { plan, unmatchedHistory };
  }

  function examsFor(hist, pending) {
    const out = [];
    (hist || []).forEach(h => {
      const t = plausibleDate(h.testDate);
      if (!t) return;
      const e = plausibleDate(h.expirationDate);
      out.push({ testDate: DS.isoDate(t), expirationDate: e ? DS.isoDate(e) : null, result: h.result || "" });
    });
    (pending || []).forEach(p => {
      const t = plausibleDate(p.date);
      if (t) out.push({ testDate: DS.isoDate(t), expirationDate: null, result: "Pending" });
    });
    return out;
  }

  /* ---------------- commit: add history only, never overwrite ----------------
     Designation: set from Safety Team Main only when still unassigned/default.
     Physicals: every exam from the reference workbook → its own record
       ("Legacy History"); a real exam competes by date like any upload.
     Master sheet due dates → separate "Master Sheet" records that count only
       when no real record exists (see data.js). Courses get an estimated
       completion date (due date minus the renewal cycle) for each required title.
     Every record is checked against what's already on file, so re-running adds
     only what's missing and never duplicates or edits existing history. */
  async function commitPlan(plan, resolutions, onProgress) {
    const cache = await DS.data.load(true);
    const renewalYears = cache.idx.courseRenewalYears;
    const requiredTitles = cache.idx.requiredTitles;

    const physKeys = new Set(), masterPhysKeys = new Set(), courseKeys = new Set();
    (await DS.spGet(L.physicals, { select: ["EmployeeId", "PhysicalDate", "ExpirationDate", "Source"] })).forEach(p => {
      const emp = String(p.EmployeeId || "").trim();
      if (p.PhysicalDate) physKeys.add(emp + "|" + DS.isoDate(p.PhysicalDate));
      if (DS.util.isFallbackSource(p) && p.ExpirationDate) masterPhysKeys.add(emp + "|" + DS.isoDate(p.ExpirationDate));
    });
    (await DS.spGet(L.courses, { select: ["EmployeeId", "CourseTitle", "DateCompleted"] })).forEach(c => {
      courseKeys.add(String(c.EmployeeId || "").trim() + "|" + String(c.CourseTitle || "").trim() + "|" + DS.isoDate(c.DateCompleted));
    });

    const ops = [];
    let skipped = 0;
    plan.forEach(entry => {
      if (entry.notOnRoster) return;
      const emp = entry.employeeId;
      const roster = cache.idx.rosterByEmp[emp];
      const name = (roster && roster.Title) || entry.name || "";
      const res = resolutions[emp];
      const desigToApply = entry.designationLocked ? null : (res && res.designation !== undefined ? res.designation : entry.suggestedDesignation);
      if (desigToApply && desigToApply !== entry.rawDesignation)
        ops.push({ list: L.roster, kind: "update", id: roster.Id, fields: { DriverStatus: desigToApply } });

      (entry.exams || []).forEach(x => {
        const k = emp + "|" + x.testDate;
        if (physKeys.has(k)) { skipped++; return; }
        physKeys.add(k);
        const f = { Title: name, EmployeeId: emp, PhysicalDate: x.testDate, Result: x.result, Source: "Legacy History" };
        if (x.expirationDate) f.ExpirationDate = x.expirationDate;
        ops.push({ list: L.physicals, kind: "create", fields: f });
      });

      if (entry.physicalDueDate) {
        const k = emp + "|" + entry.physicalDueDate;
        if (masterPhysKeys.has(k)) skipped++;
        else {
          masterPhysKeys.add(k);
          ops.push({ list: L.physicals, kind: "create", fields: { Title: name, EmployeeId: emp, ExpirationDate: entry.physicalDueDate, Source: "Master Sheet" } });
        }
      }

      if (entry.courseDueDate) {
        const est = DS.isoDate(DS.util.addYears(DS.parseDate(entry.courseDueDate), -renewalYears));
        requiredTitles.forEach(title => {
          const k = emp + "|" + title + "|" + est;
          if (courseKeys.has(k)) { skipped++; return; }
          courseKeys.add(k);
          ops.push({ list: L.courses, kind: "create", fields: { Title: name, EmployeeId: emp, CourseTitle: title,
            DateCompleted: est, CompletionStatus: "Passed", Source: "Master Sheet" } });
        });
      }
    });

    let rosterN = 0, physN = 0, courseN = 0;
    ops.forEach(o => { if (o.list === L.roster) rosterN++; else if (o.list === L.physicals) physN++; else courseN++; });

    const errors = await DS.runBatched(ops, async op => {
      if (op.kind === "update") await DS.spUpdate(op.list, op.id, op.fields);
      else await DS.spCreate(op.list, op.fields);
    }, onProgress || (() => {}));

    await DS.audit("Legacy migration committed", null, null,
      rosterN + " designations set, " + physN + " physical records, " + courseN + " course records added, " +
      skipped + " already on file, " + errors.length + " failed");
    DS.data.clear();
    return { rosterN, physN, courseN, skipped, errors };
  }

  /* ---- what each upload's real column headers look like, shown on hover ---- */
  const STATUS_FORMAT = {
    source: "Safety Team Main spreadsheet — current status roster",
    columns: ["Last Name", "First Name", "Emp#", "Badge", "Rank", "Driver Physical Date", "Defensive Driving Date", "Future verified"],
    sample: ["Smith", "Jordan", "123456", "4821", "Police Officer", "3/1/2027", "Non-Primary", ""],
    note: 'The physical/DD columns hold either a due date or a status note (e.g. "Non-Primary", "Non-Driver").',
  };
  const HISTORY_FORMAT = {
    source: "Reference physicals workbook — sheet names matched loosely",
    sections: [
      { label: "Master List / Add to Main Spreadsheet", columns: ["Test Date", "Last, First", "Employee ID", "Result", "Expiration Date"], sample: ["4/5/2022", "Smith, Jordan", "123456", "Pass", "4/4/2024"] },
      { label: "New Hires", columns: ["Test Date", "Last Name", "First Name", "Result", "Expiration Date"], sample: ["5/1/2025", "Smith", "Jordan", "Pass", "5/1/2027"] },
      { label: "Pending", columns: ["Date", "First Last", "Employee ID", "Status", "Note"], sample: ["1/25/2024", "Jordan Smith", "123456", "Pending", ""] },
    ],
    note: "No header row in these sheets — data starts on row 1. New Hires has no Employee ID, so those rows match by name.",
  };

  /* ---------------- screen ---------------- */
  let state = { statusRows: null, history: null, pendingRows: null, foundSheets: [], plan: null, unmatchedHistory: 0, resolutions: {} };

  async function renderMigrate(container) {
    container.innerHTML = "";
    container.appendChild(el("div", { class: "import-note warn", text:
      "One-time legacy migration tool. Not linked in the sidebar — reached only by this direct URL. Files are parsed entirely in this browser; nothing is uploaded anywhere but SharePoint." }));

    const dropRow = el("div", { class: "migrate-drop-row" });
    dropRow.appendChild(buildDrop("Safety Team Main spreadsheet", "Current status — Last Name, First Name, Emp#, Driver Physical Date, Defensive Driving Date", async (wb) => {
      state.statusRows = parseStatusRoster(wb);
      redraw();
    }, STATUS_FORMAT));
    dropRow.appendChild(buildDrop("Reference physicals workbook", "Master List / Add to Main Spreadsheet / New Hires / Pending — optional but recommended", async (wb) => {
      const h = parseHistoryWorkbook(wb);
      state.history = h.examLog; state.pendingRows = h.pendingRows; state.foundSheets = h.foundSheets;
      redraw();
    }, HISTORY_FORMAT));
    container.appendChild(dropRow);

    const resultWrap = el("div", { id: "migrateResult" });
    container.appendChild(resultWrap);
    container.appendChild(renderMaintenance());

    function redraw() { renderBody(resultWrap, container); }
    redraw();
  }

  /* ---------------- Maintenance: remove one source's records ----------------
     For redoing a bad import cleanly (e.g. the first course upload, whose dates
     were affected by the time-zone bug). Removes ONLY records tagged with the
     chosen source, in the chosen list. */
  function renderMaintenance() {
    const listSel = el("select", { class: "field", style: "max-width:220px" }, [
      el("option", { value: L.courses, text: "Courses" }),
      el("option", { value: L.physicals, text: "Physicals" }),
    ]);
    const srcSel = el("select", { class: "field", style: "max-width:260px" }, [
      el("option", { value: "Bulk Upload", text: "Bulk Upload (Imports screen)" }),
      el("option", { value: "Legacy Migration", text: "Legacy Migration (older migration runs)" }),
      el("option", { value: "Master Sheet", text: "Master Sheet (this tool)" }),
      el("option", { value: "Legacy History", text: "Legacy History (this tool)" }),
    ]);
    const countBtn = el("button", { class: "btn btn--ghost", type: "button", text: "Count records" });
    const status = el("div", { class: "help", style: "margin-top:10px" });
    const actWrap = el("div", { style: "margin-top:10px" });
    countBtn.addEventListener("click", async () => {
      actWrap.innerHTML = ""; status.textContent = "Counting…";
      try {
        const rows = (await DS.spGet(listSel.value, { select: ["Id", "Source"] })).filter(r => r.Source === srcSel.value);
        const listLabel = listSel.options[listSel.selectedIndex].text;
        status.textContent = rows.length + " " + listLabel + " record(s) tagged \u201C" + srcSel.value + "\u201D.";
        if (!rows.length) return;
        const del = el("button", { class: "btn", type: "button", text: "Remove these " + rows.length + " records" });
        del.addEventListener("click", async () => {
          if (!confirm("Permanently remove " + rows.length + " " + listLabel + " records tagged \u201C" + srcSel.value + "\u201D? Other records are not affected.")) return;
          del.disabled = true; countBtn.disabled = true;
          const errors = await DS.runBatched(rows, r => DS.spDelete(listSel.value, r.Id), (d, t, label) => {
            status.textContent = label || ("Removing " + d + " of " + t + "\u2026");
          });
          await DS.audit("Records removed by source", listSel.value, null, (rows.length - errors.length) + " \u201C" + srcSel.value + "\u201D records removed");
          DS.data.clear();
          status.textContent = (rows.length - errors.length) + " removed." + (errors.length ? " " + errors.length + " failed \u2014 see below." : "");
          actWrap.innerHTML = ""; countBtn.disabled = false;
          if (errors.length) actWrap.appendChild(DS.errorSummaryEl(errors));
        });
        actWrap.appendChild(del);
      } catch (e) { status.textContent = "Couldn't count: " + e.message; }
    });
    return el("div", { class: "card", style: "margin-top:28px" }, [
      el("div", { class: "card__head" }, el("h3", { text: "Maintenance \u2014 remove records from one source" })),
      el("div", { class: "card__body" }, [
        el("div", { class: "import-note warn", text: "For redoing an import cleanly. Only records with the chosen source tag are removed; everything else stays. Count first, then confirm." }),
        el("div", { style: "display:flex; gap:10px; flex-wrap:wrap; align-items:center" }, [listSel, srcSel, countBtn]),
        status, actWrap,
      ]),
    ]);
  }

  function buildDrop(title, sub, onFile, formatSpec) {
    const h3 = el("h3", { text: title });
    if (formatSpec) h3.appendChild(DS.formatHint(formatSpec));
    const dz = el("div", { class: "dropzone" }, [
      el("div", { class: "dz-ico", text: "⬆" }),
      h3,
      el("p", { text: sub }),
    ]);
    const input = el("input", { type: "file", accept: ".xlsx,.xls", style: "display:none" });
    dz.appendChild(input);
    async function handle(file) {
      dz.classList.add("done");
      dz.querySelector("h3").textContent = title + " — loaded (" + file.name + ")";
      const XLSXlib = await DS.ensureXlsx();
      const buf = await file.arrayBuffer();
      const wb = XLSXlib.read(buf, { type: "array", cellDates: true });
      try { await onFile(wb); } catch (e) { DS.toast(e.message, "error"); }
    }
    dz.addEventListener("click", () => input.click());
    dz.addEventListener("dragover", e => { e.preventDefault(); dz.classList.add("drag"); });
    dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
    dz.addEventListener("drop", e => { e.preventDefault(); dz.classList.remove("drag"); if (e.dataTransfer.files[0]) handle(e.dataTransfer.files[0]); });
    input.addEventListener("change", () => { if (input.files[0]) handle(input.files[0]); });
    return dz;
  }

  async function renderBody(wrap, container) {
    wrap.innerHTML = "";
    if (!state.statusRows) { wrap.appendChild(el("div", { class: "empty-mini", text: "Drop the Safety Team Main spreadsheet to begin." })); return; }

    DS.showLoading(wrap, "Reconciling against the live roster…");
    const cache = await DS.data.load();
    const { plan, unmatchedHistory } = reconcile(cache, state.statusRows, state.history || [], state.pendingRows || []);
    state.plan = plan; state.unmatchedHistory = unmatchedHistory;
    wrap.innerHTML = "";

    const onRoster = plan.filter(p => !p.notOnRoster && !p.historyOnly);
    const historyOnly = plan.filter(p => p.historyOnly);
    const notOnRoster = plan.filter(p => p.notOnRoster);
    const locked = onRoster.filter(p => p.designationLocked);
    const flagged = onRoster.filter(p => !p.designationLocked && p.flags.length);
    const clean = onRoster.filter(p => !p.designationLocked && !p.flags.length && p.suggestedDesignation);

    const needStat = stat(flagged.length, "need your input", flagged.length ? "overdue" : "clear");
    wrap.appendChild(el("div", { class: "migrate-summary" }, [
      stat(onRoster.length, "on current roster"),
      stat(clean.length, "clean — ready to apply"),
      needStat,
      stat(locked.length, "already set — untouched"),
      stat(notOnRoster.length, "not on current roster", notOnRoster.length ? "due" : "clear"),
    ]));

    if (state.history) {
      wrap.appendChild(el("div", { class: "import-note", text:
        "History workbook sheets found: " + (state.foundSheets.join(", ") || "none recognized") +
        ". " + unmatchedHistory + " history row(s) couldn't be matched to an employee (New Hires rows are matched by name, not ID)." +
        (historyOnly.length ? " " + historyOnly.length + " more employee(s) aren't in Safety Team Main but have exam history — their history will be logged too." : "") }));
    } else {
      wrap.appendChild(el("div", { class: "import-note warn", text:
        "No history workbook loaded — physical exam history and Pending results won't be included. Designations and due dates from Safety Team Main will still be applied." }));
    }

    // Decisions update counts in place — no full re-render, so a click is
    // visibly kept (picked button + green row) and the page doesn't jump.
    let refreshCounts = () => {};
    let flagHead = null;
    if (flagged.length) {
      const card = el("div", { class: "card", style: "margin-bottom:18px" });
      flagHead = el("h3", { text: "Needs your input (" + flagged.length + ")" });
      card.appendChild(el("div", { class: "card__head" }, [
        flagHead,
        el("span", { class: "help", text: "Dashed border = suggested. Click to decide." }),
      ]));
      const body = el("div", null);
      flagged.forEach(entry => body.appendChild(renderFlagRow(entry, () => refreshCounts())));
      card.appendChild(body);
      wrap.appendChild(card);
    }

    if (notOnRoster.length) {
      wrap.appendChild(el("div", { class: "import-note warn", text:
        notOnRoster.length + " employee(s) in Safety Team Main aren't on the current roster — likely former employees. They won't be written; nothing to do." }));
    }

    const undecidedNote = el("div", { class: "help", style: "margin-top:8px" });
    const commitBtn = el("button", { class: "btn" });
    refreshCounts = () => {
      const decided = flagged.filter(e => state.resolutions[e.employeeId]).length;
      const remaining = flagged.length - decided;
      const applying = Object.values(state.resolutions).filter(r => r && r.designation).length;
      needStat.firstChild.textContent = String(remaining);
      needStat.className = "stat stat--" + (remaining ? "overdue" : "clear");
      if (flagHead) flagHead.textContent = "Needs your input (" + remaining + " of " + flagged.length + " remaining)";
      if (!commitBtn.disabled) commitBtn.textContent = "Apply to SharePoint (" + (clean.length + applying) + " employees)";
      undecidedNote.textContent = remaining
        ? remaining + " still undecided — you can apply now; they'll be left as they are and can be decided on a later run."
        : "";
    };
    refreshCounts();
    commitBtn.addEventListener("click", async () => {
      commitBtn.disabled = true; commitBtn.textContent = "Applying…";
      // live progress — this run can be several thousand writes and take a while
      const bar = el("div", { class: "progress" }, el("div", { class: "progress__bar" }));
      const barFill = bar.firstChild;
      const status = el("div", { style: "font-size:13px; color:var(--slate)", text: "Preparing…" });
      const progCard = el("div", { class: "card", style: "margin-top:14px" }, el("div", { class: "card__body" }, [
        el("h3", { text: "Applying migration…", style: "font-size:15px; margin-bottom:6px" }),
        el("div", { class: "help", text: "Large runs can take a while. Keep this tab open and in front until it finishes." }),
        bar, status,
      ]));
      wrap.appendChild(progCard);
      progCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
      try {
        const result = await commitPlan(plan, state.resolutions, (done, total, label) => {
          barFill.style.width = (total ? done / total * 100 : 100) + "%";
          status.textContent = label || ("Writing " + done + " of " + total + "…");
        });
        wrap.innerHTML = "";
        const body = el("div", { class: "card__body" }, el("div", { class: "import-note" + (result.errors.length ? " warn" : ""), text:
          "Planned: " + result.rosterN + " designation(s), " + result.physN + " physical record(s), " + result.courseN + " course record(s) added to history; " +
          result.skipped + " already on file (skipped). " +
          (result.errors.length
            ? result.errors.length + " operation(s) failed — grouped below. Re-running is safe; it only fills in what's missing."
            : "All written. Safe to re-run later as more files come in — already-set designations won't be touched.") }));
        if (result.errors.length) body.appendChild(DS.errorSummaryEl(result.errors));
        wrap.appendChild(el("div", { class: "card" }, body));
        DS.toast("Migration applied.", result.errors.length ? "error" : "success");
      } catch (e) {
        progCard.remove();
        commitBtn.disabled = false; refreshCounts();
        DS.toast("Migration failed: " + e.message, "error");
      }
    });
    wrap.appendChild(el("div", { style: "margin-top:10px" }, [commitBtn, undecidedNote]));
  }

  function stat(n, label, kind) {
    return el("div", { class: "stat" + (kind ? " stat--" + kind : "") }, [
      el("b", { class: "tnum", text: String(n) }), el("span", { text: label }),
    ]);
  }

  function renderFlagRow(entry, onChange) {
    const flag = entry.flags[0];
    const row = el("div", { class: "flag-row" });
    row.appendChild(el("div", { class: "who" }, [
      el("b", { text: entry.name + " (" + entry.employeeId + ")" }),
      el("span", { text: flag.reason }),
    ]));
    row.appendChild(el("div", { class: "raw", text: String(flag.raw) }));
    const choices = el("div", { class: "choices" });
    const saved = state.resolutions[entry.employeeId];
    const savedLabel = saved ? (saved.designation || "Skip") : null;
    if (savedLabel) row.classList.add("resolved");
    ["Primary", "Secondary", "Non-Driver", "Skip"].forEach(opt => {
      const btn = el("button", { type: "button", text: opt });
      if (opt === savedLabel) btn.classList.add("picked");
      // A suggestion is only a hint (dashed border) — it is NOT a decision until clicked.
      if (flag.suggested === opt) { btn.classList.add("suggested"); btn.title = "Suggested"; }
      btn.addEventListener("click", (ev) => {
        ev.preventDefault(); ev.stopPropagation();
        state.resolutions[entry.employeeId] = { designation: opt === "Skip" ? null : opt };
        choices.querySelectorAll("button").forEach(b => b.classList.remove("picked"));
        btn.classList.add("picked");
        row.classList.add("resolved");
        onChange();
      });
      choices.appendChild(btn);
    });
    row.appendChild(choices);
    return row;
  }

  // Styles for decided rows and suggestion hints — injected here so this fix
  // only requires re-uploading this one file.
  (function injectStyles() {
    if (document.getElementById("migrate-legacy-styles")) return;
    const s = document.createElement("style");
    s.id = "migrate-legacy-styles";
    s.textContent =
      ".flag-row.resolved { background: var(--clear-bg); }" +
      ".flag-row .choices button.suggested { border-style: dashed; border-color: var(--navy-500); }" +
      ".flag-row .choices button.picked { background: var(--navy-700); color: #fff; border-color: var(--navy-700); border-style: solid; }";
    document.head.appendChild(s);
  })();

  DS.registerScreen("migrate", { title: "Legacy Migration", icon: "⚑", render: renderMigrate });
})();
