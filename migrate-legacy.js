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
     - Physicals/Courses records written here are tagged Source="Legacy
       Migration" and upserted against themselves on repeat runs, so
       re-running doesn't create duplicates, and a genuinely newer real
       upload still wins on its own via the app's normal max-date logic.
   ============================================================ */
(function () {
  const el = DS.el;
  const L = DS.LISTS;
  const SRC = "Legacy Migration";
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

      // merged exam history (most recent plausible test date wins)
      const hist = (historyByEmp[sr.employeeId] || []).filter(h => plausibleDate(h.testDate));
      if (hist.length) {
        hist.sort((a, b) => new Date(b.testDate) - new Date(a.testDate));
        const top = hist[0];
        entry.lastExam = { testDate: DS.isoDate(top.testDate), result: top.result, expirationDate: DS.isoDate(top.expirationDate), source: top.source };
      }
      if (pendingByEmp[sr.employeeId]) {
        const p = pendingByEmp[sr.employeeId][0];
        entry.pending = { date: DS.isoDate(p.date), note: p.note };
      }

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

    return { plan, unmatchedHistory };
  }

  /* ---------------- commit (idempotent — safe to re-run) ---------------- */
  async function commitPlan(plan, resolutions) {
    const cache = await DS.data.load(true);
    const renewalYears = cache.idx.courseRenewalYears;
    const requiredTitles = cache.idx.requiredTitles;

    const existingPhys = await DS.spGet(L.physicals, { select: ["Id", "EmployeeId", "Source"] });
    const physByEmp = {}; existingPhys.forEach(p => { if (p.Source === SRC) physByEmp[String(p.EmployeeId).trim()] = p.Id; });
    const existingCourses = await DS.spGet(L.courses, { select: ["Id", "EmployeeId", "CourseTitle", "Source"] });
    const courseByKey = {}; existingCourses.forEach(c => { if (c.Source === SRC) courseByKey[String(c.EmployeeId).trim() + "|" + c.CourseTitle] = c.Id; });

    const ops = [];
    plan.forEach(entry => {
      if (entry.notOnRoster) return;
      const roster = cache.idx.rosterByEmp[entry.employeeId];
      const res = resolutions[entry.employeeId];
      const desigToApply = entry.designationLocked ? null : (res && res.designation !== undefined ? res.designation : entry.suggestedDesignation);
      const finalDesig = desigToApply || entry.currentDesignation;

      if (desigToApply && desigToApply !== entry.currentDesignation)
        ops.push({ list: L.roster, kind: "update", id: roster.Id, fields: { DriverStatus: desigToApply } });

      if (finalDesig !== "Non-Driver" && (entry.physicalDueDate || entry.lastExam || entry.pending)) {
        const fields = { EmployeeId: entry.employeeId, Source: SRC };
        if (entry.lastExam) {
          fields.PhysicalDate = entry.lastExam.testDate;
          fields.Result = entry.lastExam.result || "";
          if (entry.lastExam.expirationDate) fields.ExpirationDate = entry.lastExam.expirationDate;
        }
        if (entry.physicalDueDate) fields.ExpirationDate = entry.physicalDueDate; // Safety Team Main wins
        if (entry.pending) { fields.Result = "Pending"; if (!fields.PhysicalDate) fields.PhysicalDate = entry.pending.date; }
        if (fields.PhysicalDate || fields.ExpirationDate) {
          const existingId = physByEmp[entry.employeeId];
          ops.push(existingId ? { list: L.physicals, kind: "update", id: existingId, fields } : { list: L.physicals, kind: "create", fields });
        }
      }

      if (finalDesig !== "Non-Driver" && entry.courseDueDate) {
        const implied = DS.isoDate(DS.util.addYears(new Date(entry.courseDueDate), -renewalYears));
        requiredTitles.forEach(title => {
          const fields = { EmployeeId: entry.employeeId, CourseTitle: title, DateCompleted: implied, CompletionStatus: "Passed", Source: SRC };
          const key = entry.employeeId + "|" + title;
          const existingId = courseByKey[key];
          ops.push(existingId ? { list: L.courses, kind: "update", id: existingId, fields } : { list: L.courses, kind: "create", fields });
        });
      }
    });

    let rosterN = 0, physN = 0, courseN = 0;
    ops.forEach(o => { if (o.list === L.roster) rosterN++; else if (o.list === L.physicals) physN++; else courseN++; });

    const errors = await DS.runBatched(ops, async op => {
      if (op.kind === "update") await DS.spUpdate(op.list, op.id, op.fields);
      else await DS.spCreate(op.list, op.fields);
    }, () => {});

    await DS.audit("Legacy migration committed", null, null,
      rosterN + " designations set, " + physN + " physical records, " + courseN + " course records, " + errors.length + " failed");
    DS.data.clear();
    return { rosterN, physN, courseN, errors };
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

    function redraw() { renderBody(resultWrap, container); }
    redraw();
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

    const onRoster = plan.filter(p => !p.notOnRoster);
    const notOnRoster = plan.filter(p => p.notOnRoster);
    const locked = onRoster.filter(p => p.designationLocked);
    const flagged = onRoster.filter(p => !p.designationLocked && p.flags.length);
    const clean = onRoster.filter(p => !p.designationLocked && !p.flags.length && p.suggestedDesignation);

    wrap.appendChild(el("div", { class: "migrate-summary" }, [
      stat(onRoster.length, "on current roster"),
      stat(clean.length, "clean — ready to apply"),
      stat(flagged.length, "need your input", flagged.length ? "overdue" : "clear"),
      stat(locked.length, "already set — untouched"),
      stat(notOnRoster.length, "not on current roster", notOnRoster.length ? "due" : "clear"),
    ]));

    if (state.history) {
      wrap.appendChild(el("div", { class: "import-note", text:
        "History workbook sheets found: " + (state.foundSheets.join(", ") || "none recognized") +
        ". " + unmatchedHistory + " history row(s) couldn't be matched to an employee (New Hires rows are matched by name, not ID)." }));
    } else {
      wrap.appendChild(el("div", { class: "import-note warn", text:
        "No history workbook loaded — physical exam history and Pending results won't be included. Designations and due dates from Safety Team Main will still be applied." }));
    }

    if (flagged.length) {
      const card = el("div", { class: "card", style: "margin-bottom:18px" });
      card.appendChild(el("div", { class: "card__head" }, el("h3", { text: "Needs your input (" + flagged.length + ")" })));
      const body = el("div", null);
      flagged.forEach(entry => body.appendChild(renderFlagRow(entry, () => renderBody(wrap, container))));
      card.appendChild(body);
      wrap.appendChild(card);
    }

    if (notOnRoster.length) {
      wrap.appendChild(el("div", { class: "import-note warn", text:
        notOnRoster.length + " employee(s) in Safety Team Main aren't on the current roster — likely former employees. They won't be written; nothing to do." }));
    }

    const commitBtn = el("button", { class: "btn", text: "Apply to SharePoint (" + (clean.length + Object.keys(state.resolutions).length) + " employees)" });
    commitBtn.addEventListener("click", async () => {
      commitBtn.disabled = true; commitBtn.textContent = "Applying…";
      try {
        const result = await commitPlan(plan, state.resolutions);
        wrap.innerHTML = "";
        wrap.appendChild(el("div", { class: "card" }, el("div", { class: "card__body" }, el("div", { class: "import-note", text:
          result.rosterN + " designation(s) set, " + result.physN + " physical record(s), " + result.courseN + " course record(s) written." +
          (result.errors.length ? " " + result.errors.length + " operation(s) failed — check the console." : " Safe to re-run later as more files come in — already-set designations won't be touched.") }))));
        DS.toast("Migration applied.", result.errors.length ? "error" : "success");
      } catch (e) {
        commitBtn.disabled = false; commitBtn.textContent = "Apply to SharePoint";
        DS.toast("Migration failed: " + e.message, "error");
      }
    });
    wrap.appendChild(el("div", { style: "margin-top:10px" }, commitBtn));
  }

  function stat(n, label, kind) {
    return el("div", { class: "stat" + (kind ? " stat--" + kind : "") }, [
      el("b", { class: "tnum", text: String(n) }), el("span", { text: label }),
    ]);
  }

  function renderFlagRow(entry, onResolved) {
    const flag = entry.flags[0];
    const row = el("div", { class: "flag-row" });
    row.appendChild(el("div", { class: "who" }, [
      el("b", { text: entry.name + " (" + entry.employeeId + ")" }),
      el("span", { text: flag.reason }),
    ]));
    row.appendChild(el("div", { class: "raw", text: String(flag.raw) }));
    const choices = el("div", { class: "choices" });
    ["Primary", "Secondary", "Non-Driver", "Skip"].forEach(opt => {
      const btn = el("button", { text: opt });
      if (flag.suggested === opt) btn.classList.add("picked");
      btn.addEventListener("click", () => {
        state.resolutions[entry.employeeId] = { designation: opt === "Skip" ? null : opt };
        choices.querySelectorAll("button").forEach(b => b.classList.remove("picked"));
        btn.classList.add("picked");
        entry.flags = []; // resolved — drop from the queue on next redraw
        onResolved();
      });
      choices.appendChild(btn);
    });
    row.appendChild(choices);
    return row;
  }

  DS.registerScreen("migrate", { title: "Legacy Migration", icon: "⚑", render: renderMigrate });
})();
