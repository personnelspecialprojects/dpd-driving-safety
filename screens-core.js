/* ============================================================
   screens-core.js — Dashboard, Roster, Audit Log
   Loaded after data.js. Each screen self-registers via
   DS.registerScreen(key, { title, icon, render }).
   ============================================================ */
(function () {
  const el = DS.el;

  /* ---- shared: build a <table> from column defs + rows ---- */
  function buildTable(columns, rows) {
    const thead = el("thead", null, el("tr", null,
      columns.map(c => el("th", { class: c.thClass || "" }, c.head))));
    const tbody = el("tbody", null, rows.map(r => {
      const tr = el("tr", null, columns.map(c => {
        const cell = c.render(r);
        return el("td", { class: c.tdClass || "" }, typeof cell === "string" ? cell : [cell]);
      }));
      if (r.__onclick) { tr.className = r.__rowClass || ""; tr.addEventListener("click", r.__onclick); }
      return tr;
    }));
    return el("table", { class: "tbl" }, [thead, tbody]);
  }

  function card(title, headerRight, body) {
    return el("div", { class: "card" }, [
      el("div", { class: "card__head" }, [
        el("h3", { text: title }),
        headerRight || null,
      ]),
      body,
    ]);
  }

  function emptyMini(msg) { return el("div", { class: "empty-mini", text: msg }); }

  function dueBadge(dueDate, overdue) {
    return DS.badge(DS.fmtDate(dueDate), overdue ? "overdue" : "due");
  }

  // dashboard physicals: urgency-aware badge
  function physBadge(row) {
    if (row.urgency === "missing") return DS.badge("Missing", "overdue");
    if (row.urgency === "overdue") return DS.badge("Overdue \u00b7 " + DS.fmtDate(row.dueDate), "overdue");
    return DS.badge(DS.fmtDate(row.dueDate), "due");
  }
  // driver-status cell (Secondary de-emphasized)
  function driverCell(required) {
    return required
      ? el("span", { text: "Primary" })
      : el("span", { style: "color:var(--muted)", text: "Secondary" });
  }
  // roster detail physical row, required-aware
  function physicalDetailBadge(phys) {
    if (!phys.applicable) return DS.badge("Not applicable (Non-Driver)", "neutral");
    if (!phys.has) return phys.required
      ? DS.badge("Required \u2014 none on record", "overdue")
      : DS.badge("Not required \u2014 none on record", "neutral");
    if (phys.overdue) return DS.badge("Overdue \u00b7 " + DS.fmtDate(phys.dueDate), "overdue");
    return DS.badge(DS.fmtDate(phys.dueDate), phys.required ? "due" : "neutral");
  }

  /* ============================================================
     DASHBOARD
     ============================================================ */
  async function renderDashboard(container) {
    const cache = await DS.data.load();
    const physicals = DS.compute.physicalsDue(cache);
    const courses = DS.compute.coursesDue(cache);
    const awards = DS.compute.awardsEligible(cache);
    const restricted = DS.compute.drivingRestricted(cache);

    container.innerHTML = "";

    // stat row
    const criticalPhys = physicals.filter(p => p.urgency === "missing" || p.urgency === "overdue").length;
    const noDriving = restricted.filter(r => r.status === "No-Driving").length;
    container.appendChild(el("div", { class: "stats" }, [
      statCard(physicals.length, "Physicals due or missing", criticalPhys ? "overdue" : "due"),
      statCard(courses.length, "Courses coming due", courses.some(c => c.overdue) ? "overdue" : "due"),
      statCard(awards.length, "Awards eligible now", "clear"),
      statCard(restricted.length, "On restricted / no-driving", noDriving ? "overdue" : "due"),
    ]));

    const grid = el("div", { class: "dash-grid" });
    container.appendChild(grid);

    // Physicals due
    grid.appendChild(card(
      "Physicals due",
      el("span", { class: "count-pill", text: physicals.length + (physicals.length === 1 ? " employee" : " employees") }),
      physicals.length ? buildTable([
        { head: "Employee", render: r => el("span", { class: "strong", text: r.name }) },
        { head: "Driver", render: r => driverCell(r.required) },
        { head: "Status", thClass: "nowrap", tdClass: "nowrap", render: r => physBadge(r) },
      ], rowsWithNav(physicals)) : emptyMini("No physicals due or missing within the alert window.")
    ));

    // Courses due
    grid.appendChild(card(
      "Defensive driving due",
      el("span", { class: "count-pill", text: courses.length + (courses.length === 1 ? " employee" : " employees") }),
      courses.length ? buildTable([
        { head: "Employee", render: r => el("span", { class: "strong", text: r.name }) },
        { head: "Status", render: r => DS.badge(r.status, r.status === "Not started" ? "overdue" : "neutral") },
        { head: "Due", thClass: "nowrap", tdClass: "nowrap", render: r => dueBadge(r.dueDate, r.overdue) },
      ], rowsWithNav(courses)) : emptyMini("No courses due within the alert window.")
    ));

    // Awards eligible
    grid.appendChild(card(
      "Awards eligible",
      el("span", { class: "count-pill", text: awards.length + (awards.length === 1 ? " employee" : " employees") }),
      awards.length ? buildTable(awardColumns(cache, container), rowsWithNav(awards, false))
        : emptyMini("No employees are award-eligible right now.")
    ));

    // Driving status (Restrictive / No-Driving)
    grid.appendChild(card(
      "Driving status",
      el("span", { class: "count-pill", text: restricted.length + (restricted.length === 1 ? " employee" : " employees") }),
      restricted.length ? buildTable([
        { head: "Employee", render: r => el("span", { class: "strong", text: r.name }) },
        { head: "Status", render: r => DS.badge(r.status, r.status === "No-Driving" ? "overdue" : "due") },
        { head: "Active points", thClass: "num", tdClass: "num", render: r => el("span", { class: "tnum", text: String(r.points) }) },
      ], rowsWithNav(restricted)) : emptyMini("No employees are on restricted or no-driving status.")
    ));
  }

  function statCard(n, label, kind) {
    return el("div", { class: "stat stat--" + kind }, [
      el("b", { class: "tnum", text: String(n) }),
      el("span", { text: label }),
    ]);
  }

  // make rows clickable → jump to roster + select (via hash param)
  function rowsWithNav(rows, clickable) {
    if (clickable === false) return rows;
    return rows.map(r => Object.assign({}, r, {
      __rowClass: "roster-row",
      __onclick: () => DS.navigate("roster/" + encodeURIComponent(r.employeeId)),
    }));
  }

  function awardColumns(cache, container) {
    const cols = [
      { head: "Employee", render: r => el("span", { class: "strong", text: r.name }) },
      { head: "Milestone", render: r => DS.badge(r.nextMilestone + "-year", "clear") },
      { head: "Eligible since", thClass: "nowrap", tdClass: "nowrap", render: r => el("span", { class: "tnum", text: DS.fmtDate(r.eligibleDate) }) },
    ];
    if (cache.idx.allowMark) {
      cols.push({
        head: "", tdClass: "nowrap", render: r => {
          const btn = el("button", { class: "btn btn--sm btn--ghost", text: "Mark awarded" });
          btn.addEventListener("click", async () => {
            btn.disabled = true; btn.textContent = "Saving…";
            try {
              await DS.spCreate(DS.LISTS.awards, {
                Title: r.name,
                EmployeeId: r.employeeId,
                AwardDate: new Date().toISOString(),
                MilestoneYears: r.nextMilestone,
                IssuedBy: (DS.me && (DS.me.mail || DS.me.userPrincipalName)) || "",
              });
              await DS.audit("Award marked", DS.LISTS.awards, r.employeeId,
                r.name + " — " + r.nextMilestone + "-year award");
              DS.toast(r.name + " marked as awarded (" + r.nextMilestone + "-year).", "success");
              DS.data.clear();
              renderDashboard(container);
            } catch (e) {
              btn.disabled = false; btn.textContent = "Mark awarded";
              DS.toast("Couldn't save the award: " + e.message, "error");
            }
          });
          return btn;
        }
      });
    }
    return cols;
  }

  /* ============================================================
     ROSTER  (route can carry a selected employee: #/roster/123456)
     ============================================================ */
  let rosterState = { search: "", selected: null };

  async function renderRoster(container) {
    const cache = await DS.data.load();
    const preselect = routeParam();
    if (preselect) rosterState.selected = preselect;

    container.innerHTML = "";

    const search = el("input", {
      class: "field", type: "search",
      placeholder: "Search by name or employee ID", value: rosterState.search,
    });
    container.appendChild(el("div", { class: "toolbar" }, [
      search,
      el("span", { class: "count-pill", id: "rosterCount" }),
    ]));

    const split = el("div", { class: "split" });
    const listWrap = el("div", { class: "card" });
    const detailWrap = el("div", { class: "card detail" });
    split.appendChild(listWrap);
    split.appendChild(detailWrap);
    container.appendChild(split);

    function paint() {
      const q = rosterState.search.trim().toLowerCase();
      let rows = cache.idx.activeRoster.filter(r => {
        if (!q) return true;
        return String(r.Title || "").toLowerCase().includes(q)
          || String(r.LastName || "").toLowerCase().includes(q)
          || String(r.EmployeeId || "").toLowerCase().includes(q);
      });
      rows.sort((a, b) => String(a.LastName || a.Title).localeCompare(String(b.LastName || b.Title)));

      document.getElementById("rosterCount").textContent =
        rows.length + (rows.length === 1 ? " employee" : " employees");

      listWrap.innerHTML = "";
      if (!rows.length) { listWrap.appendChild(emptyMini("No matching employees.")); }
      else {
        const tableRows = rows.slice(0, 400).map(r => Object.assign({}, r, {
          __rowClass: "roster-row" + (DS.util.empKey(r.EmployeeId) === rosterState.selected ? " selected" : ""),
          __onclick: () => { rosterState.selected = DS.util.empKey(r.EmployeeId); paint(); paintDetail(); },
        }));
        listWrap.appendChild(buildTable([
          { head: "Name", render: r => el("span", { class: "strong", text: r.Title || "—" }) },
          { head: "ID", tdClass: "nowrap num", thClass: "num", render: r => el("span", { class: "tnum", text: r.EmployeeId || "—" }) },
          { head: "Designation", render: r => DS.badge(DS.util.designation(r), "neutral") },
          { head: "Driving", render: r => {
              const d = DS.compute.drivingStatusFor(cache, DS.util.empKey(r.EmployeeId));
              if (d.status === "No-Driving") return DS.badge("No-Driving", "overdue");
              if (d.status === "Restrictive") return DS.badge("Restrictive", "due");
              return el("span", { style: "color:var(--muted)", text: "—" });
            } },
        ], tableRows));
        if (rows.length > 400) listWrap.appendChild(emptyMini("Showing first 400 — narrow the search to see the rest."));
      }
    }

    function paintDetail() {
      detailWrap.innerHTML = "";
      const emp = rosterState.selected;
      const r = emp ? cache.idx.rosterByEmp[emp] : null;
      if (!r) {
        detailWrap.appendChild(el("div", { class: "detail__empty", text: "Select an employee to see compliance detail." }));
        return;
      }
      const phys = DS.compute.physicalFor(cache, emp);
      const crs = DS.compute.coursesFor(cache, emp);
      const awd = DS.compute.awardFor(cache, emp);
      const drive = DS.compute.drivingStatusFor(cache, emp);
      const today = DS.util.startOfToday();

      const body = el("div", { class: "card__body" });
      body.appendChild(el("div", { class: "detail__name", text: r.Title || "—" }));
      body.appendChild(el("div", { class: "detail__sub", text: [r.EmployeeId, r.Rank].filter(Boolean).join(" · ") || "—" }));

      body.appendChild(detailRow("Division", r.Division || "—"));
      body.appendChild(detailRow("Assignment", r.Assignment || "—"));
      if (r.Supervisor) body.appendChild(detailRow("Supervisor", r.Supervisor));

      const driveKind = drive.status === "No-Driving" ? "overdue" : drive.status === "Restrictive" ? "due" : "neutral";
      const driveText = drive.status + (drive.override
        ? " \u00b7 manual" + (drive.overrideUntil ? " until " + DS.fmtDate(drive.overrideUntil) : "")
        : " \u00b7 " + drive.points + " pts");
      body.appendChild(detailRow("Driving eligibility", DS.badge(driveText, driveKind)));
      if (drive.override) {
        body.appendChild(detailRow("Points (computed)", drive.points + " active \u2192 " + drive.computed));
        if (drive.overrideNote) body.appendChild(detailRow("Override note", drive.overrideNote));
      }

      body.appendChild(detailRow("Physical", physicalDetailBadge(phys)));
      body.appendChild(detailRow("Courses",
        !crs.applicable ? DS.badge("Not applicable (Non-Driver)", "neutral")
          : (crs.dueDate && crs.dueDate < today ? DS.badge(crs.status, "overdue") : DS.badge(crs.status, "neutral"))));
      body.appendChild(detailRow("Next award",
        awd.eligible ? DS.badge(awd.nextMilestone + "-year (eligible)", "clear")
          : el("span", { class: "tnum", text: awd.nextMilestone + "-year · " + DS.fmtDate(awd.eligibleDate) })));

      // driver designation editor
      const editWrap = el("div", { class: "detail__edit" });
      editWrap.appendChild(el("span", { class: "label", text: "Driver designation" }));
      const sel = el("select", { class: "field" }, [
        el("option", { value: "Primary", text: "Primary" }),
        el("option", { value: "Secondary", text: "Secondary" }),
        el("option", { value: "Non-Driver", text: "Non-Driver" }),
      ]);
      sel.value = DS.util.designation(r);
      const saveBtn = el("button", { class: "btn btn--sm", text: "Save" });
      saveBtn.addEventListener("click", async () => {
        const val = sel.value;
        const cur = DS.util.designation(r);
        if (val === cur) { DS.toast("No change to save."); return; }
        saveBtn.disabled = true; saveBtn.textContent = "Saving…";
        try {
          await DS.spUpdate(DS.LISTS.roster, r.Id, { DriverStatus: val });
          await DS.audit("Designation changed", DS.LISTS.roster, r.EmployeeId,
            r.Title + ": " + cur + " → " + val);
          r.DriverStatus = val;                 // update in-memory cache
          DS.toast(r.Title + " set to " + val + ".", "success");
          paint(); paintDetail();
        } catch (e) {
          DS.toast("Couldn't update designation: " + e.message + " — the DriverStatus choice list may need a 'Non-Driver' option.", "error");
        } finally {
          saveBtn.disabled = false; saveBtn.textContent = "Save";
        }
      });
      editWrap.appendChild(el("div", { class: "row" }, [sel, saveBtn]));
      body.appendChild(editWrap);
      body.appendChild(recordsSection(r, emp));

      detailWrap.appendChild(body);
    }

    /* ---------- Records: manual add / correct, per employee ---------- */
    async function reload() { DS.data.clear(); await renderRoster(container); }
    const byDateDesc = f => (a, b) => (DS.parseDate(b[f]) || 0) - (DS.parseDate(a[f]) || 0);

    function recordsSection(r, emp) {
      const wrap = el("div", { class: "detail__edit" });
      wrap.appendChild(el("span", { class: "label", text: "Records" }));
      wrap.appendChild(el("div", { class: "rec-actions" }, [
        smallBtn("+ Course", () => openCourseForm(r)),
        smallBtn("+ Physical", () => openPhysicalForm(r)),
        smallBtn("+ Accident", () => openAccidentForm(r)),
        smallBtn("Driving override", () => openOverrideForm(r)),
      ]));

      const courses = (cache.idx.coursesByEmp[emp] || []).slice().sort(byDateDesc("DateCompleted"));
      wrap.appendChild(recList("Courses", courses, DS.LISTS.courses, c =>
        [c.CourseTitle || "—", DS.fmtDate(c.DateCompleted)]));

      const phys = (cache.idx.physicalsByEmp[emp] || []).slice().sort(byDateDesc("PhysicalDate"));
      wrap.appendChild(recList("Physicals", phys, DS.LISTS.physicals, p =>
        ["Tested " + DS.fmtDate(p.PhysicalDate) + (p.Result ? " \u00b7 " + p.Result : ""),
         p.ExpirationDate ? "exp " + DS.fmtDate(p.ExpirationDate) : ""]));

      const accs = (cache.idx.accidentsByEmp[emp] || []).slice().sort(byDateDesc("AccidentDate"));
      wrap.appendChild(accidentList(accs));
      return wrap;
    }

    function smallBtn(text, onClick) {
      const b = el("button", { class: "btn btn--ghost btn--sm", type: "button", text: text });
      b.addEventListener("click", onClick);
      return b;
    }

    // Generic history list. Only "Manual Entry" rows can be removed — imported
    // records are corrected by adding a newer record (newest always wins).
    function recList(title, rows, listName, cells) {
      const box = el("div", { class: "rec-list" });
      box.appendChild(el("div", { class: "rec-head", text: title + " (" + rows.length + ")" }));
      if (!rows.length) { box.appendChild(el("div", { class: "rec-empty", text: "None on record" })); return box; }
      rows.slice(0, 6).forEach(rec => {
        const c = cells(rec);
        const right = [el("span", { class: "rec-meta", text: c[1] || "" })];
        if (rec.Source === "Manual Entry") right.push(removeBtn(listName, rec, title));
        box.appendChild(el("div", { class: "rec-row" }, [
          el("span", { text: c[0] }),
          el("span", { class: "rec-right" }, right),
        ]));
      });
      if (rows.length > 6) box.appendChild(el("div", { class: "rec-empty", text: "+ " + (rows.length - 6) + " older" }));
      return box;
    }

    function removeBtn(listName, rec, what) {
      const b = el("button", { class: "rec-x", type: "button", title: "Remove this manual entry", text: "\u00d7" });
      b.addEventListener("click", async () => {
        if (!confirm("Remove this manually entered " + what.toLowerCase().replace(/s$/, "") + " record?")) return;
        try {
          await DS.spDelete(listName, rec.Id);
          await DS.audit("Manual entry removed", listName, rec.EmployeeId, what + " record " + rec.Id);
          DS.toast("Removed.", "success");
          reload();
        } catch (e) { DS.toast(permMsg(e, "remove from " + listName), "error"); }
      });
      return b;
    }

    function accidentList(rows) {
      const box = el("div", { class: "rec-list" });
      box.appendChild(el("div", { class: "rec-head", text: "Accidents (" + rows.length + ")" }));
      if (!rows.length) { box.appendChild(el("div", { class: "rec-empty", text: "None on record" })); return box; }
      rows.slice(0, 6).forEach(a => {
        const d = DS.parseDate(a.AccidentDate);
        const counts = String(a.CountsAgainstStreak || "Auto");
        const active = d && d >= cache.idx.ptsCutoff && counts !== "Force No";
        const sel = el("select", { class: "rec-sel", title: "Counts against driving points / award streak" },
          ["Auto", "Force Yes", "Force No"].map(v => el("option", { value: v, text: v })));
        sel.value = ["Auto", "Force Yes", "Force No"].includes(counts) ? counts : "Auto";
        sel.addEventListener("change", async () => {
          const val = sel.value;
          try {
            await DS.spUpdate(DS.LISTS.accidents, a.Id, { CountsAgainstStreak: val });
            await DS.audit("Accident override changed", DS.LISTS.accidents, a.EmployeeId,
              (a.IncidentNumber || a.Id) + ": " + counts + " \u2192 " + val);
            DS.toast("Accident updated.", "success");
            reload();
          } catch (e) { sel.value = counts; DS.toast(permMsg(e, "edit Accidents"), "error"); }
        });
        box.appendChild(el("div", { class: "rec-row" }, [
          el("span", { text: (a.IncidentNumber || "—") + " \u00b7 " + DS.fmtDate(a.AccidentDate) }),
          el("span", { class: "rec-right" }, [
            el("span", { class: "rec-meta", text: (Number(a.FinalPoints) || 0) + " pts" + (active ? "" : " (inactive)") }),
            sel,
          ]),
        ]));
      });
      if (rows.length > 6) box.appendChild(el("div", { class: "rec-empty", text: "+ " + (rows.length - 6) + " older" }));
      return box;
    }

    // Plain-language errors — this is also how permission problems surface.
    function permMsg(e, action) {
      if (e && e.status === 403) return "No permission to " + action + " (SharePoint 403). Check this list's permissions for your account.";
      if (e && e.status === 400) return "SharePoint rejected the save (400) — a column may be missing on the list. " + e.message;
      return "Couldn't save: " + ((e && e.message) || e);
    }

    // Shared modal-save wrapper: create/update, audit, reload, readable errors.
    function formModal(title, fields, onSave) {
      const saveBtn = el("button", { class: "btn", type: "button", text: "Save" });
      const cancelBtn = el("button", { class: "btn btn--ghost", type: "button", text: "Cancel" });
      const m = DS.modal(title, el("div", null, fields), [cancelBtn, saveBtn]);
      cancelBtn.addEventListener("click", m.close);
      saveBtn.addEventListener("click", async () => {
        saveBtn.disabled = true; saveBtn.textContent = "Saving\u2026";
        try {
          const ok = await onSave();
          if (ok === false) { saveBtn.disabled = false; saveBtn.textContent = "Save"; return; }
          m.close(); reload();
        } catch (e) {
          saveBtn.disabled = false; saveBtn.textContent = "Save";
          DS.toast(e.__action ? permMsg(e, e.__action) : permMsg(e, "save"), "error");
        }
      });
    }
    function tag(e, action) { e.__action = action; return e; }
    const today = () => DS.isoDate(new Date());
    const dateInput = v => el("input", { class: "field", type: "date", value: v || "" });

    function openCourseForm(r) {
      const titles = cache.idx.requiredTitles;
      const which = el("select", { class: "field" },
        [el("option", { value: "__all", text: "All required courses (bundled)" })]
          .concat(titles.map(t => el("option", { value: t, text: t }))));
      const date = dateInput(today());
      formModal("Add course completion \u2014 " + (r.Title || r.EmployeeId), [
        DS.formField("Course", which),
        DS.formField("Date completed", date),
      ], async () => {
        if (!date.value) { DS.toast("Enter the completion date."); return false; }
        const list = which.value === "__all" ? titles : [which.value];
        try {
          for (const t of list) {
            await DS.spCreate(DS.LISTS.courses, { Title: r.Title || "", EmployeeId: String(r.EmployeeId), CourseTitle: t,
              DateCompleted: date.value, CompletionStatus: "Passed", Source: "Manual Entry" });
          }
        } catch (e) { throw tag(e, "add to Courses"); }
        await DS.audit("Manual course entry", DS.LISTS.courses, r.EmployeeId, list.join(", ") + " on " + date.value);
        DS.toast("Course completion saved.", "success");
      });
    }

    function openPhysicalForm(r) {
      const tested = dateInput(today());
      const exp = dateInput("");
      const result = el("select", { class: "field" }, ["Pass", "Fail", "Pending"].map(v => el("option", { value: v, text: v })));
      const provider = el("input", { class: "field", type: "text", placeholder: "Clinic (optional)" });
      formModal("Add physical \u2014 " + (r.Title || r.EmployeeId), [
        DS.formField("Date tested", tested),
        DS.formField("Expiration / due date", exp, "Leave blank to use the default cycle from the test date."),
        DS.formField("Result", result),
        DS.formField("Provider", provider),
      ], async () => {
        if (!tested.value && !exp.value) { DS.toast("Enter a test date or an expiration date."); return false; }
        const rec = { Title: r.Title || "", EmployeeId: String(r.EmployeeId), Result: result.value,
          Provider: provider.value.trim(), Source: "Manual Entry" };
        if (tested.value) rec.PhysicalDate = tested.value;
        if (exp.value) rec.ExpirationDate = exp.value;
        try { await DS.spCreate(DS.LISTS.physicals, rec); } catch (e) { throw tag(e, "add to Physicals"); }
        await DS.audit("Manual physical entry", DS.LISTS.physicals, r.EmployeeId,
          "Tested " + (tested.value || "\u2014") + ", exp " + (exp.value || "default") + ", " + result.value);
        DS.toast("Physical saved.", "success");
      });
    }

    function openAccidentForm(r) {
      const inc = el("input", { class: "field", type: "text", placeholder: "Required \u2014 must be unique" });
      const date = dateInput(today());
      const pts = el("input", { class: "field", type: "number", min: "0", step: "1", value: "0" });
      const counts = el("select", { class: "field" }, ["Auto", "Force Yes", "Force No"].map(v => el("option", { value: v, text: v })));
      const note = el("input", { class: "field", type: "text", placeholder: "Decision / note (optional)" });
      formModal("Add accident \u2014 " + (r.Title || r.EmployeeId), [
        DS.formField("Incident number", inc),
        DS.formField("Accident date", date),
        DS.formField("Final points", pts),
        DS.formField("Counts against points / streak", counts, "Auto = counts if it has points. Force No = never counts."),
        DS.formField("Decision / note", note),
      ], async () => {
        const n = inc.value.trim();
        if (!n || !date.value) { DS.toast("Incident number and date are required."); return false; }
        if (cache.accidents.some(a => String(a.IncidentNumber || "").trim() === n)) {
          DS.toast("That incident number is already on file."); return false;
        }
        try {
          await DS.spCreate(DS.LISTS.accidents, { Title: r.Title || "", IncidentNumber: n, EmployeeId: String(r.EmployeeId),
            AccidentDate: date.value, FinalPoints: Number(pts.value) || 0, FinalDecision: note.value.trim(),
            CountsAgainstStreak: counts.value, Source: "Manual Entry" });
        } catch (e) { throw tag(e, "add to Accidents"); }
        await DS.audit("Manual accident entry", DS.LISTS.accidents, r.EmployeeId, n + ", " + (Number(pts.value) || 0) + " pts");
        DS.toast("Accident saved.", "success");
      });
    }

    function openOverrideForm(r) {
      const status = el("select", { class: "field" }, [
        el("option", { value: "", text: "No override \u2014 use accident points" }),
        el("option", { value: "Normal", text: "Normal" }),
        el("option", { value: "Restrictive", text: "Restrictive" }),
        el("option", { value: "No-Driving", text: "No-Driving" }),
      ]);
      status.value = String(r.DrivingOverride || "");
      const until = dateInput(DS.isoDate(r.DrivingOverrideUntil) || "");
      const note = el("input", { class: "field", type: "text", value: r.DrivingOverrideNote || "", placeholder: "Reason (e.g. medical)" });
      formModal("Driving status override \u2014 " + (r.Title || r.EmployeeId), [
        DS.formField("Override", status, "Overrides the points-based status. Use for restrictions points can't show."),
        DS.formField("Until", until, "Leave blank for no end date. After this date, the points-based status returns automatically."),
        DS.formField("Note", note),
      ], async () => {
        const fields = { DrivingOverride: status.value, DrivingOverrideUntil: status.value && until.value ? until.value : null,
          DrivingOverrideNote: status.value ? note.value.trim() : "" };
        try { await DS.spUpdate(DS.LISTS.roster, r.Id, fields); } catch (e) { throw tag(e, "edit the Roster"); }
        await DS.audit("Driving override " + (status.value ? "set" : "cleared"), DS.LISTS.roster, r.EmployeeId,
          status.value ? status.value + (until.value ? " until " + until.value : "") + (note.value ? " \u2014 " + note.value.trim() : "") : "");
        DS.toast(status.value ? "Override saved." : "Override cleared.", "success");
      });
    }

    search.addEventListener("input", () => { rosterState.search = search.value; paint(); });

    paint();
    paintDetail();
  }

  function detailRow(k, v) {
    return el("div", { class: "detail__row" }, [
      el("span", { class: "k", text: k }),
      el("span", { class: "v" }, typeof v === "string" ? v : [v]),
    ]);
  }

  // pull an employee id out of a route like #/roster/123456
  function routeParam() {
    const h = (location.hash || "").replace(/^#\/?/, "");
    const parts = h.split("/");
    return parts[1] ? decodeURIComponent(parts[1]) : null;
  }

  /* ============================================================
     AUDIT LOG
     ============================================================ */
  async function renderAudit(container) {
    const rows = await DS.spGet(DS.LISTS.audit, { orderby: "ActionTimestamp desc", top: 500 });
    container.innerHTML = "";

    const types = Array.from(new Set(rows.map(r => r.ActionType).filter(Boolean))).sort();
    const filter = el("select", { class: "field" }, [
      el("option", { value: "", text: "All action types" }),
      ...types.map(t => el("option", { value: t, text: t })),
    ]);
    container.appendChild(el("div", { class: "toolbar" }, [
      filter,
      el("span", { class: "count-pill", id: "auditCount" }),
    ]));

    const tableWrap = el("div", { class: "card" });
    container.appendChild(tableWrap);

    function paint() {
      const t = filter.value;
      const view = t ? rows.filter(r => r.ActionType === t) : rows;
      document.getElementById("auditCount").textContent =
        view.length + (view.length === 1 ? " entry" : " entries");
      tableWrap.innerHTML = "";
      if (!view.length) { tableWrap.appendChild(emptyMini("No matching activity.")); return; }
      tableWrap.appendChild(buildTable([
        { head: "When", thClass: "nowrap", tdClass: "nowrap", render: r => el("span", { class: "tnum", text: DS.fmtDateTime(r.ActionTimestamp) }) },
        { head: "Who", render: r => el("span", { text: DS.emailLocal(r.Actor) || r.Actor || "—" }) },
        { head: "Action", render: r => DS.badge(r.ActionType || "—", "neutral") },
        { head: "Detail", render: r => r.Detail || "—" },
      ], view));
    }

    filter.addEventListener("change", paint);
    paint();
  }

  (function injectRecordStyles() {
    if (document.getElementById("roster-records-styles")) return;
    const s = document.createElement("style");
    s.id = "roster-records-styles";
    s.textContent =
      ".detail{max-height:calc(100vh - 96px);overflow-y:auto}" +
      ".rec-actions{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0 14px}" +
      ".rec-list{margin-bottom:12px}" +
      ".rec-head{font-size:11.5px;font-weight:600;color:var(--slate);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}" +
      ".rec-row{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:12.5px;padding:5px 0;border-bottom:1px solid var(--line-2)}" +
      ".rec-right{display:flex;align-items:center;gap:6px;flex:none}" +
      ".rec-meta{color:var(--muted);white-space:nowrap}" +
      ".rec-empty{font-size:12.5px;color:var(--muted);padding:3px 0}" +
      ".rec-x{border:none;background:none;color:var(--muted);font-size:16px;cursor:pointer;line-height:1;padding:0 2px}" +
      ".rec-x:hover{color:var(--overdue)}" +
      ".rec-sel{font-size:12px;padding:2px 4px;border:1px solid var(--line);border-radius:6px;background:var(--paper)}";
    document.head.appendChild(s);
  })();

  /* ---- register ---- */
  DS.registerScreen("dashboard", { title: "Dashboard", icon: "◆", render: renderDashboard });
  DS.registerScreen("roster", { title: "Roster", icon: "▤", render: renderRoster });
  DS.registerScreen("audit", { title: "Audit log", icon: "◷", render: renderAudit });

})();
