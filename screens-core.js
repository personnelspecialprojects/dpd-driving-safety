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

      body.appendChild(detailRow("Driving eligibility",
        drive.status === "No-Driving" ? DS.badge("No-Driving \u00b7 " + drive.points + " pts", "overdue")
          : drive.status === "Restrictive" ? DS.badge("Restrictive \u00b7 " + drive.points + " pts", "due")
          : DS.badge("Normal \u00b7 " + drive.points + " pts", "neutral")));

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

      detailWrap.appendChild(body);
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

  /* ---- register ---- */
  DS.registerScreen("dashboard", { title: "Dashboard", icon: "◆", render: renderDashboard });
  DS.registerScreen("roster", { title: "Roster", icon: "▤", render: renderRoster });
  DS.registerScreen("audit", { title: "Audit log", icon: "◷", render: renderAudit });

})();
