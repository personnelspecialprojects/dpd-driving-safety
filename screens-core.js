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

  /* ---- smartTable: sortable columns + search + filter dropdowns ----
     o.key       remembers sort/filter per table between visits
     o.columns   [{ head, render(r), sort(r)?, thClass?, tdClass?, dir? }]
                 (columns with a sort() get a clickable header)
     o.facets    [{ label, options: [[value,text]] or fn(rows), test(r, value) }]
     o.search    r → text matched by the search box
     o.onRow / o.rowClass  optional row click + extra class
     o.countEl / o.noun    element showing "12 of 85 employees"
     Returns { bar, body, refresh(), setRows(rows) }. */
  const TABLE_STATE = {};
  function smartTable(o) {
    // o.defaultSort = [columnIndex, 1 | -1]; otherwise rows keep their given (priority) order
    const dflt = o.defaultSort || [null, 1];
    const st = TABLE_STATE[o.key] = TABLE_STATE[o.key] || { q: "", sort: dflt[0], dir: dflt[1], facets: {} };
    let rows = o.rows || [];
    const bar = el("div", { class: "st-bar" });
    const body = el("div", { class: "st-body" });
    const empty = v => v == null || v === "" || (v instanceof Date && isNaN(v));

    function buildBar() {
      bar.innerHTML = "";
      const q = el("input", { class: "st-search", type: "search", placeholder: o.placeholder || "Search name or ID", value: st.q });
      q.addEventListener("input", () => { st.q = q.value; draw(); });
      bar.appendChild(q);
      (o.facets || []).forEach(f => {
        const opts = typeof f.options === "function" ? f.options(rows) : f.options;
        const sel = el("select", { class: "st-facet", title: f.label },
          [el("option", { value: "", text: f.label + ": All" })].concat(opts.map(([v, t]) => el("option", { value: v, text: t }))));
        sel.value = st.facets[f.label] || "";
        if (sel.value !== (st.facets[f.label] || "")) st.facets[f.label] = "";   // saved choice no longer offered
        sel.addEventListener("change", () => { st.facets[f.label] = sel.value; draw(); });
        bar.appendChild(sel);
      });
      const clear = el("button", { class: "st-clear", type: "button", text: "Clear" });
      clear.addEventListener("click", () => { st.q = ""; st.facets = {}; st.sort = dflt[0]; st.dir = dflt[1]; buildBar(); draw(); });
      bar.appendChild(clear);
    }

    function visible() {
      const needle = st.q.trim().toLowerCase();
      let out = rows.filter(r =>
        (!needle || String(o.search ? o.search(r) : "").toLowerCase().includes(needle)) &&
        (o.facets || []).every(f => { const v = st.facets[f.label]; return !v || f.test(r, v); }));
      const col = st.sort != null ? o.columns[st.sort] : null;
      if (col && col.sort) {
        out = out.slice().sort((a, b) => {
          let va = col.sort(a), vb = col.sort(b);
          const ea = empty(va), eb = empty(vb);
          if (ea || eb) return ea === eb ? 0 : (ea ? 1 : -1);          // blanks always last
          if (va instanceof Date) va = va.getTime();
          if (vb instanceof Date) vb = vb.getTime();
          const c = (typeof va === "number" && typeof vb === "number") ? va - vb
            : String(va).localeCompare(String(vb), undefined, { numeric: true, sensitivity: "base" });
          return c * st.dir;
        });
      }
      return out;
    }

    function draw() {
      const vis = visible();
      if (o.countEl) {
        const n = rows.length, noun = o.noun || "employee", plural = o.nounPlural || noun + "s";
        o.countEl.textContent = (vis.length === n ? "" : vis.length + " of ") + n + " " + (n === 1 ? noun : plural);
      }
      body.innerHTML = "";
      if (!rows.length) { body.appendChild(emptyMini(o.emptyText || "Nothing to show.")); return; }
      if (!vis.length) { body.appendChild(emptyMini("No matches \u2014 adjust or clear the filters.")); return; }
      const shown = o.limit ? vis.slice(0, o.limit) : vis;
      const thead = el("thead", null, el("tr", null, o.columns.map((c, i) => {
        const arrow = st.sort === i ? (st.dir > 0 ? " \u25B2" : " \u25BC") : "";
        const th = el("th", { class: (c.thClass || "") + (c.sort ? " st-sortable" : ""), title: c.sort ? "Click to sort" : "" }, c.head + arrow);
        if (c.sort) th.addEventListener("click", () => {
          if (st.sort === i) st.dir = -st.dir; else { st.sort = i; st.dir = c.dir || 1; }
          draw();
        });
        return th;
      })));
      const tbody = el("tbody", null, shown.map(r => {
        const tr = el("tr", null, o.columns.map(c => {
          const cell = c.render(r);
          return el("td", { class: c.tdClass || "" }, typeof cell === "string" ? cell : [cell]);
        }));
        if (o.onRow) {
          tr.className = "roster-row" + (o.rowClass ? " " + o.rowClass(r) : "");
          tr.addEventListener("click", () => o.onRow(r));
        }
        return tr;
      }));
      body.appendChild(el("table", { class: "tbl" }, [thead, tbody]));
      if (o.limit && vis.length > o.limit) body.appendChild(emptyMini("Showing the first " + o.limit + " \u2014 search or filter to narrow the list."));
    }

    buildBar(); draw();
    return { bar, body, refresh: draw, setRows(r) { rows = r; buildBar(); draw(); } };
  }
  const distinctOptions = fn => rows => Array.from(new Set(rows.map(fn).filter(v => v != null && String(v).trim() !== "")))
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true })).map(v => [String(v), String(v)]);

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
  // Is a due date inside the alert window (amber) or comfortably in the future (green)?
  function withinLead(dueDate, leadDays) {
    return dueDate <= DS.util.addDays(DS.util.startOfToday(), leadDays);
  }

  function physicalDetailBadge(phys, leadDays) {
    if (!phys.applicable) return DS.badge("Not applicable (Non-Driver)", "neutral");
    if (!phys.has) return phys.required
      ? DS.badge("Required \u2014 none on record", "overdue")
      : DS.badge("Not required \u2014 none on record", "neutral");
    if (!phys.dueDate) return DS.badge("On record \u2014 no due date", "neutral");
    if (phys.overdue) return DS.badge("Overdue \u00b7 " + DS.fmtDate(phys.dueDate), "overdue");
    if (withinLead(phys.dueDate, leadDays)) return DS.badge("Due " + DS.fmtDate(phys.dueDate), phys.required ? "due" : "neutral");
    return DS.badge("Current \u00b7 next due " + DS.fmtDate(phys.dueDate), phys.required ? "clear" : "neutral");
  }

  function courseDetailBadge(crs, leadDays) {
    if (!crs.applicable) return DS.badge("Not applicable (Non-Driver)", "neutral");
    const today = DS.util.startOfToday();
    const when = crs.dueDate ? DS.fmtDate(crs.dueDate) : null;
    if (crs.doneCount === crs.requiredCount) {           // all required courses on file
      if (!crs.dueDate) return DS.badge("Complete", "clear");
      if (crs.dueDate < today) return DS.badge("Renewal overdue \u00b7 " + when, "overdue");
      if (withinLead(crs.dueDate, leadDays)) return DS.badge("Renewal due " + when, "due");
      return DS.badge("Current \u00b7 next due " + when, "clear");
    }
    const label = crs.status + " (" + crs.doneCount + " of " + crs.requiredCount + ")";
    if (!crs.dueDate) return DS.badge(label + " \u00b7 no hire date", "neutral");
    if (crs.dueDate < today) return DS.badge(label + " \u00b7 overdue", "overdue");
    return DS.badge(label + " \u00b7 due " + when, withinLead(crs.dueDate, leadDays) ? "due" : "neutral");
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
    let unassigned = DS.compute.needsDesignation(cache);

    container.innerHTML = "";

    // stat row
    const criticalPhys = physicals.filter(p => p.urgency === "missing" || p.urgency === "overdue").length;
    const noDriving = restricted.filter(r => r.status === "No-Driving").length;
    const needStat = statCard(unassigned.length, "Need a designation", unassigned.length ? "due" : "clear");
    container.appendChild(el("div", { class: "stats" }, [
      statCard(physicals.length, "Physicals due or missing", criticalPhys ? "overdue" : "due"),
      statCard(courses.length, "Courses coming due", courses.some(c => c.overdue) ? "overdue" : "due"),
      needStat,
      statCard(restricted.length, "On restricted / no-driving", noDriving ? "overdue" : "due"),
      statCard(awards.length, "Awards eligible now", "clear"),
    ]));

    const grid = el("div", { class: "dash-quad" });
    container.appendChild(grid);

    // shared helpers for these panels
    const nameSort = r => { const rr = cache.idx.rosterByEmp[r.employeeId]; return ((rr && rr.LastName) || "") + " " + (r.name || ""); };
    const nameCol = { head: "Employee", render: r => el("span", { class: "strong", text: r.name || "\u2014" }), sort: nameSort };
    const searchText = r => (r.name || "") + " " + r.employeeId;
    const toRoster = r => DS.navigate("roster/" + encodeURIComponent(r.employeeId));
    function panel(key, title, rows, cfg) {
      const pill = el("span", { class: "count-pill" });
      const t = smartTable(Object.assign({ key: "dash-" + key, rows, countEl: pill, search: searchText, onRow: toRoster }, cfg));
      grid.appendChild(quad(title, pill, t.body, t.bar));
      return t;
    }

    // Q1 — Physicals due
    const PH_RANK = { missing: 0, overdue: 1, due: 2 };
    panel("phys", "Physicals due", physicals, {
      emptyText: "No physicals due or missing within the alert window.",
      columns: [
        nameCol,
        { head: "Driver", render: r => driverCell(r.required), sort: r => (r.required ? 0 : 1) },
        { head: "Status", thClass: "nowrap", tdClass: "nowrap", render: r => physBadge(r),
          sort: r => PH_RANK[r.urgency] * 1e13 + (r.dueDate ? r.dueDate.getTime() : 0) },
      ],
      facets: [
        { label: "Driver", options: [["Primary", "Primary"], ["Secondary", "Secondary"]], test: (r, v) => (r.required ? "Primary" : "Secondary") === v },
        { label: "Status", options: [["missing", "Missing"], ["overdue", "Overdue"], ["due", "Due soon"]], test: (r, v) => r.urgency === v },
      ],
    });

    // Q2 — Courses due
    panel("courses", "Defensive driving due", courses, {
      emptyText: "No courses due within the alert window.",
      columns: [
        nameCol,
        { head: "Status", render: r => DS.badge(r.status, r.status === "Not started" ? "overdue" : "neutral"), sort: r => r.status },
        { head: "Due", thClass: "nowrap", tdClass: "nowrap", render: r => dueBadge(r.dueDate, r.overdue), sort: r => r.dueDate },
      ],
      facets: [
        { label: "Status", options: [["Not started", "Not started"], ["Incomplete", "Incomplete"], ["Renewal due", "Renewal due"]], test: (r, v) => r.status === v },
        { label: "Timing", options: [["overdue", "Overdue"], ["upcoming", "Coming due"]], test: (r, v) => (r.overdue ? "overdue" : "upcoming") === v },
      ],
    });

    // Q3 — Needs a designation (assign right here)
    let needTable = null;
    function assignedOne(row) {
      unassigned = unassigned.filter(x => x !== row);
      needTable.setRows(unassigned);
      needStat.firstChild.textContent = String(unassigned.length);
      needStat.className = "stat stat--" + (unassigned.length ? "due" : "clear");
    }
    needTable = panel("need", "Needs a designation", unassigned, {
      emptyText: "Everyone on the active roster has a designation.",
      columns: [
        nameCol,
        { head: "ID", tdClass: "nowrap", render: r => el("span", { class: "tnum", text: String(r.record.EmployeeId || r.employeeId) }), sort: r => r.employeeId },
        { head: "Rank", render: r => r.rank || "\u2014", sort: r => r.rank },
        { head: "Assign", tdClass: "nowrap", render: r => assignSelect(r, assignedOne) },
      ],
      facets: [{ label: "Rank", options: distinctOptions(r => r.rank), test: (r, v) => r.rank === v }],
    });

    // Q4 — Driving status
    const isOverride = r => !!DS.compute.drivingStatusFor(cache, r.employeeId).override;
    panel("driving", "Driving status", restricted, {
      emptyText: "No employees are on restricted or no-driving status.",
      columns: [
        nameCol,
        { head: "Status", render: r => DS.badge(r.status + (isOverride(r) ? " \u00b7 manual" : ""), r.status === "No-Driving" ? "overdue" : "due"),
          sort: r => (r.status === "No-Driving" ? 0 : 1) },
        { head: "Active points", thClass: "num", tdClass: "num", render: r => el("span", { class: "tnum", text: String(r.points) }), sort: r => r.points, dir: -1 },
      ],
      facets: [
        { label: "Status", options: [["No-Driving", "No-Driving"], ["Restrictive", "Restrictive"]], test: (r, v) => r.status === v },
        { label: "Source", options: [["points", "Accident points"], ["manual", "Manual override"]], test: (r, v) => (isOverride(r) ? "manual" : "points") === v },
      ],
    });

    // Awards — full width below the quadrants (rows not clickable: they hold a button)
    const awardsPill = el("span", { class: "count-pill" });
    const at = smartTable({ key: "dash-awards", rows: awards, countEl: awardsPill, search: searchText,
      emptyText: "No employees are award-eligible right now.", columns: awardColumns(cache, container, nameSort),
      facets: [{ label: "Milestone", options: distinctOptions(r => r.nextMilestone + "-year"), test: (r, v) => (r.nextMilestone + "-year") === v }] });
    const awardsQuad = quad("Awards eligible", awardsPill, at.body, at.bar);
    awardsQuad.classList.add("dash-wide");
    grid.appendChild(awardsQuad);
  }

  // A dashboard panel: header, filter bar, then a fixed-height scrolling body
  function quad(title, headerRight, body, bar) {
    return el("div", { class: "card dash-panel" }, [
      el("div", { class: "card__head" }, [el("h3", { text: title }), headerRight || null]),
      bar || null,
      el("div", { class: "dash-scroll" }, body),
    ]);
  }

  // Inline designation picker for the "Needs a designation" panel
  function assignSelect(row, onAssigned) {
    const sel = el("select", { class: "rec-sel", title: "Assign a designation" }, [
      el("option", { value: "", text: "Assign\u2026" }),
      el("option", { value: "Primary", text: "Primary" }),
      el("option", { value: "Secondary", text: "Secondary" }),
      el("option", { value: "Non-Driver", text: "Non-Driver" }),
    ]);
    sel.addEventListener("click", e => e.stopPropagation());   // don't open the Roster row
    sel.addEventListener("change", async e => {
      e.stopPropagation();
      const val = sel.value; if (!val) return;
      sel.disabled = true;
      try {
        await DS.spUpdate(DS.LISTS.roster, row.record.Id, { DriverStatus: val });
        await DS.audit("Designation assigned", DS.LISTS.roster, row.employeeId, (row.name || row.employeeId) + ": \u2192 " + val);
        row.record.DriverStatus = val;               // keep the in-memory roster in sync
        onAssigned(row);
        DS.toast((row.name || row.employeeId) + " set to " + val + ".", "success");
      } catch (err) {
        sel.disabled = false; sel.value = "";
        DS.toast("Couldn't assign: " + err.message + (err.status === 400 ? " \u2014 the DriverStatus choice list may need that option." : ""), "error");
      }
    });
    return sel;
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

  function awardColumns(cache, container, nameSort) {
    const cols = [
      { head: "Employee", render: r => el("span", { class: "strong", text: r.name }), sort: nameSort },
      { head: "Milestone", render: r => DS.badge(r.nextMilestone + "-year", "clear"), sort: r => r.nextMilestone },
      { head: "Eligible since", thClass: "nowrap", tdClass: "nowrap", render: r => el("span", { class: "tnum", text: DS.fmtDate(r.eligibleDate) }), sort: r => r.eligibleDate },
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
  let rosterState = { search: "", selected: null, show: "active" };

  async function renderRoster(container) {
    const cache = await DS.data.load();
    const preselect = routeParam();
    if (preselect) rosterState.selected = DS.util.empKey(preselect);

    container.innerHTML = "";

    const showSel = el("select", { class: "field", style: "max-width:200px" }, [
      el("option", { value: "active", text: "Active employees" }),
      el("option", { value: "separated", text: "Separated employees" }),
      el("option", { value: "all", text: "Everyone" }),
    ]);
    showSel.value = rosterState.show;
    const countPill = el("span", { class: "count-pill" });
    container.appendChild(el("div", { class: "toolbar" }, [showSel, countPill]));

    const split = el("div", { class: "split" });
    const listWrap = el("div", { class: "card" });
    const detailWrap = el("div", { class: "card detail" });
    split.appendChild(listWrap);
    split.appendChild(detailWrap);
    container.appendChild(split);

    const baseRows = () => rosterState.show === "active" ? cache.idx.activeRoster
      : rosterState.show === "separated" ? cache.roster.filter(r => !DS.util.isActive(r))
      : cache.roster;
    const desigLabel = r => DS.util.isUnassigned(r) ? "Unassigned" : DS.util.designation(r);
    const driving = r => DS.compute.drivingStatusFor(cache, DS.util.empKey(r.EmployeeId)).status;
    const DRIVE_RANK = { "No-Driving": 0, "Restrictive": 1, "Normal": 2 };

    const table = smartTable({
      key: "roster", rows: baseRows(), countEl: countPill, limit: 400, defaultSort: [0, 1],   // last name A→Z
      placeholder: "Search by name, employee # or badge",
      search: r => [r.Title, r.LastName, r.EmployeeId, r.Badge].join(" "),
      onRow: r => { rosterState.selected = DS.util.empKey(r.EmployeeId); table.refresh(); paintDetail(); },
      rowClass: r => DS.util.empKey(r.EmployeeId) === rosterState.selected ? "selected" : "",
      emptyText: "No employees in this view.",
      columns: [
        { head: "Name", sort: r => (r.LastName || "") + " " + (r.Title || ""),
          render: r => DS.util.isActive(r)
            ? el("span", { class: "strong", text: r.Title || "—" })
            : el("span", null, [el("span", { class: "strong", style: "color:var(--muted)", text: r.Title || "—" }), " ", DS.badge("Separated", "neutral")]) },
        { head: "ID", tdClass: "nowrap num", thClass: "num", sort: r => String(r.EmployeeId || ""),
          render: r => el("span", { class: "tnum", text: r.EmployeeId || "—" }) },
        { head: "Designation", sort: desigLabel,
          render: r => DS.util.isUnassigned(r) ? DS.badge("Unassigned", "due") : DS.badge(DS.util.designation(r), "neutral") },
        { head: "Driving", sort: r => DRIVE_RANK[driving(r)],
          render: r => {
            const st = driving(r);
            if (st === "No-Driving") return DS.badge("No-Driving", "overdue");
            if (st === "Restrictive") return DS.badge("Restrictive", "due");
            return el("span", { style: "color:var(--muted)", text: "—" });
          } },
      ],
      facets: [
        { label: "Designation", options: [["Unassigned", "Unassigned"], ["Primary", "Primary"], ["Secondary", "Secondary"], ["Non-Driver", "Non-Driver"]],
          test: (r, v) => desigLabel(r) === v },
        { label: "Driving", options: [["Normal", "Normal"], ["Restrictive", "Restrictive"], ["No-Driving", "No-Driving"]], test: (r, v) => driving(r) === v },
        { label: "Division", options: distinctOptions(r => r.Division), test: (r, v) => String(r.Division || "") === v },
      ],
    });
    listWrap.appendChild(table.bar);
    listWrap.appendChild(table.body);
    showSel.addEventListener("change", () => { rosterState.show = showSel.value; table.setRows(baseRows()); });
    function paint() { table.refresh(); }   // re-draw after an edit (e.g. designation saved)

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

      if (String(r.Badge || "").trim()) body.appendChild(detailRow("Badge", String(r.Badge)));
      body.appendChild(detailRow("Employment", employmentControl(r)));
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

      body.appendChild(detailRow("Physical", physicalDetailBadge(phys, cache.idx.physLeadDays)));
      body.appendChild(detailRow("Courses", courseDetailBadge(crs, cache.idx.courseLeadDays)));
      body.appendChild(detailRow("Next award",
        awd.eligible ? DS.badge(awd.nextMilestone + "-year (eligible)", "clear")
          : el("span", { class: "tnum", text: awd.nextMilestone + "-year · " + DS.fmtDate(awd.eligibleDate) })));

      // driver designation editor
      const editWrap = el("div", { class: "detail__edit" });
      editWrap.appendChild(el("span", { class: "label", text: "Driver designation" }));
      const sel = el("select", { class: "field" }, [
        el("option", { value: "", text: "Unassigned \u2014 choose\u2026" }),
        el("option", { value: "Primary", text: "Primary" }),
        el("option", { value: "Secondary", text: "Secondary" }),
        el("option", { value: "Non-Driver", text: "Non-Driver" }),
      ]);
      sel.value = DS.util.isUnassigned(r) ? "" : DS.util.designation(r);
      const saveBtn = el("button", { class: "btn btn--sm", text: "Save" });
      saveBtn.addEventListener("click", async () => {
        const val = sel.value;
        const cur = DS.util.isUnassigned(r) ? "" : DS.util.designation(r);
        if (!val) { DS.toast("Choose a designation to save."); return; }
        if (val === cur) { DS.toast("No change to save."); return; }
        saveBtn.disabled = true; saveBtn.textContent = "Saving…";
        try {
          await DS.spUpdate(DS.LISTS.roster, r.Id, { DriverStatus: val });
          await DS.audit("Designation changed", DS.LISTS.roster, r.EmployeeId,
            r.Title + ": " + (cur || "Unassigned") + " → " + val);
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

    /* ---------- Employment: mark separated / reactivate (history is kept) ---------- */
    function employmentControl(r) {
      const active = DS.util.isActive(r);
      const btn = el("button", { class: "btn btn--ghost btn--sm", type: "button", text: active ? "Mark separated" : "Reactivate" });
      btn.addEventListener("click", async () => {
        const msg = active
          ? "Mark " + (r.Title || r.EmployeeId) + " as separated?\n\nThey'll drop off the Dashboard and out of compliance tracking. All of their records are kept, and they can be reactivated any time."
          : "Reactivate " + (r.Title || r.EmployeeId) + "?\n\nThey'll return to the Dashboard with their full history and designation.";
        if (!confirm(msg)) return;
        btn.disabled = true;
        try {
          await DS.spUpdate(DS.LISTS.roster, r.Id, { ActiveEmployee: !active });
          await DS.audit(active ? "Employee separated" : "Employee reactivated", DS.LISTS.roster, r.EmployeeId, r.Title || "");
          DS.toast((r.Title || r.EmployeeId) + (active ? " marked separated." : " reactivated."), "success");
          reload();
        } catch (e) { btn.disabled = false; DS.toast("Couldn't update: " + e.message, "error"); }
      });
      return el("span", { style: "display:inline-flex; gap:8px; align-items:center" }, [
        active ? DS.badge("Active", "clear") : DS.badge("Separated", "neutral"), btn,
      ]);
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

      const isMaster = rec => DS.util.isFallbackSource(rec);
      const renewal = cache.idx.courseRenewalYears;
      const courseEff = cache.idx.courseInEffect[emp] || {};
      // Master-sheet course rows hold an estimated completion (due − cycle); show the original due date.
      const courseSortDate = c => DS.parseDate(c.DateCompleted);
      const courses = (cache.idx.coursesByEmp[emp] || []).slice().sort((a, b) => (courseSortDate(b) || 0) - (courseSortDate(a) || 0));
      wrap.appendChild(recList("Courses", courses, DS.LISTS.courses, c => {
        const d = DS.parseDate(c.DateCompleted);
        return [c.CourseTitle || "\u2014",
          isMaster(c) ? (d ? "due " + DS.fmtDate(DS.util.addYears(d, renewal)) : "\u2014") : DS.fmtDate(c.DateCompleted)];
      }, c => courseEff[String(c.CourseTitle || "").trim()] === c));

      const physEff = cache.idx.physInEffect[emp];
      const physSortDate = p => DS.parseDate(p.PhysicalDate) || DS.parseDate(p.ExpirationDate);
      const phys = (cache.idx.physicalsByEmp[emp] || []).slice().sort((a, b) => (physSortDate(b) || 0) - (physSortDate(a) || 0));
      wrap.appendChild(recList("Physicals", phys, DS.LISTS.physicals, p =>
        p.PhysicalDate
          ? ["Tested " + DS.fmtDate(p.PhysicalDate) + (p.Result ? " \u00b7 " + p.Result : ""), p.ExpirationDate ? "exp " + DS.fmtDate(p.ExpirationDate) : ""]
          : [isMaster(p) ? "Master sheet" : "Physical" + (p.Result ? " \u00b7 " + p.Result : ""), p.ExpirationDate ? "due " + DS.fmtDate(p.ExpirationDate) : ""],
        p => p === physEff));

      const accs = (cache.idx.accidentsByEmp[emp] || []).slice().sort(byDateDesc("AccidentDate"));
      wrap.appendChild(accidentList(accs));
      wrap.appendChild(el("div", { class: "help", style: "margin-top:6px",
        text: "\u2713 = the record currently used for status. Master-sheet dates are used only when no uploaded, manual, or exam-history record exists." }));
      return wrap;
    }

    function smallBtn(text, onClick) {
      const b = el("button", { class: "btn btn--ghost btn--sm", type: "button", text: text });
      b.addEventListener("click", onClick);
      return b;
    }

    const SOURCE_TAGS = {
      "Bulk Upload": ["upload", "Added by an upload on the Imports screen"],
      "Manual Entry": ["manual", "Entered by hand in the app"],
      "Legacy History": ["exam log", "From Julie's exam history workbook"],
      "Master Sheet": ["master", "From Julie's master spreadsheet \u2014 used only if nothing newer is on file"],
      "Legacy Migration": ["master", "From Julie's master spreadsheet (older migration run) \u2014 used only if nothing newer is on file"],
    };

    // Full employee history. Shows 6 at first with "Show all". Only "Manual Entry"
    // rows can be removed; everything else is permanent history.
    function recList(title, rows, listName, cells, inEffect) {
      const box = el("div", { class: "rec-list" });
      box.appendChild(el("div", { class: "rec-head", text: title + " (" + rows.length + ")" }));
      if (!rows.length) { box.appendChild(el("div", { class: "rec-empty", text: "None on record" })); return box; }
      const body = el("div");
      box.appendChild(body);
      function draw(all) {
        body.innerHTML = "";
        (all ? rows : rows.slice(0, 6)).forEach(rec => {
          const c = cells(rec);
          const tag = SOURCE_TAGS[rec.Source];
          const right = [el("span", { class: "rec-meta", text: c[1] || "" })];
          if (tag) right.push(el("span", { class: "rec-src", title: tag[1], text: tag[0] }));
          if (rec.Source === "Manual Entry") right.push(removeBtn(listName, rec, title));
          const eff = inEffect && inEffect(rec);
          body.appendChild(el("div", { class: "rec-row" + (eff ? " rec-eff" : ""), title: eff ? "Currently used for status" : "" }, [
            el("span", { text: (eff ? "\u2713 " : "") + c[0] }),
            el("span", { class: "rec-right" }, right),
          ]));
        });
        if (rows.length > 6) {
          const t = el("button", { class: "rec-more", type: "button", text: all ? "Show fewer" : "Show all " + rows.length });
          t.addEventListener("click", () => draw(!all));
          body.appendChild(t);
        }
      }
      draw(false);
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
    const today = () => DS.todayIso();
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
  /* ============================================================
     AUDIT LOG — two tabs: Activity (every action) and Upload results
     (one entry per upload/migration, with the rows that need review)
     ============================================================ */
  let auditTab = "activity";
  async function renderAudit(container) {
    container.innerHTML = "";
    const seg = el("div", { class: "seg" });
    [["activity", "Activity"], ["uploads", "Upload results"]].forEach(([k, label]) => {
      const b = el("button", { type: "button", text: label });
      if (k === auditTab) b.classList.add("active");
      b.addEventListener("click", () => { auditTab = k; renderAudit(container); });
      seg.appendChild(b);
    });
    container.appendChild(seg);
    const body = el("div");
    container.appendChild(body);
    DS.showLoading(body);
    if (auditTab === "uploads") await renderUploadResults(body);
    else await renderActivity(body);
  }

  async function renderActivity(container) {
    const rows = await DS.spGet(DS.LISTS.audit, { orderby: "Id desc", top: 500 });
    container.innerHTML = "";
    const pill = el("span", { class: "count-pill" });
    container.appendChild(el("div", { class: "toolbar" }, [el("span", { class: "help", text: "Most recent 500 actions" }), pill]));
    const t = smartTable({
      key: "audit-activity", rows, countEl: pill, noun: "entry", nounPlural: "entries",
      placeholder: "Search who, action, or detail",
      search: r => [r.Actor, r.ActionType, r.Detail, r.TargetId].join(" "),
      emptyText: "No activity yet.",
      columns: [
        { head: "When", thClass: "nowrap", tdClass: "nowrap", sort: r => DS.parseDate(r.ActionTimestamp), dir: -1,
          render: r => el("span", { class: "tnum", text: DS.fmtDateTime(r.ActionTimestamp) }) },
        { head: "Who", sort: r => DS.emailLocal(r.Actor), render: r => el("span", { text: DS.emailLocal(r.Actor) || r.Actor || "\u2014" }) },
        { head: "Action", sort: r => r.ActionType, render: r => DS.badge(r.ActionType || "\u2014", "neutral") },
        { head: "Detail", render: r => r.Detail || "\u2014" },
      ],
      facets: [{ label: "Action", options: distinctOptions(r => r.ActionType), test: (r, v) => r.ActionType === v }],
    });
    container.appendChild(el("div", { class: "card" }, [t.bar, t.body]));
  }

  const COUNT_LABELS = {
    rows: "Rows in file", added: "Added", updated: "Updated", inactivated: "Marked inactive", byName: "Matched by name (your choice)",
    byEmployeeNumber: "Matched by employee #", byBadge: "Matched by badge", noMatch: "Matched no one",
    matchedTwoPeople: "Two possible people (unsettled)", resolvedByName: "Two possible people (name confirmed)", alreadyInSharePoint: "Already in SharePoint",
    repeatedInFile: "Repeated in file", onlySomeCourses: "Only some required courses",
    noEmployeeNumber: "No employee number", failed: "Failed to save",
    employees: "Employees in sheet", designationsSet: "Designations set", physicalsAdded: "Physical records added",
    coursesAdded: "Course records added", alreadyOnFile: "Already on file", notOnRoster: "Not on roster",
    historyRowsUnmatched: "Exam-history rows unmatched",
  };
  const BAD_COUNTS = ["noMatch", "matchedTwoPeople", "failed", "noEmployeeNumber", "notOnRoster", "historyRowsUnmatched"];

  async function renderUploadResults(container) {
    let raw;
    try { raw = await DS.spGet(DS.LISTS.uploads, { orderby: "Id desc", top: 300 }); }
    catch (e) {
      container.innerHTML = "";
      container.appendChild(el("div", { class: "card" }, el("div", { class: "card__body" }, [
        el("h3", { text: "Upload results aren't set up yet", style: "font-size:15px; margin-bottom:8px" }),
        el("div", { class: "import-note", text: "Create a SharePoint list named DrivingSafety_UploadLog with two columns: UploadType (Single line of text) and Details (Multiple lines of text \u2014 Plain text). Every upload and migration run is recorded automatically after that." }),
        el("div", { class: "help", text: "SharePoint said: " + e.message }),
      ])));
      return;
    }
    const items = raw.map(r => {
      let d = {};
      try { d = JSON.parse(r.Details || "{}"); } catch (_) {}
      const counts = d.counts || {};
      const added = counts.added != null ? counts.added : (counts.physicalsAdded || 0) + (counts.coursesAdded || 0);
      return { id: r.Id, type: r.UploadType || d.type || "", file: r.Title || d.file || "", at: d.at || r.Created, by: d.by || "",
        counts, added, review: d.review || 0, problems: d.problems || [], truncated: !!d.truncated, warnings: d.warnings || [] };
    });
    container.innerHTML = "";
    const pill = el("span", { class: "count-pill" });
    container.appendChild(el("div", { class: "toolbar" }, [
      el("span", { class: "help", text: "One entry per upload or migration run. Click an entry to see which rows need review." }), pill]));
    const detail = el("div", { style: "margin-top:18px" });
    let selected = null;
    const t = smartTable({
      key: "audit-uploads", rows: items, countEl: pill, noun: "upload",
      placeholder: "Search file, type, or person",
      search: x => [x.file, x.type, x.by].join(" "),
      emptyText: "No uploads recorded yet. Uploads are logged from now on.",
      onRow: x => { selected = x.id; t.refresh(); showDetail(x); },
      rowClass: x => x.id === selected ? "selected" : "",
      columns: [
        { head: "When", tdClass: "nowrap", sort: x => DS.parseDate(x.at), dir: -1, render: x => el("span", { class: "tnum", text: DS.fmtDateTime(x.at) }) },
        { head: "Type", sort: x => x.type, render: x => DS.badge(x.type || "\u2014", "neutral") },
        { head: "File", sort: x => x.file, render: x => x.file || "\u2014" },
        { head: "By", sort: x => DS.emailLocal(x.by), render: x => DS.emailLocal(x.by) || "\u2014" },
        { head: "Rows", thClass: "num", tdClass: "num", sort: x => x.counts.rows || x.counts.employees || 0, render: x => String(x.counts.rows != null ? x.counts.rows : (x.counts.employees || "\u2014")) },
        { head: "Added", thClass: "num", tdClass: "num", sort: x => x.added, render: x => String(x.added) },
        { head: "Health", sort: x => x.review, dir: -1,
          render: x => x.review ? DS.badge(x.review + " to review", "overdue") : DS.badge("Clean", "clear") },
      ],
      facets: [
        { label: "Type", options: distinctOptions(x => x.type), test: (x, v) => x.type === v },
        { label: "Health", options: [["review", "Needs review"], ["clean", "Clean"]], test: (x, v) => (x.review ? "review" : "clean") === v },
      ],
    });
    container.appendChild(el("div", { class: "card" }, [t.bar, t.body]));
    container.appendChild(detail);

    function showDetail(x) {
      detail.innerHTML = "";
      const grid = el("div", { class: "migrate-summary" }, Object.keys(x.counts).filter(k => COUNT_LABELS[k]).map(k => {
        const v = x.counts[k];
        const kind = BAD_COUNTS.includes(k) && v ? "overdue" : null;
        return el("div", { class: "stat" + (kind ? " stat--" + kind : "") }, [el("b", { class: "tnum", text: String(v) }), el("span", { text: COUNT_LABELS[k] })]);
      }));
      const body = el("div", { class: "card__body" }, [grid]);
      x.warnings.forEach(w => body.appendChild(el("div", { class: "import-note warn", text: w })));
      if (x.problems.length && DS.problemsDetails) body.appendChild(DS.problemsDetails(x.problems, x.file || x.type, true));
      else body.appendChild(el("div", { class: "import-note", text: "Nothing to review \u2014 every row was matched and saved." }));
      if (x.truncated) body.appendChild(el("div", { class: "help", text: "This upload had more problem rows than the log can hold; the list above is trimmed. The counts are complete." }));
      detail.appendChild(el("div", { class: "card" }, [
        el("div", { class: "card__head" }, [el("h3", { text: (x.file || x.type) + " \u2014 " + DS.fmtDateTime(x.at) }), DS.badge(x.type, "neutral")]),
        body,
      ]));
      if (detail.scrollIntoView) detail.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }


  (function injectRecordStyles() {
    if (document.getElementById("roster-records-styles")) return;
    const s = document.createElement("style");
    s.id = "roster-records-styles";
    s.textContent =
      ".detail{max-height:calc(100vh - 96px);overflow-y:auto}" +
      ".dash-quad{display:grid;grid-template-columns:1fr 1fr;gap:18px;align-items:start}" +
      ".st-bar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:10px 14px;border-bottom:1px solid var(--line-2)}" +
      ".st-search{flex:1 1 150px;min-width:120px;padding:6px 10px;border:1px solid var(--line);border-radius:6px;font-size:13px;font-family:var(--font)}" +
      ".st-facet{padding:5px 6px;border:1px solid var(--line);border-radius:6px;font-size:12.5px;background:var(--paper);max-width:170px;font-family:var(--font)}" +
      ".st-clear{border:none;background:none;color:var(--navy-500);font-size:12.5px;cursor:pointer;padding:4px 2px}" +
      ".st-sortable{cursor:pointer;user-select:none}" +
      ".st-sortable:hover{color:var(--navy-700)}" +
      ".dash-wide{grid-column:1/-1}" +
      ".dash-scroll{height:340px;overflow-y:auto}" +
      ".dash-wide .dash-scroll{height:auto;max-height:300px}" +
      ".dash-scroll thead th{position:sticky;top:0;background:var(--paper);z-index:1;padding-top:10px}" +
      ".dash-scroll .empty-mini{padding-top:60px}" +
      ".rec-src{font-size:10.5px;color:var(--slate);background:var(--line-2);border-radius:4px;padding:1px 5px}" +
      ".rec-eff{font-weight:600}" +
      ".rec-more{border:none;background:none;color:var(--navy-500);font-size:12.5px;cursor:pointer;padding:6px 0}" +
      "@media (max-width:900px){.dash-quad{grid-template-columns:1fr}}" +
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
