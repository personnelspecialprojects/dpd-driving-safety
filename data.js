/* ============================================================
   data.js — shared data layer + compliance engine
   Loaded after the foundation (index.html). Depends on DS.*

   FIELD NAMES: these come from the documented schema and the field
   names that worked in the Power Automate flows. The one field not yet
   confirmed against a live list is Physicals → "ExpirationDate" (the new
   per-employee expiry). To avoid the "$select on a missing field returns
   400" landmine, loads fetch ALL fields (no $select) — robust but heavier.
   Once fields are verified with DS.inspectList(...), we can add $select
   for performance. If a field name here is wrong, screens will show blanks
   rather than error, and the fix is a one-word edit in this file.
   ============================================================ */
(function () {

  /* ---- date helpers ---- */
  function addYears(d, n) { const x = new Date(d); x.setFullYear(x.getFullYear() + n); return x; }
  function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

  DS.fmtDateTime = function (v) {
    const d = DS.parseDate(v);
    if (!d) return "—";
    return d.toLocaleString("en-US", { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  };

  /* ---- misc row helpers ---- */
  function empKey(v) { return String(v == null ? "" : v).trim(); }
  function isActive(r) {
    const v = r.ActiveEmployee;
    return v === true || v === 1 || String(v).toLowerCase() === "yes";
  }
  function cfgNum(config, key, dflt) {
    const v = config[key];
    return (v === 0 || v) ? Number(v) : dflt;
  }

  const DEFAULT_COURSES = "Defensive Driving Basics;Defensive Driving Principles;Distracted Driving For Law Enforcement";

  /* ============================================================
     DS.data — load + cache the core lists, build lookup indexes
     ============================================================ */
  DS.data = {
    cache: null,

    async load(force) {
      if (this.cache && !force) return this.cache;
      const L = DS.LISTS;
      const [roster, courses, physicals, accidents, awards, configArr] = await Promise.all([
        DS.spGet(L.roster),
        DS.spGet(L.courses),
        DS.spGet(L.physicals),
        DS.spGet(L.accidents),
        DS.spGet(L.awards),
        DS.spGet(L.config, { top: 1 }),
      ]);
      const config = configArr[0] || {};
      const cache = { roster, courses, physicals, accidents, awards, config };
      cache.idx = buildIndexes(cache);
      this.cache = cache;
      return cache;
    },

    clear() { this.cache = null; },
  };

  function buildIndexes(cache) {
    const config = cache.config;
    const idx = {
      courseLeadDays: cfgNum(config, "CourseAlertLeadDays", 30),
      physLeadDays: cfgNum(config, "PhysicalAlertLeadDays", 30),
      requiredTitles: String(config.RequiredCourseTitles || DEFAULT_COURSES)
        .split(";").map(t => t.trim()).filter(Boolean),
      allowMark: config.AllowMarkAsAwarded === true || String(config.AllowMarkAsAwarded).toLowerCase() === "yes",
      rosterByEmp: {},
      activeRoster: [],
      courseLatest: {},     // emp -> { title -> Date }
      physLatest: {},       // emp -> physical record (most recent by PhysicalDate)
      qualResetDate: {},    // emp -> Date (latest qualifying accident)
      highestAward: {},     // emp -> number
    };

    cache.roster.forEach(r => {
      const emp = empKey(r.EmployeeId);
      if (emp) idx.rosterByEmp[emp] = r;
      if (isActive(r)) idx.activeRoster.push(r);
    });

    cache.courses.forEach(c => {
      const emp = empKey(c.EmployeeId);
      const title = String(c.CourseTitle || "").trim();
      const d = DS.parseDate(c.DateCompleted);
      if (!emp || !title || !d) return;
      const m = (idx.courseLatest[emp] = idx.courseLatest[emp] || {});
      if (!m[title] || d > m[title]) m[title] = d;
    });

    cache.physicals.forEach(p => {
      const emp = empKey(p.EmployeeId);
      const d = DS.parseDate(p.PhysicalDate);
      if (!emp) return;
      const cur = idx.physLatest[emp];
      const curD = cur ? DS.parseDate(cur.PhysicalDate) : null;
      if (!cur || (d && (!curD || d > curD))) idx.physLatest[emp] = p;
    });

    cache.accidents.forEach(a => {
      const emp = empKey(a.EmployeeId);
      const streak = String(a.CountsAgainstStreak || "").trim();
      const fp = Number(a.FinalPoints) || 0;
      const qualifies = streak === "Force Yes" || (streak === "Auto" && fp > 0);
      if (!emp || !qualifies) return;
      const d = DS.parseDate(a.AccidentDate);
      if (!d) return;
      if (!idx.qualResetDate[emp] || d > idx.qualResetDate[emp]) idx.qualResetDate[emp] = d;
    });

    cache.awards.forEach(a => {
      const emp = empKey(a.EmployeeId);
      const m = Number(a.MilestoneYears) || 0;
      if (!emp) return;
      if (!idx.highestAward[emp] || m > idx.highestAward[emp]) idx.highestAward[emp] = m;
    });

    return idx;
  }

  /* ============================================================
     DS.compute — per-employee status + dashboard "due" lists.
     Per-employee functions power the Roster detail panel; the
     due-list functions power the Dashboard.
     ============================================================ */
  DS.compute = {

    /* Physical: due date is the per-employee expiration date if present,
       otherwise last physical + 2 years. Employees with no physical on
       record return null (not flagged here — that's a separate
       "who is required to have one" question). */
    physicalFor(cache, emp) {
      const p = cache.idx.physLatest[emp];
      if (!p) return { dueDate: null, has: false };
      const exp = DS.parseDate(p.ExpirationDate);
      const pd = DS.parseDate(p.PhysicalDate);
      const dueDate = exp || (pd ? addYears(pd, 2) : null);
      return { dueDate, has: true, lastDate: pd, fromExpiry: !!exp };
    },

    /* Courses: cycle status + due date, using the required-course set. */
    coursesFor(cache, emp) {
      const req = cache.idx.requiredTitles;
      const cm = cache.idx.courseLatest[emp] || {};
      const dates = req.map(t => cm[t] || null);
      const anyBlank = dates.some(d => !d);
      const allBlank = dates.every(d => !d);
      const hire = DS.parseDate((cache.idx.rosterByEmp[emp] || {}).HireDate);
      let status, dueDate;
      if (anyBlank) {
        status = allBlank ? "Not started" : "Incomplete";
        dueDate = hire ? addYears(hire, 1) : null;   // 1-year grace from hire
      } else {
        status = "Renewal due";
        const minTime = Math.min.apply(null, dates.map(d => d.getTime()));
        dueDate = addYears(new Date(minTime), 3);      // 3-year renewal cycle
      }
      return { status, dueDate };
    },

    /* Award: next milestone + the date it's reached. */
    awardFor(cache, emp) {
      const hire = DS.parseDate((cache.idx.rosterByEmp[emp] || {}).HireDate);
      const reset = cache.idx.qualResetDate[emp] || hire;
      const highest = cache.idx.highestAward[emp] || 0;
      const next = highest + 5;
      const eligibleDate = reset ? addYears(reset, next) : null;
      const eligible = !!eligibleDate && eligibleDate <= startOfToday();
      return { nextMilestone: next, highest, resetDate: reset, eligibleDate, eligible };
    },

    /* ---- Dashboard due-lists (active employees only) ---- */
    physicalsDue(cache) {
      const today = startOfToday();
      const cutoff = addDays(today, cache.idx.physLeadDays);
      const out = [];
      cache.idx.activeRoster.forEach(r => {
        const emp = empKey(r.EmployeeId);
        const s = this.physicalFor(cache, emp);
        if (s.dueDate && s.dueDate <= cutoff) {
          out.push({ employeeId: emp, name: r.Title, dueDate: s.dueDate, overdue: s.dueDate < today });
        }
      });
      out.sort((a, b) => a.dueDate - b.dueDate);
      return out;
    },

    coursesDue(cache) {
      const today = startOfToday();
      const cutoff = addDays(today, cache.idx.courseLeadDays);
      const out = [];
      cache.idx.activeRoster.forEach(r => {
        const emp = empKey(r.EmployeeId);
        const s = this.coursesFor(cache, emp);
        if (s.dueDate && s.dueDate <= cutoff) {
          out.push({ employeeId: emp, name: r.Title, dueDate: s.dueDate, status: s.status, overdue: s.dueDate < today });
        }
      });
      out.sort((a, b) => a.dueDate - b.dueDate);
      return out;
    },

    awardsEligible(cache) {
      const out = [];
      cache.idx.activeRoster.forEach(r => {
        const emp = empKey(r.EmployeeId);
        const s = this.awardFor(cache, emp);
        if (s.eligible) {
          out.push({ employeeId: emp, name: r.Title, nextMilestone: s.nextMilestone, eligibleDate: s.eligibleDate });
        }
      });
      out.sort((a, b) => a.eligibleDate - b.eligibleDate);
      return out;
    },
  };

  /* small shared exports for screens */
  DS.util = { addYears, addDays, startOfToday, empKey, isActive };

})();
