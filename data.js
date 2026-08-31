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
  function addMonths(d, n) { const x = new Date(d); x.setMonth(x.getMonth() + n); return x; }
  function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

  DS.fmtDateTime = function (v) {
    const d = DS.parseDate(v);
    if (!d) return "—";
    return d.toLocaleString("en-US", { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  };

  /* ---- driver designation (job-duty; manually maintained) ---- */
  function designation(r) {
    const s = String((r && r.DriverStatus) || "Primary").trim();
    return (s === "Secondary" || s === "Non-Driver") ? s : "Primary";  // default Primary
  }
  function physicalRequired(r) { return designation(r) === "Primary"; }             // only Primary
  function courseRequired(r) { return designation(r) !== "Non-Driver"; }            // Primary + Secondary

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
      // alert timing
      courseLeadDays: cfgNum(config, "CourseAlertLeadDays", 30),
      physLeadDays: cfgNum(config, "PhysicalAlertLeadDays", 30),
      // compliance rules (safe defaults — app works before these columns exist)
      courseRenewalYears: cfgNum(config, "CourseRenewalYears", 3),
      courseGraceMonths: cfgNum(config, "CourseGraceMonths", 12),
      physicalDefaultYears: cfgNum(config, "PhysicalDefaultYears", 2),
      awardMilestoneYears: cfgNum(config, "AwardMilestoneYears", 5),
      awardResetPointThreshold: cfgNum(config, "AwardResetPointThreshold", 0),
      // driving-eligibility (point-based, rolling window)
      pointRolloffMonths: cfgNum(config, "PointRolloffMonths", 24),
      restrictivePoints: cfgNum(config, "RestrictivePoints", 4),
      noDrivingPoints: cfgNum(config, "NoDrivingPoints", 5),
      activePoints: {},     // emp -> sum of FinalPoints within the roll-off window
      // required courses + tickets
      requiredTitles: String(config.RequiredCourseTitles || DEFAULT_COURSES)
        .split(";").map(t => t.trim()).filter(Boolean),
      ticketCategories: String(config.TicketCategories || "")
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

    const ptsCutoff = addMonths(startOfToday(), -idx.pointRolloffMonths);
    cache.accidents.forEach(a => {
      const emp = empKey(a.EmployeeId);
      if (!emp) return;
      const streak = String(a.CountsAgainstStreak || "").trim();
      const fp = Number(a.FinalPoints) || 0;
      const d = DS.parseDate(a.AccidentDate);
      if (!d) return;
      // awards: streak reset on a qualifying accident
      const qualifies = streak === "Force Yes" || (streak === "Auto" && fp > idx.awardResetPointThreshold);
      if (qualifies && (!idx.qualResetDate[emp] || d > idx.qualResetDate[emp])) idx.qualResetDate[emp] = d;
      // driving points: sum within roll-off window, excluding accidents forced not to count
      if (streak !== "Force No" && d >= ptsCutoff) idx.activePoints[emp] = (idx.activePoints[emp] || 0) + fp;
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
    /* Physical: Primary drivers (the default) are required to have one, so a
       Primary with none on record is "Missing" (overdue-level). Secondary
       drivers aren't required — a missing physical isn't a violation, but an
       expiring one is still surfaced (lower priority). Due date is the
       per-employee expiration if present, else last physical + fallback years. */
    physicalFor(cache, emp) {
      const r = cache.idx.rosterByEmp[emp] || {};
      const desig = designation(r);
      const required = desig === "Primary";        // only Primary requires a physical
      const applicable = desig !== "Non-Driver";   // Non-Driver: physicals don't apply at all
      const p = cache.idx.physLatest[emp];
      const today = startOfToday();
      if (!p) {
        return { has: false, required, applicable, desig, dueDate: null,
                 missing: required, overdue: required,
                 status: required ? "Missing" : (applicable ? "None on record" : "Not applicable") };
      }
      const exp = DS.parseDate(p.ExpirationDate);
      const pd = DS.parseDate(p.PhysicalDate);
      const dueDate = exp || (pd ? addYears(pd, cache.idx.physicalDefaultYears) : null);
      const overdue = !!dueDate && dueDate < today;
      return { has: true, required, applicable, desig, dueDate, missing: false, overdue,
               lastDate: pd, fromExpiry: !!exp, status: overdue ? "Overdue" : "Current" };
    },

    /* Courses: cycle status + due date, using the required-course set. */
    coursesFor(cache, emp) {
      const r = cache.idx.rosterByEmp[emp] || {};
      const applicable = designation(r) !== "Non-Driver";   // courses don't apply to Non-Drivers
      const req = cache.idx.requiredTitles;
      const cm = cache.idx.courseLatest[emp] || {};
      const dates = req.map(t => cm[t] || null);
      const anyBlank = dates.some(d => !d);
      const allBlank = dates.every(d => !d);
      const hire = DS.parseDate(r.HireDate);
      let status, dueDate;
      if (anyBlank) {
        status = allBlank ? "Not started" : "Incomplete";
        dueDate = hire ? addMonths(hire, cache.idx.courseGraceMonths) : null;   // grace from hire
      } else {
        status = "Renewal due";
        const minTime = Math.min.apply(null, dates.map(d => d.getTime()));
        dueDate = addYears(new Date(minTime), cache.idx.courseRenewalYears);      // renewal cycle
      }
      return { status, dueDate, applicable };
    },

    /* Driving eligibility from active points (rolling window). */
    drivingStatusFor(cache, emp) {
      const pts = cache.idx.activePoints[emp] || 0;
      let status = "Normal";
      if (pts >= cache.idx.noDrivingPoints) status = "No-Driving";
      else if (pts >= cache.idx.restrictivePoints) status = "Restrictive";
      return { points: pts, status };
    },

    /* Award: next milestone + the date it's reached. */
    awardFor(cache, emp) {
      const hire = DS.parseDate((cache.idx.rosterByEmp[emp] || {}).HireDate);
      const reset = cache.idx.qualResetDate[emp] || hire;
      const highest = cache.idx.highestAward[emp] || 0;
      const next = highest + cache.idx.awardMilestoneYears;
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
        if (!s.applicable) return;                                     // Non-Driver: physicals don't apply
        let urgency = null;
        if (s.required && s.missing) urgency = "missing";               // Primary, none on record
        else if (s.dueDate && s.dueDate <= cutoff) urgency = s.overdue ? "overdue" : "due";
        if (!urgency) return;                                            // Secondary w/ no physical is not flagged here
        out.push({
          employeeId: emp, name: r.Title, dueDate: s.dueDate,
          required: s.required, urgency,
          overdue: urgency === "overdue" || urgency === "missing",
        });
      });
      const rank = u => ({ missing: 0, overdue: 1, due: 2 }[u]);
      out.sort((a, b) => {
        if (a.required !== b.required) return a.required ? -1 : 1;      // Primary first
        if (rank(a.urgency) !== rank(b.urgency)) return rank(a.urgency) - rank(b.urgency);
        const ad = a.dueDate ? a.dueDate.getTime() : Infinity;
        const bd = b.dueDate ? b.dueDate.getTime() : Infinity;
        return ad - bd;
      });
      return out;
    },

    coursesDue(cache) {
      const today = startOfToday();
      const cutoff = addDays(today, cache.idx.courseLeadDays);
      const out = [];
      cache.idx.activeRoster.forEach(r => {
        const emp = empKey(r.EmployeeId);
        const s = this.coursesFor(cache, emp);
        if (!s.applicable) return;                                     // Non-Driver: courses don't apply
        if (s.dueDate && s.dueDate <= cutoff) {
          out.push({ employeeId: emp, name: r.Title, dueDate: s.dueDate, status: s.status, overdue: s.dueDate < today });
        }
      });
      out.sort((a, b) => a.dueDate - b.dueDate);
      return out;
    },

    /* Employees currently on Restrictive or No-Driving status. */
    drivingRestricted(cache) {
      const out = [];
      cache.idx.activeRoster.forEach(r => {
        const emp = empKey(r.EmployeeId);
        const s = this.drivingStatusFor(cache, emp);
        if (s.status !== "Normal") out.push({ employeeId: emp, name: r.Title, points: s.points, status: s.status });
      });
      out.sort((a, b) => {
        if (a.status !== b.status) return a.status === "No-Driving" ? -1 : 1;  // No-Driving first
        return b.points - a.points;
      });
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
  DS.util = { addYears, addMonths, addDays, startOfToday, empKey, isActive, designation, physicalRequired, courseRequired };

})();
