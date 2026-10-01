/* ============================================================
   record-export.js — per-employee driving record (PDF)
   ------------------------------------------------------------
   Roster → employee → "Export record (PDF)". The user picks any
   combination of sections:
     • Driving history (accidents, points, awards)
     • Defensive driving course history
     • Physical history
     • Ticket history (tickets whose subject is this employee)
   The PDF is built in the browser with jsPDF (+ AutoTable), loaded
   like SheetJS: local repo copy first, then CDNs. If the library can't
   load (e.g. a blocked CDN), a printable page opens instead so the
   user can choose "Save as PDF". Every export is written to the
   Audit log (who, whose record, which sections).
   Loaded after screens-admin.js (uses DS.modal / DS.formField).
   ============================================================ */
(function () {
  const el = DS.el;
  const L = DS.LISTS;

  const SECTIONS = [
    { key: "driving", label: "Driving history", help: "Accidents, points, driving status, safe-driving awards" },
    { key: "courses", label: "Defensive driving course history", help: "Every course completion on file" },
    { key: "physicals", label: "Physical history", help: "Every driver physical on file" },
    { key: "tickets", label: "Service ticket history", help: "Requests to the Driving Safety team where this employee was the subject (not traffic citations)" },
  ];

  const SOURCE_LABELS = {
    "Bulk Upload": "Upload",
    "Manual Entry": "Entered manually",
    "Legacy History": "Exam history (program records)",
    "Master Sheet": "Program spreadsheet (estimate)",
    "Legacy Migration": "Program spreadsheet (estimate)",
  };
  const srcLabel = rec => SOURCE_LABELS[rec.Source] || (rec.Source || "\u2014");
  const fmt = v => { const d = DS.parseDate(v); return d ? DS.fmtDate(d) : "\u2014"; };
  const byDateDesc = get => (a, b) => (DS.parseDate(get(b)) || 0) - (DS.parseDate(get(a)) || 0);
  const K = v => DS.util.empKey(v);

  /* ============================================================
     1) MODEL — plain data describing the report (easy to test)
     ============================================================ */
  function buildModel(cache, r, tickets, opts, now) {
    now = now || new Date();
    const emp = K(r.EmployeeId);
    const today = DS.util.startOfToday();
    const designation = DS.util.isUnassigned(r) ? "Unassigned" : DS.util.designation(r);
    const who = (DS.me && (DS.me.displayName || DS.me.mail || DS.me.userPrincipalName)) || "";
    const m = {
      title: "Dallas Police Department",
      subtitle: "Driving Safety Program \u2014 Employee Driving Record",
      generated: DS.fmtDate(now) + (who ? " by " + who : ""),
      name: r.Title || ("Employee " + r.EmployeeId),
      meta: [
        ["Employee #", String(r.EmployeeId || "\u2014")],
        ["Badge", String(r.Badge || "") || "\u2014"],
        ["Rank", r.Rank || "\u2014"],
        ["Division", r.Division || "\u2014"],
        ["Assignment", r.Assignment || "\u2014"],
        ["Hire / service date", fmt(r.HireDate)],
        ["Driver designation", designation],
        ["Employment", DS.util.isActive(r) ? "Active" : "Separated"],
      ],
      sections: [],
    };

    /* ---- Driving history ---- */
    if (opts.driving) {
      const drive = DS.compute.drivingStatusFor(cache, emp);
      const awd = DS.compute.awardFor(cache, emp);
      const coverage = cache.idx.accidentRecordsStart;
      // Same terms as the app: the designation, plus a restriction only if one applies.
      let statusText = designation;
      if (drive.status === "Restrictive" || drive.status === "No-Driving") {
        statusText += " \u2014 " + drive.status + (drive.override
          ? " (set by the Driving Safety team" + (drive.overrideUntil ? " until " + DS.fmtDate(drive.overrideUntil) : "") +
            (drive.overrideNote ? ": " + drive.overrideNote : "") + ")"
          : " (" + drive.points + " active accident point" + (drive.points === 1 ? "" : "s") + ")");
      }
      const awards = (cache.awards || []).filter(a => K(a.EmployeeId) === emp).sort(byDateDesc(a => a.AwardDate));
      const nextAward = awd.paused ? "Not available \u2014 accident records have not been loaded"
        : awd.noHire ? "Not available \u2014 no hire date on file"
        : awd.eligible ? "Eligible for the " + awd.nextMilestone + "-year award (since " + DS.fmtDate(awd.eligibleDate) + ")"
        : awd.nextMilestone + "-year award on " + DS.fmtDate(awd.eligibleDate);
      const accidents = (cache.idx.accidentsByEmp[emp] || []).slice().sort(byDateDesc(a => a.AccidentDate));
      m.sections.push({
        key: "driving", title: "Driving History",
        summary: [
          ["Driving status", statusText],
          ["Active accident points", String(drive.points) + " (points count for " + cache.idx.pointRolloffMonths + " months)"],
          ["Accident records on file", coverage
            ? "From " + DS.fmtDate(coverage) + (cache.idx.accidentRecordsSource === "setting" ? " (program setting)" : " (oldest accident on file)")
            : "No accident records have been loaded"],
          ["Safe-driving awards received", awards.length
            ? awards.map(a => (a.MilestoneYears || "?") + "-year (" + fmt(a.AwardDate) + ")").join(", ") : "None on record"],
          ["Next safe-driving award", nextAward],
        ],
        note: coverage ? "Accidents before " + DS.fmtDate(coverage) + " are not in the program's records, so this history only covers that date forward." : null,
        head: ["Incident #", "Date", "Final points", "Counts", "Status now", "Decision / location"],
        widths: { 0: 88, 1: 72, 2: 62, 3: 52, 4: 92 },
        rows: accidents.map(a => {
          const d = DS.parseDate(a.AccidentDate);
          const counts = String(a.CountsAgainstStreak || "Auto");
          const state = counts === "Force No" ? "Excluded"
            : (d && d >= cache.idx.ptsCutoff ? "Counts toward points" : "Rolled off");
          return [a.IncidentNumber || "\u2014", fmt(a.AccidentDate), String(Number(a.FinalPoints) || 0), counts, state,
            [a.FinalDecision, a.Location].filter(Boolean).join(" \u2014 ") || "\u2014"];
        }),
        emptyText: "No accidents on record" + (coverage ? " since " + DS.fmtDate(coverage) : "") + ".",
      });
    }

    /* ---- Defensive driving courses ---- */
    if (opts.courses) {
      const crs = DS.compute.coursesFor(cache, emp);
      let status;
      if (!crs.applicable) status = "Not required (Non-Driver)";
      else if (crs.doneCount === crs.requiredCount) {
        status = !crs.dueDate ? "Complete"
          : crs.dueDate < today ? "Renewal overdue since " + DS.fmtDate(crs.dueDate)
          : "Current \u2014 next renewal due " + DS.fmtDate(crs.dueDate);
      } else {
        status = crs.status + " (" + crs.doneCount + " of " + crs.requiredCount + " required courses)" +
          (crs.dueDate ? (crs.dueDate < today ? " \u2014 overdue since " : " \u2014 due ") + DS.fmtDate(crs.dueDate) : "");
      }
      const courses = (cache.idx.coursesByEmp[emp] || []).slice().sort(byDateDesc(c => c.DateCompleted));
      const hasEst = courses.some(c => DS.util.isFallbackSource(c));
      m.sections.push({
        key: "courses", title: "Defensive Driving Course History",
        summary: [
          ["Requirement", crs.applicable ? "Required" : "Not required (Non-Driver)"],
          ["Status", status],
          ["Required courses", cache.idx.requiredTitles.join("; ")],
        ],
        note: hasEst ? "Entries marked \u201Cestimate\u201D come from the program's earlier spreadsheet, which recorded due dates rather than completion dates; the completion date shown is estimated from that due date." : null,
        head: ["Course", "Completed", "Source"],
        widths: { 1: 110, 2: 180 },
        rows: courses.map(c => [c.CourseTitle || "\u2014",
          (DS.util.isFallbackSource(c) ? "Est. " : "") + fmt(c.DateCompleted),
          srcLabel(c)]),
        emptyText: "No course completions on record.",
      });
    }

    /* ---- Physicals ---- */
    if (opts.physicals) {
      const ph = DS.compute.physicalFor(cache, emp);
      const status = !ph.applicable ? "Not required (Non-Driver)"
        : ph.inGrace ? "New hire \u2014 first physical due " + DS.fmtDate(ph.dueDate)
        : !ph.has ? (ph.required ? "No physical on record" + (ph.graceEnd ? " (was due " + DS.fmtDate(ph.graceEnd) + ")" : "") : "None on record (not required)")
        : !ph.dueDate ? "On record \u2014 no expiration date"
        : ph.overdue ? "Expired " + DS.fmtDate(ph.dueDate)
        : "Current \u2014 expires " + DS.fmtDate(ph.dueDate);
      const phys = (cache.idx.physicalsByEmp[emp] || []).slice()
        .sort(byDateDesc(p => p.PhysicalDate || p.ExpirationDate));
      m.sections.push({
        key: "physicals", title: "Physical History",
        summary: [
          ["Requirement", ph.required ? "Required (Primary driver)" : ph.applicable ? "Not required (Secondary driver)" : "Not required (Non-Driver)"],
          ["Status", status],
        ],
        note: null,
        head: ["Date tested", "Result", "Expires / due"],
        widths: { 0: 150, 1: 150 },
        rows: phys.map(p => [p.PhysicalDate ? fmt(p.PhysicalDate) : "\u2014",
          p.Result || (p.PhysicalDate ? "\u2014" : "Due date on file"), fmt(p.ExpirationDate)]),
        emptyText: "No physicals on record.",
      });
    }

    /* ---- Tickets ---- */
    if (opts.tickets) {
      const failed = tickets && tickets.error;
      const mine = failed ? [] : (tickets || []).filter(t => K(t.EmployeeId) === emp).sort(byDateDesc(t => t.OpenedOn || t.Created));
      m.sections.push({
        key: "tickets", title: "Service Ticket History",
        summary: [["Service tickets on record", failed ? "Could not be loaded" : String(mine.length)]],
        note: failed ? "Service tickets could not be loaded from SharePoint (" + tickets.error + "). Try the export again."
          : "Service tickets are requests handled by the Driving Safety team. They are not traffic citations.",
        head: ["Opened", "Subject", "Status", "Resolved", "Details"],
        widths: { 0: 72, 1: 130, 2: 58, 3: 72 },
        rows: mine.map(t => [fmt(t.OpenedOn || t.Created), t.Title || "\u2014", t.Status || "Open", fmt(t.ResolvedOn),
          [t.Category, t.Description, t.ResolutionNotes ? "Resolution: " + t.ResolutionNotes : ""].filter(Boolean).join(" \u2014 ") || "\u2014"]),
        emptyText: failed ? "" : "No service tickets on record for this employee.",
      });
    }
    return m;
  }

  /* ============================================================
     2) PDF — jsPDF + AutoTable
     ============================================================ */
  // The PDF's built-in fonts only cover Western-European characters; swap anything else.
  function T(s) {
    return String(s == null ? "" : s)
      .replace(/[\u2192\u2794]/g, "->").replace(/[\u2713\u2714]/g, "")
      .replace(/[^\x00-\xFF\u2013\u2014\u2018\u2019\u201C\u201D\u2022\u2026\u00B7]/g, ch => {
        const base = ch.normalize("NFD").replace(/[\u0300-\u036f]/g, "");     // e.g. \u1EC5 -> e
        return /^[\x00-\xFF]*$/.test(base) ? base : "";
      });
  }

  function renderPdf(jsPDF, m) {
    const doc = new jsPDF({ unit: "pt", format: "letter" });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight(), M = 40;
    const NAVY = [28, 53, 89], BRASS = [176, 141, 87], SLATE = [90, 100, 120], INK = [26, 31, 43];

    // Header band
    doc.setFillColor(...NAVY); doc.rect(0, 0, W, 62, "F");
    doc.setDrawColor(...BRASS); doc.setLineWidth(2.5); doc.line(0, 62, W, 62);
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold"); doc.setFontSize(15); doc.text(T(m.title), M, 29);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.text(T(m.subtitle), M, 46);
    doc.text(T("Generated " + m.generated), W - M, 46, { align: "right" });

    // Employee block
    let y = 90;
    doc.setTextColor(...INK); doc.setFont("helvetica", "bold"); doc.setFontSize(14);
    doc.text(T(m.name), M, y);
    const pairs = [];
    for (let i = 0; i < m.meta.length; i += 2) {
      const a = m.meta[i], b = m.meta[i + 1] || ["", ""];
      pairs.push([T(a[0]), T(a[1]), T(b[0]), T(b[1])]);
    }
    doc.autoTable({
      startY: y + 8, theme: "plain", body: pairs, margin: { left: M, right: M },
      styles: { fontSize: 9, cellPadding: { top: 2, bottom: 2, left: 0, right: 6 }, textColor: INK },
      columnStyles: { 0: { fontStyle: "bold", textColor: SLATE, cellWidth: 100 }, 1: { cellWidth: 170 },
                      2: { fontStyle: "bold", textColor: SLATE, cellWidth: 105 } },
    });
    y = doc.lastAutoTable.finalY + 22;

    m.sections.forEach(sec => {
      if (y > H - 140) { doc.addPage(); y = 54; }
      doc.setFont("helvetica", "bold"); doc.setFontSize(12.5); doc.setTextColor(...NAVY);
      doc.text(T(sec.title), M, y);
      doc.setDrawColor(...BRASS); doc.setLineWidth(1); doc.line(M, y + 5, W - M, y + 5);
      doc.autoTable({
        startY: y + 11, theme: "plain", body: sec.summary.map(p => [T(p[0]), T(p[1])]), margin: { left: M, right: M },
        styles: { fontSize: 9.5, cellPadding: { top: 2.5, bottom: 2.5, left: 0, right: 8 }, textColor: INK },
        columnStyles: { 0: { fontStyle: "bold", textColor: SLATE, cellWidth: 170 } },
      });
      y = doc.lastAutoTable.finalY + 6;
      if (sec.note) {
        doc.setFont("helvetica", "italic"); doc.setFontSize(8.5); doc.setTextColor(...SLATE);
        const lines = doc.splitTextToSize(T(sec.note), W - 2 * M);
        if (y + lines.length * 11 > H - 60) { doc.addPage(); y = 54; }
        doc.text(lines, M, y + 8); y += lines.length * 11 + 6;
      }
      if (sec.rows.length) {
        doc.autoTable({
          startY: y + 4, head: [sec.head.map(T)], body: sec.rows.map(row => row.map(T)),
          margin: { left: M, right: M, top: 50, bottom: 50 }, theme: "striped",
          styles: { fontSize: 8.5, cellPadding: 4, textColor: INK, overflow: "linebreak" },
          headStyles: { fillColor: NAVY, textColor: 255, fontStyle: "bold" },
          alternateRowStyles: { fillColor: [244, 246, 249] },
          columnStyles: Object.fromEntries(Object.entries(sec.widths || {}).map(([k, v]) => [k, { cellWidth: v }])),
        });
        y = doc.lastAutoTable.finalY + 26;
      } else if (sec.emptyText) {
        doc.setFont("helvetica", "italic"); doc.setFontSize(9.5); doc.setTextColor(...SLATE);
        doc.text(T(sec.emptyText), M, y + 12); y += 34;
      } else y += 20;
    });

    // Footer on every page
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      doc.setDrawColor(223, 227, 234); doc.setLineWidth(0.5); doc.line(M, H - 38, W - M, H - 38);
      doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...SLATE);
      doc.text(T(m.name + " \u2014 Employee #" + (m.meta[0][1] || "") + " \u2014 Official use only"), M, H - 24);
      doc.text("Page " + i + " of " + n, W - M, H - 24, { align: "right" });
    }
    return doc;
  }

  /* ---- Load jsPDF + AutoTable (local copy, then CDNs) ---- */
  function loadScript(srcs) {
    return new Promise((resolve, reject) => {
      (function tryAt(i) {
        if (i >= srcs.length) { reject(new Error("couldn't load " + srcs[srcs.length - 1].split("/").pop())); return; }
        const s = document.createElement("script");
        s.src = srcs[i]; s.onload = () => resolve(); s.onerror = () => { s.remove(); tryAt(i + 1); };
        document.head.appendChild(s);
      })(0);
    });
  }
  let pdfLibPromise = null;
  function ensurePdfLib() {
    const ready = () => window.jspdf && window.jspdf.jsPDF && window.jspdf.jsPDF.API && window.jspdf.jsPDF.API.autoTable;
    if (ready()) return Promise.resolve(window.jspdf.jsPDF);
    if (!pdfLibPromise) {
      pdfLibPromise = (async () => {
        if (!(window.jspdf && window.jspdf.jsPDF)) await loadScript([
          "./jspdf.umd.min.js",
          "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
          "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
        ]);
        await loadScript([
          "./jspdf.plugin.autotable.min.js",
          "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js",
          "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js",
        ]);
        if (!ready()) throw new Error("PDF library loaded incompletely");
        return window.jspdf.jsPDF;
      })().catch(e => { pdfLibPromise = null; throw e; });
    }
    return pdfLibPromise;
  }

  /* ============================================================
     3) PRINT FALLBACK — same content as a printable page
     ============================================================ */
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function renderHtml(m) {
    const kv = rows => "<table class=kv>" + rows.map(p => "<tr><th>" + esc(p[0]) + "</th><td>" + esc(p[1]) + "</td></tr>").join("") + "</table>";
    return "<!DOCTYPE html><html><head><meta charset=utf-8><title>" + esc(m.name) + " \u2014 Driving Record</title><style>" +
      "body{font-family:Segoe UI,Arial,sans-serif;color:#1a1f2b;margin:32px;font-size:12px}" +
      "header{background:#1c3559;color:#fff;padding:14px 18px;border-bottom:3px solid #b08d57;margin:-32px -32px 20px}" +
      "header b{font-size:18px;display:block}h1{font-size:18px;margin:0 0 8px}h2{font-size:14px;color:#1c3559;border-bottom:1px solid #b08d57;padding-bottom:3px;margin:22px 0 8px}" +
      "table{border-collapse:collapse;width:100%;margin:4px 0}.kv th{text-align:left;color:#5a6478;width:190px;padding:2px 8px 2px 0;vertical-align:top}.kv td{padding:2px 0}" +
      ".grid th{background:#1c3559;color:#fff;text-align:left;padding:5px}.grid td{padding:5px;border-bottom:1px solid #e3e7ee;vertical-align:top}" +
      ".note{font-style:italic;color:#5a6478;font-size:11px}footer{margin-top:24px;color:#5a6478;font-size:10px}" +
      "@media print{body{margin:0}header{margin:0 0 16px}}</style></head><body>" +
      "<header><b>" + esc(m.title) + "</b>" + esc(m.subtitle) + " \u00b7 Generated " + esc(m.generated) + "</header>" +
      "<h1>" + esc(m.name) + "</h1>" + kv(m.meta) +
      m.sections.map(sec => "<h2>" + esc(sec.title) + "</h2>" + kv(sec.summary) +
        (sec.note ? "<p class=note>" + esc(sec.note) + "</p>" : "") +
        (sec.rows.length
          ? "<table class=grid><thead><tr>" + sec.head.map(h => "<th>" + esc(h) + "</th>").join("") + "</tr></thead><tbody>" +
            sec.rows.map(r => "<tr>" + r.map(c => "<td>" + esc(c) + "</td>").join("") + "</tr>").join("") + "</tbody></table>"
          : (sec.emptyText ? "<p class=note>" + esc(sec.emptyText) + "</p>" : ""))).join("") +
      "<footer>Official use only.</footer></body></html>";
  }
  function printHtml(html) {
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    document.body.appendChild(frame);
    const d = frame.contentWindow.document;
    d.open(); d.write(html); d.close();
    setTimeout(() => { try { frame.contentWindow.focus(); frame.contentWindow.print(); } catch (_) {} setTimeout(() => frame.remove(), 60000); }, 300);
  }

  /* ============================================================
     4) EXPORT — gather data, build, save, audit
     ============================================================ */
  async function loadTickets() {
    try { return await DS.spGet(L.tickets); }            // all columns; no Person expansion needed here
    catch (e) { return { error: e.message }; }
  }

  function fileName(r) {
    const nm = String(r.Title || "").trim().split(/\s+/);
    const last = r.LastName || nm[nm.length - 1] || "";
    const first = nm.filter(x => x !== last).join(" ");
    const who = (last && first ? last + ", " + first : (r.Title || "Employee")).replace(/[\\/:*?"<>|]/g, "");
    return "Driving Record - " + who + " - " + String(r.EmployeeId || "").replace(/[\\/:*?"<>|]/g, "") + " - " + DS.todayIso() + ".pdf";
  }

  async function exportRecord(r, opts) {
    const cache = await DS.data.load();
    const tickets = opts.tickets ? await loadTickets() : null;
    const model = buildModel(cache, r, tickets, opts);
    const chosen = SECTIONS.filter(s => opts[s.key]).map(s => s.label).join(", ");
    let method = "pdf";
    try {
      const jsPDF = await ensurePdfLib();
      renderPdf(jsPDF, model).save(fileName(r));
    } catch (e) {
      console.warn("PDF library unavailable, using print view:", e.message);
      method = "print";
      printHtml(renderHtml(model));
    }
    await DS.audit("Record exported", L.roster, r.EmployeeId, (r.Title || r.EmployeeId) + ": " + chosen + (method === "print" ? " (print view)" : ""));
    return method;
  }

  /* ---- Dialog: pick any combination of sections ---- */
  function openDialog(r) {
    const boxes = {};
    const list = el("div", null, SECTIONS.map(s => {
      const cb = el("input", { type: "checkbox" }); cb.checked = true; boxes[s.key] = cb;
      return el("label", { class: "check", style: "align-items:flex-start; margin-bottom:12px" }, [cb,
        el("span", null, [el("b", { text: s.label, style: "display:block; font-size:14px; font-weight:600" }),
          el("span", { class: "help", text: s.help })])]);
    }));
    const note = el("div", { class: "help", style: "margin-top:4px", text: "Choose one or more sections. The PDF downloads to this computer and the export is recorded in the Audit log." });
    const goBtn = el("button", { class: "btn", type: "button", text: "Create PDF" });
    const cancelBtn = el("button", { class: "btn btn--ghost", type: "button", text: "Cancel" });
    const m = DS.modal("Export driving record \u2014 " + (r.Title || r.EmployeeId), el("div", null, [list, note]), [cancelBtn, goBtn]);
    cancelBtn.addEventListener("click", m.close);
    const sync = () => { goBtn.disabled = !SECTIONS.some(s => boxes[s.key].checked); };
    Object.values(boxes).forEach(cb => cb.addEventListener("change", sync));
    goBtn.addEventListener("click", async () => {
      const opts = {}; SECTIONS.forEach(s => { opts[s.key] = boxes[s.key].checked; });
      if (!SECTIONS.some(s => opts[s.key])) { DS.toast("Choose at least one section."); return; }
      goBtn.disabled = true; goBtn.textContent = "Creating\u2026";
      try {
        const method = await exportRecord(r, opts);
        m.close();
        DS.toast(method === "pdf" ? "Driving record PDF downloaded." : "Print view opened \u2014 choose \u201CSave as PDF\u201D as the printer.", "success");
      } catch (e) {
        goBtn.disabled = false; goBtn.textContent = "Create PDF";
        DS.toast("Couldn't create the record: " + e.message, "error");
      }
    });
  }

  DS.recordExport = { open: openDialog, exportRecord, buildModel, renderPdf, renderHtml, SECTIONS };
})();
