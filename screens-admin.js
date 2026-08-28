/* ============================================================
   screens-admin.js — Admin, Tickets, Reports (Excel export)
   Loaded after imports.js (uses DS.ensureXlsx for the export).
   ============================================================ */
(function () {
  const el = DS.el;
  const L = DS.LISTS;

  /* ---- reusable modal ---- */
  function modal(title, bodyNode, footNodes) {
    const backdrop = el("div", { class: "modal-backdrop" });
    const close = () => backdrop.remove();
    const box = el("div", { class: "modal" }, [
      el("div", { class: "modal__head" }, [
        el("h3", { text: title }),
        el("button", { class: "x-btn", text: "×", onclick: close }),
      ]),
      el("div", { class: "modal__body" }, [bodyNode]),
      footNodes ? el("div", { class: "modal__foot" }, footNodes) : null,
    ]);
    box.addEventListener("click", e => e.stopPropagation());
    backdrop.addEventListener("click", close);
    document.body.appendChild(backdrop);
    return { close, backdrop };
  }

  function field(labelText, input, help) {
    return el("div", { class: "form-group" }, [
      el("label", { class: "label", text: labelText }),
      input,
      help ? el("div", { class: "help", text: help }) : null,
    ]);
  }
  function emptyMini(msg) { return el("div", { class: "empty-mini", text: msg }); }

  /* ============================================================
     ADMIN — edit the single Config item
     ============================================================ */
  async function renderAdmin(container) {
    const rows = await DS.spGet(L.config, { top: 1 });
    const cfg = rows[0] || null;
    container.innerHTML = "";

    const courseLead = el("input", { class: "field", type: "number", min: "0", value: cfg && cfg.CourseAlertLeadDays != null ? cfg.CourseAlertLeadDays : 30 });
    const physLead = el("input", { class: "field", type: "number", min: "0", value: cfg && cfg.PhysicalAlertLeadDays != null ? cfg.PhysicalAlertLeadDays : 30 });
    const required = el("textarea", { class: "field", text: cfg && cfg.RequiredCourseTitles ? cfg.RequiredCourseTitles : "Defensive Driving Basics;Defensive Driving Principles;Distracted Driving For Law Enforcement" });
    const recipients = el("input", { class: "field", type: "text", value: cfg && cfg.DigestRecipients ? cfg.DigestRecipients : "", placeholder: "name@dallaspolice.gov; name2@dallaspolice.gov" });
    const allowMark = el("input", { type: "checkbox" });
    allowMark.checked = !!(cfg && (cfg.AllowMarkAsAwarded === true || String(cfg.AllowMarkAsAwarded).toLowerCase() === "yes"));

    const saveBtn = el("button", { class: "btn", text: "Save settings" });

    const body = el("div", { class: "card__body", style: "max-width:560px" }, [
      field("Course alert lead time (days)", courseLead, "How far ahead a course renewal shows as due."),
      field("Physical alert lead time (days)", physLead, "How far ahead a physical shows as due."),
      field("Required courses", required, "Semicolon-separated. These are the courses tracked for compliance."),
      field("Digest recipients", recipients, "Used by the alert email (a later phase). Semicolon-separated addresses."),
      el("div", { class: "form-group" }, [
        el("label", { class: "check" }, [allowMark, el("span", { text: "Allow \u201CMark awarded\u201D on the Dashboard" })]),
      ]),
      el("div", { class: "form-actions" }, [saveBtn]),
    ]);
    container.appendChild(el("div", { class: "card" }, [
      el("div", { class: "card__head" }, el("h3", { text: "Program settings" })),
      body,
    ]));

    saveBtn.addEventListener("click", async () => {
      saveBtn.disabled = true; saveBtn.textContent = "Saving…";
      const fields = {
        CourseAlertLeadDays: Number(courseLead.value) || 0,
        PhysicalAlertLeadDays: Number(physLead.value) || 0,
        RequiredCourseTitles: required.value.trim(),
        DigestRecipients: recipients.value.trim(),
        AllowMarkAsAwarded: allowMark.checked,
      };
      try {
        if (cfg) await DS.spUpdate(L.config, cfg.Id, fields);
        else { fields.Title = "Config"; await DS.spCreate(L.config, fields); }
        await DS.audit("Settings saved", L.config, cfg ? cfg.Id : null, "");
        DS.data.clear();
        DS.toast("Settings saved.", "success");
      } catch (e) {
        DS.toast("Couldn't save settings: " + e.message, "error");
      } finally {
        saveBtn.disabled = false; saveBtn.textContent = "Save settings";
      }
    });
  }

  /* ============================================================
     TICKETS — manual entry + resolve (activity tracking)
     ============================================================ */
  const STATUSES = ["Open", "In Progress", "Resolved"];

  async function renderTickets(container) {
    const tickets = await DS.spGet(L.tickets, { orderby: "OpenedOn desc", top: 1000 });
    container.innerHTML = "";

    const openCount = tickets.filter(t => String(t.Status) !== "Resolved").length;
    const resolvedCount = tickets.filter(t => String(t.Status) === "Resolved").length;
    container.appendChild(el("div", { class: "stats" }, [
      statCard(openCount, "Open", "due"),
      statCard(resolvedCount, "Resolved", "clear"),
      statCard(tickets.length, "Total logged", "neutral"),
    ]));

    const filter = el("select", { class: "field" }, [
      el("option", { value: "open", text: "Open + In progress" }),
      el("option", { value: "all", text: "All tickets" }),
      el("option", { value: "Resolved", text: "Resolved" }),
    ]);
    const newBtn = el("button", { class: "btn", text: "New ticket" });
    newBtn.addEventListener("click", () => openTicketForm(null, container));
    container.appendChild(el("div", { class: "toolbar" }, [
      filter, el("div", { class: "toolbar__spacer" }), newBtn,
    ]));

    const tableWrap = el("div", { class: "card" });
    container.appendChild(tableWrap);

    function paint() {
      const f = filter.value;
      const view = tickets.filter(t => {
        const s = String(t.Status || "Open");
        if (f === "all") return true;
        if (f === "Resolved") return s === "Resolved";
        return s !== "Resolved";
      });
      tableWrap.innerHTML = "";
      if (!view.length) { tableWrap.appendChild(emptyMini("No tickets to show.")); return; }
      const thead = el("thead", null, el("tr", null,
        ["Subject", "Employee", "Status", "Opened", "Assigned to"].map(h => el("th", { text: h }))));
      const tbody = el("tbody", null, view.map(t => {
        const tr = el("tr", { class: "roster-row" }, [
          el("td", null, el("span", { class: "strong", text: t.Title || "—" })),
          el("td", { text: t.EmployeeId || "—" }),
          el("td", null, DS.badge(t.Status || "Open", String(t.Status) === "Resolved" ? "clear" : (String(t.Status) === "In Progress" ? "due" : "neutral"))),
          el("td", { class: "nowrap", text: t.OpenedOn ? DS.fmtDate(t.OpenedOn) : "—" }),
          el("td", { text: DS.emailLocal(t.AssignedTo) || t.AssignedTo || "—" }),
        ]);
        tr.addEventListener("click", () => openTicketForm(t, container));
        return tr;
      }));
      tableWrap.appendChild(el("table", { class: "tbl" }, [thead, tbody]));
    }
    filter.addEventListener("change", paint);
    paint();
  }

  function openTicketForm(ticket, container) {
    const isNew = !ticket;
    const title = el("input", { class: "field", type: "text", value: isNew ? "" : (ticket.Title || ""), placeholder: "Short subject / reason for contact" });
    const emp = el("input", { class: "field", type: "text", value: isNew ? "" : (ticket.EmployeeId || ""), placeholder: "Optional" });
    const desc = el("textarea", { class: "field", text: isNew ? "" : (ticket.Description || "") });
    const assigned = el("input", { class: "field", type: "text", value: isNew ? ((DS.me && (DS.me.mail || DS.me.userPrincipalName)) || "") : (ticket.AssignedTo || "") });

    const statusSel = el("select", { class: "field" }, STATUSES.map(s => el("option", { value: s, text: s })));
    if (!isNew) statusSel.value = STATUSES.includes(ticket.Status) ? ticket.Status : "Open";
    const resNotes = el("textarea", { class: "field", text: isNew ? "" : (ticket.ResolutionNotes || "") });

    const bodyParts = [
      field("Subject", title),
      field("Employee ID", emp, "Who the inquiry is about (optional)."),
      field("Description", desc),
      field("Assigned to", assigned),
    ];
    if (!isNew) {
      bodyParts.push(field("Status", statusSel));
      bodyParts.push(field("Resolution notes", resNotes));
    }
    const body = el("div", null, bodyParts);

    const saveBtn = el("button", { class: "btn", text: isNew ? "Create ticket" : "Save" });
    const cancelBtn = el("button", { class: "btn btn--ghost", text: "Cancel" });
    const m = modal(isNew ? "New ticket" : "Ticket", body, [cancelBtn, saveBtn]);
    cancelBtn.addEventListener("click", m.close);

    saveBtn.addEventListener("click", async () => {
      if (!title.value.trim()) { DS.toast("A subject is required."); return; }
      saveBtn.disabled = true; saveBtn.textContent = "Saving…";
      try {
        if (isNew) {
          await DS.spCreate(L.tickets, {
            Title: title.value.trim(),
            EmployeeId: emp.value.trim(),
            Description: desc.value.trim(),
            AssignedTo: assigned.value.trim(),
            Status: "Open",
            OpenedOn: new Date().toISOString(),
          });
          await DS.audit("Ticket opened", L.tickets, null, title.value.trim());
          DS.toast("Ticket created.", "success");
        } else {
          const wasResolved = String(ticket.Status) === "Resolved";
          const nowResolved = statusSel.value === "Resolved";
          const fields = {
            Title: title.value.trim(),
            EmployeeId: emp.value.trim(),
            Description: desc.value.trim(),
            AssignedTo: assigned.value.trim(),
            Status: statusSel.value,
            ResolutionNotes: resNotes.value.trim(),
          };
          if (nowResolved && !wasResolved) fields.ResolvedOn = new Date().toISOString();
          if (!nowResolved && wasResolved) fields.ResolvedOn = null;
          await DS.spUpdate(L.tickets, ticket.Id, fields);
          await DS.audit("Ticket updated", L.tickets, ticket.Id, title.value.trim() + " → " + statusSel.value);
          DS.toast("Ticket saved.", "success");
        }
        m.close();
        renderTickets(container);
      } catch (e) {
        saveBtn.disabled = false; saveBtn.textContent = isNew ? "Create ticket" : "Save";
        DS.toast("Couldn't save the ticket: " + e.message, "error");
      }
    });
  }

  function statCard(n, label, kind) {
    return el("div", { class: "stat stat--" + kind }, [
      el("b", { class: "tnum", text: String(n) }),
      el("span", { text: label }),
    ]);
  }

  /* ============================================================
     REPORTS — flexible multi-sheet Excel export
     Each selected list becomes its own tab in one workbook.
     (To combine into a single sheet instead, change buildWorkbook.)
     ============================================================ */
  const EXPORTS = [
    { key: "roster",    label: "Roster",    list: L.roster,    note: "Employees + assignments" },
    { key: "courses",   label: "Courses",   list: L.courses,   note: "Defensive driving records" },
    { key: "physicals", label: "Physicals", list: L.physicals, note: "Driver physicals" },
    { key: "accidents", label: "Accidents", list: L.accidents, note: "Equipment incidents" },
    { key: "awards",    label: "Awards",    list: L.awards,    note: "Safe-driving awards" },
    { key: "tickets",   label: "Tickets",   list: L.tickets,   note: "Activity log" },
    { key: "audit",     label: "Audit log", list: L.audit,     note: "System history" },
  ];

  const SYS_FIELDS = new Set([
    "FileSystemObjectType", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl", "ContentTypeId",
    "ComplianceAssetId", "GUID", "AttachmentFiles", "Attachments", "ID", "LinkTitleNoMenu", "LinkTitle",
    "ItemChildCount", "FolderChildCount", "AuthorId", "EditorId", "AppAuthorId", "AppEditorId",
    "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
    "_UIVersionString", "OData__UIVersionString", "OData__ColorTag",
  ]);
  function cleanRow(row) {
    const o = {};
    Object.keys(row).forEach(k => {
      if (SYS_FIELDS.has(k) || k.startsWith("OData__")) return;
      let v = row[k];
      if (v && typeof v === "object") return;  // skip nested/lookup objects
      o[k] = v;
    });
    return o;
  }

  async function renderReports(container) {
    container.innerHTML = "";
    const selected = new Set(EXPORTS.map(e => e.key));  // default: all on

    const list = el("div", { class: "export-list" });
    EXPORTS.forEach(e => {
      const cb = el("input", { type: "checkbox" }); cb.checked = true;
      const item = el("label", { class: "export-item on" }, [
        cb, el("span", { class: "lbl" }, [el("b", { text: e.label }), el("span", { text: e.note })]),
      ]);
      cb.addEventListener("change", () => {
        if (cb.checked) { selected.add(e.key); item.classList.add("on"); }
        else { selected.delete(e.key); item.classList.remove("on"); }
        updateBtn();
      });
      list.appendChild(item);
    });

    const allBtn = el("button", { class: "btn btn--ghost btn--sm", text: "Select all" });
    const noneBtn = el("button", { class: "btn btn--ghost btn--sm", text: "Clear" });
    allBtn.addEventListener("click", () => setAll(true));
    noneBtn.addEventListener("click", () => setAll(false));
    function setAll(on) {
      list.querySelectorAll("input").forEach((cb, i) => {
        cb.checked = on; const e = EXPORTS[i];
        if (on) { selected.add(e.key); cb.parentNode.classList.add("on"); }
        else { selected.delete(e.key); cb.parentNode.classList.remove("on"); }
      });
      updateBtn();
    }

    const exportBtn = el("button", { class: "btn", text: "Export to Excel" });
    function updateBtn() {
      exportBtn.disabled = selected.size === 0;
      exportBtn.textContent = "Export to Excel (" + selected.size + ")";
    }
    updateBtn();

    const status = el("div", { class: "help", style: "margin-top:12px" });

    exportBtn.addEventListener("click", async () => {
      exportBtn.disabled = true;
      const chosen = EXPORTS.filter(e => selected.has(e.key));
      try {
        status.textContent = "Loading the spreadsheet engine…";
        const XLSXlib = await DS.ensureXlsx();
        const wb = XLSXlib.utils.book_new();
        for (const e of chosen) {
          status.textContent = "Fetching " + e.label + "…";
          const rows = await DS.spGet(e.list, { top: 5000 });
          const cleaned = rows.map(cleanRow);
          const ws = XLSXlib.utils.json_to_sheet(cleaned.length ? cleaned : [{ "(no records)": "" }]);
          XLSXlib.utils.book_append_sheet(wb, ws, e.label.slice(0, 31));
        }
        const stamp = new Date().toISOString().slice(0, 10);
        XLSXlib.writeFile(wb, "DrivingSafety_Export_" + stamp + ".xlsx");
        await DS.audit("Report exported", null, null, chosen.map(e => e.label).join(", "));
        status.textContent = "Downloaded DrivingSafety_Export_" + stamp + ".xlsx";
        DS.toast("Export ready.", "success");
      } catch (err) {
        status.textContent = "";
        DS.toast("Export failed: " + err.message, "error");
      } finally {
        exportBtn.disabled = false; updateBtn();
      }
    });

    container.appendChild(el("div", { class: "card" }, [
      el("div", { class: "card__head" }, [
        el("h3", { text: "Export to Excel" }),
        el("div", { style: "display:flex; gap:8px" }, [allBtn, noneBtn]),
      ]),
      el("div", { class: "card__body" }, [
        el("div", { class: "import-note", text: "Pick any combination of lists. Each becomes its own tab in a single Excel workbook." }),
        list,
        el("div", { style: "display:flex; align-items:center; gap:14px" }, [exportBtn, status]),
      ]),
    ]));
  }

  /* ---- register ---- */
  DS.registerScreen("admin",   { title: "Admin",     icon: "◈", render: renderAdmin });
  DS.registerScreen("tickets", { title: "Tickets",   icon: "▣", render: renderTickets });
  DS.registerScreen("reports", { title: "Reports",   icon: "▦", render: renderReports });

})();
