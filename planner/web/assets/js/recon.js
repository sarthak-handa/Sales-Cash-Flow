/* Forecast vs Actual: SALES & CASH FLOW rows matched to SAP billing. */
(() => {
  const { inr, cr, esc, getJSON, postJSON, toast, loading } = YD;
  let ov = null, proj = null, issues = [], tables = {}, curGroup = null;
  let projScope = "all", groupStatus = "", issueStatus = "";

  const ISSUE_ORDER = ["billed_not_disp", "disp_not_billed", "sap_only", "part_billed", "value_diff"];
  const HELP = {
    billed_not_disp: "SAP has billed it, but the sales file still says Not Disp. Mark it Disp. in the sales file.",
    disp_not_billed: "The sales file says Disp., but SAP has no bill for it. Either it was billed under another name (link it), or it is not billed yet and should be Not Disp.",
    sap_only: "SAP billed it, but no sales-file row matches. Link it to a row, or add the row to the sales file.",
    part_billed: "SAP has billed less than 90% of the forecast value. The balance is still to bill.",
    value_diff: "Billed, but the SAP value differs from the forecast by more than 5% (and ₹50,000). SAP is the actual.",
    ok: "Forecast and SAP agree (value within 5% or ₹50,000).",
    slipped: "Not billed and not marked Disp.; its forecast month has passed.",
    pending: "Not billed yet; its forecast month is this month or later.",
    service: "Engineering / erection / commissioning bill. Not part of the supply forecast.",
    zero: "Sales-file row with no billing value and nothing billed.",
  };

  // ------------------------------------------------------------------ helpers
  const chip = (tone, label) => `<span class="st-chip tone-${tone}">${esc(label)}</span>`;
  const dispChip = (s) => `<span class="badge-st ${s === "Disp." ? "st-dispatched" : "st-notdisp"}">${esc(s)}</span>`;
  const signed = (v) => {
    if (v === null || v === undefined || isNaN(v) || Math.abs(v) < 1) return `<span class="text-muted">—</span>`;
    return `<span class="${v > 0 ? "pos" : "neg"}">${v > 0 ? "+" : "−"}${inr(Math.abs(v))}</span>`;
  };
  const moneyCol = (get) => ({
    render: (d, t, r) => { const v = get(r); return t === "display" ? (v === null || v === undefined ? "" : inr(v)) : (v ?? -1); },
    className: "num",
  });
  const uniq = (a) => [...new Set(a.filter((x) => x !== "" && x !== null && x !== undefined))];

  function dt(id, opts) {
    if (tables[id]) { tables[id].clear().rows.add(opts.data).draw(); return tables[id]; }
    tables[id] = $("#" + id).DataTable({
      pageLength: 25, lengthMenu: [10, 25, 50, 100, 500], autoWidth: false, deferRender: true,
      language: { search: "", searchPlaceholder: "Search…", emptyTable: "Nothing here" }, ...opts,
    });
    $("#" + id + " tbody").on("click", "tr", function () {
      const r = tables[id].row(this).data();
      if (r && opts.onRow) opts.onRow(r);
    });
    return tables[id];
  }

  // status filters for the two group tables (DataTables custom search)
  $.fn.dataTable.ext.search.push((settings, _row, idx) => {
    const id = settings.nTable.id;
    if (id === "tblGroups" && groupStatus) return tables.tblGroups.row(idx).data().status === groupStatus;
    if (id === "tblIssues" && issueStatus) return tables.tblIssues.row(idx).data().status === issueStatus;
    return true;
  });

  function pills(el, counts, cur, onPick, withAll = true) {
    const order = [...ISSUE_ORDER, "ok", "slipped", "pending", "service", "zero"].filter((k) => counts[k]);
    const total = order.reduce((s, k) => s + counts[k], 0);
    let h = withAll ? `<button class="pill ${cur ? "" : "on"}" data-k="">All <b>${total}</b></button>` : "";
    for (const k of order) h += `<button class="pill tone-${ov.tones[k]} ${cur === k ? "on" : ""}" data-k="${k}">${esc(ov.labels[k])} <b>${counts[k]}</b></button>`;
    $(el).html(h).find(".pill").on("click", function () { onPick($(this).data("k")); });
  }

  // ------------------------------------------------------------------ overview
  async function loadOverview(refresh = false) {
    loading(true, refresh ? "Re-reading the sales file and SAP…" : "Matching forecast to SAP…");
    try {
      ov = await getJSON("/api/recon/overview" + (refresh ? "?refresh=true" : ""));
      issues = await getJSON("/api/recon/issues");
    } catch (e) {
      $("#sapChip").addClass("err").find(".t").text("SAP read error");
      toast("error", "Could not read the data", e.message);
      return;
    } finally { loading(false); }
    const fmt = (d) => (d ? new Date(d).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }) : "");
    $("#sapChip").removeClass("err").attr("title", `${ov.sap_file}, as of ${ov.sap_saved}\nSales file: ${ov.sales_source}`)
      .find(".t").text(`SAP ${fmt(ov.sap_from)} – ${fmt(ov.sap_to)}`);
    $("#pageSub").html(`Sales &amp; Cash Flow forecast compared with SAP billing (read-only), assembly by assembly · SAP invoices ${fmt(ov.sap_from)} – ${fmt(ov.sap_to)}`);
    $("#howSap").text(`SAP: ${ov.sap_file}, as of ${ov.sap_saved}. Sales file: ${ov.sales_source}. ${ov.sap_unlinked_rows} invoice lines have no project (scrap, IT, services) — ${cr(ov.sap_unlinked_inr)} — and are left out.`);
    $("#statusHelp").html(Object.keys(HELP).map((k) => `<div class="mb-2">${chip(ov.tones[k], ov.labels[k])}<div class="small mt-1">${esc(HELP[k])}</div></div>`).join(""));
    renderKpis();
    renderProjects();
    renderIssues();
    fillProjectSelect();
    loadSide();
  }

  function renderKpis() {
    const c = ov.counts, v = ov.values;
    const P = ov.projects.filter((p) => p.in_sales);
    const disp = P.reduce((s, p) => s + p.marked_disp, 0), sap = P.reduce((s, p) => s + p.sap_billed, 0);
    const tile = (cls, k, label, value, sub) =>
      `<div class="kpi ${cls} ${k ? "click" : ""}" ${k ? `data-k="${k}" role="button" tabindex="0"` : ""}><div class="kpi-label">${label}</div>
        <div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div></div>`;
    $("#kpis").html(
      tile("", "", "SAP billed vs marked Disp.", cr(sap), `sales file marks <b>${cr(disp)}</b> as Disp. · gap <b>${cr(sap - disp)}</b>`) +
      tile("red", "billed_not_disp", "Billed in SAP, Not Disp.", c.billed_not_disp, `<b>${cr(v.billed_not_disp)}</b> billed · status not updated`) +
      tile("red", "disp_not_billed", "Marked Disp., no SAP bill", c.disp_not_billed, `<b>${cr(v.disp_not_billed)}</b> of forecast`) +
      tile("amber", "sap_only", "Billed, not in sales file", c.sap_only, `<b>${cr(v.sap_only)}</b> billed in SAP`) +
      tile("amber", "value_diff", "Value differs / part billed", c.value_diff + c.part_billed, `<b>${cr(v.value_diff + v.part_billed)}</b> total difference`) +
      tile("green", "", "Matches", c.ok, `<b>${cr(v.ok)}</b> agree within 5%`));
    $("#kpis .kpi.click").on("click keydown", function (e) {
      if (e.type === "keydown" && e.key !== "Enter") return;
      issueStatus = $(this).data("k");
      renderIssuePills();
      tables.tblIssues && tables.tblIssues.draw();
      bootstrap.Tab.getOrCreateInstance(document.getElementById("tabIssuesBtn")).show();
    });
  }

  function renderProjects() {
    const rows = ov.projects.filter((p) => projScope === "all" || Math.abs(p.sap_billed) > 1 || p.marked_disp > 0);
    $("#cntProjects").text(rows.length);
    dt("tblProjects", {
      data: rows, order: [[6, "desc"]],
      columns: [
        { data: "project", render: (d, t, r) => t === "display"
            ? `<b>${esc(d)}</b>${r.in_sales ? "" : ' <span class="line-chip">SAP only</span>'}<div class="small text-muted">${esc(r.customer || "")}</div>` : d },
        { data: "pm", render: (d) => esc(d || "") },
        { data: "forecast", ...moneyCol((r) => r.forecast) },
        { data: "marked_disp", ...moneyCol((r) => r.marked_disp) },
        { data: "sap_billed", ...moneyCol((r) => r.sap_billed) },
        { data: null, className: "num", render: (d, t, r) => { const g = r.sap_billed - r.marked_disp; return t === "display" ? signed(g) : g; } },
        { data: "red", className: "num", render: (d, t) => (t === "display" ? (d ? `<span class="cnt red">${d}</span>` : '<span class="text-muted">0</span>') : d) },
        { data: null, className: "num", render: (d, t, r) => { const n = r.counts.sap_only + r.counts.part_billed + r.counts.value_diff; return t === "display" ? (n ? `<span class="cnt amber">${n}</span>` : '<span class="text-muted">0</span>') : n; } },
        { data: null, render: (d, t, r) => {
            const n = ISSUE_ORDER.reduce((s, k) => s + r.counts[k], 0) + r.counts.ok;
            const pct = n ? Math.round((r.counts.ok / n) * 100) : 0;
            return t === "display" ? (n ? `<div class="small">${r.counts.ok} of ${n} billed items</div><div class="prog"><span style="width:${pct}%"></span></div>` : '<span class="text-muted small">nothing billed yet</span>') : pct;
          } },
      ],
      createdRow: (tr) => tr.classList.add("clickable"),
      onRow: (r) => openProject(r.pno),
    });
  }

  function fillProjectSelect() {
    const rows = [...ov.projects].sort((a, b) => b.red - a.red || b.issues - a.issues || a.pno.localeCompare(b.pno));
    $("#projSelect").html('<option value="">Choose a project…</option>' + rows.map((p) =>
      `<option value="${esc(p.pno)}">${esc(p.project)}${p.red ? ` · ${p.red} red` : ""}${p.issues - p.red ? ` · ${p.issues - p.red} to check` : ""}</option>`).join(""));
    if (proj) $("#projSelect").val(proj.pno);
  }

  // ------------------------------------------------------------------ one project
  async function openProject(pno, gid = null) {
    if (!pno) return;
    bootstrap.Tab.getOrCreateInstance(document.getElementById("tabProjectBtn")).show();
    if (!proj || proj.pno !== pno) {
      loading(true, "Matching project " + pno + "…");
      try { proj = await getJSON("/api/recon/project?pno=" + encodeURIComponent(pno)); }
      catch (e) { toast("error", "Could not load the project", e.message); return; }
      finally { loading(false); }
      groupStatus = "";
    }
    renderProject();
    if (gid !== null) { const g = proj.groups.find((x) => x.gid === gid); if (g) openGroup(g); }
  }

  function renderProject() {
    const s = proj.summary;
    $("#projSelect").val(proj.pno);
    $("#btnExportProject").attr("href", `/api/recon/export?scope=all&pno=${encodeURIComponent(proj.pno)}`);
    $("#projNote").text(`${s.rows} sales-file rows · ${s.wbs} SAP items · ${s.matched} matched groups`);
    const gap = s.sap_billed - s.marked_disp;
    $("#projKpis").html(`
      <div class="kpi"><div class="kpi-label">Forecast (sales file)</div><div class="kpi-value">${cr(s.forecast)}</div><div class="kpi-sub">${s.rows} rows${s.pm ? " · PM " + esc(s.pm) : ""}</div></div>
      <div class="kpi"><div class="kpi-label">Marked Disp.</div><div class="kpi-value">${cr(s.marked_disp)}</div><div class="kpi-sub">${s.forecast ? Math.round((s.marked_disp / s.forecast) * 100) : 0}% of forecast</div></div>
      <div class="kpi"><div class="kpi-label">SAP billed (net)</div><div class="kpi-value">${cr(s.sap_billed)}</div><div class="kpi-sub">gap to Disp. ${signed(gap)}</div></div>
      <div class="kpi red"><div class="kpi-label">Red flags</div><div class="kpi-value">${s.red}</div><div class="kpi-sub">status in the sales file looks wrong</div></div>
      <div class="kpi amber"><div class="kpi-label">To check</div><div class="kpi-value">${s.issues - s.red}</div><div class="kpi-sub">value differs, part billed, or not in sales file</div></div>`);
    renderGroupPills();
    dt("tblGroups", {
      data: proj.groups, order: [],
      columns: [
        { data: "status", render: (d, t, r) => (t === "display" ? chip(r.tone, r.label) : ISSUE_ORDER.indexOf(d) + 1 || 99) },
        { data: null, render: (d, t, r) => {
            if (!r.sales.length) return t === "display" ? '<span class="text-muted">— not in sales file —</span>' : "";
            const names = r.sales.map((x) => esc(x.name));
            return t === "display" ? `${names.join('<span class="plus">+</span>')}<div class="small text-muted">row ${r.sales.map((x) => x.row).join(", ")}${r.sales.some((x) => x.note) ? ' · <i class="fa-regular fa-note-sticky"></i>' : ""}</div>` : names.join(" ");
          } },
        { data: null, ...moneyCol((r) => (r.sales.length ? r.f_value : null)) },
        { data: "f_month", render: (d, t, r) => (t === "display" ? esc(uniq(r.sales.map((x) => x.month_label)).join(", ")) : d || "") },
        { data: null, render: (d, t, r) => { const u = uniq(r.sales.map((x) => x.status)); return t === "display" ? u.map(dispChip).join(" ") : u.join(); } },
        { data: null, render: (d, t, r) => {
            if (!r.sap.length) return t === "display" ? '<span class="text-muted">— no SAP bill —</span>' : "";
            return t === "display" ? r.sap.map((w) => `${esc(w.desc)}<div class="small text-muted">${esc(w.wbs.replace(/^ITEM:/, ""))}${w.cancelled > 1 ? " · had cancellations" : ""}</div>`).join("") : r.sap.map((w) => w.desc + " " + w.wbs).join(" ");
          } },
        { data: null, ...moneyCol((r) => (r.sap.length ? r.s_value : null)) },
        { data: "s_month", render: (d, t, r) => {
            if (t !== "display") return d || "";
            const md = r.month_diff;
            const lag = md ? ` <span class="lag ${md > 0 ? "late" : "early"}">${md > 0 ? "+" : ""}${md} mo</span>` : "";
            return esc(r.s_month_label || "") + (r.sales.length && r.sap.length ? lag : "");
          } },
        { data: "diff", className: "num", render: (d, t, r) => (t === "display" ? (r.sales.length && r.sap.length ? signed(d) : "") : d) },
        { data: "basis", render: (d, t, r) => (t === "display" ? (d ? `<span class="basis b-${d}">${esc(r.basis_label)}${d === "name" && r.score !== null ? ` · ${r.score}` : ""}</span>` : "") : d) },
      ],
      createdRow: (tr) => tr.classList.add("clickable"),
      onRow: (r) => openGroup(r),
    });
  }

  function renderGroupPills() {
    const counts = {};
    proj.groups.forEach((g) => (counts[g.status] = (counts[g.status] || 0) + 1));
    pills("#statusPills", counts, groupStatus, (k) => { groupStatus = k; renderGroupPills(); tables.tblGroups.draw(); });
  }

  // ------------------------------------------------------------------ all mismatches
  function renderIssuePills() {
    const counts = {};
    issues.forEach((g) => (counts[g.status] = (counts[g.status] || 0) + 1));
    pills("#issuePills", counts, issueStatus, (k) => { issueStatus = k; renderIssuePills(); tables.tblIssues.draw(); });
  }

  function renderIssues() {
    $("#cntIssues").text(issues.length);
    renderIssuePills();
    dt("tblIssues", {
      data: issues, order: [[1, "asc"], [9, "desc"]],
      columns: [
        { data: "project", render: (d) => `<b>${esc(d)}</b>` },
        { data: "status", render: (d, t, r) => (t === "display" ? chip(r.tone, r.label) : ISSUE_ORDER.indexOf(d)) },
        { data: "sales", render: (d, t, r) => (t === "display" ? (d ? `${esc(d)}<div class="small text-muted">row ${r.rows.join(", ")}</div>` : '<span class="text-muted">— not in sales file —</span>') : d) },
        { data: "f_value", ...moneyCol((r) => r.f_value) },
        { data: "f_month_label", render: (d) => esc(d || "") },
        { data: "sap", render: (d, t) => (t === "display" ? (d ? esc(d) : '<span class="text-muted">— no SAP bill —</span>') : d) },
        { data: "s_value", ...moneyCol((r) => r.s_value) },
        { data: "s_month_label", render: (d) => esc(d || "") },
        { data: "pm", render: (d) => esc(d || "") },
        { data: null, visible: false, render: (d, t, r) => Math.max(Math.abs(r.f_value || 0), Math.abs(r.s_value || 0)) },
      ],
      createdRow: (tr) => tr.classList.add("clickable"),
      onRow: (r) => openProject(r.pno, r.gid),
    });
  }

  // ------------------------------------------------------------------ group detail
  const gm = () => bootstrap.Modal.getOrCreateInstance(document.getElementById("groupModal"));

  function summaryText(g) {
    const rows = g.sales.map((x) => x.row).join(", ");
    switch (g.status) {
      case "billed_not_disp": return `SAP billed <b>${inr(g.s_value)}</b> in <b>${esc(g.s_month_label)}</b>, but sales-file row ${rows} still says <b>Not Disp.</b> → mark it Disp.`;
      case "disp_not_billed": return `Sales-file row ${rows} says <b>Disp.</b> (${inr(g.f_value)}, ${esc(g.f_month_label)}) but no SAP bill was found. If SAP billed it under another name, link it below; if not, it is not billed yet.`;
      case "sap_only": return `SAP billed <b>${inr(g.s_value)}</b> in ${esc(g.s_month_label)}, and no sales-file row matches it. Link it to the right row, or add it to the sales file.`;
      case "part_billed": return `SAP has billed <b>${inr(g.s_value)}</b> of the <b>${inr(g.f_value)}</b> forecast (${Math.round((g.s_value / g.f_value) * 100)}%). <b>${inr(g.f_value - g.s_value)}</b> is still to bill.`;
      case "value_diff": return `Billed, but SAP's value is ${signed(g.diff)} against the forecast of ${inr(g.f_value)}. SAP is the actual.`;
      case "ok": return `Forecast and SAP agree${g.month_diff ? ` (billed ${Math.abs(g.month_diff)} month${Math.abs(g.month_diff) > 1 ? "s" : ""} ${g.month_diff > 0 ? "after" : "before"} the forecast month)` : ""}.`;
      default: return esc(HELP[g.status] || "");
    }
  }

  function openGroup(g) {
    curGroup = g;
    $("#gmTitle").text(`${proj.summary.project} · ${g.sales[0]?.name || g.sap[0]?.desc || ""}`);
    $("#gmSummary").html(`<div class="gm-sum tone-${g.tone}">${chip(g.tone, g.label)} <span>${summaryText(g)}</span></div>`);
    $("#gmSales").html(g.sales.length ? g.sales.map((s) => `
      <div class="item">
        <div class="d-flex justify-content-between gap-2">
          <div><b>${esc(s.name)}</b><div class="small text-muted">Row ${s.row}${s.boi ? " · " + esc(s.boi) : ""}${s.pm ? " · PM " + esc(s.pm) : ""}</div></div>
          <div class="text-end"><div class="fw-bold num">${inr(s.bill)}</div><div class="small">${esc(s.month_label)} ${dispChip(s.status)}</div></div>
        </div>
        ${s.risk ? `<div class="small mt-1"><span class="risk-chip">${esc(s.risk)}</span></div>` : ""}
        ${s.remarks ? `<div class="small text-muted mt-1">${esc(s.remarks)}</div>` : ""}
        <div class="d-flex gap-2 mt-2">
          <input class="form-control form-control-sm note" data-skey="${esc(s.skey)}" value="${esc(s.note || "")}" placeholder="Your note (saved)…" />
          <button class="yd-btn btn-sm pick-sap" data-skey="${esc(s.skey)}" title="Link this row to a SAP item"><i class="fa-solid fa-link"></i> SAP item…</button>
        </div>
      </div>`).join("") : '<div class="text-muted small">No sales-file row is matched to this SAP item.</div>');
    $("#gmSap").html(g.sap.length ? g.sap.map((w) => `
      <div class="item">
        <div class="d-flex justify-content-between gap-2">
          <div><b>${esc(w.desc)}</b><div class="small text-muted">${esc(w.wbs.replace(/^ITEM:/, "no WBS · "))}${w.cur && w.cur !== "INR" ? " · " + esc(w.cur) : ""}</div></div>
          <div class="text-end"><div class="fw-bold num">${inr(w.net)}</div><div class="small">${esc(w.last_label)}</div></div>
        </div>
        ${w.cancelled > 1 ? `<div class="small mt-1 text-muted">Invoiced ${inr(w.gross)} · cancelled / credited ${inr(w.cancelled)}</div>` : ""}
        <table class="table table-sm lines mt-2 mb-1"><thead><tr><th>Date</th><th>Doc</th><th>Type</th><th>Product</th><th class="num">₹ (net of tax)</th></tr></thead><tbody>
          ${w.lines.map((l) => `<tr class="${l.inr < 0 ? "neg-line" : ""}"><td>${esc(l.date)}</td><td>${esc(l.doc)}</td><td>${esc(l.type)}</td><td>${esc(l.product)}</td><td class="num">${inr(l.inr)}</td></tr>`).join("")}
        </tbody></table>
        <div class="d-flex gap-2 mt-2">
          <input class="form-control form-control-sm note" data-wbs="${esc(w.wbs)}" value="${esc(w.note || "")}" placeholder="Your note (saved)…" />
          <button class="yd-btn btn-sm pick-sales" data-wbs="${esc(w.wbs)}" title="Link this SAP item to a sales row"><i class="fa-solid fa-link"></i> Sales row…</button>
        </div>
      </div>`).join("") : '<div class="text-muted small">No SAP bill is matched to this row.</div>');
    $("#gmPicker").addClass("d-none");
    $("#gmUnlink").toggle(!!(g.sales.length && g.sap.length));
    $("#gmBasis").text(g.basis ? `Matched by: ${g.basis_label}${g.basis === "name" && g.score !== null ? ` (name score ${g.score})` : ""}` : "");
    gm().show();
  }

  const pairsOf = (g) => {
    const p = [];
    g.sales.forEach((s) => g.sap.forEach((w) => p.push([s.skey, w.wbs])));
    return p;
  };
  const itemsOf = (g) => [...g.sales.map((s) => [s.skey, null]), ...g.sap.map((w) => [null, w.wbs])];

  async function override(body, keep) {
    loading(true, "Saving…");
    try {
      proj = await (await postJSON("/api/recon/override", { pno: proj.pno, ...body })).json();
    } catch (e) { toast("error", "Not saved", e.message); return; }
    finally { loading(false); }
    renderProject();
    // keep the modal on the group that now holds the same item
    const g = proj.groups.find((x) => x.sales.some((s) => s.skey === keep.skey) || x.sap.some((w) => w.wbs === keep.wbs));
    if (g) openGroup(g); else gm().hide();
    refreshOverviewQuietly();
  }

  async function refreshOverviewQuietly() {
    try {
      ov = await getJSON("/api/recon/overview");
      issues = await getJSON("/api/recon/issues");
      renderKpis(); renderProjects(); renderIssues(); fillProjectSelect(); loadSide();
    } catch (_) {}
  }

  async function picker(kind, key) {
    const q = kind === "sap" ? `skey=${encodeURIComponent(key)}` : `wbs=${encodeURIComponent(key)}`;
    let c;
    try { c = await getJSON(`/api/recon/candidates?pno=${encodeURIComponent(proj.pno)}&${q}`); }
    catch (e) { toast("error", "Could not list items", e.message); return; }
    const fit = (x) => (x.fit >= 0.98 ? ' <span class="basis b-value">same value</span>' : "");
    const head = kind === "sap"
      ? "<tr><th>SAP item</th><th class='num'>Billed</th><th>Billed in</th><th class='num'>Name match</th><th></th></tr>"
      : "<tr><th>Sales-file row</th><th class='num'>Forecast</th><th>Month</th><th>Status</th><th class='num'>Name match</th><th></th></tr>";
    const body = c.map((x) => kind === "sap"
      ? `<tr data-q="${esc((x.desc + " " + x.wbs).toLowerCase())}"><td>${esc(x.desc)}<div class="small text-muted">${esc(x.wbs)}</div></td><td class="num">${inr(x.net)}${fit(x)}</td><td>${esc(x.last)}</td><td class="num">${x.score}</td>
          <td><button class="yd-btn btn-sm do-link" data-skey="${esc(key)}" data-wbs="${esc(x.wbs)}">Link</button></td></tr>`
      : `<tr data-q="${esc(x.name.toLowerCase())}"><td>${esc(x.name)}</td><td class="num">${inr(x.bill)}${fit(x)}</td><td>${esc(x.month)}</td><td>${dispChip(x.status)}</td><td class="num">${x.score}</td>
          <td><button class="yd-btn btn-sm do-link" data-skey="${esc(x.skey)}" data-wbs="${esc(key)}">Link</button></td></tr>`).join("");
    $("#gmPickerTitle").text(kind === "sap" ? "Pick the SAP item this row was billed as" : "Pick the sales-file row this SAP item belongs to");
    $("#gmPickerTbl thead").html(head);
    $("#gmPickerTbl tbody").html(body || '<tr><td colspan="6" class="text-muted">Nothing to link in this project.</td></tr>');
    $("#gmPickerQ").val("");
    $("#gmPicker").removeClass("d-none")[0].scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  // ------------------------------------------------------------------ side by side (every project, three columns)
  let side = [], sideScope = "all";
  const open = new Set();
  const dmy = (d) => (d ? new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) : "");
  const sumOf = (a, f) => a.reduce((t, x) => t + (f(x) || 0), 0);

  async function loadSide() {
    try { side = await getJSON("/api/recon/side"); } catch (e) { toast("error", "Could not load the projects", e.message); return; }
    renderSide();
  }

  function itemBoth(p, x) {
    const sn = x.sales.map((y) => esc(y.name)).join('<span class="plus">+</span>');
    const sap = x.sap.map((y) => esc(y.desc)).join(" + ");
    const flag = x.status === "ok" ? "" : `<div class="mt-1">${chip(x.tone, x.label)}</div>`;
    return `<div class="si ${x.tone === "red" ? "flag" : ""}" data-p="${esc(p)}" data-g="${x.gid}">
      <div class="si-name">${sn} <span class="si-row">row ${x.sales.map((y) => y.row).join(", ")}</span></div>
      <div class="si-sub">SAP: ${sap}</div>
      <div class="si-vals"><span><i>Forecast</i> ${inr(x.f_value)} · ${esc(x.f_month || "—")}</span>
        <span><i>Billed</i> <b>${inr(x.s_value)}</b> · ${dmy(x.s_date)}${x.bills > 1 ? ` <small>(${x.bills} bills)</small>` : ""}</span></div>${flag}</div>`;
  }
  function itemSap(p, x) {
    return `<div class="si" data-p="${esc(p)}" data-g="${x.gid}">
      <div class="si-name">${esc(x.desc)}${x.status === "service" ? ' <span class="line-chip">service</span>' : ""}</div>
      <div class="si-sub">${esc(String(x.wbs).replace(/^ITEM:/, ""))}</div>
      <div class="si-vals"><span><i>Billed</i> <b>${inr(x.value)}</b></span><span><i>Billing date</i> ${dmy(x.date)}${x.bills > 1 ? ` <small>(${x.bills} bills)</small>` : ""}</span></div></div>`;
  }
  function itemSales(p, x) {
    const red = x.status === "disp_not_billed";
    const tag = red ? chip("red", "Marked Disp., no SAP bill") : x.status === "slipped" ? chip("grey", "Month passed, not billed") : x.status === "zero" ? chip("grey", "₹0 row") : "";
    return `<div class="si ${red ? "flag" : ""}" data-p="${esc(p)}" data-g="${x.gid}">
      <div class="si-name">${esc(x.name)} <span class="si-row">row ${x.row}</span></div>
      <div class="si-vals"><span><i>Tentative</i> <b>${inr(x.value)}</b></span><span><i>Month</i> ${esc(x.month_label || "—")}</span><span>${dispChip(x.disp)}</span></div>
      ${tag ? `<div class="mt-1">${tag}</div>` : ""}</div>`;
  }

  function filtered(d, q) {
    if (!q) return d;
    const m = (t) => String(t || "").toLowerCase().includes(q);
    if (m(d.summary.project) || m(d.summary.pm) || m(d.summary.customer)) return d;
    return {
      ...d,
      both: d.both.filter((x) => x.sales.some((y) => m(y.name)) || x.sap.some((y) => m(y.desc) || m(y.wbs))),
      sap_only: d.sap_only.filter((x) => m(x.desc) || m(x.wbs)),
      sales_only: d.sales_only.filter((x) => m(x.name)),
    };
  }

  function sideBody(d) {
    const p = d.summary.pno;
    const col = (cls, title, items, f, totals) => `<div class="col-lg-4"><div class="side-col ${cls}">
        <div class="side-col-h"><b>${title}</b> <span>${items.length} · ${totals}</span></div>
        ${items.length ? items.map((x) => f(p, x)).join("") : '<div class="text-muted small p-2">None</div>'}</div></div>`;
    return `<div class="row g-2">
      ${col("both", "In both", d.both, itemBoth, `forecast ${cr(sumOf(d.both, (x) => x.f_value))} · billed ${cr(sumOf(d.both, (x) => x.s_value))}`)}
      ${col("sap", "Only in SAP", d.sap_only, itemSap, `billed ${cr(sumOf(d.sap_only, (x) => x.value))}`)}
      ${col("sales", "Only in sales file", d.sales_only, itemSales, `tentative ${cr(sumOf(d.sales_only, (x) => x.value))}`)}</div>`;
  }

  function renderSide() {
    const q = ($("#sideQ").val() || "").trim().toLowerCase();
    let list = side.map((d) => filtered(d, q)).filter((d) => d.both.length + d.sap_only.length + d.sales_only.length);
    if (sideScope === "diff") list = list.filter((d) => d.sap_only.length || d.sales_only.some((x) => x.status !== "zero"));
    $("#cntSide").text(list.length);
    $("#sideList").html(list.map((d) => {
      const s = d.summary, p = s.pno, isOpen = open.has(p) || (q && list.length <= 15);
      const red = d.both.filter((x) => x.tone === "red").length + d.sales_only.filter((x) => x.tone === "red").length;
      return `<div class="side-proj" data-p="${esc(p)}">
        <div class="side-h" role="button" tabindex="0">
          <i class="fa-solid fa-chevron-${isOpen ? "down" : "right"} me-2"></i>
          <div class="flex-grow-1"><b>${esc(s.project)}</b>${s.in_sales ? "" : ' <span class="line-chip">not in sales file</span>'}
            <div class="small text-muted">${esc([s.customer, s.pm ? "PM " + s.pm : ""].filter(Boolean).join(" · "))}
              · forecast ${cr(s.forecast)} · SAP billed ${cr(s.sap_billed)}</div></div>
          <span class="sc both">${d.both.length} in both</span>
          <span class="sc sap">${d.sap_only.length} only SAP</span>
          <span class="sc sales">${d.sales_only.length} only sales file</span>
          ${red ? `<span class="cnt red ms-1" title="Red flags">${red}</span>` : ""}
        </div>
        <div class="side-b ${isOpen ? "" : "d-none"}">${isOpen ? sideBody(d) : ""}</div></div>`;
    }).join("") || '<div class="text-muted p-3">No project matches.</div>');
  }

  function toggleProj(el, force) {
    const $p = $(el).closest(".side-proj"), p = String($p.data("p"));
    const show = force ?? !open.has(p);
    if (show) open.add(p); else open.delete(p);
    const q = ($("#sideQ").val() || "").trim().toLowerCase();
    const d = side.find((x) => x.summary.pno === p);
    $p.find(".side-b").toggleClass("d-none", !show).html(show && d ? sideBody(filtered(d, q)) : "");
    $p.find(".side-h > i").attr("class", `fa-solid fa-chevron-${show ? "down" : "right"} me-2`);
  }

  async function openGroupIn(pno, gid) {
    if (!proj || proj.pno !== pno) {
      loading(true, "Loading project " + pno + "…");
      try { proj = await getJSON("/api/recon/project?pno=" + encodeURIComponent(pno)); }
      catch (e) { toast("error", "Could not load the project", e.message); return; }
      finally { loading(false); }
      groupStatus = "";
      renderProject();
    }
    const g = proj.groups.find((x) => x.gid === gid);
    if (g) openGroup(g);
  }

  $("#sideList").on("click keydown", ".side-h", function (e) {
    if (e.type === "keydown" && e.key !== "Enter") return;
    toggleProj(this);
  });
  $("#sideList").on("click", ".si", function () { openGroupIn(String($(this).data("p")), Number($(this).data("g"))); });
  let sideT;
  $("#sideQ").on("input", () => { clearTimeout(sideT); sideT = setTimeout(renderSide, 200); });
  $("#sideScope").on("click", "button", function () {
    sideScope = $(this).data("scope");
    $("#sideScope button").removeClass("primary");
    $(this).addClass("primary");
    renderSide();
  });
  $("#sideExpand").on("click", () => { side.forEach((d) => open.add(d.summary.pno)); renderSide(); });
  $("#sideCollapse").on("click", () => { open.clear(); renderSide(); });

  // ------------------------------------------------------------------ events
  $(document).on("click", ".pick-sap", function () { picker("sap", $(this).data("skey")); });
  $(document).on("click", ".pick-sales", function () { picker("sales", $(this).data("wbs")); });
  $(document).on("click", ".do-link", function () {
    const skey = String($(this).data("skey")), wbs = String($(this).data("wbs"));
    override({ action: "link", pairs: [[skey, wbs]] }, { skey, wbs });
  });
  $("#gmPickerQ").on("input", function () {
    const q = this.value.toLowerCase();
    $("#gmPickerTbl tbody tr").each(function () { $(this).toggle(!q || String($(this).data("q") || "").includes(q)); });
  });
  $("#gmUnlink").on("click", async () => {
    const g = curGroup;
    const ok = await Swal.fire({
      icon: "question", title: "Unlink these items?", confirmButtonColor: "#ad283f", showCancelButton: true, confirmButtonText: "Unlink",
      text: "The sales-file row(s) and the SAP item(s) will no longer be matched to each other. You can undo this with \"Back to automatic\".",
    });
    if (ok.isConfirmed) override({ action: "unlink", pairs: pairsOf(g) }, { skey: g.sales[0]?.skey, wbs: g.sap[0]?.wbs });
  });
  $("#gmReset").on("click", () => {
    const g = curGroup;
    override({ action: "reset", pairs: itemsOf(g) }, { skey: g.sales[0]?.skey, wbs: g.sap[0]?.wbs });
  });
  $(document).on("change", "#groupModal input.note", async function () {
    const skey = $(this).data("skey"), wbs = $(this).data("wbs");
    try {
      proj = await (await postJSON("/api/recon/override", { pno: proj.pno, action: "note", skey: skey ? String(skey) : null, wbs: wbs ? String(wbs) : null, note: this.value.trim() })).json();
      renderProject();
      toast("success", "Note saved");
    } catch (e) { toast("error", "Note not saved", e.message); }
  });
  $("#projSelect").on("change", function () { openProject(this.value); });
  $("#projScope").on("click", "button", function () {
    projScope = $(this).data("scope");
    $("#projScope button").removeClass("primary");
    $(this).addClass("primary");
    renderProjects();
  });
  $("#btnSync").on("click", () => { proj = null; loadOverview(true); });
  // DataTables inside hidden tabs need a column re-measure when shown
  $('button[data-bs-toggle="tab"]').on("shown.bs.tab", () => Object.values(tables).forEach((t) => t.columns.adjust()));

  loadOverview();
})();
