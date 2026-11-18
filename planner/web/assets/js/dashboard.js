/* Sales & Cash Flow dashboard. */
(() => {
  const { inr, cr, esc, getJSON } = YD;
  const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  let chart, dash, tables = {};

  const statusBadge = (s) =>
    `<span class="badge-st st-${s.toLowerCase()}">${esc(s)}</span>`;
  const riskChip = (r) => (r ? `<span class="risk-chip">${esc(r)}</span>` : "");
  const money = (v) => ({ display: inr(v), sort: v });

  function dt(id, opts) {
    if (tables[id]) { tables[id].clear().rows.add(opts.data).draw(); return tables[id]; }
    tables[id] = $("#" + id).DataTable({
      pageLength: 25, lengthMenu: [10, 25, 50, 100], autoWidth: false, deferRender: true,
      language: { search: "", searchPlaceholder: "Search…" }, ...opts,
    });
    return tables[id];
  }

  // ------------------------------------------------------------------ KPIs
  function renderKpis(k, fyLabel) {
    const pct = k.fy_total ? Math.round((k.fy_dispatched / k.fy_total) * 100) : 0;
    const mpct = k.month_plan ? Math.round((k.month_dispatched / k.month_plan) * 100) : 0;
    $("#kpis").html(`
      <div class="kpi"><div class="kpi-label">${esc(fyLabel)} billing plan</div><div class="kpi-value">${cr(k.fy_total)}</div>
        <div class="kpi-sub"><b>${cr(k.fy_dispatched)}</b> dispatched · ${pct}%</div><div class="bar"><span style="width:${pct}%"></span></div>
        ${dash.fy_target ? `<div class="kpi-sub mt-1">FY target <b>${cr(dash.fy_target)}</b> · plan is ${Math.round((k.fy_total / dash.fy_target) * 100)}% of it</div>` : ""}</div>
      <div class="kpi"><div class="kpi-label">${esc(dash.current_month)} plan</div><div class="kpi-value">${cr(k.month_plan)}</div>
        <div class="kpi-sub"><b>${cr(k.month_dispatched)}</b> dispatched so far · ${mpct}%</div><div class="bar"><span style="width:${mpct}%"></span></div></div>
      <div class="kpi"><div class="kpi-label">Next 3 months · ${esc(k.next3_label)}</div><div class="kpi-value">${cr(k.next3)}</div>
        <div class="kpi-sub">still to dispatch</div></div>
      <div class="kpi red"><div class="kpi-label">Slipped (month passed)</div><div class="kpi-value">${cr(k.slipped_value)}</div>
        <div class="kpi-sub"><b>${k.slipped_count}</b> assemblies not dispatched in their month</div></div>
      <div class="kpi"><div class="kpi-label">Active projects</div><div class="kpi-value">${k.projects_active}</div>
        <div class="kpi-sub">with billing still to dispatch</div></div>`);
  }

  // ------------------------------------------------------------------ chart
  function renderChart(series, fyLabel) {
    $("#chartTitle").text(`Month-wise billing · ${fyLabel}`);
    const labels = series.map((s) => s.month.replace(" 20", " '"));
    const keys = ["dispatched", "slipped", "planned"];
    const colors = [css("--s-dispatched"), css("--s-slipped"), css("--s-planned")];
    const names = ["Dispatched", "Slipped", "Planned"];
    const topIndex = (i) => { for (let d = keys.length - 1; d >= 0; d--) if (series[i][keys[d]] > 0) return d; return -1; };
    const datasets = keys.map((k, d) => ({
      label: names[d], data: series.map((s) => s[k] / 1e7), backgroundColor: colors[d],
      borderColor: "#ffffff", borderWidth: { top: 2, left: 0, right: 0, bottom: 0 }, borderSkipped: false,
      borderRadius: (ctx) => (topIndex(ctx.dataIndex) === d ? { topLeft: 4, topRight: 4 } : 0),
      maxBarThickness: 46, stack: "s",
    }));
    const cur = series.findIndex((s) => s.is_current);
    const cfg = {
      type: "bar", data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: "#1f2937", padding: 10, titleFont: { weight: "700" },
            callbacks: {
              title: (it) => `${series[it[0].dataIndex].month} · ${series[it[0].dataIndex].week}`,
              label: (it) => ` ${it.dataset.label}: ₹${it.raw.toFixed(2)} Cr`,
              footer: (it) => {
                const s = series[it[0].dataIndex];
                return `Total ₹${((s.dispatched + s.slipped + s.planned) / 1e7).toFixed(2)} Cr · ${s.assemblies} assemblies`;
              },
            },
          },
        },
        scales: {
          x: { stacked: true, grid: { display: false },
               ticks: { color: (c) => (c.index === cur ? css("--yd-red") : "#6c757d"), font: (c) => ({ weight: c.index === cur ? "800" : "600" }) } },
          y: { stacked: true, beginAtZero: true, grid: { color: "rgba(0,0,0,0.06)" }, border: { display: false },
               ticks: { color: "#6c757d", callback: (v) => "₹" + v + " Cr" } },
        },
        onClick: (_, els) => { if (els.length) openMonth(series[els[0].index].key); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? "pointer" : "default"; },
      },
    };
    if (chart) chart.destroy();
    chart = new Chart(document.getElementById("monthChart"), cfg);
  }

  function renderQuarters(q) {
    $("#quarters").html(q.map((x) => {
      const p = x.total ? Math.round((x.dispatched / x.total) * 100) : 0;
      return `<div class="q-tile"><div class="q-label">${esc(x.label)}</div><div class="q-val">${cr(x.total)}</div>
        <div class="q-sub">${cr(x.dispatched)} dispatched · ${p}%</div></div>`;
    }).join(""));
  }

  // ------------------------------------------------------------------ tables
  async function loadProjects() {
    const rows = await getJSON("/api/projects");
    $("#cntProjects").text(rows.length);
    dt("tblProjects", {
      data: rows, order: [[6, "desc"]],
      createdRow: (tr) => tr.classList.add("clickable"),
      columns: [
        { data: "project", render: (v) => `<b>${esc(v)}</b>` },
        { data: "line", render: (v) => `<span class="line-chip">${esc(v)}</span>` },
        { data: "pm", render: esc },
        { data: "total", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "dispatched", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "pct", render: (v, t) => (t === "display" ? `<div class="d-flex align-items-center gap-2"><div class="prog flex-grow-1"><span style="width:${Math.round(v * 100)}%"></span></div><small>${Math.round(v * 100)}%</small></div>` : v) },
        { data: "pending", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "slipped", className: "num", render: (v, t, r) => (t === "display" ? (v > 0 ? `<span class="text-danger fw-bold">${inr(v)}</span> <small class="text-muted">(${r.slipped_count})</small>` : "—") : v) },
        { data: "next_key", render: (v, t, r) => (t === "display" ? esc(r.next_month || "—") : v) },
        { data: "risks", render: (v, t) => (t === "display" ? (v ? `<span class="risk-chip">${v}</span>` : "") : v) },
      ],
    });
    $("#tblProjects tbody").off("click").on("click", "tr", function () {
      const r = tables.tblProjects.row(this).data();
      if (r) openProject(r);
    });
  }

  async function openProject(p) {
    $("#pmTitle").text(p.project);
    $("#pmKpis").html(`
      <div class="kpi"><div class="kpi-label">Total billing</div><div class="kpi-value">${cr(p.total)}</div><div class="kpi-sub">${p.assemblies} assemblies · ${esc(p.line)}</div></div>
      <div class="kpi"><div class="kpi-label">Dispatched</div><div class="kpi-value">${cr(p.dispatched)}</div><div class="kpi-sub">${Math.round(p.pct * 100)}% of total</div></div>
      <div class="kpi"><div class="kpi-label">Pending</div><div class="kpi-value">${cr(p.pending)}</div><div class="kpi-sub">next: ${esc(p.next_month || "—")}</div></div>
      <div class="kpi red"><div class="kpi-label">Slipped</div><div class="kpi-value">${cr(p.slipped)}</div><div class="kpi-sub">${p.slipped_count} assemblies</div></div>`);
    const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById("projectModal"));
    modal.show();
    const rows = await getJSON("/api/project?name=" + encodeURIComponent(p.project));
    dt("tblProjectRows", {
      data: rows, order: [[3, "asc"]], pageLength: 50,
      columns: [
        { data: "assembly", render: esc }, { data: "type", render: esc },
        { data: "bill", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "key", render: (v, t, r) => (t === "display" ? esc(r.month) : v) },
        { data: "week", className: "text-nowrap", render: esc },
        { data: "status", render: (v, t) => (t === "display" ? statusBadge(v) : v) },
        { data: "risk", render: (v, t) => (t === "display" ? riskChip(v) : v) },
        { data: "remarks", render: esc },
      ],
    });
  }

  async function openMonth(key) {
    $("#monthSelect").val(key);
    bootstrap.Tab.getOrCreateInstance(document.getElementById("tabMonthBtn")).show();
    await loadMonth(key);
    document.getElementById("tabMonthBtn").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function loadMonth(key) {
    const rows = await getJSON("/api/month?key=" + key);
    const tot = rows.reduce((a, r) => a + r.bill, 0);
    const disp = rows.filter((r) => r.status === "Dispatched").reduce((a, r) => a + r.bill, 0);
    $("#monthSummary").html(`${rows.length} assemblies · <b>${cr(tot)}</b> planned · <b>${cr(disp)}</b> dispatched`);
    dt("tblMonth", {
      data: rows, order: [[4, "desc"]],
      columns: [
        { data: "project", render: (v) => `<b>${esc(v)}</b>` }, { data: "assembly", render: esc }, { data: "type", render: esc },
        { data: "pm", render: esc },
        { data: "bill", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "status", render: (v, t) => (t === "display" ? statusBadge(v) : v) },
        { data: "risk", render: (v, t) => (t === "display" ? riskChip(v) : v) },
        { data: "remarks", render: esc },
      ],
    });
  }

  async function loadAttention() {
    const rows = await getJSON("/api/attention");
    $("#cntAttention").text(rows.length);
    dt("tblAttention", {
      data: rows, order: [[2, "asc"]],
      columns: [
        { data: "project", render: (v) => `<b>${esc(v)}</b>` }, { data: "assembly", render: esc },
        { data: "key", render: (v, t, r) => (t === "display" ? esc(r.month) : v) },
        { data: "bill", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "status", render: (v, t) => (t === "display" ? statusBadge(v) : v) },
        { data: "risk", render: (v, t) => (t === "display" ? riskChip(v) : v) },
        { data: "pm", render: esc }, { data: "remarks", render: esc },
      ],
    });
  }

  async function loadCash() {
    const c = await getJSON("/api/cashflow");
    if (!c.rows.length) { $("#cashWrap").html('<p class="text-muted">No CASH FLOW sheet found in the workbook.</p>'); return; }
    const head = `<tr><th>Project</th>${c.columns.map((x) => `<th class="num">${esc(x)}</th>`).join("")}</tr>`;
    const body = c.rows.map((r) => `<tr${/grand total/i.test(r.project) ? ' class="fw-bold"' : ""}><td>${esc(r.project)}</td>${r.values.map((v) => `<td class="num">${v ? inr(v) : ""}</td>`).join("")}</tr>`).join("");
    $("#cashWrap").html(`<table class="table table-bordered table-hover yd-table align-middle mb-0"><thead>${head}</thead><tbody>${body}</tbody></table>`);
  }

  async function lookup() {
    const q = $("#lookupQ").val().trim();
    if (!q) return;
    const rows = await getJSON("/api/lookup?q=" + encodeURIComponent(q));
    const vals = rows.map((r) => r.bill).sort((a, b) => a - b);
    const med = vals.length ? vals[Math.floor(vals.length / 2)] : null;
    $("#lookupSummary").html(rows.length ? `${rows.length} matches · median <b>${inr(med)}</b> · range ${inr(vals[0])} – ${inr(vals[vals.length - 1])}` : "No match");
    dt("tblLookup", {
      data: rows, order: [[4, "desc"]],
      columns: [
        { data: "project", render: (v) => `<b>${esc(v)}</b>` }, { data: "assembly", render: esc }, { data: "type", render: esc },
        { data: "bill", className: "num", render: (v, t) => (t === "display" ? inr(v) : v) },
        { data: "key", render: (v, t, r) => (t === "display" ? esc(r.month) : v) },
        { data: "status", render: (v, t) => (t === "display" ? statusBadge(v) : v) },
      ],
    });
  }

  // ------------------------------------------------------------------ boot
  async function loadDashboard(fy) {
    dash = await getJSON("/api/dashboard" + (fy ? "?fy=" + fy : ""));
    const sel = $("#fySelect");
    if (!sel.children().length) {
      dash.fys.forEach((y) => sel.append(`<option value="${y}">FY ${y}-${String(y + 1).slice(-2)}</option>`));
    }
    sel.val(dash.fy);
    $("#pageSub").text(`Billing plan, dispatch status and forecast for every project · as of ${dash.today}`);
    renderKpis(dash.kpi, dash.fy_label);
    renderChart(dash.series, dash.fy_label);
    renderQuarters(dash.quarters);
    const ms = $("#monthSelect").empty();
    dash.series.forEach((s) => ms.append(`<option value="${s.key}">${s.month} · ${s.week}</option>`));
    const cur = dash.series.find((s) => s.is_current) || dash.series[0];
    ms.val(cur.key);
    await loadMonth(cur.key);
  }

  async function boot(refresh = false) {
    YD.loading(true, refresh ? "Re-reading the sales workbook…" : "Loading sales & cash flow…");
    try {
      await YD.syncChip(refresh);
      await loadDashboard($("#fySelect").val());
      await Promise.all([loadProjects(), loadAttention(), loadCash()]);
    } catch (e) {
      YD.toast("error", "Could not load data", e.message);
    } finally {
      YD.loading(false);
    }
  }

  $(function () {
    $("#fySelect").on("change", function () { YD.loading(true); loadDashboard(this.value).finally(() => YD.loading(false)); });
    $("#monthSelect").on("change", function () { loadMonth(this.value); });
    $("#btnSync").on("click", () => boot(true));
    $("#btnAttention").on("click", () => {
      bootstrap.Tab.getOrCreateInstance(document.getElementById("tabAttentionBtn")).show();
      document.getElementById("tabAttentionBtn").scrollIntoView({ behavior: "smooth" });
    });
    $("#lookupGo").on("click", lookup);
    $("#lookupQ").on("keydown", (e) => { if (e.key === "Enter") lookup(); });
    document.querySelectorAll('button[data-bs-toggle="tab"]').forEach((b) =>
      b.addEventListener("shown.bs.tab", () => Object.values(tables).forEach((t) => t.columns.adjust())));
    boot();
  });
})();
