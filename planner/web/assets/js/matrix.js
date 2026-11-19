/* Assembly Matrix: one project, every assembly × month. Billed → SAP value & date; otherwise the forecast. */
(() => {
  const { inr, cr, esc, getJSON, toast, loading } = YD;
  const KIND = {
    both: { label: "In both", cls: "both", title: "In the sales file and billed in SAP — SAP value and billing date" },
    sap: { label: "Only in SAP", cls: "sap", title: "Billed in SAP, not in the sales file" },
    sales: { label: "Only in sales file", cls: "sales", title: "In the sales file, not billed in SAP — tentative value and month" },
  };
  let ov = null, data = null, kind = "";

  const mLabel = (ym) => new Date(ym + "-01T00:00:00").toLocaleDateString("en-IN", { month: "short", year: "2-digit" });
  const dLabel = (d, k) => (!d ? "—" : k === "sales"
    ? new Date(d).toLocaleDateString("en-IN", { month: "short", year: "numeric" })
    : new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }));
  const short = (v) => (v ? cr(v, 2).replace("₹", "") : "");

  // ------------------------------------------------------------------ project list
  async function loadList() {
    if (!ov) {
      loading(true, "Reading the sales file and SAP…");
      try { ov = await getJSON("/api/recon/overview"); }
      catch (e) { toast("error", "Could not load the projects", e.message); return; }
      finally { loading(false); }
      $("#srcChip .t").text("Sales: " + (ov.sales_source.startsWith("Live") ? "live" : "local") + " · " + (ov.sap_file.startsWith("SAP live") ? "SAP live" : "SAP file"));
      $("#srcChip").attr("title", `Sales file: ${ov.sales_source}\nSAP: ${ov.sap_file}, as of ${ov.sap_saved}`);
    }
    renderList();
  }

  function renderList() {
    const q = ($("#listQ").val() || "").trim().toLowerCase();
    const rows = ov.projects
      .filter((p) => !q || [p.project, p.customer, p.pm].some((t) => String(t || "").toLowerCase().includes(q)))
      .sort((a, b) => b.pno.localeCompare(a.pno));
    $("#listCount").text(`${rows.length} projects`);
    $("#projGrid").html(rows.map((p) => `
      <a class="mx-card" href="/matrix?pno=${encodeURIComponent(p.pno)}" data-p="${esc(p.pno)}">
        <div class="mx-card-t">${esc(p.project)}</div>
        <div class="small text-muted">${esc([p.customer, p.pm ? "PM " + p.pm : ""].filter(Boolean).join(" · ") || " ")}</div>
        <div class="mx-card-n">
          <span><i>Forecast</i> ${cr(p.forecast)}</span><span><i>SAP billed</i> ${cr(p.sap_billed)}</span>
        </div>
        ${p.in_sales ? "" : '<span class="line-chip">not in sales file</span>'}
        ${p.red ? `<span class="cnt red" title="Red flags">${p.red}</span>` : ""}
      </a>`).join("") || '<div class="text-muted">No project matches.</div>');
  }

  // ------------------------------------------------------------------ one project
  async function openProject(pno, push = true) {
    loading(true, "Loading " + pno + "…");
    try { data = await getJSON("/api/recon/matrix?pno=" + encodeURIComponent(pno)); }
    catch (e) { toast("error", "Could not load the project", e.message); return; }
    finally { loading(false); }
    if (push) history.pushState({ pno }, "", "/matrix?pno=" + encodeURIComponent(pno));
    kind = "";
    $("#rowQ").val("");
    $("#viewList").addClass("d-none");
    $("#viewProject").removeClass("d-none");
    const s = data.summary;
    $("#pTitle").text(s.project);
    $("#pSub").text([s.customer, s.pm ? "PM " + s.pm : ""].filter(Boolean).join(" · "));
    $("#btnExport").attr("href", "/api/recon/matrix_export?pno=" + encodeURIComponent(s.pno));
    document.title = s.pno + " · Assembly Matrix | YOGIJI DIGI";
    renderKpis();
    renderPills();
    renderMatrix();
    window.scrollTo(0, 0);
  }

  function renderKpis() {
    const R = data.rows, sum = (k) => R.filter((r) => r.kind === k).reduce((t, r) => t + (r.value || 0), 0);
    const n = (k) => R.filter((r) => r.kind === k).length;
    const s = data.summary;
    $("#pKpis").html(`
      <div class="kpi"><div class="kpi-label">Assemblies</div><div class="kpi-value">${R.length}</div><div class="kpi-sub">sales file ${s.forecast ? cr(s.forecast) : "—"} · SAP billed ${cr(s.sap_billed)}</div></div>
      <div class="kpi green"><div class="kpi-label">In both</div><div class="kpi-value">${n("both")}</div><div class="kpi-sub"><b>${cr(sum("both"))}</b> billed in SAP</div></div>
      <div class="kpi amber"><div class="kpi-label">Only in SAP</div><div class="kpi-value">${n("sap")}</div><div class="kpi-sub"><b>${cr(sum("sap"))}</b> billed, not in sales file</div></div>
      <div class="kpi"><div class="kpi-label">Only in sales file</div><div class="kpi-value">${n("sales")}</div><div class="kpi-sub"><b>${cr(sum("sales"))}</b> tentative, not billed</div></div>
      <div class="kpi"><div class="kpi-label">Expected total</div><div class="kpi-value">${cr(sum("both") + sum("sap") + sum("sales"))}</div><div class="kpi-sub">billed + still to bill</div></div>`);
  }

  function renderPills() {
    const c = (k) => data.rows.filter((r) => !k || r.kind === k).length;
    $("#kindPills").html([["", "All"], ["both", "In both"], ["sap", "Only in SAP"], ["sales", "Only in sales file"]]
      .map(([k, l]) => `<button class="pill ${k ? "k-" + k : ""} ${kind === k ? "on" : ""}" data-k="${k}">${l} <b>${c(k)}</b></button>`).join(""));
  }

  function renderMatrix() {
    const q = ($("#rowQ").val() || "").trim().toLowerCase();
    const rows = data.rows.filter((r) => (!kind || r.kind === kind) &&
      (!q || (r.name + " " + r.other + " " + r.ref).toLowerCase().includes(q)));
    const M = data.months, hasNone = rows.some((r) => !r.month);
    const cols = [...M, ...(hasNone ? [""] : [])];
    const cur = new Date().toISOString().slice(0, 7);
    $("#mx thead").html(`<tr><th class="stick">Assembly</th><th>Where</th><th class="num">Value</th><th>Billing date / month</th>
      ${cols.map((m) => `<th class="num mcol ${m === cur ? "now" : ""}">${m ? mLabel(m) : "No month"}</th>`).join("")}</tr>`);
    const tot = Object.fromEntries(cols.map((m) => [m, 0]));
    $("#mx tbody").html(rows.map((r) => {
      tot[r.month || ""] += r.value || 0;
      const K = KIND[r.kind];
      const flag = ["ok", "pending", "sap_only"].includes(r.status) ? "" : `<div><span class="st-chip tone-${r.tone}">${esc(r.label)}</span></div>`;
      const sub = r.kind === "both"
        ? `<div class="mx-sub">SAP: ${esc(r.other)} · forecast ${inr(r.f_value)}, ${esc(r.f_month || "—")}${r.disp ? " · " + esc(r.disp) : ""}</div>`
        : r.kind === "sap" ? `<div class="mx-sub">${esc(String(r.ref).replace(/^ITEM:/, ""))}</div>`
        : `<div class="mx-sub">row ${esc(r.ref)} · ${esc(r.disp)}</div>`;
      return `<tr class="k-${K.cls}" data-g="${r.gid}">
        <td class="stick"><div class="mx-name">${esc(r.name)}</div>${sub}${flag}</td>
        <td><span class="kind-chip ${K.cls}" title="${esc(K.title)}">${K.label}</span></td>
        <td class="num"><b>${inr(r.value)}</b></td>
        <td class="nowrap">${dLabel(r.date, r.kind)}${r.bills > 1 ? ` <small class="text-muted">(${r.bills} bills)</small>` : ""}</td>
        ${cols.map((m) => (m === (r.month || "") ? `<td class="num cell ${K.cls}" title="${esc(K.title)}">${short(r.value)}</td>` : `<td class="${m === cur ? "now" : ""}"></td>`)).join("")}
      </tr>`;
    }).join("") || `<tr><td class="stick text-muted" colspan="${4 + cols.length}">Nothing here.</td></tr>`);
    const all = rows.reduce((t, r) => t + (r.value || 0), 0);
    $("#mx tfoot").html(`<tr><td class="stick">Total (${rows.length})</td><td></td><td class="num">${inr(all)}</td><td></td>
      ${cols.map((m) => `<td class="num">${short(tot[m])}</td>`).join("")}</tr>`);
  }

  // ------------------------------------------------------------------ events
  $("#listQ").on("input", renderList);
  $("#projGrid").on("click", ".mx-card", function (e) { e.preventDefault(); openProject(String($(this).data("p"))); });
  $("#kindPills").on("click", ".pill", function () { kind = $(this).data("k"); renderPills(); renderMatrix(); });
  $("#rowQ").on("input", renderMatrix);
  $("#backLink").on("click", (e) => { e.preventDefault(); history.pushState({}, "", "/matrix"); showList(); });
  function showList() {
    $("#viewProject").addClass("d-none");
    $("#viewList").removeClass("d-none");
    document.title = "Assembly Matrix | YOGIJI DIGI";
    loadList();
  }
  window.addEventListener("popstate", () => {
    const p = new URLSearchParams(location.search).get("pno");
    if (p) openProject(p, false); else showList();
  });

  const start = new URLSearchParams(location.search).get("pno");
  loadList().then(() => { if (start) openProject(start, false); });
})();
