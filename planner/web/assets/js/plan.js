/* New plan wizard: LoE -> project & PO details -> Add Plan (dispatch week, billing value, category). */
(() => {
  const { inr, plain, cr, parseNum, esc, getJSON, postJSON } = YD;
  const CATS = ["", "0", "1", "2", "3", "4", "3/4", "5"];
  let loe = null;          // /api/loe response
  let plan = null;         // /api/plan response
  let rows = [];           // working rows (edited)
  let projectEdited = false;

  // ------------------------------------------------------------------ stepper
  function go(step) {
    [1, 2, 3].forEach((s) => $("#step" + s).toggleClass("d-none", s !== step));
    $("#stepper .stepper-item").each(function () {
      const s = +this.dataset.step;
      this.classList.toggle("active", s === step);
      this.classList.toggle("completed", s < step);
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  $("#stepper").on("click", ".stepper-item.completed", function () { go(+this.dataset.step); });

  // ------------------------------------------------------------------ step 1: LoE
  async function upload(file) {
    if (!file) return;
    const fd = new FormData();
    fd.append("file", file);
    YD.loading(true, "Reading the LoE…");
    try {
      loe = await (await YD.api("/api/loe", { method: "POST", body: fd })).json();
      showLoe();
      prefill();
    } catch (e) {
      YD.toast("error", "Couldn't read this LoE", e.message);
    } finally { YD.loading(false); }
  }

  function showLoe() {
    $("#loeMeta").removeClass("d-none");
    $("#loeName").text(loe.filename);
    $("#loeSheet").html(loe.sheets.map((s) => `<option${s === loe.sheet ? " selected" : ""}>${esc(s)}</option>`).join(""));
    $("#loeCounts").html(`<b>${loe.rows}</b> rows read${loe.bought_out_rows ? ` · ${loe.bought_out_rows} in the bought-out list` : ""}`);
    const info = Object.entries(loe.info).map(([k, v]) => `<div><b>${esc(k.charAt(0) + k.slice(1).toLowerCase())}:</b> ${esc(v)}</div>`).join("");
    $("#loeInfo").html(info || "No header details found in the LoE.");
    $("#previewTable tbody").html(loe.preview.map((r) => `<tr><td class="c">${esc(r.sr)}</td><td>${esc(r.section)}</td><td>${esc(r.name)}</td><td class="c">${esc(r.drawing)}</td><td class="c">${esc(r.qty)}</td></tr>`).join(""));
    $("#toStep2").prop("disabled", !loe.rows);
  }

  function prefill() {
    const g = loe.guess;
    if (g.project_no) $("#fPno").val(g.project_no);
    if (g.client) $("#fClient").val(g.client);
    if (g.family) $("#fFamily").val(g.family);
    projectEdited = false;
    autoProject();
  }

  function autoProject() {
    if (projectEdited) return;
    const pno = $("#fPno").val().trim();
    const client = ($("#fClient").val().trim().split(/[\s,]+/)[0] || "").toUpperCase().replace(/[^A-Z0-9&]/g, "");
    const fam = $("#fFamily").val();
    $("#fProject").val([pno, client, fam].filter(Boolean).join("-"));
  }

  // ------------------------------------------------------------------ step 2: details
  function hints() {
    const cur = $("#fCur").val();
    $(".fx-only").toggleClass("d-none", cur !== "USD");
    const v = parseNum($("#fValue").val()), fx = parseNum($("#fFx").val()) || 95, pct = parseNum($("#fPct").val()) || 100;
    if (!isNaN(v)) {
      const inrV = cur === "USD" ? v * fx : v;
      $("#fValueHint").html(cur === "USD"
        ? `$${plain(v)} × ₹${fx} = <b>${inr(inrV)}</b> (${cr(inrV)})${pct !== 100 ? ` · billable ${pct}% = <b>${cr(inrV * pct / 100)}</b>` : ""}`
        : `<b>${cr(inrV)}</b>${pct !== 100 ? ` · billable ${pct}% = <b>${cr(inrV * pct / 100)}</b>` : ""}`);
    } else $("#fValueHint").text("");
    const pd = $("#fPoDate").val(), dd = $("#fDel").val();
    let h = "";
    if (pd && dd) {
      const months = (new Date(dd) - new Date(pd)) / (30.44 * 864e5);
      if (months < 0) h = '<span class="text-danger fw-bold">Delivery date is before the PO date.</span>';
      else if (months > 24) h = `<span class="text-danger fw-bold">${Math.round(months)} months after PO — check the year.</span>`;
      else h = `${months.toFixed(1)} months after PO`;
      if (new Date(dd) < new Date()) h += ' · <span class="text-warning-emphasis fw-bold">already passed — dispatches will be planned from this month</span>';
    }
    $("#fDelHint").html(h);
    $(".lots-only").toggleClass("d-none", $("input[name=fMode]:checked").val() !== "lots");
  }

  function readForm() {
    const req = {
      token: loe.token, sheet: loe.sheet,
      project: $("#fProject").val().trim(), project_no: $("#fPno").val().trim(), client: $("#fClient").val().trim(),
      family: $("#fFamily").val(), pm: $("#fPm").val().trim(), po_no: $("#fPoNo").val().trim(),
      po_value: parseNum($("#fValue").val()), currency: $("#fCur").val(), fx: parseNum($("#fFx").val()) || 95,
      billable_pct: parseNum($("#fPct").val()) || 100, po_date: $("#fPoDate").val(), delivery_date: $("#fDel").val(),
      start_month: $("#fStart").val() ? $("#fStart").val() + "-01" : null,
      end_month: $("#fEnd").val() ? $("#fEnd").val() + "-01" : null,
      mode: $("input[name=fMode]:checked").val(), lots: +$("#fLots").val() || 3, payment_terms: $("#fTerms").val().trim(),
    };
    const missing = [];
    if (!req.project_no) missing.push("Project No");
    if (!req.project) missing.push("Project name");
    if (!(req.po_value > 0)) missing.push("PO value");
    if (!req.po_date) missing.push("PO date");
    if (!req.delivery_date) missing.push("Delivery date");
    return { req, missing };
  }

  async function generate() {
    const { req, missing } = readForm();
    if (missing.length) { YD.toast("warning", "Please fill in", missing.join(", ")); return; }
    if (new Date(req.delivery_date) < new Date(req.po_date)) { YD.toast("warning", "Check dates", "Delivery date is before the PO date."); return; }
    const months = (new Date(req.delivery_date) - new Date(req.po_date)) / (30.44 * 864e5);
    if (months > 24) {
      const ok = await Swal.fire({ icon: "question", title: "Delivery is " + Math.round(months) + " months after PO",
        text: "Is the delivery year right?", showCancelButton: true, confirmButtonText: "Yes, continue", confirmButtonColor: "#2a5071" });
      if (!ok.isConfirmed) return;
    }
    YD.loading(true, "Matching every assembly against past projects…");
    try {
      plan = await (await postJSON("/api/plan", req)).json();
      indexWeeks();
      rows = plan.rows.map((r) => ({ ...r, _orig: { v: r.billable ? r.billing_value : 0, m: r.dispatch_week, c: r.category } }));
      renderPlan();
      go(3);
    } catch (e) {
      YD.toast("error", "Could not generate the plan", e.message);
    } finally { YD.loading(false); }
  }

  // ------------------------------------------------------------------ step 3: Add Plan
  // Dispatch week = an actual week (WW). Weeks are grouped by the month they bill in; the last week of each
  // month is his month-end label, which the engine proposes by default.
  const weekMonth = {};
  const indexWeeks = () => plan.months.forEach((m) => (m.weeks || []).forEach((w) => { weekMonth[w.week] = m.key; }));
  const weekOf = (r) => r.dispatch_week || (plan.months.find((m) => m.key === r.dispatch_month) || {}).week || "";
  const monthOpt = (r) => {
    const cur = weekOf(r);
    return plan.months.map((m) => `<optgroup label="${m.label}">` + (m.weeks || []).map((w) =>
      `<option value="${w.week}"${w.week === cur ? " selected" : ""}>${w.week} · ${w.from}–${w.to}</option>`).join("") + "</optgroup>").join("");
  };
  const isUsd = () => plan && plan.po.currency === "USD";
  const setUsd = (r) => { if (isUsd()) r.billing_usd = r.billing_value > 0 ? Math.round(r.billing_value / (plan.po.fx || 95)) : null; };
  const catOpt = (c) => CATS.map((x) => `<option value="${x}"${x === (c || "") ? " selected" : ""}>${x || "—"}</option>`).join("");
  const valued = (r) => r.billable && r.billing_value > 0;

  function confChip(r) {
    if (!valued(r)) return '<span class="conf na">—</span>';
    const s = r.match_score || 0;
    if (!s) return '<span class="conf na" title="No similar assembly in history - typical price used">new</span>';
    return `<span class="conf ${s >= 85 ? "hi" : s >= 70 ? "mid" : "lo"}">${s}</span>`;
  }

  function rowHtml(r, i) {
    const v = valued(r);
    const tip = [`Why: ${r.reason || ""}`, r.basis ? `Value & timing from: ${r.basis}` : "Value: typical price (no close history)",
      `Category: ${r.category_basis || ""}`].join("\n");
    const pad = r.level === 1 ? "padding-left:22px" : "";
    const edited = r._orig && ((v ? r.billing_value : 0) !== r._orig.v || (v && r.dispatch_week !== r._orig.m) || r.category !== r._orig.c);
    return `<tr data-i="${i}" class="${v ? "" : "novalue"} ${edited ? "edited" : ""}">
      <td class="idx">${i + 1}</td>
      <td class="c">${esc(r.sr)}</td>
      <td class="asm"><div style="${pad}">${esc(r.name)}${r.has_children ? ' <span class="line-chip">group</span>' : ""}</div>
        <div class="why" style="${pad}">${esc(r.reason || "")}</div></td>
      <td class="c small">${esc(r.drawing || "")}</td>
      <td class="edit-cell"><select class="wk"${v ? "" : " disabled"}><option value="">—</option>${v ? monthOpt(r) : ""}</select></td>
      <td class="edit-cell"><input class="bv" value="${v ? plain(r.billing_value) : ""}" placeholder="—" inputmode="numeric" />${v && isUsd() && r.billing_usd ? `<div class="why text-end">US$ ${plain(r.billing_usd)}</div>` : ""}</td>
      <td class="edit-cell c"><select class="cat">${catOpt(v ? r.category : "")}</select></td>
      <td class="c text-nowrap">${confChip(r)} <i class="fa-solid fa-circle-info info-btn ms-1" title="${esc(tip)}"></i></td>
    </tr>`;
  }

  function visible(r) {
    const q = $("#apSearch").val().trim().toLowerCase();
    if (q && !(`${r.name} ${r.sr} ${r.drawing || ""}`.toLowerCase().includes(q))) return false;
    const f = $("input[name=apf]:checked").val();
    if (f === "valued") return valued(r);
    if (f === "none") return !valued(r);
    if (f === "low") return (valued(r) && (r.match_score || 0) < 70) || (!valued(r) && /no history/.test(r.reason || ""));
    return true;
  }

  function renderRows() {
    const html = rows.map((r, i) => (visible(r) ? rowHtml(r, i) : "")).join("");
    $("#apTable tbody").html(html || '<tr><td colspan="8" class="text-center text-muted py-4">No rows match</td></tr>');
  }

  function renderPlan() {
    const po = plan.po;
    $("#apSub").text(`Project: ${po.project} · ${rows.length} assemblies · PO ${po.currency} ${plain(po.po_value)}${po.po_no ? " · PO No " + po.po_no : ""}`);
    const vm = plan.value_mode === "project shares"
      ? "each assembly's typical <b>share of a project</b> in past " + esc(po.family) + " projects × PO value"
      : plan.value_mode === "item prices" ? "<b>typical prices</b> of the same items in past projects, scaled to the PO value (small / revamp order)"
      : plan.value_mode === "equal split (USD)" ? `the <b>USD PO value split evenly</b> over the assemblies (bills are cut in USD; INR at ₹${plan.po.fx} is for reference only)`
      : "<b>equal lots</b>";
    $("#methodNote").html(`<i class="fa-solid fa-circle-info me-1"></i> Billing values: ${vm}. Totals always equal the PO value
      (${cr(plan.target)}). Dispatch window: <b>${esc(plan.window.start)} – ${esc(plan.window.end)}</b>; each assembly is placed where it
      sits in past projects' dispatch sequence, never before its lead time. Category from the <b>Priority Index</b> where listed,
      otherwise from bought-out / long-lead rules. Rows marked <span class="conf lo">&lt;70</span> or <span class="conf na">new</span> need your check.`);
    renderRows();
    totals();
  }

  function totals() {
    const v = rows.filter(valued);
    const sum = v.reduce((a, r) => a + r.billing_value, 0);
    const diff = sum - plan.target;
    $("#totTarget").text(inr(plan.target));
    $("#totPlanned").text(inr(sum));
    $("#totDiff").html(Math.abs(diff) < 1 ? '<span class="ok"><i class="fa-solid fa-circle-check"></i> matches PO</span>' : `<span class="bad">${diff > 0 ? "+" : ""}${inr(diff)}</span>`);
    $("#totCount").text(`${v.length} / ${rows.length}`);
    $("#btnRebalance").prop("disabled", Math.abs(diff) < 1 || !v.length);
    const by = {};
    v.forEach((r) => { by[r.dispatch_month] = (by[r.dispatch_month] || 0) + r.billing_value; });
    const chips = Object.keys(by).sort().map((k) => {
      const m = plan.months.find((x) => x.key === k);
      return `<span class="month-chip">${m ? m.week : k} · <b>${cr(by[k])}</b></span>`;
    });
    $("#monthChips").html(chips.join(""));
  }

  function niceRound(v) {
    if (!v || v <= 0) return 0;
    const step = v < 1e6 ? 5e4 : v < 1e7 ? 1e5 : v < 1e8 ? 5e5 : 1e6;
    return Math.max(step, Math.round(v / step) * step);
  }

  function rebalance() {
    const v = rows.filter(valued);
    const sum = v.reduce((a, r) => a + r.billing_value, 0);
    if (!sum) return;
    v.forEach((r) => { r.billing_value = niceRound((r.billing_value / sum) * plan.target); });
    const big = v.reduce((a, r) => (r.billing_value > a.billing_value ? r : a), v[0]);
    big.billing_value += plan.target - v.reduce((a, r) => a + r.billing_value, 0);
    v.forEach(setUsd);
    renderRows(); totals();
  }

  // edits
  $("#apTable").on("change", "input.bv", function () {
    const i = +$(this).closest("tr").data("i"), r = rows[i];
    const n = parseNum(this.value);
    if (isNaN(n) || n <= 0) { r.billable = false; r.billing_value = 0; }
    else {
      r.billable = true; r.billing_value = Math.round(n); setUsd(r);
      if (!r.dispatch_month) r.dispatch_month = plan.po.start_month.slice(0, 7);
      if (!r.dispatch_week) r.dispatch_week = weekOf(r);
      if (!r.category) r.category = r._orig.c || "1";
    }
    $(this).closest("tr").replaceWith(rowHtml(r, i));
    totals();
  });
  $("#apTable").on("change", "select.wk", function () {
    const i = +$(this).closest("tr").data("i");
    if (this.value) { rows[i].dispatch_week = this.value; rows[i].dispatch_month = weekMonth[this.value] || rows[i].dispatch_month; }
    $(this).closest("tr").replaceWith(rowHtml(rows[i], i));
    totals();
  });
  $("#apTable").on("change", "select.cat", function () {
    const i = +$(this).closest("tr").data("i");
    rows[i].category = this.value;
    $(this).closest("tr").replaceWith(rowHtml(rows[i], i));
  });
  $("#apTable").on("focus", "input.bv", function () { this.value = this.value.replace(/,/g, ""); this.select(); });
  $("#apSearch").on("input", renderRows);
  $("input[name=apf]").on("change", renderRows);

  // exports
  const payload = (extra = {}) => ({ po: plan.po, rows: rows.map(({ _orig, ...r }) => r), token: loe.token, sheet: loe.sheet, ...extra });
  async function guarded(fn, label) {
    YD.loading(true, label);
    try { await fn(); } catch (e) { YD.toast("error", "Failed", e.message); } finally { YD.loading(false); }
  }
  $("#btnExport").on("click", () => guarded(() => YD.download("/api/export/plan", payload(), "Add Plan.xlsx"), "Preparing Add Plan…"));
  $("#btnLoe").on("click", () => {
    if (plan.po.mode === "lots") { YD.toast("info", "Lot-wise plan", "A lot-wise plan has no LoE rows — use Download Add Plan."); return; }
    guarded(() => YD.download("/api/export/loe", payload(), "planning.xlsx"), "Filling the LoE…");
  });
  $("#btnRef").on("click", async () => {
    if (plan.po.mode === "lots") { YD.toast("info", "Lot-wise plan", "Only assembly-wise plans can be saved as reference."); return; }
    const save = async (overwrite) => {
      const res = await fetch("/api/reference", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload({ overwrite })) });
      if (res.status === 409) {
        const j = await res.json();
        const ok = await Swal.fire({ icon: "warning", title: `${j.file} already exists`, text: "Replace it with this plan? Future plans will learn from the new one.",
          showCancelButton: true, confirmButtonText: "Replace", confirmButtonColor: "#ad283f" });
        if (ok.isConfirmed) return save(true);
        return;
      }
      if (!res.ok) throw new Error((await res.json()).detail || res.statusText);
      const j = await res.json();
      YD.toast("success", "Saved " + j.saved, "Future plans will learn from it.");
      YD.syncChip();
    };
    guarded(() => save(false), "Saving reference plan…");
  });
  $("#btnRebalance").on("click", rebalance);

  // ------------------------------------------------------------------ boot
  $(async function () {
    const s = await YD.syncChip();
    $("#fFamily").html((s ? s.families : ["CGL", "CRM", "CCL", "PICKLING", "SPM", "FINISHING", "ARP", "OTHER"]).map((f) => `<option>${f}</option>`).join(""));
    $("#loeFile").on("change", (e) => upload(e.target.files[0]));
    const dz = document.getElementById("dropzone");
    ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => upload(e.dataTransfer.files[0]));
    $("#loeSheet").on("change", async function () {
      YD.loading(true, "Reading sheet…");
      try { loe = await getJSON(`/api/loe/${loe.token}?sheet=${encodeURIComponent(this.value)}`); showLoe(); prefill(); }
      catch (e) { YD.toast("error", "Couldn't read this sheet", e.message); }
      finally { YD.loading(false); }
    });
    $("#toStep2").on("click", () => go(2));
    $("#back1").on("click", () => go(1));
    $("#back2").on("click", () => go(2));
    $("#generate").on("click", generate);
    $("#fPno, #fClient, #fFamily").on("input change", autoProject);
    $("#fProject").on("input", () => { projectEdited = true; });
    $("#fCur, #fValue, #fFx, #fPct, #fPoDate, #fDel, input[name=fMode]").on("input change", hints);
    $("#fValue").on("blur", function () { const n = parseNum(this.value); if (!isNaN(n)) this.value = plain(n); });
    $("#btnRestart").on("click", () => location.reload());
    hints();
  });
})();
