/* Shared helpers for the planner pages. */
const YD = (() => {
  const inrFmt = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 });

  function inr(v) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    return "₹" + inrFmt.format(Math.round(v));
  }
  function plain(v) {
    if (v === null || v === undefined || isNaN(v) || v === "") return "";
    return inrFmt.format(Math.round(v));
  }
  /** ₹ 12.34 Cr / ₹ 45.6 L */
  function cr(v, digits = 2) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    const a = Math.abs(v);
    if (a >= 1e7) return "₹" + (v / 1e7).toFixed(digits) + " Cr";
    if (a >= 1e5) return "₹" + (v / 1e5).toFixed(1) + " L";
    return inr(v);
  }
  function parseNum(s) {
    if (s === null || s === undefined) return NaN;
    const t = String(s).replace(/[^\d.\-]/g, "");
    return t === "" ? NaN : parseFloat(t);
  }
  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  async function api(path, opts = {}) {
    const res = await fetch(path, opts);
    if (!res.ok) {
      let msg = res.statusText;
      try { const j = await res.json(); msg = j.detail || JSON.stringify(j); } catch (_) {}
      const err = new Error(msg); err.status = res.status; throw err;
    }
    return res;
  }
  async function getJSON(path) { return (await api(path)).json(); }
  async function postJSON(path, body) {
    return api(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  }
  async function download(path, body, fallbackName) {
    const res = await postJSON(path, body);
    const blob = await res.blob();
    const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename="([^"]+)"/);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = m ? m[1] : fallbackName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function loading(on, text = "Working…") {
    let el = document.getElementById("loadingOverlay");
    if (!el) {
      el = document.createElement("div");
      el.id = "loadingOverlay"; el.className = "loading-overlay";
      el.innerHTML = '<div class="box"><span class="spinner-border spinner-border-sm me-2"></span><span class="t"></span></div>';
      document.body.appendChild(el);
    }
    el.querySelector(".t").textContent = text;
    el.classList.toggle("show", !!on);
  }
  function toast(icon, title, text) {
    if (window.Swal) Swal.fire({ icon, title, text, toast: icon === "success", position: icon === "success" ? "top-end" : "center",
      timer: icon === "success" ? 2600 : undefined, showConfirmButton: icon !== "success", confirmButtonColor: "#2a5071" });
    else alert(title + (text ? "\n" + text : ""));
  }
  async function syncChip(refresh = false) {
    const chip = document.getElementById("syncChip");
    if (!chip) return null;
    try {
      const s = await getJSON("/api/status" + (refresh ? "?refresh=true" : ""));
      chip.classList.remove("err");
      chip.querySelector(".t").textContent = `${s.source} · saved ${s.saved}`;
      chip.title = `${s.file}\n${s.projects} projects · ${s.assembly_rows.toLocaleString("en-IN")} valued assemblies\n` +
        `${s.reference_plans} reference plans · ${s.priority_index_rows} priority-index rows\nRead at ${s.loaded}`;
      return s;
    } catch (e) {
      chip.classList.add("err");
      chip.querySelector(".t").textContent = "Data source error";
      chip.title = e.message;
      return null;
    }
  }
  return { inr, plain, cr, parseNum, esc, api, getJSON, postJSON, download, loading, toast, syncChip };
})();
