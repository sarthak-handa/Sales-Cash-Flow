# Sales & Cash Flow Planner — handover for any agent

> **Read this whole file before you do anything in this repo.** It records what the user told us, what we
> found in the data, what has been built, and what is still open. It is written so another agent (Claude Code,
> Google Antigravity / Gemini, Codex, …) can continue without re-asking. **Keep it current:** when you learn
> something new or change the system, update the matching section and add a line to the [session log](#10-session-log).

## Contents

1. [The user and the goal](#1-the-user-and-the-goal)
2. [How the planning manager works](#2-how-the-planning-manager-works-business-domain)
3. [Files in this repo](#3-files-in-this-repo)
4. [Findings from the data](#4-findings-from-the-data)
5. [The tool we built (`planner/`)](#5-the-tool-we-built-planner)
6. [The user's existing YD Ascend dashboard](#6-the-users-existing-yd-ascend-dashboard-file-dashboard)
7. [How the user wants us to work](#7-how-the-user-wants-us-to-work-preferences--feedback)
8. [Open questions for Shivendu sir](#8-open-questions-for-shivendu-sir-the-planner)
9. [Backlog / next steps](#9-backlog--next-steps-in-priority-order)
10. [Session log](#10-session-log)

---

## 1. The user and the goal

**The user** works at **Yogiji Digi Ltd (YD)**, Faridabad / Palwal, a maker of steel-processing lines: CGL, CRM, CCL,
pickling, SPM, trimming, ARP and others. Their **VP Projects** asked them to take over **all the work of the
planning manager, Shivendu sir**, who is leaving.

**Shivendu sir's background.** He joined as a consultant when YD planned to adopt the *Realization* streamliner
planning tool, and built the planning process described below.

**Immediate goal.** Upload a project's **LoE**, enter the PO and TS details, and get a **Dispatch Week**, **Billing
Value** and **Category** for every assembly, as accurately as possible. Accuracy is the stated top priority.

**Wider goal.** One place to see the live sales and cash-flow plan and forecast. The user also wants:

- the plan generated straight from a **TS (tech spec)**, before the LoE exists;
- **drawing release week** prediction later.

**What the user can provide:**

- more **TS / PO / LoE / completed planning sheets**, on request;
- answers from **Shivendu sir**, relayed by message (see [§8](#8-open-questions-for-shivendu-sir-the-planner));
- a **Gemini API key**, if AI processing is needed.

## 2. How the planning manager works (business domain)

### Project structure and full kit

Project → **LoE** (List of Equipment) → assemblies (with sub-assemblies, or independent) → BOM → components.

- **Full kit:** all components of all sub-assemblies of an assembly are manufactured and gathered. Only then can
  assembly start.

### Capacity

- Productivity is measured by **weight**.
- Shop-floor storage is limited by the **size** of finished assemblies.
- He has **no capacity data**; the user confirmed this.

### Categories

The 0–5 category is mostly *in his head*. Partly written down in `Priority Index for Lines (1) (8).xlsx`, which
was given to the user about a year ago: it covers every line type but is **reference only, not kept up to date**.

| Cat | Meaning |
|---|---|
| 0 | Not a full assembly. Items required by the long-lead Category 2 assemblies: mandrel drum, mandrel shaft, gearbox, motor, rolls (and back-up roll assemblies in SPM). |
| 1 | Manufactured, short lead (typically 2–3 months) |
| 2 | Manufactured, long lead (uncoilers, recoilers, bridle rolls — gearbox / mandrel / rolls / castings) |
| 3 | Bought out, shipped by the vendor **straight to the customer** — never comes in to the YD plant. |
| 4 | Bought out, **received at the YD plant**, then dispatched to the customer. |
| 5 | Long-lead bought out. Must be ordered very early (pot equipment, induction heater, tension leveller, welder…) |

### BOI / MANF. column values (sales file)

- `BOI` — bought out
- `BOI-LL` — long-lead bought out
- `UNIT-1`, `UNIT-2`, `UNIT-3` — YD manufacturing units
- `STRUCTURE`, `ELECTRICAL`

### How he plans dispatch and billing

For each assembly he sets a tentative **dispatch week** and **billing value**. He weighs:

- lead time;
- when the customer needs it;
- billing of similar past assemblies;
- the monthly revenue target;
- whether the customer will pay.

The **RISK** column names who is responsible for a risk: FB, PAYMENT, DRAWING, CUSTOMER, or a vendor name. An
assembly that misses its month moves to the next month, or sometimes later.

**Who does what (Shivendu sir, 2026-11-18).**

- **The PM** fills the RISK column and clears it when the risk is cleared. The PM also fills the long-lead flags,
  updates DISP. STATUS, and sends the status.
- **The PM** says whether something can be dispatched, e.g. when the customer is refusing delivery or not paying.
- **The design team** gives the drawing release week.
- **First and last dispatch months** are set by YD (his team and the PMs), not by the customer. The customer's
  required dates are in the TS.
- **Design time** after the PO is about **10 days at most** before work starts.
- **How he moves dates when something slips:**
  - **Customer refuses this month** → he postpones the item to the specific month the customer gives.
  - **Ordered item not received** → he moves it to the month we expect it to arrive.
  - **Made, but the customer says "hold"** → he still puts it in **next month**, every month, even knowing the
    customer won't take it, so the PM or leadership follow up. (So "on hold" rows look like normal next-month rows.)
- **DRG. REQ.** = ticked when the assembly still needs its drawing. **PENDING BOI ORDERING** = ticked when the item
  still has to be ordered. The PM unticks each one when it is done.
- **DESIGN CUTOFF DAYS** and **NATURE OF PAYMENT**: ignore (user, 2026-11-18).
- **Best reference project per line type** (his choice): **CGL 2428, CRM 2414, pickling 2533, CCL 2520,
  trimming 2239**. SPM: not given.
- **Typical PO → delivery period:** mill (CRM/SPM) **9 months**, CGL **12**, pickling **7**, trimming **6**.
- **No project starts dispatching before PO + 3 months** (drawings, ordering, etc.).
- **FY billing target:** ₹1,000 Cr (his words: "fy monthly target was set at 1k crore but it is somehow
  decreasing"). We read it as the annual target. It is set in `planner/config.json` → `fy_target_cr`.
- **YD Ascend** took its plan data from the sales file, so **the sales file is the master**. 2616 and 2619 are in
  YD Ascend but not active yet. There are no other open POs missing from the sales file.
- **Realization** will be discontinued: ignore it. There is no capacity data.

### The sales file

`SALES & CASH FLOW.xlsx` is the master forecast. The planner reads it through `planner/config.json`. On
2026-11-18 the user fixed the slipped rows (324 assemblies / ₹63 Cr → 26 / ₹4.8 Cr) in a "(1)" copy; on 2026-11-19
they replaced `SALES & CASH FLOW.xlsx` with it and deleted the copy. It has these sheets:

- **PROJECTS** — about 6,300 assembly rows across 90 projects, one row per assembly.
- **CASH FLOW** — a pivot of the payment requirement by month.
- **Q3 / Q4** — pivots by quarter.
- **Detail1–3** — the rows behind pivot drill-downs.

PROJECTS columns:

`PROJECT, BOI/MANF., PM, ASSEMBLY, RISK, BILLING VALUE ($), BILLING VALUE, DISP. STATUS (Disp./Not Disp.),
DISPATCH MONTH (stored as the 28th of the month), REMARKS, DESIGN CUTOFF DAYS, WEIGHT, VENDOR, DRG. REQ.,
PENDING BOI ORDERING, GEARBOX, MOTOR, MANDREL, MANDREL SHAFT, ROLLS, CASTING, FORGING, BEARING, LOAD CELLS,
COUPLING` (the last ten are 0/1 long-lead flags).

**Future data source.** The user will sync the live workbook through the **Microsoft Graph API**. The workbook is
updated regularly and reflects live status.

## 3. Files in this repo

> **Visibility.** The GitHub repo `github.com/sarthak-handa/Sales-Cash-Flow` is **public**. On 2026-11-18 the user
> confirmed this is intentional ("it's fine anyone can see this data"). Pushing to it is OK; don't raise it again.
>
> Sending files to *other* outside services (for example Gemini) still needs the user's OK.

| File | What it is |
|---|---|
| `SALES & CASH FLOW.xlsx` | Master forecast (see [§2](#the-sales-file)), slipped rows fixed by the user. The planner's knowledge base. |
| `reference projects.xlsx` | **His reference projects for billing value division** (added 2026-11-19). One sheet per line type — `cgl` = 2428 (₹59.73 Cr), `crm` = 2414 (₹29.37 Cr), `pickling` = 2533 (₹14.94 Cr), `ccl` = 2520 (₹41.85 Cr), `trimming` = 2239 (₹3.44 Cr) — each exactly that project's sales-file rows (BOI/MANF., ASSEMBLY, BILLING VALUE, DISPATCH MONTH…). See [§4](#reference-projects-billing-value-division). |
| `revenue data sap.xlsx` | **SAP billing export** (sheet `Query1`, ~6,100 invoice lines, Apr 2025 → 18 Nov 2026). Columns include Project, WBSDescription, ProductName (e.g. `2611-BRIDLE ROLL ASSY.#1`), Billing_date, doc_type (F2 invoice / S1 cancellation / G2 credit memo / S2), TransactionCurrency, Exchange_Rate, `TOTAL_SALE_WITHOUT_TAX` (INR, net of tax). Often open in Excel, so copy it before reading. The user committed it to the repo on 2026-11-18. See [§4](#sap-actual-billing). |
| `Priority Index for Lines (1) (8).xlsx` | One sheet per line type (Skin Pass Mill, Trimming, CRM, CCL, Pickling, CGL), each a reference project's LoE with **Priority (= category 0–5), Lead Time ("2-3M"), Remarks, BOI/Trading**. Reference projects: CGL 2214, CRM 2205, CCL 2217, Pickling 2302, Trimming 2401, SPM 2218. |
| `Time Sheet (1) (6).xlsx` | **V2:** engineering / fabrication / assembly days per sub-assembly, 481 rows across 8 line types. **V1:** a machine-hours template with only 4 rows. |
| `2611 planning.xlsx`, `2618 planning.xlsx` | **His completed plans** (CGL). LoE plus `Dispatch Week` ("WW49 2026") and `Billing Value (INR)` stored as text ("20,00,000"). Category column empty. These are the *reference plans* the engine learns from. |
| `2621 planning.xlsx`, `2630 planning.xlsx` | **His completed CRM plans** (2621 Prompt Steel 6HI CRM: ₹26.91 Cr over 76 rows, Apr–Sep 2027; 2630 JSW 6HI CRM: ₹25.26 Cr over 94 rows, Nov 26–May 27). Added 2026-11-18 (uploaded with random names, renamed). They value **only the mechanical LoE**: the sales file books ELECTRICAL, HT MOTORS and SPARES on top (2621 ₹36.5 Cr, 2630 ₹32 Cr in total). |
| `2619 planning.xlsx` | AMNS CGL modification. Engineering documents plus a bought-out list; no dispatch or billing filled in yet. |
| `*_2611_r1.pdf`, `*_2618_r1.pdf` | **Scanned** POs (images, no text layer) |
| `*po_5300063359*.pdf` | AMNS PO for 2619 (has text) |
| `2611 ts.pdf` | Text-readable tech spec |
| `2619 ts.pdf` | **Scanned** tech spec |
| `new projects/2616 *`, `new projects/2636 *` | LoE, PO and TS for the new projects |
| `new projects/2616 plan - dispatch & billing.xlsx` | Hand-made plan for 2616 (see [§4](#2616--hand-made-plan)) |
| `ydascend live` | Saved HTML of the user's **YD Ascend Projects Panel**. The UI style reference. |
| `dashboard` | Saved HTML of the user's own **YD Ascend Sales & Cash Flow dashboard** (see [§6](#6-the-users-existing-yd-ascend-dashboard-file-dashboard)) |
| `YDLOGO.png` | Company logo |
| `planner/` | The tool we built ([§5](#5-the-tool-we-built-planner)) |
| `Start Planner.bat` | Double-click to run the tool |

## 4. Findings from the data

### The PO is the billing anchor

**Confirmed by Shivendu sir:** the PO value always equals the total of the billing values. A billing value is the
**selling price** (what we bill the customer), not our cost. Erection, commissioning and supervision are secondary
and **not** part of this plan.

- **2618:** ₹79.80 Cr, which is 100% of the PO.
- **2611:** ₹57.97 Cr against a ₹59.71 Cr PO. The ₹1.7 Cr gap is **service bills** the company has already cut.
  SAP shows a ₹2.5 Cr "Supply of engineering services" invoice on 2611 in Sep 2026, under the E&C WBS.
- **Buffer.** He also keeps about a **₹3 Cr buffer** that he can include or exclude later.
- **Vendor price** comes from our PO to the vendor; he uses it as a guide for bought-out items.
- **Manufactured items** are valued by judgement plus past projects.
- **Group vs child rows:** "see the sales file for the rule". The engine already learns this from sales history
  and his plans.

### PO facts

| Project | Client | PO value | PO date | Delivery | Payment terms |
|---|---|---|---|---|---|
| 2611 | APL Apollo Sikandrabad (Al-Zn line 750 mm) | ₹59.71 Cr | 07-03-2026 (revised 28-05-2026) | 8 months (10-11-2026) | 10% advance / 10% civil drawings / 10% utility data / 62.5% against delivery / 7.5% after commissioning |
| 2618 | SG Mart Raipur (GI line 1500 mm) | ₹79.80 Cr | 13-03-2026 (revised 28-05-2026) | 15-03-2027 | Same as 2611 |
| 2619 | AMNS Gandhidham (Optigal conversion) | ₹3.73 Cr | 17-04-2026 | 04-11-2026 | 10% advance against ABG / 70% within 90 days of receipt / 10% at commissioning / 10% at PAC. Liquidated damages 0.5% per week, maximum 5%. |
| 2616 | APL Apollo Raipur (GP → GL revamp of project 2112) | ₹2,29,38,000 | 11-05-2026 (revised 02-06-2026) | 20-08-2026 | 30% advance / 60% against PI / balance after installation |
| 2636 | Constech Trading FZE, for MMI Uganda | US$1.5M ARP (+ US$1.2M 3-tank pickling line = project **2635**) | 31-07-2026 | 10 months | 10/10/10/60/5/5 |

### Billing is not cash

Cash follows the payment terms. On 2611 and 2618, only 62.5% of each dispatch's billing comes in at delivery.

### Dispatch week label

**The dispatch week is an actual week (WW).** In his 2611/2618 plans every label happens to be the month-end
week: the **ISO week number of (last day of the month + 1 day)**, e.g. Sep-26 → WW40 2026, Dec-26 → WW53 2026,
Jan-27 → WW5 2027, May-27 → WW22 2027. So the engine proposes the month-end week, and the user can pick any
week. A week belongs to the month of the Sunday before its ISO Monday (`engine.weeks_of_month` / `week_month`).
The sales file still stores the month, as the 28th.

### Dispatch window

His plans run from **PO date + 0.75 × the delivery period** to **PO delivery date + 2 months**, but never
before **PO + 3 months**. Long-lead bought-out items come last. When there is no delivery date, the typical period
for the line type is used (CRM/SPM 9, CGL/CCL 12, pickling 7, trimming 6 months). When the delivery date has
passed, the window moves to start **next month**.

### Lots outside the LoE (CRM)

On CRM, the electrical package, HT motors and spares are **not in the LoE**, but they are part of the PO and are
booked as separate sales-file rows. Over the CRM projects in history they are a median of **18%**, **9–10%** and
**5%** of a project. His CRM plans (2621, 2630) cover only the LoE, i.e. 74–79% of the project. On CGL, the
electrical lot is the LoE's "ELECTRICAL CONTROL ROOM" row, which he values.

### Typical order of dispatch on a CGL

1. Coil cars, drag chains, saddles, OBBS (Cat 1)
2. Tables and pit covers
3. Uncoilers (Cat 2)
4. Bridle rolls, accumulators, recoiler
5. Furnace and cleaning
6. TLL, SPM, pot, air knife, induction heater (Cat 4/5), last

### Which LoE rows get a value

He values the **dispatchable unit**:

- **Group rows:** "UNCOILER ASSY#1" is valued as one unit; its mandrel, gearbox and snubber get nothing. But a
  heading like "BRIDLE ROLL ASSY." is valued through each child, #1–#10.
- **Zero-value rows:** engineering documents, refractory, drive sub-assemblies and most "GA of …" rows.
- **Detail lists:** the skin pass mill's detailed parts list (un-sectioned rows inside a CGL LoE) is valued as
  one item, "SKIN PASS MILL".
- **Inconsistent between 2611 and 2618:** CAG and furnace structure, the welder, and the hydraulic power pack.

### Billing values are judgement

He picks round numbers. Big bought-out items such as the pot and inductors are priced from vendor cost: both POs
were revised over the price of Wisdom China vs Inductotherm India inductors.

### USD orders

**Rule (Shivendu sir):** the USD PO value is **divided evenly across the assemblies**, and **bills are cut in
USD**. The INR figure is only for reference in the plan, converted at the current rate, rounded (he used **₹95**
when the rate was moving between 93 and 96).

In the sales file, 2636 is booked as **3 equal lots of ₹3.8 Cr (Jan–Mar 2027) = ₹11.4 Cr**, which is 80% of
$1.5M × 95. The 80% is still unexplained; it is minor, since the INR is for reference only. 2635 is booked at
₹9.5 Cr.

### Gaps found

- **2616 and 2619 are not in the sales file at all.** That is ₹6 Cr missing from the forecast, and both PO
  delivery dates have passed. They are in YD Ascend but not active yet.
- **Slipped:** ₹63 Cr / 324 assemblies in the old file. **The user fixed this: now 26 / ₹4.8 Cr.**

### Reference projects (billing value division)

He gave `reference projects.xlsx`: "these are the reference projects for billing value division". A new
project's PO is divided the way its line type's reference project was billed. Tested 2026-11-19 against his own
2611/2618/2621/2630 plans:

- **His plans do not copy the reference rows.** Only 3–26% of his values equal the reference value; after scaling,
  31–54% are within ±30%. Same items get different values (pot equipment ₹2 Cr in 2611/2618 vs ₹1.04 Cr in 2428;
  TLL ₹1.4–1.8 Cr vs ₹2.05 Cr), and the reference is booked at a **finer level** (2428 books the entry accumulator as
  fixed frame ₹95 L + moving frame ₹75 L + structure ₹25 L; his plans have one "ENTRY ACCUMULATOR ASSY." ₹1.7–4 Cr;
  2414's back-up roll is split into rolls, chocks, bearings).
- **So the engine uses his own plans of the same line type first, then the reference project, then history.**
  Making the reference project the first source lowered CRM values within ±30% from 77% to 37–45%; summing
  "part" rows of the reference into one LoE assembly lowered CGL from 40% to 29%. Both were rejected.
- The reference file's DISPATCH MONTHs are **actual** dispatches (slips included), not his plan, so timing is
  not taken from it when history exists.
- For pickling, CCL and trimming (no plans of his yet) the reference project is the main value source.

### SAP actual billing

`revenue data sap.xlsx`, analysed on 2026-11-18:

- **FY 26-27 to date:** SAP billed **₹326 Cr** (₹314 Cr linked to a project), against ₹252.5 Cr marked "Disp."
  in the sales file for the same period.
- **Per project, most match closely:** 2428, 2517, 2527, 2601, 2540 and 2516 are within ±₹0.4 Cr.
- **Bigger gaps** (SAP above the sales file): 2518 +₹12.5 Cr, 2519 +₹7.2 Cr, 2514 +₹3.7 Cr, 2611 +₹3.3 Cr (the
  engineering services bill), 2529 +₹2.8 Cr. These are likely spares or services, or items billed but not yet
  marked "Disp.".
- **Item names** follow `<project>-<ASSEMBLY>`. Cancellations (S1) and credit memos (G2) are negative, so sum net.
- **Uses:** actual-vs-plan by month on the dashboard; checking "Disp." status; real item prices as extra history.

### 2616 — hand-made plan

Values sum to exactly the PO value, ₹2,29,38,000:

| Month | Week | Billing | Items |
|---|---|---|---|
| Dec 2026 | WW53 2026 | ₹88 L | Bought-out and fabricated items |
| Jan 2027 | WW5 2027 | ₹1.41 Cr | Pot roll system ₹1.31 Cr (Cat 5); water-cooled CAG roll ₹10 L (Cat 2) |

- **Pot roll system value is uncertain.** It is the balancing figure; ask for the vendor quote.
- **No value:** client-scope (ABPL) rows, and row 14 (design only).

### TS → LoE (tested 2026-11-18)

**The source.** TS documents are YD's own offers (offer numbers like `YD/CGL/046/2025/R2`). Chapter 3,
"SCOPE & SUPPLY – DIVISION LIST", is an equipment list:

- numbered `10 / 10.1 …`;
- with QTY and UNIT;
- with scope codes for BSD, BSE, DTE, **SUP**, INS, SIN, COM and SCO, where S = supplier (us) and B = buyer.

**What the prototype does** (`planner/ts_parse.py`):

- 2611: extracted **240 rows**, covering **60% of his billing value**.
- A plan built from the TS alone spread the money over the months about as well as the plan built from the LoE.

**What it misses:**

- electrical, furnace and bought-out items in other chapters, or under other names (ECR, skin pass mill, air knife);
- other layouts — 2636 uses lettered sub-items (A, B, C), and 2616's descriptions wrap onto two lines;
- 2619's TS, which is a scan.

**Status.** We asked the user (a) for the **source Word/Excel offers** from the proposals team, which would beat
PDF parsing, and (b) whether **Gemini** may be used for OCR / parsing.

## 5. The tool we built (`planner/`)

### Run it

Double-click `Start Planner.bat`, then open http://localhost:8600.

Or run it directly:

```bash
python -m uvicorn server:app --app-dir planner --port 8600
```

- **Python 3.14** with `pandas 3`, `openpyxl`, `rapidfuzz`, `fastapi`, `uvicorn` and `python-multipart`.
- There is **no auto-reload**: restart the server after changing Python code.

### UI

Built to **match YD Ascend**, which the user requires:

- Bootstrap 5.3, jQuery 3.6, DataTables 1.13.4, Font Awesome 6.5.2, SweetAlert2, and the Nunito font.
- Colours: navy `#2a5071`, red `#ad283f`, background `#f5f7fa`.
- Components: logo header with pill nav, a white rounded "table-container" card, red page title, navy toolbar
  buttons, navy table heads, a stepper.
- Chart series colours: `#2a78d6`, `#eb6834`, `#1baf7a` (validated colour-blind safe).

### Pages

**`/` Sales & Cash Flow:**

- FY KPI tiles: plan (with **% of the FY target**, from `config.json` → `fy_target_cr`), dispatched, current
  month, next 3 months, slipped, active projects.
- A month-wise stacked chart (Dispatched / Slipped / Planned). Clicking a month drills into it.
- Quarter tiles.
- Tabs: Projects (row opens a modal with the project's assemblies), Month detail, Needs attention (slipped or
  risk-flagged), Cash flow (the CASH FLOW sheet), Assembly lookup.

**`/plan` New Plan** — a three-step wizard:

1. Upload the LoE. Shows a preview and auto-fills the project number, client and line type.
2. Project & PO details. It validates the dates: it warns if delivery is more than 24 months after the PO (the
   user once typed 2027 for 2026) or already in the past.
3. **"Add Plan — Dispatch Week, Billing Value & Category"** table, the same columns as the YD Ascend "Add Plan"
   modal: SR No, Assembly, Drawing/Doc No, Dispatch Week, Billing Value (INR), Category.
   - **Basis column:** a match-score chip and an ⓘ tooltip.
   - **Filters:** all, with value, without value, "Check these".
   - **Live total vs PO**, with a "Rebalance to PO" button.
   - **Exports:** the Add Plan xlsx (sheets: Add Plan, Sales file rows, Month-wise, PO inputs); the LoE with the
     plan filled in; and **"Save as reference"**, which writes `<pno> planning.xlsx` to the repo root so future
     plans learn from it.

### Code

| File | Role |
|---|---|
| `engine.py` | Planning logic (below) |
| `sales.py` | Dashboard numbers. Uses the Indian financial year (Apr–Mar): slipped = not dispatched and month < current month. |
| `datasource.py` + `config.json` | `sales_source.type` is `local` (default) or `graph`. Graph mode downloads the workbook with client credentials (tenant_id, client_id, drive_id, item_id; secret in the `YD_GRAPH_CLIENT_SECRET` environment variable) to `cache/sales_live.xlsx` every 15 minutes. **Not yet configured — needs the user's IT team to create an Azure app.** |
| `server.py` | FastAPI. Knowledge is re-read automatically when the sales file, a `*planning*.xlsx` or the Priority Index changes. Endpoints: `/api/status`, `/dashboard`, `/projects`, `/project`, `/month`, `/attention`, `/cashflow`, `/lookup`, `/loe` (upload), `/plan`, `/export/plan`, `/export/loe`, `/reference`. |
| `backtest.py` | Re-plans 2611 and 2618 with each project hidden from the engine, and compares the result with his plans. **Run it before and after every engine change.** |
| `ts_parse.py` | TS division-list parser (prototype) |

### Engine: `plan(loe, po, kb)`

**1. Which rows get a value.** Checked in this order:

1. Engineering, document or drawing rows → no value.
2. **Scope-matrix LoE** (revamps like 2616: three or more rows say "supply…") → only rows marked "supply", or
   "YD", get a value.
3. A near-identical row (match ≥85) in his reference plans → valued if he valued it anywhere.
4. Strong sales history (≥80) → valued.
5. "SUPPLY" in the scope column → valued.
6. A line type with fewer than 50 history rows → everything valued by default.
7. Un-sectioned top-level rows in a sectioned LoE → treated as an equipment detail list, no value.

Then two clean-up steps:

- **Group rows:** follow how he handled the same group (valued as one unit vs valued through its children). If
  the group is valued as one unit, its children are dropped unless he valued them separately.
- **Duplicates:** differently-named rows in one section that match the same past item (CAG Structure Columns /
  Beams / Staircases) → only the best match keeps its value.

**2. Value.**

- **Whole lines:** each assembly's median **share of a past project** (same line type) × PO value.
- **Small orders** (PO < 20% of a typical project of that line type, e.g. revamps): median **item price** from
  history; where there is no match, the typical price for a long-lead bought-out, bought-out or manufactured item.
- Values are then scaled to the PO value, rounded his way (50k / 1 L / 5 L / 10 L steps), and the remainder is
  put on the largest item, so the **total always equals the PO value**.

**3. Category.**

1. Priority Index match (same line type, ≥85).
2. Otherwise inherited from the parent assembly.
3. Otherwise rules: bought-out and long-lead → 5; bought-out → 3/4; manufactured with gearbox / mandrel / roll /
   casting → 2; anything else → 1.

**4. Dispatch month / week.** The assembly's typical position in past projects' dispatch order (his plans first,
then history), mapped onto the window from [§4](#dispatch-window). Never earlier than PO + `DESIGN_DAYS` (10) +
lead time. If the delivery date has passed, the window shifts to start this month. The proposed week is the
month-end week; in the Add Plan table the user can pick any actual week (grouped by month), and the exports keep it.

**5. USD orders** (`usd_split`). When the currency is USD, the USD PO value is split evenly over the valued
assemblies (`billing_usd`), and INR = USD × rate for reference. The Add Plan export gains a "Billing Value (USD)"
column, and the sales-file rows fill BILLING VALUE ($).

**6. Lot-wise mode.** N equal lots across the window (for export and trading orders).

**7. Value sources, in order** (`REF_MODE = "after_labels"`):
1. **His own plans of the same line type** (match ≥85). Plans of another line type are never used for values;
   the line type comes from the sales file's project name.
2. **The line type's reference project** from `reference projects.xlsx` (`load_reference_projects`, match ≥85):
   that row's share of the reference project. Timing from same-line history where available. If the file is
   missing or locked, the same projects' rows are taken from the sales file (`BEST_REF`).
3. **Sales history** of the line type (median share), preferring rows of the reference project.
`backtest.py ref=first|value|off` and `nogroup` reproduce the rejected variants (see [§4](#reference-projects-billing-value-division)).

**8. Lots outside the LoE** (`typical_lots`). For whole-line orders, if at least half of the line type's past
projects book ELECTRICAL / HT MOTORS / SPARES separately (median ≥3% of a project), and no valued LoE row already
covers it, the engine adds that lot as a row in a "NOT IN LoE" section, valued at its typical share. Today this
fires for CRM (all three lots), not for CGL. The "LoE with plan" export writes these rows below the LoE.

**Timing constants:** `DESIGN_DAYS` = 10, `MIN_FIRST_DISPATCH_MONTHS` = 3, `DELIVERY_MONTHS` per line type.

### Accuracy (backtest, 18-Nov-2026)

| Check | 2611 CGL | 2618 CGL | 2621 CRM | 2630 CRM |
|---|---|---|---|---|
| Rows he valued that the engine also valued | 95% | 92% | 99% | 88% |
| Rows the engine valued that he also valued | 83% | 92% | 91% | 98% |
| Assembly values within ±30% of his | 40% | 42% | 77% | 77% |
| Assembly values within ±50% of his | 60% | 60% | 88% | 82% |
| Dispatch month within ±1 month | 96% | 83% | 99%* | 94%* |
| Project total equal to PO value | always | always | always | always |

\* CRM POs are not available. The backtest uses the sales-file project total as the PO value and his first and
last month as the window, so it tests valuation and ordering only. Without the "lots outside the LoE" rule, CRM
values within ±30% were only 17% / 69%.

## 6. The user's existing YD Ascend dashboard (file `dashboard`)

The user has already built a **"YD PROJECTS — Sales & Cash Flow"** page in YD Ascend (Bootstrap, Tailwind,
Chart.js). It has:

- KPIs: projects, assemblies, revenue, average;
- charts: forecasted billing allocation, billing by production line, monthly billing;
- an assembly search, quick filters (bought-out, in-house, dispatched, not dispatched) and dispatch tracking;
- **Project Category Configuration** — here "category" means **production line** (CGL, CRM…), *not* the 0–5
  planning category, so it is irrelevant to the engine;
- **Revenue Project Configuration**, which maps SAP WBS codes to projects.

**Backend endpoints** (on their `API_URL`):

- `/excel/forecast/dashboard`
- `/excel/forecast/category-mappings`
- `/excel/actuals/config`
- `/excel/actuals/sap` — **revenue comes from SAP; the SAP data is not flowing yet.**

**Decision (user, 2026-11-18).** The YD Ascend backend is large and **not relevant — ignore it**. Do not plan to
integrate with it or ask for it. The planner stays a **standalone tool** in this repo, styled like YD Ascend. Their
dashboard is the user's own; it only tells us the user already has forecast search and SAP revenue there.

## 7. How the user wants us to work (preferences & feedback)

- **UI must look like their website (YD Ascend).** They rejected the first Streamlit UI. The plan table must look
  like the YD Ascend "Add Plan" screen.
- **Accuracy over everything.** Show the basis for each number. Ask when unsure; they will ask Shivendu sir.
- **Numbers:** Indian format (₹2,29,38,000; lakh / crore).
- **Show the result directly.** When they ask for values for specific items, give the filled table.
- **Never lose context.** They work in Claude Code inside Google Antigravity and sessions may log out. Keep this
  file updated and committed.

## 8. Open questions for Shivendu sir (the planner)

He answered most of the list on 2026-11-18 (two rounds). His answers are recorded in §2 and §4. **Still open:**

1. **2616 pot roll system** (reworded): "In our draft plan for 2616, the Three Pot Roll System gets ₹1.31 Cr.
   That is simply what is left of the PO after the other items are priced — it isn't a figure from you. Does the
   2616 PO list a price per item? If not, what would you put on the pot roll system?"
2. **The ₹3 Cr buffer:** is it inside the PO total or on top of it, and how does he show it in the sales file?
3. **SPM best reference project** (he named CGL, CRM, pickling, CCL and trimming only).
4. **2621 / 2630 PO value and PO date / delivery date** — to backtest CRM dispatch timing properly.
5. He has said he will send reference plans for **CRM, CGL, SPM, CCL and trimming**. Waiting.
6. **2636 at 80%** of $1.5M × 95 — minor (INR is reference only).

## 9. Backlog / next steps (in priority order)

1. **Plan 2616 and 2619** and get them into the sales file.
2. **"Start from TS" option** on the New Plan page. Parse the division list into an editable LoE, then plan.
   Needs the source offers (Word/Excel) or Gemini approval for scanned and messy PDFs.
3. **More reference plans**: he will send CRM, CGL, SPM, CCL and trimming plans; drop them in as
   `<pno> planning.xlsx` and add each to `backtest.py` CASES. **SPM** has no reference project yet.
3b. **Small orders (revamps like 2616):** items with no history get a flat "typical BOI price" (₹13 L each on
   2616, pot roll only ₹63 L). Price long-lead bought-outs from similar past items instead (pot equipment median
   ₹1.37 Cr). Offered to the user 2026-11-19; awaiting a yes.
3a. **Needs attention:** show "drawing pending" (DRG. REQ. ticked) and "BOI ordering pending" (PENDING BOI ORDERING
   ticked) for assemblies due in the next 1–3 months.
4. **Graph API live sync** (Azure app needed). **SAP revenue** (their dashboard; data not flowing yet).
5. **Cash flow from payment terms:** billing × PO milestone terms → expected receipts per month.
6. **SAP actuals on the dashboard:** add billed-per-month (SAP) next to planned and dispatched, and flag projects
   where SAP and "Disp." disagree (see [§4](#sap-actual-billing)).
7. **Drawing release week.** The design team gives it (DESIGN CUTOFF DAYS is to be ignored). A prediction, if
   still wanted, could use the Plan Date / Issued Date columns in the planning LoEs.

## 10. Session log

- **2026-11-18 (Claude Code, Opus).**
  - Understood the domain.
  - Analysed the sales file, Priority Index, Time Sheet, his 2611/2618/2619 plans and the POs.
  - Built `planner/`: first in Streamlit, then rebuilt as FastAPI plus a YD Ascend-styled UI.
  - Added category inference and small-order pricing.
  - Made the hand-made 2616 plan.
  - Tested TS → LoE parsing.
  - Found that the GitHub repo is public.
  - Wrote this handover.
  - **User decisions:** the public repo is fine; "category" in their YD Ascend dashboard means production line, so
    ignore it; the YD Ascend backend is irrelevant, so ignore it and keep the planner standalone.
  - Compiled the full 30-question list for Shivendu sir (§8) for the user to relay.
  - **Shivendu sir's answers** recorded (§2, §4, §8): categories 3 vs 4, Cat 0 examples; billing = selling price;
    PO = total billing; 2611 gap = service bills; ₹3 Cr buffer; USD split evenly with bills in USD; dispatch week is an
    actual week; window set by YD; ~10 days design; FY target ₹1,000 Cr; PM owns RISK and status; sales file is the master.
  - Switched the planner to `SALES & CASH FLOW (1).xlsx` (slipped fixed by the user). Engine: any actual week can be
    picked (`weeks_of_month`), design time 10 days, USD even split (`usd_split`), "diagram/schematic" rows treated as
    documents. Dashboard: FY target. Backtest unchanged (2611/2618 same as before).
  - Analysed `revenue data sap.xlsx` (SAP invoices); added to the backlog as actual-vs-plan.
  - **Second round of answers** (§2, §4, §8): how slips are moved (customer month / expected receipt month /
    "hold" rolls to next month for follow-up); DRG. REQ. and PENDING BOI ORDERING meanings; ignore DESIGN CUTOFF DAYS
    and NATURE OF PAYMENT; best references CGL 2428, CRM 2414, pickling 2533, CCL 2520, trimming 2239; delivery
    periods; no dispatch before PO + 3 months. The user re-saved `SALES & CASH FLOW (1).xlsx` (a filter had been on).
  - Added his CRM plans `2621 planning.xlsx` and `2630 planning.xlsx`. Engine: best-reference preference;
    same-line-type plan matching (line type from the sales file); lots outside the LoE (CRM electrical / HT motors /
    spares); 3-month minimum; typical delivery periods; slipped windows start next month. Backtest now covers 2 CGL
    and 2 CRM plans (table in §5).
- **2026-11-19 (Claude Code, Opus).**
  - The user replaced `SALES & CASH FLOW.xlsx` with the fixed copy (deleted the "(1)" file) and committed the SAP export;
    `config.json` points back to `SALES & CASH FLOW.xlsx`.
  - Added `reference projects.xlsx` (his reference projects for billing value division). Engine: `load_reference_projects`,
    value order his plans → reference project → history; values never taken from his plans of another line type (bug
    found: a pickling plan was using CGL/CRM plan values). Server reloads when the file changes. Backtest unchanged
    (reference-first and part-row summing tested and rejected, §4).
