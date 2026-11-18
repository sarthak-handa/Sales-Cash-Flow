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

The 0–5 category is mostly *in his head*. Partly written down in `Priority Index for Lines (1) (8).xlsx`.

| Cat | Meaning |
|---|---|
| 0 | Not a full assembly. Items required by the long-lead Category 2 assemblies (e.g. back-up roll assemblies in SPM). |
| 1 | Manufactured, short lead (typically 2–3 months) |
| 2 | Manufactured, long lead (uncoilers, recoilers, bridle rolls — gearbox / mandrel / rolls / castings) |
| 3 / 4 | Bought out. Ordered and dispatched straight to the client's site, or received at YD and re-dispatched. **The difference between 3 and 4 is still unconfirmed.** |
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
assembly that misses its month moves to the next month.

### The sales file

`SALES & CASH FLOW.xlsx` is the master forecast. It has these sheets:

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

> **CONFIDENTIAL.** These files hold client POs with GSTINs, phone numbers and e-mails; YD's sales forecast; and
> tech specs marked "exclusive property of Yogiji Digi".
>
> - **The GitHub repo must be private.** On 2026-11-18 we found `github.com/sarthak-handa/Sales-Cash-Flow` was
>   **PUBLIC** and told the user to make it private.
> - **Never push data to a public remote.**
> - **Never send these files to outside services** without the user's explicit OK. Gemini needs approval.

| File | What it is |
|---|---|
| `SALES & CASH FLOW.xlsx` | Master forecast (see [§2](#the-sales-file)). The planner's knowledge base. |
| `Priority Index for Lines (1) (8).xlsx` | One sheet per line type (Skin Pass Mill, Trimming, CRM, CCL, Pickling, CGL), each a reference project's LoE with **Priority (= category 0–5), Lead Time ("2-3M"), Remarks, BOI/Trading**. Reference projects: CGL 2214, CRM 2205, CCL 2217, Pickling 2302, Trimming 2401, SPM 2218. |
| `Time Sheet (1) (6).xlsx` | **V2:** engineering / fabrication / assembly days per sub-assembly, 481 rows across 8 line types. **V1:** a machine-hours template with only 4 rows. |
| `2611 planning.xlsx`, `2618 planning.xlsx` | **His completed plans.** LoE plus `Dispatch Week` ("WW49 2026") and `Billing Value (INR)` stored as text ("20,00,000"). Category column empty. These are the *reference plans* the engine learns from. |
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

His billing values for a project add up to the **PO value excluding GST**:

- **2618:** ₹79.80 Cr, which is 100% of the PO.
- **2611:** ₹57.97 Cr against a ₹59.71 Cr PO, which is 97%. The gap may be the erection & commissioning portion,
  or a later PO or TS revision; the user isn't sure.

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

He labels a dispatch month by the **ISO week number of (last day of the month + 1 day)**. This reproduced all 9 of
his labels: Sep-26 → WW40 2026, Dec-26 → WW53 2026, Jan-27 → WW5 2027, May-27 → WW22 2027. In effect his plan
is monthly.

### Dispatch window

His plans run from **PO date + 0.75 × the delivery period** to **PO delivery date + 2 months**. Long-lead
bought-out items come last.

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

Converted at the **current rate, rounded**. He used **₹95** when the rate was moving between 93 and 96.

In the sales file, 2636 is booked as **3 equal lots of ₹3.8 Cr (Jan–Mar 2027) = ₹11.4 Cr**. That is exactly
**80%** of $1.5M × 95, and the reason is unknown. 2635 is booked at ₹9.5 Cr. Export and trading orders are
booked lot-wise.

### Gaps found

- **2616 and 2619 are not in the sales file at all.** That is ₹6 Cr missing from the forecast, and both PO
  delivery dates have passed.
- **They have no plan in YDAscend either.**
- **₹63 Cr is slipped:** 324 assemblies past their month and not dispatched (as of 18-Nov-2026).

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

- FY KPI tiles: plan, dispatched, current month, next 3 months, slipped, active projects.
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

**4. Dispatch month.** The assembly's typical position in past projects' dispatch order (his plans first, then
history), mapped onto the window from [§4](#dispatch-window). Never earlier than PO + 1 month (design) + lead time.
If the delivery date has passed, the window shifts to start this month.

**5. Lot-wise mode.** N equal lots across the window (for export and trading orders).

### Accuracy (backtest, 18-Nov-2026)

| Check | 2611 | 2618 |
|---|---|---|
| Rows he valued that the engine also valued | 95% | 92% |
| Rows the engine valued that he also valued | 83% | 92% |
| Assembly values within ±50% of his | 60% | 60% |
| Dispatch month within ±1 month | 96% | 83% |
| Project total equal to PO value | always | always |

## 6. The user's existing YD Ascend dashboard (file `dashboard`)

The user has already built a **"YD PROJECTS — Sales & Cash Flow"** page in YD Ascend (Bootstrap, Tailwind,
Chart.js). It has:

- KPIs: projects, assemblies, revenue, average;
- charts: forecasted billing allocation, billing by production line, monthly billing;
- an assembly search, quick filters (bought-out, in-house, dispatched, not dispatched) and dispatch tracking;
- **Project Category Configuration**;
- **Revenue Project Configuration**, which maps SAP WBS codes to projects.

**Backend endpoints** (on their `API_URL`):

- `/excel/forecast/dashboard`
- `/excel/forecast/category-mappings`
- `/excel/actuals/config`
- `/excel/actuals/sap` — **revenue comes from SAP; the SAP data is not flowing yet.**

**Implication.** Our dashboard duplicates part of theirs. The planner engine (LoE → week / value / category)
should end up **inside YD Ascend**, e.g. as an "Add Plan → auto-fill" button calling our `/api/plan`. Their
category-mappings endpoint may hold category knowledge we should reuse. We have **not** seen their backend code or
`config.js` (`API_URL`) yet; ask the user.

## 7. How the user wants us to work (preferences & feedback)

- **UI must look like their website (YD Ascend).** They rejected the first Streamlit UI. The plan table must look
  like the YD Ascend "Add Plan" screen.
- **Accuracy over everything.** Show the basis for each number. Ask when unsure; they will ask Shivendu sir.
- **Numbers:** Indian format (₹2,29,38,000; lakh / crore).
- **Show the result directly.** When they ask for values for specific items, give the filled table.
- **Never lose context.** They work in Claude Code inside Google Antigravity and sessions may log out. Keep this
  file updated and committed.

## 8. Open questions for Shivendu sir (the planner)

Asked through the user on 2026-11-18; **no answers yet**.

1. How does he set a bought-out item's billing value? Vendor price plus a margin — and what margin?
2. What is the difference between Category 3 and 4? Can he give an example of Category 0?
3. Is there a monthly revenue target or cap? What are the targets?
4. When a PO delivery date has passed, how does he choose the new dispatch month?
5. Why is 2636 booked at ₹11.4 Cr, which is 80% of $1.5M × 95?
6. For 2611, is the ₹1.7 Cr between his billing total and the PO the erection & commissioning portion?
7. (From the earlier list) the monthly routine: who sends status, and who receives the plan? What do "DESIGN
   CUTOFF DAYS" and "NATURE OF PAYMENT" mean? How far did the Realization rollout get?

## 9. Backlog / next steps (in priority order)

1. **Security.** Confirm the GitHub repo is private before any further push.
2. **Plan 2616 and 2619** and get them into the sales file.
3. **"Start from TS" option** on the New Plan page. Parse the division list into an editable LoE, then plan.
   Needs the source offers (Word/Excel) or Gemini approval for scanned and messy PDFs.
4. **More reference plans** for CRM, pickling, SPM, CCL and trimming (his completed plans plus POs) to improve the
   engine beyond CGL.
5. **Integrate with YD Ascend** (their backend and `API_URL`) and their category mappings.
6. **Graph API live sync** (Azure app needed). **SAP revenue** (their dashboard; data not flowing yet).
7. **Cash flow from payment terms:** billing × PO milestone terms → expected receipts per month.
8. **Drawing release week prediction.** Learn from the Plan Date / Issued Date columns in the planning LoEs and
   from the Time Sheet V2 engineering days.

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
