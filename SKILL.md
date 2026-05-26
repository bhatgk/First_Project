---
name: ariba-project-plan
description: Generate a fully pre-populated SAP Ariba implementation project plan in Excel covering Supplier Lifecycle and Performance (SLP) and Ariba Buying. Use this skill whenever the user asks for an Ariba project plan, SLP/Buying rollout plan, SAP Ariba implementation schedule, WBS, Gantt, RACI, or RAID log for Ariba — including casual phrasings like "I need an Ariba plan in Excel", "create the SLP rollout plan", "draft the Ariba Buying implementation timeline", or "give me a project plan for our Ariba project". Also trigger when the user uploads a folder of Ariba project source documents (project charter, BRD, SOW, kickoff MoM, stakeholder list, vendor master extract, etc.) and asks Claude to "build the plan", "consolidate these", or "turn this into a project plan" — the skill knows how to read these sources, reconcile dates and scope, and produce the workbook. Trigger even if the user does not explicitly say "Excel" — the deliverable is always an .xlsx workbook. Do NOT use this skill for Ariba Sourcing, Contracts, or full Source-to-Pay scopes; this skill is scoped to SLP + Buying only.
---

# SAP Ariba (SLP + Buying) Project Plan Generator

## What this skill produces

A single `.xlsx` workbook with **five sheets**, fully pre-populated with standard SAP Ariba SLP + Buying implementation content. Claude tailors the dates, names, and any user-specific details on top of the defaults — the user should never see an empty template.

| # | Sheet name | Purpose |
|---|---|---|
| 1 | `WBS` | Phased work breakdown (Prepare → Explore → Realize → Deploy → Run) with tasks, durations, owners, dates, status |
| 2 | `Gantt` | Visual timeline with bars across weekly columns |
| 3 | `RACI` | Workstreams × roles responsibility matrix |
| 4 | `RAID Log` | Risks, Assumptions, Issues, Dependencies |
| 5 | `Milestones` | Major milestones with target dates and status |

The workbook stacks on top of the `xlsx` skill — read `/mnt/skills/public/xlsx/SKILL.md` first for fonts, formula handling, recalculation, and verification rules. The content defaults below are specific to Ariba SLP + Buying.

## Reading source documents (if provided)

If the user uploads files or points to a folder (e.g. `/mnt/user-data/uploads/`), **always read the source documents before asking any questions**. Most of the "required inputs" below — and a lot of the project-specific content — can be lifted directly from these documents, and asking the user for things that are sitting in their own files is annoying.

### Step 1 — Inventory the folder

List every file the user has provided. Don't skip files because the name looks unfamiliar; Ariba project folders typically mix several document types:

| Likely source | Typical filename hints | What to extract |
|---|---|---|
| Project Charter | "charter", "PID", "project initiation" | Client name, partner, sponsor, scope confirmation, dates, budget, assumptions, constraints |
| BRD / SOW / Requirements | "BRD", "SOW", "requirements", "FRS" | Module scope confirmation (SLP + Buying), integration details, approval matrix, catalog strategy, training population sizes, open items |
| Stakeholder list / RACI input | "stakeholder", "team", "org", "RACI" | Real names for the Owner column and the RACI sheet |
| Kickoff MoM / meeting minutes | "MoM", "minutes", "kickoff", "SteerCo" | Confirmed decisions, action items (→ RAID), risks raised in meetings |
| Vendor master / data extract | "vendor", "supplier", "master", "extract" | Data quality issues (duplicates, missing tax IDs, dormant records) → feeds Data workstream tasks and RAID risks |
| Integration / architecture docs | "integration", "architecture", "CIG", "API" | SAP version, integration pattern, SSO/IdP — feeds Technical workstream and assumptions |
| Cutover / training plans | "cutover", "training", "OCM", "change" | Wave plans, training population, hypercare staffing |

Read each one with the appropriate tool: markdown/text/CSV files can be `view`'d directly; `.docx`/`.pdf` files need the `docx` / `pdf` skills (consult their SKILL.md files first). For CSVs, use pandas if the file is large.

### Step 2 — Extract and reconcile

Build a small in-memory summary covering:

- **Client name** (from Charter or BRD)
- **Implementation partner** (from Charter or SOW)
- **Project start date** (from Charter; cross-check against Kickoff MoM)
- **Go-live target date** (from Charter; cross-check against MoM)
- **Sponsor + Program Managers + key role holders** (from Charter + Stakeholder list)
- **Scope confirmation** — verify the project really is SLP + Buying only. If the sources mention Sourcing / Contracts / Invoicing as in-scope, **stop and flag this to the user** before generating; this skill does not cover those modules.
- **Project-specific risks/assumptions/dependencies** to add on top of the defaults (e.g., a kickoff MoM that flags CIG-version risk, a vendor master with high duplicate rate, an open item about Brazil NF-e fields).
- **Project-specific milestones or constraints** (e.g., year-end no-deploy windows, parallel programs, regulatory deadlines).

**Reconcile contradictions explicitly.** If the Charter says a 24-week timeline but the kickoff date to go-live is only 19 weeks, that's a contradiction the user needs to know about — don't silently pick one. Default behavior:

- If start date and go-live date are both given but the gap is shorter than the default 24 weeks, compress the phases proportionally rather than the defaults. Tell the user in the final summary message which phases were compressed and by how much.
- If the gap is longer than 24 weeks, extend Explore and Realize proportionally (not Prepare/Deploy/Run, which are more fixed).
- If sources disagree on any factual item (e.g., partner name appears differently in two files), prefer the **Project Charter** as the source of truth, and mention the discrepancy in the summary.

### Step 3 — Only ask the user for what's actually missing

After reading the sources, ask the user *only* for items you couldn't find. If the Charter contains every required input, skip the question step entirely and proceed to generation, but **tell the user what you extracted and from which file** before generating, so they can correct any misreads. Example:

> "I read the Charter, BRD, Stakeholder list, Kickoff MoM, and Vendor master extract. Here's what I'll use: Client = Acme Industrial, Partner = Bluewave Consulting, Kickoff = 15 Jun 2026, Go-live = 30 Oct 2026 (≈19 weeks — I'll compress Explore and Realize by ~1 week each to fit). I'll also fold in 3 project-specific risks I found in the MoM. Proceed?"

### Step 4 — Apply source-derived content to the right sheets

| Sheet | What to pull from sources |
|---|---|
| WBS | Real owner names from the Stakeholder list (replace the default role labels); add a few project-specific tasks if the BRD calls out unusual scope items (e.g., a Brazil tax-fields branch in SLP registration) |
| Gantt | Adjusts automatically once WBS dates are set |
| RACI | Replace role labels with real names where the Stakeholder list gives a single owner per role; if multiple people hold a role (e.g., two Category Managers), add them as additional columns rather than merging |
| RAID Log | Add project-specific risks/assumptions/dependencies on top of the defaults. Tag the source in the Notes column (e.g., "Source: Kickoff MoM, item AI-9") |
| Milestones | Override default milestone dates with the Charter's dates if provided |

If no source documents are provided, fall back to the question-based flow below.

## Required inputs from the user (when no sources are provided)

If the user has not uploaded source documents, confirm these four items. Ask once in a single message:

1. **Project start date** (defaults to first Monday after today if not given)
2. **Client / company name** (used in the title row; defaults to "Client")
3. **Implementation partner name** (defaults to "Implementation Partner")
4. **Any specific go-live target date** (otherwise computed from the phase durations below)

Don't ask for the task list, durations, or RACI assignments — those are pre-populated from this skill. Confirming scope beyond SLP + Buying is also unnecessary; this skill is intentionally scoped to those two modules.

## Phase durations (default)

Total elapsed: **24 weeks** from kickoff to hypercare exit.

| Phase | Duration | Notes |
|---|---|---|
| Prepare | 2 weeks | Kickoff, governance, mobilization |
| Explore | 5 weeks | Design workshops, fit-gap, integration design |
| Realize | 9 weeks | Configuration, integration build, test cycles |
| Deploy | 4 weeks | UAT, training, cutover, go-live |
| Run | 4 weeks | Hypercare, transition to support |

If the user gives a go-live date, work backward: go-live = end of Deploy. Run begins the day after go-live.

## Sheet 1: WBS — standard task list

Columns: `Task ID | Phase | Workstream | Task | Owner Role | Start | End | Duration (days) | % Complete | Status | Notes`

Use these tasks (pre-populated; Claude fills owners and dates):

### Prepare (Weeks 1–2)
- P-01 | PMO | Project kickoff meeting | Program Manager
- P-02 | PMO | Governance model & steering committee setup | Program Manager
- P-03 | PMO | Project charter sign-off | Program Manager
- P-04 | PMO | Communication plan | PMO Lead
- P-05 | Technical | Ariba realm provisioning (test + prod) | Ariba Admin
- P-06 | Technical | SAP integration prerequisites checklist | Integration Lead
- P-07 | Change | Stakeholder map & change impact baseline | Change Lead

### Explore (Weeks 3–7)
- E-01 | SLP | Supplier data discovery & cleansing strategy | SLP Functional Lead
- E-02 | SLP | Supplier registration form design workshop | SLP Functional Lead
- E-03 | SLP | Supplier qualification questionnaire design | SLP Functional Lead
- E-04 | SLP | Supplier segmentation & approval workflow design | SLP Functional Lead
- E-05 | Buying | Procurement policy review & catalog strategy | Buying Functional Lead
- E-06 | Buying | Guided Buying landing page & tile design | Buying Functional Lead
- E-07 | Buying | Approval rules & spend authorization matrix design | Buying Functional Lead
- E-08 | Buying | Catalog enablement plan (CIF, PunchOut, BMEcat) | Buying Functional Lead
- E-09 | Technical | Integration design (SLP master data, PR/PO, GR, invoice) | Integration Lead
- E-10 | Technical | SSO & user provisioning design | Security Lead
- E-11 | Data | Vendor master cleansing & migration approach | Data Lead
- E-12 | Change | Training needs analysis | Change Lead

### Realize (Weeks 8–16)
- R-01 | SLP | Configure registration & qualification forms | SLP Functional Lead
- R-02 | SLP | Configure supplier segmentation & approval flows | SLP Functional Lead
- R-03 | SLP | Configure supplier 360 view & modular questionnaires | SLP Functional Lead
- R-04 | Buying | Configure Guided Buying tiles, forms, and policies | Buying Functional Lead
- R-05 | Buying | Configure approval flows & substitute approvers | Buying Functional Lead
- R-06 | Buying | Configure commodity codes & catalog views | Buying Functional Lead
- R-07 | Buying | Onboard pilot catalogs (CIF + PunchOut) | Buying Functional Lead
- R-08 | Technical | Build SAP ↔ Ariba integration (CIG / API) | Integration Lead
- R-09 | Technical | Configure SSO & user roles | Security Lead
- R-10 | Data | Execute vendor master migration to dev | Data Lead
- R-11 | Testing | SIT cycle 1 | Test Lead
- R-12 | Testing | SIT cycle 2 (defect fix retest) | Test Lead
- R-13 | Change | Develop training materials & quick reference guides | Change Lead

### Deploy (Weeks 17–20)
- D-01 | Testing | UAT preparation & test script walkthrough | Test Lead
- D-02 | Testing | UAT execution | Business UAT Lead
- D-03 | Testing | UAT sign-off | Business Sponsor
- D-04 | Change | End-user training delivery (buyers, requesters, approvers) | Change Lead
- D-05 | Change | Supplier onboarding communications | Change Lead
- D-06 | Data | Production vendor master migration | Data Lead
- D-07 | Technical | Production cutover (config transport, integration switchover) | Integration Lead
- D-08 | PMO | Go-live readiness review | Program Manager
- D-09 | PMO | Go-live | Program Manager

### Run (Weeks 21–24)
- N-01 | Support | Hypercare ticket triage (Tier 1/2/3) | Support Lead
- N-02 | Support | Daily hypercare standup | Support Lead
- N-03 | Change | Adoption metrics tracking | Change Lead
- N-04 | PMO | Lessons learned workshop | Program Manager
- N-05 | PMO | Transition to BAU support & project closure | Program Manager

**Formula notes:**
- `End = Start + Duration - 1` (Excel: `=B2+D2-1` adjusted for the column layout)
- `% Complete` defaults to 0; `Status` defaults to "Not Started"
- Apply conditional formatting on Status: green (Complete), yellow (In Progress), red (Blocked), grey (Not Started)

## Sheet 2: Gantt

Build a Gantt visualization on the same task list. Layout:
- Column A: Task ID
- Column B: Task name
- Column C: Start date
- Column D: End date
- Columns E onward: one column per **week** for 24 weeks, headers as week-start dates (e.g., "W1 — 06-Jan", "W2 — 13-Jan", …)
- For each task row, fill the cells whose week falls within `[Start, End]` with a solid colored fill keyed to the phase:
  - Prepare: light blue
  - Explore: light green
  - Realize: light orange
  - Deploy: light purple
  - Run: light grey
- Use a conditional-format formula like `=AND(E$1>=$C2, E$1<=$D2)` so the bars update if the WBS dates change. Apply per-phase by filtering on the Phase column or by using separate conditional rules per phase color.

Freeze panes at row 2 and column E so the task list stays visible when scrolling the timeline.

## Sheet 3: RACI

Rows: workstreams. Columns: roles. Cells: R / A / C / I.

Standard roles (columns):
`Business Sponsor | Program Manager | PMO Lead | SLP Functional Lead | Buying Functional Lead | Integration Lead | Data Lead | Security Lead | Test Lead | Change Lead | Support Lead | Ariba Admin | Business UAT Lead`

Standard workstreams (rows) and RACI assignments:

| Workstream | Sponsor | PM | PMO | SLP Lead | Buying Lead | Integration | Data | Security | Test | Change | Support | Admin | UAT |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Project governance | A | R | C | I | I | I | I | I | I | I | I | I | I |
| SLP — registration & qualification design | I | A | I | R | I | C | C | I | I | C | I | C | I |
| SLP — supplier segmentation & workflows | I | A | I | R | I | C | I | I | I | I | I | C | I |
| Buying — Guided Buying & policies | I | A | I | I | R | C | I | I | I | C | I | C | I |
| Buying — approvals & catalogs | I | A | I | I | R | C | I | I | I | I | I | C | I |
| Integration (SAP ↔ Ariba) | I | A | I | C | C | R | C | C | I | I | I | C | I |
| Vendor master & data migration | I | A | I | C | C | C | R | I | I | I | I | I | I |
| Security, SSO, user provisioning | I | A | I | I | I | C | I | R | I | I | I | C | I |
| Testing (SIT) | I | A | I | C | C | C | I | I | R | I | I | I | I |
| UAT | A | C | I | C | C | I | I | I | C | C | I | I | R |
| Training & change management | I | A | I | C | C | I | I | I | I | R | I | I | I |
| Cutover & go-live | A | R | C | C | C | C | C | C | C | C | C | C | C |
| Hypercare | I | A | I | C | C | C | I | I | I | C | R | C | I |

Apply a legend at the top: R = Responsible, A = Accountable, C = Consulted, I = Informed. Color R/A cells distinctly (R = green fill, A = blue fill).

## Sheet 4: RAID Log

Columns: `ID | Type | Title | Description | Likelihood (H/M/L) | Impact (H/M/L) | Owner | Due Date | Status | Mitigation / Notes`

Pre-populate with these Ariba-specific RAID items:

**Risks**
- R01 | Risk | Vendor master data quality | Poor data quality in source SAP delays SLP migration | H | H | Data Lead | (set during Explore) | Open | Run data profiling in Explore; allocate cleansing sprints
- R02 | Risk | Supplier adoption of registration | Suppliers slow to complete registration; impacts SLP go-live | M | H | Change Lead | (set during Explore) | Open | Supplier comms wave plan; concierge desk for top 50 suppliers
- R03 | Risk | CIG integration capacity | Cloud Integration Gateway throughput constraints during peak | M | M | Integration Lead | (set during Realize) | Open | Performance test integration early; SAP support ticket ready
- R04 | Risk | Catalog content readiness | Suppliers can't deliver CIF/PunchOut catalogs on time | M | M | Buying Functional Lead | (set during Realize) | Open | Stagger catalog onboarding; identify priority categories
- R05 | Risk | Approver hierarchy drift | SAP HR org changes invalidate approval matrix | L | H | Buying Functional Lead | (set during Deploy) | Open | Lock org snapshot for cutover; daily sync pre-go-live

**Assumptions**
- A01 | Assumption | Ariba realm provisioning by SAP within 5 business days of request | — | — | Ariba Admin | (Week 1) | Open | —
- A02 | Assumption | SAP backend (ECC/S4) is on a supported release for CIG | — | — | Integration Lead | (Week 2) | Open | —
- A03 | Assumption | Vendor master will be source-of-truth in SAP, not Ariba | — | — | Data Lead | (Week 3) | Open | —
- A04 | Assumption | Single SSO IdP for all user populations | — | — | Security Lead | (Week 3) | Open | —

**Issues** (placeholder — Claude leaves blank rows for the team to fill)
- I01 | Issue | (open) | | | | | | Open | —

**Dependencies**
- D01 | Dependency | SAP basis team availability for CIG setup | Blocks integration build | — | — | Integration Lead | (Week 6) | Open | Confirm resource plan in Explore
- D02 | Dependency | Procurement policy sign-off | Blocks Guided Buying design freeze | — | — | Business Sponsor | (Week 5) | Open | —
- D03 | Dependency | Master data governance decisions | Blocks vendor migration scope | — | — | Data Lead | (Week 4) | Open | —

## Sheet 5: Milestones

Columns: `Milestone | Target Date | Phase | Owner | Status`

Pre-populate:
1. Project kickoff — end of Week 1 — Prepare — Program Manager — Not Started
2. Design sign-off (SLP + Buying) — end of Week 7 — Explore — Business Sponsor — Not Started
3. Integration build complete — end of Week 13 — Realize — Integration Lead — Not Started
4. SIT exit — end of Week 16 — Realize — Test Lead — Not Started
5. UAT sign-off — end of Week 19 — Deploy — Business Sponsor — Not Started
6. **Go-live** — end of Week 20 — Deploy — Program Manager — Not Started
7. Hypercare exit / BAU transition — end of Week 24 — Run — Support Lead — Not Started

Bold the Go-live row. Apply conditional formatting on Status column (same scheme as the WBS sheet).

## Generation workflow

1. **Read `/mnt/skills/public/xlsx/SKILL.md`** for the formula/recalc rules — this skill inherits all of them (Arial font, formulas not hardcoded values, run `scripts/recalc.py` after save, verify zero formula errors).
2. **Check for source documents.** Look at `/mnt/user-data/uploads/` and anywhere else the user has pointed to. If sources exist, follow the "Reading source documents" section above to inventory, extract, and reconcile before doing anything else. Show the user what you extracted and from which file, and ask them to confirm before generating.
3. **If no sources are provided**, confirm the four required inputs with the user (one combined message, not four).
4. Compute the phase date ranges from the start date (or backward from the go-live date). If the gap between start and go-live differs from the default 24 weeks, compress or extend per the rules in "Reading source documents — Step 2".
5. Build the workbook with `openpyxl`:
   - Create the five sheets in the order above.
   - Add a title row on each sheet: "SAP Ariba SLP + Buying Implementation Plan — {Client Name}" in bold, merged across the sheet header.
   - Populate the WBS task list, computing Start/End per phase. Distribute task durations evenly within each phase unless a task obviously needs a specific slot (e.g., kickoff = Day 1, go-live = end of Deploy). Use real owner names from the Stakeholder source if available; otherwise use role labels.
   - Generate the Gantt with weekly columns and conditional-format bars.
   - Populate the RACI. If real names are available per role, replace role labels with names; if multiple people hold one role, add extra columns rather than merging.
   - Populate the RAID log with the standard items **plus** any project-specific risks/assumptions/dependencies extracted from the source documents. Tag the source in the Notes column.
   - Populate the milestones with computed dates, overriding defaults with Charter dates where given.
6. Apply formatting: Arial font throughout, bold headers, frozen panes on each sheet, column widths sized so content is readable (typically 12–30 depending on column).
7. Save to `/mnt/user-data/outputs/Ariba_SLP_Buying_Project_Plan_{Client}.xlsx`.
8. Run `python /mnt/skills/public/xlsx/scripts/recalc.py <path>` and verify the JSON returns zero errors. Fix any errors and re-run.
9. Call `present_files` with the workbook path. Keep the message short — one or two sentences summarizing what's in the workbook, and note any reconciliations or compressions you applied so the user can spot misreads.

## Things to be careful about

- **Do not invent additional Ariba modules.** If the user asks for Sourcing, Contracts, or Invoicing tasks, tell them this skill is scoped to SLP + Buying and offer to add a small custom section rather than silently expanding scope.
- **Dates are working-day calendar dates, not business-day adjusted.** If the user needs working-day logic, mention that as a follow-up enhancement.
- **Owner names vs. roles.** The pre-populated owners are *roles* (e.g., "SLP Functional Lead"). If the user provides actual names, do a global replace before saving.
- **Status defaults.** Everything starts as "Not Started" with 0% complete. Don't pre-mark anything as in progress.
- **Don't dump the full task list in chat.** The workbook is the deliverable; the chat message after `present_files` should be a brief summary, not a re-listing of every task.