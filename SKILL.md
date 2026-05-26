---
name: ariba-project-plan
description: Generate a fully pre-populated SAP Ariba implementation project plan in Excel covering Supplier Lifecycle and Performance (SLP) and Ariba Buying. Use this skill whenever the user asks for an Ariba project plan, SLP/Buying rollout plan, SAP Ariba implementation schedule, WBS, Gantt, RACI, or RAID log for Ariba — including casual phrasings like "I need an Ariba plan in Excel", "create the SLP rollout plan", "draft the Ariba Buying implementation timeline", or "give me a project plan for our Ariba project". Trigger even if the user does not explicitly say "Excel" — the deliverable is always an .xlsx workbook. Do NOT use this skill for Ariba Sourcing, Contracts, or full Source-to-Pay scopes; this skill is scoped to SLP + Buying only.
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

## Required inputs from the user

Before generating, confirm these four items. If the user hasn't provided them, ask once in a single message:

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
2. Confirm the four required inputs with the user (one combined message, not four).
3. Compute the phase date ranges from the start date (or backward from the go-live date).
4. Build the workbook with `openpyxl`:
   - Create the five sheets in the order above.
   - Add a title row on each sheet: "SAP Ariba SLP + Buying Implementation Plan — {Client Name}" in bold, merged across the sheet header.
   - Populate the WBS task list, computing Start/End per phase. Distribute task durations evenly within each phase unless a task obviously needs a specific slot (e.g., kickoff = Day 1, go-live = end of Deploy).
   - Generate the Gantt with weekly columns and conditional-format bars.
   - Populate the RACI exactly as specified.
   - Populate the RAID log with the standard items, leaving issue rows blank.
   - Populate the milestones with computed dates.
5. Apply formatting: Arial font throughout, bold headers, frozen panes on each sheet, column widths sized so content is readable (typically 12–30 depending on column).
6. Save to `/mnt/user-data/outputs/Ariba_SLP_Buying_Project_Plan_{Client}.xlsx`.
7. Run `python /mnt/skills/public/xlsx/scripts/recalc.py <path>` and verify the JSON returns zero errors. Fix any errors and re-run.
8. Call `present_files` with the workbook path. Keep the message short — one or two sentences summarizing what's in the workbook.

## Things to be careful about

- **Do not invent additional Ariba modules.** If the user asks for Sourcing, Contracts, or Invoicing tasks, tell them this skill is scoped to SLP + Buying and offer to add a small custom section rather than silently expanding scope.
- **Dates are working-day calendar dates, not business-day adjusted.** If the user needs working-day logic, mention that as a follow-up enhancement.
- **Owner names vs. roles.** The pre-populated owners are *roles* (e.g., "SLP Functional Lead"). If the user provides actual names, do a global replace before saving.
- **Status defaults.** Everything starts as "Not Started" with 0% complete. Don't pre-mark anything as in progress.
- **Don't dump the full task list in chat.** The workbook is the deliverable; the chat message after `present_files` should be a brief summary, not a re-listing of every task.