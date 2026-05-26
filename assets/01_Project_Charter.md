# Project Charter — Acme Industrial — SAP Ariba Rollout

**Document version:** 1.2
**Date:** 05 May 2026
**Status:** Approved by Steering Committee

---

## 1. Project Overview

Acme Industrial (hereafter "Acme") is implementing SAP Ariba to modernize its
procurement landscape. The Phase 1 rollout covers two modules:

- **Supplier Lifecycle and Performance (SLP)** — supplier registration,
  qualification, segmentation, and ongoing performance management.
- **Ariba Buying (with Guided Buying)** — requisitions, approvals, catalogs,
  and PO transmission.

Sourcing, Contracts, and Invoice Management are explicitly **out of scope** for
Phase 1 and will be evaluated for a future Phase 2.

The implementation partner is **Bluewave Consulting**.

## 2. Business Drivers

- Decentralized supplier onboarding across 14 plants leads to duplicate vendor
  records and inconsistent qualification standards.
- Maverick spend estimated at 18% of indirect spend; no enforceable channel
  for catalog buying.
- Audit findings (FY25) flagged weak segregation of duties in the current
  PR-to-PO process.

## 3. Scope

### In scope
- SLP for all direct and indirect suppliers globally (~6,200 active vendors).
- Guided Buying for indirect categories at the 4 North America sites in Phase 1a,
  followed by EMEA and APAC waves (out of this project's scope; tracked separately).
- Integration with SAP S/4HANA 2023 (on-premise) via SAP Cloud Integration Gateway (CIG).
- Single Sign-On via Okta.
- Vendor master migration from SAP S/4HANA into Ariba SLP.

### Out of scope
- Ariba Sourcing, Contracts, Invoice Management.
- Direct material catalogs.
- Supplier risk management module (SAP Ariba Supplier Risk).
- Non-SAP ERP integration (Acme has no other ERPs in scope).

## 4. Key Dates

| Milestone | Target Date |
|---|---|
| Project kickoff | 15 June 2026 |
| Design sign-off | 31 July 2026 |
| SIT exit | 02 October 2026 |
| UAT sign-off | 23 October 2026 |
| **Go-live** | 30 October 2026 |
| Hypercare exit / BAU transition | 27 November 2026 |

Total elapsed: ~24 weeks from kickoff to BAU transition.

## 5. Governance

- **Executive Sponsor:** Priya Raman, CPO
- **Business Sponsor:** Daniel O'Connell, VP Indirect Procurement
- **Program Manager (Acme):** Sara Lindqvist
- **Program Manager (Bluewave):** Marcus Chen
- **Steering Committee:** monthly; chaired by the Executive Sponsor.

## 6. Assumptions

- Ariba realms (test + prod) will be provisioned by SAP within 5 business
  days of the SOW signature.
- S/4HANA is on release 2023 FPS01 and is supported by current CIG version.
- Vendor master remains the system of record in S/4HANA; Ariba SLP is the
  system of engagement for suppliers.
- Okta is the single identity provider for all user populations.
- A single global Ariba realm will be used (no per-region realms).

## 7. Constraints

- Year-end financial close window (15 Dec — 7 Jan) is a hard no-deploy period.
- Acme IT change advisory board meets weekly on Tuesdays; cutover must align.
- Bluewave resources are committed; substitutions require Acme PM approval.

## 8. Budget

Approved budget: **USD 1.85M** (Bluewave fees + SAP subscription year 1 +
internal allocation). Detailed cost breakdown in the SOW.
