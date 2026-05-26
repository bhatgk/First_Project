# Business Requirements Document — Ariba SLP + Buying (Phase 1)

**Project:** Acme Industrial — SAP Ariba Phase 1
**Author:** Sara Lindqvist (Acme PM) & Marcus Chen (Bluewave PM)
**Version:** 0.9 (Draft for Explore-phase workshops)
**Date:** 12 May 2026

---

## 1. SLP — Supplier Lifecycle and Performance

### 1.1 Registration
- External registration via supplier-facing portal.
- Self-registration form fields: legal name, tax IDs (configurable per country),
  registered address, primary contact, banking, certifications, diversity flags.
- Country-specific question branches required for US, Germany, Brazil, India, UAE.
- Duplicate check against SAP vendor master and existing Ariba records.

### 1.2 Qualification
- Modular qualification process by commodity:
  - Indirect MRO — light questionnaire (anti-bribery, insurance, basic financial).
  - IT/Software — security questionnaire, data processing addendum.
  - Logistics — DOT compliance, fleet info, hazmat (US only).
  - Professional services — code of conduct, insurance, references.
- Approval flow: Procurement → Category Manager → Compliance → Finance.
- Three rejection reasons must be captured; suppliers can be re-invited.

### 1.3 Segmentation
- Auto-segment by annual spend tier and category criticality.
- Tiers: Strategic, Preferred, Approved, Transactional.

### 1.4 Performance
- KPI scorecards for Strategic and Preferred suppliers only in Phase 1.
- Quarterly review cycle; Procurement owns scorecard generation.

## 2. Ariba Buying — Guided Buying

### 2.1 Landing page
- Tiles by category: IT & Software, MRO, Office Supplies, Facilities,
  Travel-adjacent (vouchers etc), Professional Services, Marketing,
  Lab & Sample, Other.
- Policy tile prominent at top: "How to buy at Acme".

### 2.2 Forms
- Custom forms for: contingent labor, professional services SOW,
  software license requests, sample/lab purchases.

### 2.3 Catalogs
- CIF catalogs for top 8 MRO suppliers (priority loading).
- PunchOut for Staples, CDW, Dell, Grainger.
- BMEcat: out of scope for Phase 1.

### 2.4 Approvals
- Spend authorization matrix imported from SAP HR org structure.
- Thresholds:
  - < $2,500: requester self-approves
  - $2,500 – $25,000: cost center manager
  - $25,000 – $100,000: VP
  - $100,000 – $500,000: SVP + Finance Controller
  - > $500,000: CFO + CEO
- Capex flag triggers parallel Capex committee approval irrespective of amount.
- Substitute approvers must be configurable per user (vacation coverage).

### 2.5 PO transmission
- cXML to suppliers with PunchOut.
- Email PDF for all others.
- Order confirmation tracking enabled but not enforced in Phase 1.

## 3. Integration

- SAP CIG to S/4HANA 2023 FPS01.
- Master data: vendor master replication SAP → Ariba (one-way for Phase 1).
- Transactional: PR, PO, GR, invoice header data flows. Invoice processing
  itself remains in S/4HANA (Ariba Invoice Management out of scope).
- Mid-day delta sync for vendor master; real-time for transactions.

## 4. Security

- SSO via Okta; SAML 2.0.
- User provisioning: automated via SCIM from Okta to Ariba.
- Role-based access: Requester, Approver, Buyer, Catalog Manager,
  Category Manager, Compliance, Finance, Admin.

## 5. Data Migration

- ~6,200 active vendors in scope.
- Cleansing required: estimated 12-15% duplicate rate, 8% incomplete records.
- Cleansing approach: 6-week sprint in Explore + early Realize.
- Cutover: one-time full load 48 hours before go-live; deltas via integration thereafter.

## 6. Training

- Estimated user populations:
  - Requesters: ~3,400 (NA only in Phase 1)
  - Approvers: ~620
  - Buyers: ~85
  - Category Managers: ~22
  - Suppliers: ~6,200 (registration training only)
- Format: short e-learning + role-based virtual classroom.
- Quick reference guides per role.

## 7. Open Items

- O-1: Confirm whether Brazil tax registration questions need NF-e fields.
- O-2: Approval matrix for Capex committee — Finance to confirm members.
- O-3: PunchOut go-live sequencing — which supplier first?
- O-4: Hypercare staffing model (Bluewave + Acme split TBD).
