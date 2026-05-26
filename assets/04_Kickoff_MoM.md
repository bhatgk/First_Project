# Minutes of Meeting — Project Kickoff

**Meeting:** Acme × Bluewave — SAP Ariba Phase 1 Kickoff
**Date:** Monday, 15 June 2026, 09:00–11:30 GST (virtual + Dubai HQ)
**Chair:** Priya Raman (Executive Sponsor)
**Minutes by:** Helena Ruiz (PMO Lead)
**Distribution:** Project team + Steering Committee

---

## Attendees

Acme: Priya Raman, Daniel O'Connell, Sara Lindqvist, Helena Ruiz, Rina Takahashi,
Tom Becker, Olivia Park, Ravi Subramanian, Karen Faulkner, Hector Alvarez,
Lena Brandt, Felix Nguyen.

Bluewave: Marcus Chen, Anika Patel, James Whitaker, Yusuf Khalil, Maya Greenfield,
David Mensah.

Apologies: none.

## 1. Welcome & objectives (Priya)

Priya framed the strategic intent: Phase 1 is the foundation for a multi-year
S2P transformation. Success criteria for Phase 1:
- Go-live by end of October 2026 with NA Guided Buying live.
- Vendor master in Ariba SLP cleansed and authoritative for engagement.
- Audit findings on PR-to-PO segregation closed before year-end.

## 2. Confirmed dates

The dates in the Charter were confirmed with no changes:
- Kickoff: today (15 Jun 2026).
- Design sign-off: 31 Jul 2026.
- SIT exit: 02 Oct 2026.
- UAT sign-off: 23 Oct 2026.
- Go-live: 30 Oct 2026.
- Hypercare exit: 27 Nov 2026.

Daniel emphasized that go-live must hold — slipping into November runs into
Thanksgiving week and the year-end no-deploy window.

## 3. Decisions

- **D-1:** Single global Ariba realm (no regional realms). Owner: Marcus.
- **D-2:** Vendor master remains in S/4HANA as system of record; Ariba SLP is
  system of engagement. Confirmed by Rina.
- **D-3:** Catalog priority order — Grainger, Staples, CDW, Dell PunchOut first.
  CIF top-8 MRO suppliers identified by Lena offline; due 26 June.
- **D-4:** Hypercare staffing — Bluewave provides L2/L3 for 4 weeks; Acme
  ramps Ravi's team for L1 from week 2.
- **D-5:** Steering Committee cadence: monthly, first Wednesday, 14:00 GST.

## 4. Action items

| # | Action | Owner | Due |
|---|---|---|---|
| AI-1 | Provision Ariba test + prod realms via SAP ticket | Marcus | 22 Jun |
| AI-2 | Share S/4HANA FPS01 confirmation with Bluewave Integration team | Tom + Ravi | 19 Jun |
| AI-3 | Confirm top-8 CIF MRO suppliers and supplier contacts | Lena | 26 Jun |
| AI-4 | Draft design workshop schedule (4 weeks of workshops) | Anika & James | 22 Jun |
| AI-5 | Set up project SharePoint and RAID log | Helena | 18 Jun |
| AI-6 | Schedule SteerCo recurring invite | Helena | 18 Jun |
| AI-7 | Confirm Brazil NF-e tax fields requirement (BRD open item O-1) | Hector + Rina | 03 Jul |
| AI-8 | Vendor master profiling — initial duplicate & completeness scan | Rina | 03 Jul |
| AI-9 | Okta SCIM connector scope review | Tom + Yusuf | 26 Jun |
| AI-10 | Approval matrix Capex committee members (BRD O-2) | Felix | 03 Jul |

## 5. Risks raised

- Yusuf flagged that CIG version compatibility with FPS01 should be
  re-verified before integration design lock; logged as R-006 in RAID.
- Olivia flagged that 3,400 requesters across NA plants is a large training
  population for a 4-week Deploy window; she will propose a wave plan in week 3.

## 6. Next checkpoint

- Weekly PM sync: Mondays 16:00 GST starting 22 June.
- First design workshop: SLP Registration & Qualification, week of 29 June.
- Next SteerCo: 01 July 2026.

Meeting closed 11:28.
