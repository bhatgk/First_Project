# ============================================================
#  generate_plan.ps1
#  SAP Ariba SLP + Buying Project Plan — Acme Industrial
#  Excel workbook via COM automation
# ============================================================

$outFile = "d:\AI Learning\Project Management\First_Project\Ariba_SLP_Buying_Project_Plan_Acme_Industrial.xlsx"
$client  = "Acme Industrial"
$mainTitle = "SAP Ariba SLP + Buying Implementation Plan - Acme Industrial"

# ── Color helper (Excel Interior.Color = R + G*256 + B*65536) ──
function rgb([int]$r,[int]$g,[int]$b){ $r + $g*256 + $b*65536 }

$C = @{
    HdrBg   = rgb 68  114 196
    HdrFg   = 16777215
    SubHdr  = rgb 91  155 213
    Prepare = rgb 173 216 230
    Explore = rgb 144 238 144
    Realize = rgb 255 200 124
    Deploy  = rgb 216 191 216
    Run     = rgb 211 211 211
    GoLive  = rgb 255 235 156
    Complete= rgb 198 239 206
    InProg  = rgb 255 235 156
    Blocked = rgb 255 199 206
    RaciR   = rgb 198 239 206
    RaciA   = rgb 189 215 238
    RaidR   = rgb 255 199 206
    RaidA   = rgb 255 235 156
    RaidD   = rgb 189 215 238
    RaidI   = rgb 255 223 186
}

$PhaseClr = @{
    Prepare = $C.Prepare; Explore = $C.Explore
    Realize = $C.Realize; Deploy  = $C.Deploy; Run = $C.Run
}

# ── Task data: Phase, Workstream, ID, Task, Owner, Start, End ──
$TASKS = @(
    # Prepare
    @("Prepare","PMO","P-01","Project kickoff meeting","Sara Lindqvist / Marcus Chen","2026-06-15","2026-06-15"),
    @("Prepare","PMO","P-02","Governance model & steering committee setup","Sara Lindqvist","2026-06-15","2026-06-19"),
    @("Prepare","PMO","P-03","Project charter sign-off","Sara Lindqvist","2026-06-18","2026-06-22"),
    @("Prepare","PMO","P-04","Communication plan","Helena Ruiz","2026-06-22","2026-06-26"),
    @("Prepare","Technical","P-05","Ariba realm provisioning (test + prod)","Ravi Subramanian","2026-06-15","2026-06-22"),
    @("Prepare","Technical","P-06","SAP integration prerequisites checklist","Yusuf Khalil","2026-06-22","2026-06-26"),
    @("Prepare","Change","P-07","Stakeholder map & change impact baseline","Olivia Park","2026-06-15","2026-06-26"),
    # Explore
    @("Explore","SLP","E-01","Supplier data discovery & cleansing strategy","Anika Patel","2026-06-29","2026-07-03"),
    @("Explore","SLP","E-02","Supplier registration form design workshop","Anika Patel","2026-07-06","2026-07-10"),
    @("Explore","SLP","E-03","Supplier qualification questionnaire design","Anika Patel","2026-07-07","2026-07-14"),
    @("Explore","SLP","E-04","Supplier segmentation & approval workflow design","Anika Patel","2026-07-13","2026-07-17"),
    @("Explore","Buying","E-05","Procurement policy review & catalog strategy","James Whitaker","2026-06-29","2026-07-03"),
    @("Explore","Buying","E-06","Guided Buying landing page & tile design","James Whitaker","2026-07-06","2026-07-10"),
    @("Explore","Buying","E-07","Approval rules & spend authorization matrix design","James Whitaker","2026-07-13","2026-07-17"),
    @("Explore","Buying","E-08","Catalog enablement plan (CIF, PunchOut, BMEcat)","James Whitaker","2026-07-20","2026-07-24"),
    @("Explore","Technical","E-09","Integration design (SLP master data, PR/PO, GR, invoice)","Yusuf Khalil","2026-06-29","2026-07-17"),
    @("Explore","Technical","E-10","SSO & user provisioning design","Tom Becker","2026-07-14","2026-07-17"),
    @("Explore","Data","E-11","Vendor master cleansing & migration approach","Rina Takahashi","2026-06-29","2026-07-24"),
    @("Explore","Change","E-12","Training needs analysis","Olivia Park","2026-07-27","2026-07-31"),
    # Realize
    @("Realize","SLP","R-01","Configure registration & qualification forms","Anika Patel","2026-08-03","2026-08-14"),
    @("Realize","SLP","R-02","Configure supplier segmentation & approval flows","Anika Patel","2026-08-17","2026-08-21"),
    @("Realize","SLP","R-03","Configure supplier 360 view & modular questionnaires","Anika Patel","2026-08-24","2026-08-28"),
    @("Realize","Buying","R-04","Configure Guided Buying tiles, forms, and policies","James Whitaker","2026-08-03","2026-08-14"),
    @("Realize","Buying","R-05","Configure approval flows & substitute approvers","James Whitaker","2026-08-17","2026-08-21"),
    @("Realize","Buying","R-06","Configure commodity codes & catalog views","James Whitaker","2026-08-24","2026-08-28"),
    @("Realize","Buying","R-07","Onboard pilot catalogs (CIF + PunchOut)","James Whitaker","2026-08-31","2026-09-04"),
    @("Realize","Technical","R-08","Build SAP <-> Ariba integration (CIG / API)","Yusuf Khalil","2026-08-03","2026-09-11"),
    @("Realize","Technical","R-09","Configure SSO & user roles","Tom Becker","2026-08-17","2026-08-21"),
    @("Realize","Data","R-10","Execute vendor master migration to dev","Rina Takahashi","2026-08-03","2026-09-04"),
    @("Realize","Testing","R-11","SIT cycle 1","Maya Greenfield","2026-09-07","2026-09-18"),
    @("Realize","Testing","R-12","SIT cycle 2 (defect fix retest)","Maya Greenfield","2026-09-21","2026-10-02"),
    @("Realize","Change","R-13","Develop training materials & quick reference guides","Olivia Park","2026-08-03","2026-09-04"),
    # Deploy
    @("Deploy","Testing","D-01","UAT preparation & test script walkthrough","Maya Greenfield","2026-10-05","2026-10-07"),
    @("Deploy","Testing","D-02","UAT execution","Karen Faulkner","2026-10-08","2026-10-17"),
    @("Deploy","Testing","D-03","UAT sign-off","Daniel O'Connell","2026-10-22","2026-10-23"),
    @("Deploy","Change","D-04","End-user training delivery (buyers, requesters, approvers)","Olivia Park","2026-10-07","2026-10-16"),
    @("Deploy","Change","D-05","Supplier onboarding communications","Olivia Park","2026-10-07","2026-10-16"),
    @("Deploy","Data","D-06","Production vendor master migration","Rina Takahashi","2026-10-22","2026-10-28"),
    @("Deploy","Technical","D-07","Production cutover (config transport, integration switchover)","Yusuf Khalil","2026-10-26","2026-10-29"),
    @("Deploy","PMO","D-08","Go-live readiness review","Sara Lindqvist","2026-10-28","2026-10-29"),
    @("Deploy","PMO","D-09","Go-live","Sara Lindqvist","2026-10-30","2026-10-30"),
    # Run
    @("Run","Support","N-01","Hypercare ticket triage (Tier 1/2/3)","David Mensah","2026-11-02","2026-11-27"),
    @("Run","Support","N-02","Daily hypercare standup","David Mensah","2026-11-02","2026-11-27"),
    @("Run","Change","N-03","Adoption metrics tracking","Olivia Park","2026-11-02","2026-11-27"),
    @("Run","PMO","N-04","Lessons learned workshop","Sara Lindqvist","2026-11-23","2026-11-25"),
    @("Run","PMO","N-05","Transition to BAU support & project closure","Sara Lindqvist","2026-11-25","2026-11-27")
)

# ── RACI data ──
# Roles: Sponsor=Daniel O'Connell, PM=Sara Lindqvist, PMO=Helena Ruiz,
#        SLP=Anika Patel, Buying=James Whitaker, Integration=Yusuf Khalil,
#        Data=Rina Takahashi, Security=Tom Becker, Test=Maya Greenfield,
#        Change=Olivia Park, Support=David Mensah, Admin=Ravi Subramanian, UAT=Karen Faulkner
$RACI_ROLES = @(
    "Daniel O'Connell`n(Business Sponsor)",
    "Sara Lindqvist`n(Program Manager)",
    "Helena Ruiz`n(PMO Lead)",
    "Anika Patel`n(SLP Lead)",
    "James Whitaker`n(Buying Lead)",
    "Yusuf Khalil`n(Integration Lead)",
    "Rina Takahashi`n(Data Lead)",
    "Tom Becker`n(Security Lead)",
    "Maya Greenfield`n(Test Lead)",
    "Olivia Park`n(Change Lead)",
    "David Mensah`n(Support Lead)",
    "Ravi Subramanian`n(Ariba Admin)",
    "Karen Faulkner`n(UAT Lead)"
)
# Workstream, then RACI values per role (13 values)
$RACI_DATA = @(
    @("Project governance",                    "A","R","C","I","I","I","I","I","I","I","I","I","I"),
    @("SLP - Registration & qualification",    "I","A","I","R","I","C","C","I","I","C","I","C","I"),
    @("SLP - Supplier segmentation & workflows","I","A","I","R","I","C","I","I","I","I","I","C","I"),
    @("Buying - Guided Buying & policies",      "I","A","I","I","R","C","I","I","I","C","I","C","I"),
    @("Buying - Approvals & catalogs",          "I","A","I","I","R","C","I","I","I","I","I","C","I"),
    @("Integration (SAP <-> Ariba)",            "I","A","I","C","C","R","C","C","I","I","I","C","I"),
    @("Vendor master & data migration",         "I","A","I","C","C","C","R","I","I","I","I","I","I"),
    @("Security, SSO, user provisioning",       "I","A","I","I","I","C","I","R","I","I","I","C","I"),
    @("Testing (SIT)",                          "I","A","I","C","C","C","I","I","R","I","I","I","I"),
    @("UAT",                                    "A","C","I","C","C","I","I","I","C","C","I","I","R"),
    @("Training & change management",           "I","A","I","C","C","I","I","I","I","R","I","I","I"),
    @("Cutover & go-live",                      "A","R","C","C","C","C","C","C","C","C","C","C","C"),
    @("Hypercare",                              "I","A","I","C","C","C","I","I","I","C","R","C","I")
)

# ── RAID data: ID, Type, Title, Description, Likelihood, Impact, Owner, DueDate, Status, Notes ──
$RAID_DATA = @(
    # Standard risks
    @("R-01","Risk","Vendor master data quality","Poor data quality in SAP delays SLP migration","H","H","Rina Takahashi","Explore phase","Open","Run profiling in Explore (AI-8); allocate cleansing sprints. Source: BRD Sec 5, Vendor Master Extract"),
    @("R-02","Risk","Supplier adoption of registration","Suppliers slow to complete self-registration; impacts SLP go-live","M","H","Olivia Park","Explore phase","Open","Supplier comms wave plan; concierge desk for top 50 suppliers"),
    @("R-03","Risk","CIG integration capacity","Cloud Integration Gateway throughput constraints during peak","M","M","Yusuf Khalil","Realize phase","Open","Performance test integration early; SAP support ticket on standby"),
    @("R-04","Risk","Catalog content readiness","Suppliers unable to deliver CIF/PunchOut catalogs on time","M","M","James Whitaker","Realize phase","Open","Stagger catalog onboarding; identify priority categories. Source: BRD Sec 2.3"),
    @("R-05","Risk","Approver hierarchy drift","SAP HR org changes invalidate approval matrix before go-live","L","H","James Whitaker","Deploy phase","Open","Lock org snapshot for cutover; daily sync pre-go-live"),
    # Project-specific risks from MoM and Vendor Master
    @("R-06","Risk","CIG version compatibility with FPS01","CIG version may not fully support S/4HANA 2023 FPS01 - to be re-verified before integration design lock","M","H","Yusuf Khalil","26 Jun 2026","Open","Source: Kickoff MoM item R-006, flagged by Yusuf Khalil. Follow-up via AI-2"),
    @("R-07","Risk","Training population size","3,400 requesters across NA plants is large for 4-week Deploy window","M","M","Olivia Park","03 Jul 2026","Open","Source: Kickoff MoM, Olivia Park. Wave plan to be proposed by Week 3"),
    @("R-08","Risk","Duplicate vendor records","~4 confirmed duplicate pairs in Vendor Master Extract (Grainger, Siemens, TCS, Apex Lab Reagents); 2 records missing tax IDs","H","M","Rina Takahashi","03 Jul 2026","Open","Source: Vendor Master Extract. Run dedup sprint in Explore. V-101455 (test record) flagged for deletion"),
    # Assumptions
    @("A-01","Assumption","Ariba realm provisioning SLA","SAP will provision Ariba test + prod realms within 5 business days of SOW signature","","","Ravi Subramanian","22 Jun 2026","Open","Source: Project Charter Sec 6; AI-1 (Marcus, due 22 Jun)"),
    @("A-02","Assumption","S/4HANA FPS01 on supported CIG release","SAP backend is on S/4HANA 2023 FPS01 and is supported by current CIG version","","","Yusuf Khalil","19 Jun 2026","Open","Source: Charter Sec 6; AI-2 (Tom + Ravi, due 19 Jun). Critical for integration design"),
    @("A-03","Assumption","Vendor master stays in S/4HANA","Vendor master remains system of record in S/4HANA; Ariba SLP is system of engagement","","","Rina Takahashi","03 Jul 2026","Open","Source: Charter Sec 6, confirmed in Kickoff MoM D-2"),
    @("A-04","Assumption","Single SSO IdP (Okta)","Single Okta instance is the IdP for all user populations; SCIM provisioning in scope","","","Tom Becker","26 Jun 2026","Open","Source: Charter Sec 6; BRD Sec 4; AI-9 (Tom + Yusuf, due 26 Jun)"),
    @("A-05","Assumption","Brazil NF-e fields not required in Phase 1","Brazil tax registration NF-e fields are out of scope for Phase 1 SLP registration form","","","Hector Alvarez","03 Jul 2026","Open","Source: BRD O-1 (open item); AI-7 (Hector + Rina, due 3 Jul). If required, adds configuration effort"),
    # Issues
    @("I-01","Issue","(placeholder)","Open issue slot for team to populate","","","","","Open",""),
    @("I-02","Issue","Capex committee approval matrix pending","Finance has not confirmed Capex committee members for approval matrix configuration","","","Felix Nguyen","03 Jul 2026","Open","Source: BRD O-2; AI-10 (Felix, due 3 Jul). Blocks Guided Buying approval rule config"),
    @("I-03","Issue","PunchOut go-live sequencing not confirmed","Priority order for PunchOut supplier go-live not agreed","","","James Whitaker","26 Jun 2026","Open","Source: BRD O-3. Decision D-3 (Kickoff MoM) confirmed Grainger, Staples, CDW, Dell as priority order"),
    # Dependencies
    @("D-01","Dependency","SAP basis team availability for CIG setup","SAP basis resources required for CIG configuration; blocks integration build","","","Yusuf Khalil","29 Jun 2026","Open","Confirm resource plan in Explore; align with AI-2"),
    @("D-02","Dependency","Procurement policy sign-off","Approved procurement policy required before Guided Buying design can be frozen","","","Daniel O'Connell","03 Jul 2026","Open","Source: BRD Sec 2.1"),
    @("D-03","Dependency","Master data governance decisions","Vendor master governance decisions required before migration scope can be locked","","","Rina Takahashi","03 Jul 2026","Open","Source: BRD Sec 5; MoM D-2"),
    @("D-04","Dependency","Top-8 CIF MRO supplier list","Lena Brandt to confirm top-8 MRO suppliers and contacts for CIF catalog loading","","","Lena Brandt","26 Jun 2026","Open","Source: Kickoff MoM AI-3. Blocks R-07 (onboard pilot catalogs)"),
    @("D-05","Dependency","Hypercare staffing model agreement","Bluewave/Acme split for hypercare L1/L2/L3 support must be agreed before Deploy phase","","","Sara Lindqvist","Realize phase","Open","Source: BRD O-4; MoM D-4 (Bluewave L2/L3 for 4 wks; Ravi team ramps L1)")
)

# ── Milestone data: Milestone, TargetDate, Phase, Owner, Status, IsGoLive ──
$MILESTONES = @(
    @("Project kickoff","2026-06-15","Prepare","Sara Lindqvist","Not Started",$false),
    @("Design sign-off (SLP + Buying)","2026-07-31","Explore","Daniel O'Connell","Not Started",$false),
    @("Integration build complete","2026-09-11","Realize","Yusuf Khalil","Not Started",$false),
    @("SIT exit","2026-10-02","Realize","Maya Greenfield","Not Started",$false),
    @("UAT sign-off","2026-10-23","Deploy","Daniel O'Connell","Not Started",$false),
    @("GO-LIVE","2026-10-30","Deploy","Sara Lindqvist","Not Started",$true),
    @("Hypercare exit / BAU transition","2026-11-27","Run","David Mensah","Not Started",$false)
)

# ======================================================================
# Excel creation
# ======================================================================
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

# Remove extra default sheets (keep 1)
while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

# ── Inline helpers ──
function xcell($ws,$r,$c) { $ws.Cells.Item($r,$c) }

# Use Formula (string) to avoid COM VARIANT type-cast issues in PS 5.1
function wcv($ws,$r,$c,$v) { $ws.Cells.Item($r,$c).Formula = "$v" }

function wcd($ws,$r,$c,$d) {
    $dt = [datetime]$d
    $ce = $ws.Cells.Item($r,$c)
    $ce.Formula = "=DATE(" + $dt.Year + "," + $dt.Month + "," + $dt.Day + ")"
    $ce.NumberFormat = "dd-mmm-yy"
}

function title_row($ws,$ncols,$txt) {
    $rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$ncols))
    $rng.Merge() | Out-Null
    $ce = $ws.Cells.Item(1,1)
    $ce.Value2 = $txt
    $ce.Font.Bold = $true; $ce.Font.Size = 14; $ce.Font.Name = "Arial"
    $ce.Interior.Color = $C.HdrBg; $ce.Font.Color = $C.HdrFg
    $ce.HorizontalAlignment = -4108   # xlCenter
    $ws.Rows.Item(1).RowHeight = 32
}

function hdr_row($ws,$row,[string[]]$hdrs) {
    for ($h = 0; $h -lt $hdrs.Count; $h++) {
        $ce = $ws.Cells.Item($row, $h+1)
        $ce.Value2 = $hdrs[$h]
        $ce.Font.Bold = $true; $ce.Font.Name = "Arial"; $ce.Font.Size = 10
        $ce.Interior.Color = $C.SubHdr; $ce.Font.Color = $C.HdrFg
        $ce.HorizontalAlignment = -4108
        $ce.WrapText = $true
    }
    $ws.Rows.Item($row).RowHeight = 30
}

function set_col_widths($ws,[int[]]$widths) {
    for ($w = 0; $w -lt $widths.Count; $w++) {
        $ws.Columns.Item($w+1).ColumnWidth = $widths[$w]
    }
}

function freeze_at($ws,$row,$col) {
    $ws.Activate()
    $ws.Cells.Item($row,$col).Select() | Out-Null
    $xl.ActiveWindow.FreezePanes = $true
}

function thin_border($ws,$r1,$c1,$r2,$c2) {
    $rng = $ws.Range($ws.Cells.Item($r1,$c1),$ws.Cells.Item($r2,$c2))
    # outer borders
    foreach ($idx in @(7,8,9,10)) { $rng.Borders.Item($idx).LineStyle = 1; $rng.Borders.Item($idx).Weight = 2 }
    # inner grid
    foreach ($idx in @(11,12)) { $rng.Borders.Item($idx).LineStyle = 1; $rng.Borders.Item($idx).Weight = 1 }
}

function add_status_cf($ws,$r1,$r2,$col) {
    $rng = $ws.Range($ws.Cells.Item($r1,$col),$ws.Cells.Item($r2,$col))
    $cf1 = $rng.FormatConditions.Add(2,$null,'="Complete"');   $cf1.Interior.Color = $C.Complete
    $cf2 = $rng.FormatConditions.Add(2,$null,'="In Progress"');$cf2.Interior.Color = $C.InProg
    $cf3 = $rng.FormatConditions.Add(2,$null,'="Blocked"');    $cf3.Interior.Color = $C.Blocked
}

# ======================================================================
# SHEET 1 — WBS
# ======================================================================
Write-Host "Building WBS sheet..."
$wsW = $wb.Sheets.Item(1)
$wsW.Name = "WBS"

$wbsHdrs = @("Task ID","Phase","Workstream","Task","Owner","Start","End","Duration (days)","% Complete","Status","Notes")
title_row $wsW 11 $mainTitle
hdr_row   $wsW 2 $wbsHdrs

$row = 3
foreach ($t in $TASKS) {
    $phase = $t[0]; $wstream = $t[1]; $tid = $t[2]; $tname = $t[3]; $owner = $t[4]
    $sD = [datetime]::Parse($t[5]); $eD = [datetime]::Parse($t[6])
    $dur = ($eD - $sD).Days + 1
    $bgc = $PhaseClr[$phase]

    wcv $wsW $row 1 $tid
    wcv $wsW $row 2 $phase
    wcv $wsW $row 3 $wstream
    wcv $wsW $row 4 $tname
    wcv $wsW $row 5 $owner
    wcd $wsW $row 6 $sD
    wcd $wsW $row 7 $eD
    wcv $wsW $row 8 $dur
    $wsW.Cells.Item($row,9).Formula = "0"
    $wsW.Cells.Item($row,9).NumberFormat = "0%"
    wcv $wsW $row 10 "Not Started"
    wcv $wsW $row 11 ""

    # Phase color on Phase column
    $wsW.Cells.Item($row,2).Interior.Color = $bgc

    # Go-live gold highlight
    if ($tid -eq "D-09") {
        $wsW.Range($wsW.Cells.Item($row,1),$wsW.Cells.Item($row,11)).Interior.Color = $C.GoLive
        $wsW.Range($wsW.Cells.Item($row,1),$wsW.Cells.Item($row,11)).Font.Bold = $true
    }

    # Arial on all cells
    for ($c = 1; $c -le 11; $c++) { $wsW.Cells.Item($row,$c).Font.Name = "Arial"; $wsW.Cells.Item($row,$c).Font.Size = 10 }

    $row++
}

$lastDataRow = $row - 1
set_col_widths $wsW @(8,10,12,46,26,12,12,14,12,12,30)
thin_border    $wsW 2 1 $lastDataRow 11
add_status_cf  $wsW 3 $lastDataRow 10
freeze_at      $wsW 3 1

# ======================================================================
# SHEET 2 — GANTT
# ======================================================================
Write-Host "Building Gantt sheet..."
$wsG = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$wsG.Name = "Gantt"

# Week-start dates (24 Mondays from Jun 15)
$weekStarts = @()
$wkBase = [datetime]"2026-06-15"
for ($w = 0; $w -lt 24; $w++) { $weekStarts += $wkBase.AddDays($w * 7) }

$ganttCols = 5 + 24   # A-E fixed, F-AC weeks
title_row $wsG $ganttCols $mainTitle

# Header row 2: Task ID, Phase, Task, Start, End, then week labels
wcv $wsG 2 1 "Task ID"
wcv $wsG 2 2 "Phase"
wcv $wsG 2 3 "Task"
wcv $wsG 2 4 "Start"
wcv $wsG 2 5 "End"
for ($w = 0; $w -lt 24; $w++) {
    $wLabel = "W" + ($w+1) + " " + $weekStarts[$w].ToString("dd-MMM")
    wcv $wsG 2 ($w+6) $wLabel
    $hce = $wsG.Cells.Item(2,$w+6)
    $hce.Font.Bold = $true; $hce.Font.Name = "Arial"; $hce.Font.Size = 8
    $hce.Interior.Color = $C.SubHdr; $hce.Font.Color = $C.HdrFg
    $hce.HorizontalAlignment = -4108
    $hce.WrapText = $true
    $wsG.Columns.Item($w+6).ColumnWidth = 6
}
# Style the first 5 header cells
foreach ($h in @(1,2,3,4,5)) {
    $hce = $wsG.Cells.Item(2,$h)
    $hce.Font.Bold = $true; $hce.Font.Name = "Arial"; $hce.Font.Size = 10
    $hce.Interior.Color = $C.SubHdr; $hce.Font.Color = $C.HdrFg
    $hce.HorizontalAlignment = -4108
}
$wsG.Rows.Item(2).RowHeight = 38

$grow = 3
foreach ($t in $TASKS) {
    $phase = $t[0]; $tid = $t[2]; $tname = $t[3]
    $sD = [datetime]::Parse($t[5]); $eD = [datetime]::Parse($t[6])
    $bgc = $PhaseClr[$phase]

    wcv $wsG $grow 1 $tid
    wcv $wsG $grow 2 $phase
    wcv $wsG $grow 3 $tname
    wcd $wsG $grow 4 $sD
    wcd $wsG $grow 5 $eD
    $wsG.Cells.Item($grow,2).Interior.Color = $bgc

    for ($c = 1; $c -le 5; $c++) {
        $wsG.Cells.Item($grow,$c).Font.Name = "Arial"
        $wsG.Cells.Item($grow,$c).Font.Size = 10
    }

    # Fill Gantt bars
    for ($w = 0; $w -lt 24; $w++) {
        $wkS = $weekStarts[$w]
        $wkE = $wkS.AddDays(6)
        if ($wkS -le $eD -and $wkE -ge $sD) {
            $wsG.Cells.Item($grow, $w+6).Interior.Color = $bgc
        }
    }

    $grow++
}

$gLastRow = $grow - 1
# Column widths for fixed cols
$wsG.Columns.Item(1).ColumnWidth = 8
$wsG.Columns.Item(2).ColumnWidth = 9
$wsG.Columns.Item(3).ColumnWidth = 32
$wsG.Columns.Item(4).ColumnWidth = 11
$wsG.Columns.Item(5).ColumnWidth = 11

# Phase legend at bottom
$legendRow = $gLastRow + 2
wcv $wsG $legendRow 1 "Phase Legend:"
$wsG.Cells.Item($legendRow,1).Font.Bold = $true
$wsG.Cells.Item($legendRow,1).Font.Name = "Arial"
$phases = @("Prepare","Explore","Realize","Deploy","Run")
for ($p = 0; $p -lt $phases.Count; $p++) {
    $lc = 2 + $p
    $wsG.Cells.Item($legendRow,$lc).Value2 = $phases[$p]
    $wsG.Cells.Item($legendRow,$lc).Interior.Color = $PhaseClr[$phases[$p]]
    $wsG.Cells.Item($legendRow,$lc).Font.Name = "Arial"
    $wsG.Cells.Item($legendRow,$lc).Font.Size = 10
    $wsG.Cells.Item($legendRow,$lc).HorizontalAlignment = -4108
}

thin_border $wsG 2 1 $gLastRow 5
freeze_at   $wsG 3 6

# ======================================================================
# SHEET 3 — RACI
# ======================================================================
Write-Host "Building RACI sheet..."
$wsR = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$wsR.Name = "RACI"

$raciTotalCols = 1 + $RACI_ROLES.Count   # workstream + 13 roles = 14
title_row $wsR $raciTotalCols $mainTitle

# Legend row
$wsR.Range($wsR.Cells.Item(2,1),$wsR.Cells.Item(2,$raciTotalCols)).Merge() | Out-Null
$wsR.Cells.Item(2,1).Value2 = "R = Responsible   |   A = Accountable   |   C = Consulted   |   I = Informed"
$wsR.Cells.Item(2,1).Font.Bold = $true; $wsR.Cells.Item(2,1).Font.Name = "Arial"; $wsR.Cells.Item(2,1).Font.Size = 10
$wsR.Cells.Item(2,1).HorizontalAlignment = -4108
$wsR.Rows.Item(2).RowHeight = 20

# Header row 3: Workstream + roles
$wsR.Cells.Item(3,1).Value2 = "Workstream"
$wsR.Cells.Item(3,1).Font.Bold = $true; $wsR.Cells.Item(3,1).Font.Name = "Arial"; $wsR.Cells.Item(3,1).Font.Size = 10
$wsR.Cells.Item(3,1).Interior.Color = $C.SubHdr; $wsR.Cells.Item(3,1).Font.Color = $C.HdrFg
$wsR.Cells.Item(3,1).HorizontalAlignment = -4108

for ($rr = 0; $rr -lt $RACI_ROLES.Count; $rr++) {
    $ce = $wsR.Cells.Item(3,$rr+2)
    $ce.Value2 = $RACI_ROLES[$rr]
    $ce.Font.Bold = $true; $ce.Font.Name = "Arial"; $ce.Font.Size = 8
    $ce.Interior.Color = $C.SubHdr; $ce.Font.Color = $C.HdrFg
    $ce.HorizontalAlignment = -4108; $ce.WrapText = $true
}
$wsR.Rows.Item(3).RowHeight = 50

$rrow = 4
foreach ($rd in $RACI_DATA) {
    $wsR.Cells.Item($rrow,1).Value2 = $rd[0]
    $wsR.Cells.Item($rrow,1).Font.Bold = $true; $wsR.Cells.Item($rrow,1).Font.Name = "Arial"; $wsR.Cells.Item($rrow,1).Font.Size = 10

    for ($v = 1; $v -le 13; $v++) {
        $val = $rd[$v]
        $ce = $wsR.Cells.Item($rrow,$v+1)
        $ce.Value2 = $val
        $ce.Font.Bold = ($val -eq "R" -or $val -eq "A")
        $ce.Font.Name = "Arial"; $ce.Font.Size = 11
        $ce.HorizontalAlignment = -4108
        if ($val -eq "R") { $ce.Interior.Color = $C.RaciR }
        if ($val -eq "A") { $ce.Interior.Color = $C.RaciA }
    }

    $rrow++
}

$rLastRow = $rrow - 1
# Column widths
$wsR.Columns.Item(1).ColumnWidth = 32
for ($rc = 2; $rc -le $raciTotalCols; $rc++) { $wsR.Columns.Item($rc).ColumnWidth = 12 }

thin_border $wsR 3 1 $rLastRow $raciTotalCols
freeze_at   $wsR 4 2

# ======================================================================
# SHEET 4 — RAID LOG
# ======================================================================
Write-Host "Building RAID Log sheet..."
$wsRaid = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$wsRaid.Name = "RAID Log"

$raidHdrs = @("ID","Type","Title","Description","Likelihood (H/M/L)","Impact (H/M/L)","Owner","Due Date","Status","Mitigation / Notes")
title_row $wsRaid 10 $mainTitle
hdr_row   $wsRaid 2 $raidHdrs

$drow = 3
foreach ($rd in $RAID_DATA) {
    $rtype = $rd[1]
    $bgc = switch ($rtype) {
        "Risk"       { $C.RaidR }
        "Assumption" { $C.RaidA }
        "Issue"      { $C.RaidI }
        "Dependency" { $C.RaidD }
        default      { -1 }
    }

    for ($dc = 0; $dc -lt 10; $dc++) {
        $ce = $wsRaid.Cells.Item($drow,$dc+1)
        $ce.Value2 = $rd[$dc]
        $ce.Font.Name = "Arial"; $ce.Font.Size = 10
        $ce.WrapText = $true
        if ($bgc -ge 0 -and ($dc -eq 0 -or $dc -eq 1)) { $ce.Interior.Color = $bgc }
    }
    $drow++
}

$raidLastRow = $drow - 1
$raidWidths = @(6,12,28,38,14,12,22,14,10,42)
set_col_widths $wsRaid $raidWidths
thin_border    $wsRaid 2 1 $raidLastRow 10
add_status_cf  $wsRaid 3 $raidLastRow 9

# Row heights for readability
for ($dr = 3; $dr -le $raidLastRow; $dr++) { $wsRaid.Rows.Item($dr).RowHeight = 45 }
freeze_at $wsRaid 3 1

# ======================================================================
# SHEET 5 — MILESTONES
# ======================================================================
Write-Host "Building Milestones sheet..."
$wsM = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$wsM.Name = "Milestones"

$milHdrs = @("Milestone","Target Date","Phase","Owner","Status")
title_row $wsM 5 $mainTitle
hdr_row   $wsM 2 $milHdrs

$mrow = 3
foreach ($ms in $MILESTONES) {
    $mname = $ms[0]; $mdate = [datetime]::Parse($ms[1]); $mphase = $ms[2]
    $mown = $ms[3]; $mstat = $ms[4]; $isGL = [bool]$ms[5]

    wcv $wsM $mrow 1 $mname
    wcd $wsM $mrow 2 $mdate
    wcv $wsM $mrow 3 $mphase
    wcv $wsM $mrow 4 $mown
    wcv $wsM $mrow 5 $mstat

    $pbc = if ($PhaseClr.ContainsKey($mphase)) { $PhaseClr[$mphase] } else { -1 }

    for ($mc = 1; $mc -le 5; $mc++) {
        $ce = $wsM.Cells.Item($mrow,$mc)
        $ce.Font.Name = "Arial"; $ce.Font.Size = 10
        if ($isGL) {
            $ce.Font.Bold = $true
            $ce.Interior.Color = $C.GoLive
            $ce.Font.Size = 11
        }
    }
    if (-not $isGL -and $pbc -ge 0) {
        $wsM.Cells.Item($mrow,3).Interior.Color = $pbc
    }
    $mrow++
}

$mLastRow = $mrow - 1
set_col_widths $wsM @(38,14,10,22,12)
thin_border    $wsM 2 1 $mLastRow 5
add_status_cf  $wsM 3 $mLastRow 5
freeze_at      $wsM 3 1

# ======================================================================
# Reorder sheets and save
# ======================================================================
Write-Host "Saving workbook..."

# Set WBS as first/active sheet
$wsW.Move($wb.Sheets.Item(1)) | Out-Null

# Remove any completely empty extra sheet that may have been there from Add()
try {
    foreach ($s in $wb.Sheets) {
        if ($s.UsedRange.Rows.Count -eq 1 -and $s.UsedRange.Columns.Count -eq 1 -and
            $s.Cells.Item(1,1).Value2 -eq $null -and
            $s.Name -notin @("WBS","Gantt","RACI","RAID Log","Milestones")) {
            $s.Delete()
        }
    }
} catch {}

$wb.SaveAs($outFile, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
$wb.Close($false)
$xl.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host ""
Write-Host "Done! Workbook saved to:"
Write-Host $outFile
