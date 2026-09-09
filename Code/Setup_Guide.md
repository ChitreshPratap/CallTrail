# CallTrail — Setup Guide

## 1. Two files, two roles
- **Claim_Calling_Tracker.xlsx** → put this on the shared drive ONCE. It is pure data (Claims, History, Users, Config sheets). No macros live here.
- **CallTrail_App.xlsm** → the file you build below, using the 4 `.bas` modules provided. Distribute a COPY of this to every caller/admin's machine.

This split is deliberate: if the data file ever got macros in it, every user opening it from the share would need to enable macros on a file others are actively writing to — riskier and slower. Keep logic local, data central.

## Current Claims schema
`ClaimID | ClaimSite | ClaimProviderName | ClaimQuery | ClaimCreationDate | ClaimStatus | Attempt | ClaimUpdatedSite | ClaimInsertionDate | ClaimInsertedBy | LastUpdatedDate | LastUpdatedBy | LastComment | ClaimClosedDate`

- **ClaimID, ClaimSite, ClaimProviderName, ClaimQuery, ClaimCreationDate** — entered by the user, mandatory.
- **ClaimStatus, Attempt, ClaimUpdatedSite, ClaimInsertionDate, ClaimInsertedBy** — auto-filled at insertion (Status=Pending, Attempt=0).
- **LastUpdatedDate, LastUpdatedBy, LastComment, ClaimClosedDate** — auto-filled as calls get logged; power the dashboards (last-touched view, time-to-close) without needing to scan the full History sheet.

**Adding columns later:** the data-access code (`modDataAccess.bas`) looks up every column by its header text in row 1 (via `GetColIndex` in `modUtils.bas`), not by a fixed position. So to add a new field later (e.g. `Priority`, `ClaimAmount`):
1. Add the header to row 1 of the Claims sheet on the central DB.
2. Add one line in the relevant `modDataAccess` function to read/write it.
3. Nothing else in the app breaks — existing columns keep working regardless of where the new one sits.

## 2. Build the app file
1. Create a new blank workbook, save as **CallTrail_App.xlsm** (macro-enabled).
2. Open VBA editor (Alt+F11) → File → Import File → import each of:
   - `modConfig.bas`
   - `modUtils.bas`
   - `modDataAccess.bas`
   - `modMain.bas`
   - **Optional, object-oriented layer** (see section 2a below): `clsClaim.cls`, `clsHistoryEntry.cls`, `clsClaimRepository.cls`
3. In `modConfig.bas`, change `DB_PATH` to the real UNC path of the shared `Claim_Calling_Tracker.xlsx`, e.g. `\\CorpShare\Claims\Claim_Calling_Tracker.xlsx`.
4. In the `Users` sheet of the central DB, replace the sample rows with real Windows usernames (`Environ("USERNAME")` values) and mark exactly the right people as `Admin`.

## 2a. Optional: class-module (object-oriented) layer
`clsClaim.cls` and `clsHistoryEntry.cls` are entity classes — a claim/history row as an object with named properties instead of an array index. `clsClaimRepository.cls` is an object-oriented alternative to `modDataAccess.bas` that returns/accepts these objects. Both `.cls` files import the same way as `.bas` (File → Import File).

**You can use either layer, or both side by side** — they read/write the same sheets via the same `modUtils` helpers, so there's no conflict. If you use the repository class, form code looks like this instead of calling `modDataAccess` functions directly:
```vb
Private repo As New clsClaimRepository

Private Sub cmdSubmit_Click()
    If repo.LogCallAndUpdateStatus(Trim(txtClaimID.Value), Trim(txtComment.Value), cboStatus.Value) Then
        MsgBox "Call logged.", vbInformation
    End If
End Sub
```
And a dashboard/report sub can loop claims as real objects:
```vb
Dim repo As New clsClaimRepository
Dim c As clsClaim
For Each c In repo.GetAllClaims()
    If Not c.IsClosed And c.DaysOpen > 7 Then
        Debug.Print c.ClaimID & " - open " & c.DaysOpen & " days, site " & c.ClaimSite
    End If
Next c
```
`DaysOpen` and `IsClosed` on `clsClaim` are computed on the fly from the dates already loaded — nothing extra to store or keep in sync.

**One thing to know:** `clsClaimRepository.FindClaim` currently calls `GetAllClaims` internally (simplest correct implementation) — fine for the claim volumes this app is built for, but if your Claims sheet grows into the tens of thousands of rows, that's the first thing to optimize (a direct row lookup, same as `modDataAccess.FindClaimRow` does).

## 3. Build the 4 forms (UserForm, not HTML — see note below)
Insert → UserForm for each. Modern flat look for all forms: white background, `Segoe UI` 10pt font, one accent color (e.g. `#1F4E78`) for buttons/headers, no default gray Windows controls where avoidable (use flat command buttons, `BackStyle = Opaque`, no 3D borders).

### frmAddClaim
Controls: `txtClaimID`, `txtClaimSite`, `txtProviderName`, `txtClaimQuery` (MultiLine), `dtCreationDate` (TextBox, or a DTPicker if you have that control installed), `cmdSave`, `cmdCancel`. All five are mandatory — validate before saving.
```vb
Private Sub cmdSave_Click()
    If Trim(txtClaimID.Value) = "" Or Trim(txtClaimSite.Value) = "" _
       Or Trim(txtProviderName.Value) = "" Or Trim(txtClaimQuery.Value) = "" _
       Or Trim(dtCreationDate.Value) = "" Then
        MsgBox "ClaimID, Site, Provider Name, Query and Creation Date are all required.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(dtCreationDate.Value) Then
        MsgBox "Creation Date is not a valid date.", vbExclamation
        Exit Sub
    End If

    If AddClaim(Trim(txtClaimID.Value), Trim(txtClaimSite.Value), _
                Trim(txtProviderName.Value), Trim(txtClaimQuery.Value), _
                CDate(dtCreationDate.Value)) Then
        MsgBox "Claim added.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
```

### frmUpdateClaim (the form every caller uses daily)
Controls: `txtClaimID` (TextBox), `cboStatus` (ComboBox — set `.List = Array("Pending","Closed")` in `Initialize`), `txtComment` (TextBox, MultiLine), `cmdSubmit` (Button).
```vb
Private Sub UserForm_Initialize()
    cboStatus.List = Array("Pending", "Closed")
    cboStatus.ListIndex = 0
End Sub

Private Sub cmdSubmit_Click()
    If Trim(txtClaimID.Value) = "" Or Trim(txtComment.Value) = "" Then
        MsgBox "Claim ID and comment are required.", vbExclamation
        Exit Sub
    End If
    If LogCallAndUpdateStatus(Trim(txtClaimID.Value), Trim(txtComment.Value), cboStatus.Value) Then
        MsgBox "Call logged.", vbInformation
        txtComment.Value = ""
    End If
End Sub
```
`LogCallAndUpdateStatus` now also stamps `LastUpdatedDate`, `LastUpdatedBy`, and `ClaimClosedDate` automatically behind the scenes — no form changes needed for that.

### frmAdminSite (Admin-only, shown only if IsCurrentUserAdmin) — was frmAdminLocation
Controls: `txtClaimID`, `txtNewSite`, `cmdUpdate`.
```vb
Private Sub cmdUpdate_Click()
    If ChangeUpdatedSite(Trim(txtClaimID.Value), Trim(txtNewSite.Value)) Then
        MsgBox "Site updated.", vbInformation
        Unload Me
    End If
End Sub
```

### frmClaimHistory (read-only viewer)
Controls: `txtClaimID`, `cmdLoad`, `lstHistory` (ListBox, `ColumnCount = 4`, wide enough for Caller/DateTime/Comment/Status).
```vb
Private Sub cmdLoad_Click()
    Dim data As Variant
    data = GetHistoryForClaim(Trim(txtClaimID.Value))
    lstHistory.Clear
    If IsEmpty(data) Then
        MsgBox "No history found for this claim.", vbInformation
        Exit Sub
    End If
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstHistory.AddItem data(i, 1)
        lstHistory.List(lstHistory.ListCount - 1, 1) = data(i, 2)
        lstHistory.List(lstHistory.ListCount - 1, 2) = data(i, 3)
        lstHistory.List(lstHistory.ListCount - 1, 3) = data(i, 4)
    Next i
End Sub
```

## 4. Wire up buttons on Sheet1 of the app file (or a ribbon)
Simplest: put 4 buttons on Sheet1 (Insert → Shapes or Form Controls), right-click → Assign Macro →
`modMain.ShowAddClaimForm`, `ShowUpdateClaimForm`, `ShowAdminSiteForm`, `ShowClaimHistoryViewer`.

## 5. Distribute
Save `CallTrail_App.xlsm`, digitally sign it or have IT trust the location, and send a copy to each caller. Everyone points at the same `DB_PATH`.

---

## Important platform-constraint notes (read before rollout)

**"HTML" UI ask, honestly assessed:** VBA's only native way to show real HTML is the `WebBrowser` control, which runs on the Internet Explorer engine — deprecated by Microsoft, and on newer Windows/Office builds it can be missing, blocked by group policy, or need the IE11 emulation registry key set per machine. It's workable in a controlled corporate image, but it's the least reliable part of this design and the first thing to break as machines get patched. Given your stated preference for accuracy over optimism on platform constraints, I built the primary app on styled UserForms instead — 100% supported, and looks clean if you keep it flat/one-accent-color as above. If you still want the HTML route as a stretch goal, the safer variant is an **HTA file launched via `Shell()`** that talks back to Excel through a COM automation object — separate from the WebBrowser-control approach, more control over styling, but is its own project. I can build this as an add-on later if you decide it's worth it.

**The real risk in this design is Excel-as-shared-database concurrency**, not the UI. A single `.xlsx` has no row-level locking: two people saving in the same second will occasionally cause a "file in use" error, and — worse — a Save from user B *after* user A opened but before user A wrote could silently overwrite user A's unrelated changes elsewhere in the sheet, because a full-workbook Save writes the whole file, not just the changed cells. The `OpenCentralDB`/`CloseCentralDB` retry pattern above minimizes the collision window (open → single write → save → close, every time) and will be fine for a small team (roughly under 10–15 concurrent users making occasional calls). If call volume or headcount grows, or if simultaneous-write errors start showing up, the next step is swapping `modDataAccess` for calls to a real database (Access, or SQL Server via ADO) — because all data access is isolated in that one module, the forms and business logic above don't need to change at all.

**CallingAttempted** currently counts every logged call (Pending or Closed) toward the total. If you instead want it to count only calls up to and including the closing call, that's a one-line change in `LogCallAndUpdateStatus` — let me know which you want.
