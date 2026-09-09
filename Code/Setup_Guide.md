# CallTrail — Setup Guide

## 1. Two files, two roles
- **Claim_Calling_Tracker.xlsx** → put this on the shared drive ONCE. It is pure data (Claims, History, Users, Config sheets). No macros live here.
- **CallTrail_App.xlsm** → the file you build below, using the 4 `.bas` modules provided. Distribute a COPY of this to every caller/admin's machine.

This split is deliberate: if the data file ever got macros in it, every user opening it from the share would need to enable macros on a file others are actively writing to — riskier and slower. Keep logic local, data central.

## 2. Build the app file
1. Create a new blank workbook, save as **CallTrail_App.xlsm** (macro-enabled).
2. Open VBA editor (Alt+F11) → File → Import File → import each of:
   - `modConfig.bas`
   - `modUtils.bas`
   - `modDataAccess.bas`
   - `modMain.bas`
3. In `modConfig.bas`, change `DB_PATH` to the real UNC path of the shared `Claim_Calling_Tracker.xlsx`, e.g. `\\CorpShare\Claims\Claim_Calling_Tracker.xlsx`.
4. In the `Users` sheet of the central DB, replace the sample rows with real Windows usernames (`Environ("USERNAME")` values) and mark exactly the right people as `Admin`.

## 3. Build the 4 forms (UserForm, not HTML — see note below)
Insert → UserForm for each. Modern flat look for all forms: white background, `Segoe UI` 10pt font, one accent color (e.g. `#1F4E78`) for buttons/headers, no default gray Windows controls where avoidable (use flat command buttons, `BackStyle = Opaque`, no 3D borders).

### frmAddClaim
Controls: `txtClaimID` (TextBox), `txtCallerLocation` (TextBox), `cmdSave` (Button), `cmdCancel` (Button).
```vb
Private Sub cmdSave_Click()
    If Trim(txtClaimID.Value) = "" Or Trim(txtCallerLocation.Value) = "" Then
        MsgBox "Claim ID and Caller Location are required.", vbExclamation
        Exit Sub
    End If
    If AddClaim(Trim(txtClaimID.Value), Trim(txtCallerLocation.Value)) Then
        MsgBox "Claim added.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
```

### frmUpdateClaim (the form every caller uses daily)
Controls: `txtClaimID` (TextBox), `cboStatus` (ComboBox, RowSource list: Pending/Closed — set `.List = Array("Pending","Closed")` in `Initialize`), `txtComment` (TextBox, MultiLine), `cmdSubmit` (Button).
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

### frmAdminLocation (Admin-only, shown only if IsCurrentUserAdmin)
Controls: `txtClaimID`, `txtNewLocation`, `cmdUpdate`.
```vb
Private Sub cmdUpdate_Click()
    If ChangeCallerLocation(Trim(txtClaimID.Value), Trim(txtNewLocation.Value)) Then
        MsgBox "Location updated.", vbInformation
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
`modMain.ShowAddClaimForm`, `ShowUpdateClaimForm`, `ShowAdminLocationForm`, `ShowClaimHistoryViewer`.

## 5. Distribute
Save `CallTrail_App.xlsm`, digitally sign it or have IT trust the location, and send a copy to each caller. Everyone points at the same `DB_PATH`.

---

## Important platform-constraint notes (read before rollout)

**"HTML" UI ask, honestly assessed:** VBA's only native way to show real HTML is the `WebBrowser` control, which runs on the Internet Explorer engine — deprecated by Microsoft, and on newer Windows/Office builds it can be missing, blocked by group policy, or need the IE11 emulation registry key set per machine. It's workable in a controlled corporate image, but it's the least reliable part of this design and the first thing to break as machines get patched. Given your stated preference for accuracy over optimism on platform constraints, I built the primary app on styled UserForms instead — 100% supported, and looks clean if you keep it flat/one-accent-color as above. If you still want the HTML route as a stretch goal, the safer variant is an **HTA file launched via `Shell()`** that talks back to Excel through a COM automation object — separate from the WebBrowser-control approach, more control over styling, but is its own project. I can build this as an add-on later if you decide it's worth it.

**The real risk in this design is Excel-as-shared-database concurrency**, not the UI. A single `.xlsx` has no row-level locking: two people saving in the same second will occasionally cause a "file in use" error, and — worse — a Save from user B *after* user A opened but before user A wrote could silently overwrite user A's unrelated changes elsewhere in the sheet, because a full-workbook Save writes the whole file, not just the changed cells. The `OpenCentralDB`/`CloseCentralDB` retry pattern above minimizes the collision window (open → single write → save → close, every time) and will be fine for a small team (roughly under 10–15 concurrent users making occasional calls). If call volume or headcount grows, or if simultaneous-write errors start showing up, the next step is swapping `modDataAccess` for calls to a real database (Access, or SQL Server via ADO) — because all data access is isolated in that one module, the forms and business logic above don't need to change at all.

**CallingAttempted** currently counts every logged call (Pending or Closed) toward the total. If you instead want it to count only calls up to and including the closing call, that's a one-line change in `LogCallAndUpdateStatus` — let me know which you want.
