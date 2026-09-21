# CallTrail — Setup Guide

## 1. Two files, two roles
- **Claim_Calling_Tracker.xlsx** → put this on the shared drive ONCE. It is pure data (Claims, History, Users, Config sheets). No macros live here.
- **CallTrail_App.xlsm** → the file you build below, using the 4 `.bas` modules provided. Distribute a COPY of this to every caller/admin's machine.

This split is deliberate: if the data file ever got macros in it, every user opening it from the share would need to enable macros on a file others are actively writing to — riskier and slower. Keep logic local, data central.

## Database sheets
`Claims`, `History`, `ArchivedClaims`, `ArchivedHistory`, `Users`, `Config`.

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
   - `modBulkImport.bas`
   - `modClaimFilter.bas`
   - `modPageHelpers.bas`
   - `modDownload.bas`
   - `modRepair.bas` (one-off repair tool — see section 7f)
   - For the tabbed shell (section 7): `IPage.cls`, `clsAppContext.cls`, `clsFormStyler.cls`, `clsArchiveService.cls`, `clsPageHome.cls`, `clsPageLogCall.cls`, `clsPageSearch.cls`, `clsPageView.cls`, `clsPageAddClaim.cls`, `clsPageAdmin.cls`
   - **Optional, object-oriented layer** (see section 2a below): `clsClaim.cls`, `clsHistoryEntry.cls`, `clsClaimRepository.cls`

   Note: `modBulkImport.bas` uses `clsClaimRepository`, so import the three `.cls` files even if you otherwise stick to the procedural layer.
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

## 3. Build the core forms (UserForm, not HTML — see note below)
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
Simplest: put buttons on Sheet1 (Insert → Shapes or Form Controls), right-click → Assign Macro →

**Recommended — one button for the whole app:**
`modMain.ShowApp` opens the tabbed shell (section 7). The individual entry points below still work if you'd rather wire single-purpose buttons.

**Single-claim / daily use:**
`modMain.ShowAddClaimForm`, `ShowUpdateClaimForm`, `ShowAdminSiteForm`, `ShowClaimHistoryViewer`.

**Browse / inspect (see sections 5 and 5a):**
`modMain.ShowViewClaims`, `modMain.ShowSearchClaim`.

**Bulk claim entry:**
| Button label | Macro | Notes |
|---|---|---|
| Setup Bulk Sheet | `modMain.BulkSheetSetup` | Run once to create the sheet |
| Load From File | `modMain.BulkLoadFromFile` | Optional — pulls a CSV/xlsx into the sheet |
| Check Claims | `modMain.BulkCheck` | Optional dry run — flags problems, changes nothing |
| **Add Claims to Database** | `modMain.BulkAddClaims` | The main action |

## 4a. How bulk claim entry works
`modBulkImport.bas` adds a local **bulkClaimAdd** sheet to the app file. Nothing touches the central DB until the user clicks *Add Claims to Database*.

The sheet has six columns: the five mandatory fields the user fills (ClaimID, ClaimSite, ClaimProviderName, ClaimQuery, ClaimCreationDate), plus a **ValidationError** column the app writes into.

**One click does everything.** *Add Claims to Database* walks every row and:
- **Valid rows** → written to the central database, then **deleted from the sheet**.
- **Invalid rows** → **left exactly where they are**, with the reason written next to them in the ValidationError column and the cell shaded red.

So the sheet always shows just the work that still needs attention. The user corrects those rows and clicks the same button again — corrected rows go in and disappear, anything still wrong stays flagged.

**What gets flagged:** missing fields, a date that isn't a valid date, a creation date in the future, a duplicate ClaimID within the sheet itself, and a ClaimID that already exists in the database. Multiple problems on one row are listed together, separated by semicolons.

**Adding your own validation rules:** they all live in one place — the `ValidateRow` function in `modBulkImport.bas`. Add a check there and both the dry run and the real import pick it up, so the two can never disagree about what's valid.

### Two implementation notes worth knowing

**Rows are written in one batch, not one at a time.** The user experiences this as "check each row, move it, delete it", but internally all valid rows go to the shared workbook in a single write via `clsClaimRepository.AddClaimsBulk`. Writing row-by-row would mean opening and saving the shared file once per claim — 500 claims would be 500 network round trips and 500 chances to collide with another user mid-save. The visible outcome is identical; it's just far faster and safer on a share drive.

**Deletion only happens after a confirmed successful insert.** If the database insert returns anything unexpected — for example some IDs were skipped because another user added them between validation and import — the rows are left in the sheet rather than deleted. Losing a claim that never reached the database is much worse than leaving a duplicate row for the user to review, so the code errs that way deliberately.

## 5. frmViewClaims — browse, filter and inspect claims

Insert → UserForm, name it **frmViewClaims**, then paste in the contents of `frmViewClaims_code.txt`. Also import `modClaimFilter.bas` (the filtering logic lives there, not in the form).

Wire a button to `modMain.ShowViewClaims`.

### Controls to place (names must match exactly)

**Filter bar — across the top**
| Control | Name | Notes |
|---|---|---|
| ComboBox | `cboStatus` | Populated in code: All / Pending / Closed |
| ComboBox | `cboDateField` | Which date the range applies to |
| TextBox | `txtDateFrom` | Blank = no lower bound |
| TextBox | `txtDateTo` | Blank = no upper bound |
| TextBox | `txtSearch` | Free-text across ID, site, provider, query |
| CommandButton | `cmdApply` | Caption "Apply" |
| CommandButton | `cmdClear` | Caption "Clear Filters" |
| CommandButton | `cmdRefresh` | Caption "Refresh" |
| CommandButton | `cmdClose` | Caption "Close" |

Add plain Labels next to each (Status:, Date field:, From:, To:, Search:) — those aren't referenced in code, so name them anything.

**Claims list — left/centre, make it the largest control on the form**
| Control | Name |
|---|---|
| ListBox | `lstClaims` |

Columns are configured in code: ClaimID, Site, Provider, Status, Attempt, CreationDate. Scrolling is native — just size the ListBox tall.

**Detail panel — right side (all Labels, empty Caption at design time)**
`lblClaimID`, `lblSite`, `lblProvider`, `lblQuery`, `lblStatus`, `lblAttempt`, `lblCreated`, `lblUpdatedSite`, `lblInsertedOn`, `lblInsertedBy`, `lblLastUpdated`, `lblLastUpdatedBy`, `lblLastComment`, `lblClosedOn`, `lblDaysOpen`

Put a static caption Label beside each ("Claim ID:", "Site:", …). For `lblQuery` and `lblLastComment`, set `WordWrap = True` and give them height — those hold free text.

**Call history for the selected claim — below the detail panel**
| Control | Name |
|---|---|
| ListBox | `lstHistory` |

**Status bar — bottom**
| Control | Name | Notes |
|---|---|---|
| Label | `lblStatusBar` | Messages and warnings |
| Label | `lblCount` | Shows "shown of total" |

### How it behaves
- **Opens showing everything.** No filter values = all records, as you asked.
- **Filters combine.** Status and date range apply together; leaving either blank drops that criterion.
- **Selecting a row** fills the detail panel with every column of that claim, and loads its full call history underneath.
- **Clear Filters** returns to the full list.

### Design decisions worth knowing

**Claims load once, filtering happens in memory.** The form reads the shared database on open and on Refresh only. Every filter change then works on that in-memory copy, so filtering is instant and — more importantly — the shared workbook isn't being opened and locked repeatedly while someone browses. The trade-off: claims added by other callers *after* you opened the form won't appear until you hit **Refresh**. That's the right trade for a browse screen, but it's why the Refresh button exists rather than being optional.

**Call history is fetched per selected claim, not preloaded.** Loading every claim's history upfront would mean pulling the entire History sheet into memory just to show one claim's calls.

**"Which date?" is an explicit choice.** Claims carry four dates (creation, insertion, last updated, closed) and "filter by date" means something different for each. Rather than guess, `cboDateField` lets the user pick. Note that filtering on *Closed Date* or *Last Updated* excludes claims where that field is still blank — an unworked claim genuinely has no last-updated date, so it can't fall inside a range.

**The list is capped at 10,000 rows.** A UserForm ListBox degrades badly well before Excel's row limit. Past the cap the form shows the first 10,000 and tells the user to narrow the filter, rather than freezing. If you routinely need to look at more than that at once, a worksheet-based view (native AutoFilter, no row cap) is the better tool — worth switching to if you hit this regularly.

## 5a. frmSearchClaim — look up one claim, view history, correct details

Insert → UserForm named **frmSearchClaim**, paste in `frmSearchClaim_code.txt`. Wire a button to `modMain.ShowSearchClaim`.

### Controls to place (names must match exactly)

**Search bar — top**
| Control | Name | Notes |
|---|---|---|
| TextBox | `txtSearchID` | Claim ID to look up; Enter key also searches |
| CommandButton | `cmdSearch` | Caption "Search" |

**Read-only fields — Labels, empty Caption at design time**
`lblClaimID`, `lblStatus`, `lblAttempt`, `lblInsertedOn`, `lblInsertedBy`, `lblLastUpdated`, `lblLastUpdatedBy`, `lblLastComment`, `lblClosedOn`, `lblDaysOpen`

Set `WordWrap = True` on `lblLastComment`.

**Editable fields — TextBoxes**
| Control | Name | Notes |
|---|---|---|
| TextBox | `txtSite` | |
| TextBox | `txtProvider` | |
| TextBox | `txtQuery` | `MultiLine = True`, give it height |
| TextBox | `txtCreationDate` | |
| TextBox | `txtUpdatedSite` | Auto-disabled for non-Admins |
| Label | `lblAdminNote` | Caption "Updated Site can only be changed by an Admin." Code shows/hides it |

**Call history — read-only**
| Control | Name | Notes |
|---|---|---|
| ListBox | `lstHistory` | 4 columns, configured in code |
| Label | `lblHistoryCount` | Shows "N call(s) logged" |

**Actions — bottom**
| Control | Name | Caption |
|---|---|---|
| CommandButton | `cmdSave` | "Save Changes" |
| CommandButton | `cmdRevert` | "Discard Changes" |
| CommandButton | `cmdClose` | "Close" |
| Label | `lblStatusBar` | Messages |

### What's editable, and why the rest isn't

| Field | Editable | Reason |
|---|---|---|
| ClaimSite, ClaimProviderName, ClaimQuery, ClaimCreationDate | Yes | These are what a user typed at insert time, so a user can correct them |
| ClaimUpdatedSite | Admin only | Matches the existing admin rule; non-admins see the value greyed out rather than hidden |
| **ClaimStatus, Attempt, ClaimClosedDate** | **No** | Maintained by `LogCallAndUpdateStatus` so status, attempt count and the call trail stay in sync. A textbox that set status directly would let someone close a claim with no call behind it, and `Attempt` would stop meaning anything. Status changes belong on the update/call form — and reopening a closed claim is Admin-only, via `ReopenClaim` (section 7c) |
| ClaimID | No | It's the key that links the claim to its history |
| ClaimInsertionDate, ClaimInsertedBy, LastUpdatedDate, LastUpdatedBy | No | Audit fields — an audit trail you can type over isn't an audit trail |
| **Call history** | **No** | Append-only. Rows are added by logging a call, never edited — that's what makes the trail trustworthy |

### Behaviour details
- Editable boxes stay **locked until a claim is loaded**, so nobody types into a blank form and wonders why Save does nothing.
- **Save is disabled until something actually changes**, and re-enables on the first edit.
- Closing with unsaved changes (button or the X) prompts first.
- After a successful save the claim is **re-read from the database** rather than trusting the form's copy, so the audit fields on screen show what was really written.
- A non-admin who edits Updated Site doesn't get their whole save rejected — that one field is skipped and the status bar says so, and their other corrections still go through.

## 5b. Home page, registration and guest lockdown

`clsPageHome.cls` is the landing tab. It shows who's signed in, whether they're registered, and what the app does.

### Registration model

On startup the app looks up the Windows username in the `Users` sheet.

| Found? | Role | Access |
|---|---|---|
| Yes, `Role = Admin` | Admin | Everything including the Admin tab |
| Yes, any other role | User | Everything except Admin |
| **No** | **Guest** | **Home only — every other tab greyed and labelled "(locked)"** |

The Home tab names the user explicitly: *"User 'jdoe' is not registered in the database."* That matters on shared machines and with multiple domain accounts — "you are not registered" leaves people guessing which account the app actually saw.

**This changes previous behaviour.** `GetCurrentUserRole` used to default an unknown user to `"User"`, meaning anyone who opened the file could add and edit claims. Unknown users are now Guests.

A **Re-check Registration** button clears the cached lookup so someone just added to the `Users` sheet can confirm it without restarting Excel. Unlocking the tabs still needs a restart — tab state is applied once at startup, and rebuilding it live would mean re-running arrangement for pages that may already hold unsaved input.

**Locking is convenience, not security.** Anyone who can open the VBE can re-enable a tab. The real enforcement is that every write re-checks the role server-side.

### Add Claim site autopopulation

`ClaimSite` pre-fills from the signed-in user's `DefaultLocation` in the `Users` sheet, and **stays editable**. A caller can legitimately enter a claim for another site; locking the field would push a routine action to an admin.

If `DefaultLocation` is blank or the column is missing, the field is simply empty — nothing breaks.

### Controls on `pgHome`

| Name | Type |
|---|---|
| `lblHomeWelcome` | Label (header) |
| `lblCapHomeUser`, `lblCapHomeRole`, `lblCapHomeReg` | Labels (captions) |
| `lblHomeUserName`, `lblHomeRole`, `lblHomeRegStatus` | Labels (values) |
| `lblHomeGuestNotice` | Label |
| `lblHomeFeaturesHeader` | Label (header) |
| `lblHomeFeature1` … `lblHomeFeature6` | Labels |
| `imgHomeBanner` | **Image** — set its Picture in the designer |
| `cmdHomeStart`, `cmdHomeRefreshUser` | CommandButtons (optional) |

`pgHome` must be the **first** page on the MultiPage, and registered first in `RegisterPages`.

### The banner image

`imgHomeBanner` shows a static picture — set its `Picture` property in the designer.

**If you were expecting an animated GIF here, it won't work.** MSForms `Image` controls render only the first frame of a GIF and stop; there is no property that changes this. It's a limitation of the control, not a setting you're missing. The two workarounds, if you ever want motion:

- **Frame swapping** on an `Application.OnTime` tick. Reliable, but UserForms have no timer control and `OnTime` won't tick below about a second, so it only suits deliberate, slow effects.
- **A WebBrowser control**, which genuinely animates GIFs but runs on the deprecated Internet Explorer engine — often absent, policy-blocked, or needing a per-machine registry key on current builds.

A static image avoids both problems and renders identically everywhere.

### Cross-tab navigation

The Home tab's **Start Calling** button switches to the Log Call tab by posting `GOTO:LogCall` through the status channel, which the shell interprets. Controllers stay decoupled from the shell rather than holding a reference to the form and driving it directly.

## 7. frmMain — the tabbed application shell

This is the recommended way to run CallTrail. One window, five tabs, shared data cache.

### Build it

1. Import the new class modules: `IPage.cls`, `clsAppContext.cls`, `clsPageHome.cls`, `clsPageLogCall.cls`, `clsPageSearch.cls`, `clsPageView.cls`, `clsPageAddClaim.cls`, `clsPageAdmin.cls`, and `modPageHelpers.bas`.
2. Insert a UserForm named **frmMain**. Paste in `frmMain_code.txt`.
3. Add these three controls directly on the form (not inside the MultiPage):

| Control | Name |
|---|---|
| MultiPage | `MultiPage1` |
| Label | `lblStatusBar` |
| CommandButton | `cmdClose` (Caption "Close") |

4. On `MultiPage1`, create five pages and set their **Name** property (the Caption is set automatically in code):

| Page name | Tab shows as |
|---|---|
| `pgLogCall` | Log Call |
| `pgSearch` | Search & Edit |
| `pgView` | View Claims |
| `pgAdd` | Add Claim |
| `pgAdmin` | Admin |

5. Place the controls below on each page. **Names must match exactly** — the controllers look them up by name, and a mismatch raises an error naming the control and the tab.

#### Tab: Log Call (`pgLogCall`) — 10 controls

| Control name | Type | Notes |
|---|---|---|
| `txtLogClaimID` | TextBox |  |
| `cmdLogFind` | CommandButton |  |
| `cboLogStatus` | ComboBox |  |
| `txtLogComment` | TextBox | MultiLine |
| `cmdLogSubmit` | CommandButton |  |
| `lstLogHistory` | ListBox | Columns set in code |
| `lblLogClaimInfo` | Label |  |
| `lblLogCurrentStatus` | Label |  |
| `lblLogAttempt` | Label |  |
| `lblLogLastComment` | Label | Optional — omit if you don't want it; MultiLine |

#### Tab: Search & Edit (`pgSearch`) — 21 controls

| Control name | Type | Notes |
|---|---|---|
| `txtSrchID` | TextBox |  |
| `cmdSrchFind` | CommandButton |  |
| `txtSrchSite` | TextBox |  |
| `txtSrchProvider` | TextBox |  |
| `txtSrchQuery` | TextBox | MultiLine |
| `txtSrchCreationDate` | TextBox |  |
| `txtSrchUpdatedSite` | TextBox |  |
| `cmdSrchSave` | CommandButton |  |
| `cmdSrchRevert` | CommandButton |  |
| `lstSrchHistory` | ListBox | Columns set in code |
| `lblSrchClaimID` | Label |  |
| `lblSrchStatus` | Label |  |
| `lblSrchAttempt` | Label |  |
| `lblSrchInsertedOn` | Label |  |
| `lblSrchInsertedBy` | Label |  |
| `lblSrchLastUpdated` | Label |  |
| `lblSrchLastUpdatedBy` | Label |  |
| `lblSrchClosedOn` | Label |  |
| `lblSrchDaysOpen` | Label |  |
| `lblSrchHistoryCount` | Label |  |
| `lblSrchAdminNote` | Label | Optional — omit if you don't want it |

#### Tab: View Claims (`pgView`) — 26 controls

| Control name | Type | Notes |
|---|---|---|
| `cboViewStatus` | ComboBox |  |
| `cboViewDateField` | ComboBox |  |
| `txtViewDateFrom` | TextBox |  |
| `txtViewDateTo` | TextBox |  |
| `txtViewSearch` | TextBox |  |
| `cmdViewApply` | CommandButton |  |
| `cmdViewClear` | CommandButton |  |
| `cmdViewRefresh` | CommandButton |  |
| `lstViewClaims` | ListBox | Columns set in code |
| `lstViewHistory` | ListBox | Columns set in code |
| `lblViewCount` | Label |  |
| `lblViewClaimID` | Label |  |
| `lblViewSite` | Label |  |
| `lblViewProvider` | Label |  |
| `lblViewQuery` | Label | MultiLine |
| `lblViewStatus` | Label |  |
| `lblViewAttempt` | Label |  |
| `lblViewCreated` | Label |  |
| `lblViewUpdatedSite` | Label |  |
| `lblViewInsertedOn` | Label |  |
| `lblViewInsertedBy` | Label |  |
| `lblViewLastUpdated` | Label |  |
| `lblViewLastUpdatedBy` | Label |  |
| `lblViewLastComment` | Label | MultiLine |
| `lblViewClosedOn` | Label |  |
| `lblViewDaysOpen` | Label |  |

#### Tab: Add Claim (`pgAdd`)

Place these in the designer as usual; `clsPageAddClaim` positions and styles them at run time (see section 7a), so their designer position and size don't matter. Note that tab also needs caption labels — full list in section 7a.

| Control name | Type | Notes |
|---|---|---|
| `txtAddClaimID` | TextBox |  |
| `txtAddSite` | TextBox |  |
| `txtAddProvider` | TextBox |  |
| `txtAddQuery` | TextBox | MultiLine |
| `txtAddCreationDate` | TextBox |  |
| `cmdAddSave` | CommandButton |  |
| `cmdAddClear` | CommandButton |  |
| `lblAddInfo` | Label | Optional — omit if you don't want it |

#### Tab: Admin (`pgAdmin`) — 9 controls

| Control name | Type | Notes |
|---|---|---|
| `txtAdmClaimID` | TextBox |  |
| `txtAdmNewSite` | TextBox |  |
| `cmdAdmChangeSite` | CommandButton |  |
| `cmdAdmBulkSetup` | CommandButton |  |
| `cmdAdmBulkLoad` | CommandButton |  |
| `cmdAdmBulkCheck` | CommandButton |  |
| `cmdAdmBulkProcess` | CommandButton |  |
| `lblAdmUser` | Label | Optional — omit if you don't want it |
| `lblAdmBulkInfo` | Label | Optional — omit if you don't want it |

Add plain caption Labels beside the fields ("Claim ID:", "Site:", …). Those aren't referenced in code, so name them anything.

### How the architecture works

**`IPage`** is an interface — VBA has no inheritance, but `Implements` gives you contracts. Every tab controller implements three members:

| Member | Called | Purpose |
|---|---|---|
| `InitPage` | Once, at startup | Bind controls to the class's `WithEvents` variables |
| `OnActivate` | Every tab switch | Refresh data — *not* in InitPage |
| `PageTitle` | As needed | Tab caption and error messages |

`frmMain` holds a `Collection` of `IPage` and never knows which concrete class is behind a tab.

**`clsAppContext`** is the shared state — one repository, one cached claim list, one cached role check, handed to every controller. This is what makes the tabbed shell *cheaper* than five separate forms rather than more expensive: without it, each screen would hit the shared workbook independently. Any tab that writes calls `InvalidateClaims`, and the next tab to need claims reloads. Nothing reloads speculatively.

**`WithEvents` in a class module** is the mechanism that makes this modular at all. A class can hold event handlers for MSForms controls, so each tab's ~20 controls and all their event code live in that tab's own file instead of one 1,500-line form module.

### Adding a sixth tab

1. Add a Page to `MultiPage1` in the designer, with its controls.
2. Write `clsPageXxx` that `Implements IPage`.
3. Add **one** `AddPage` line to `RegisterPages` in `frmMain`, in tab order.

Nothing else changes — not the shell, not the other five tabs.

### Decisions worth knowing

**Bulk claim entry stays on the worksheet.** The Admin tab has buttons that *launch* it, but the actual work happens on the `bulkClaimAdd` sheet. Pasting 200 rows is a worksheet job; a form control would be worse at the thing the user actually needs to do.

**The Admin tab being hidden is convenience, not security.** Non-admins don't see it, but every admin action re-checks the role in the data layer — anyone who can open the VBE could unhide a tab. The enforcement that matters is in `clsClaimRepository`.

**Every control on every tab is created at form load.** VBA's MultiPage doesn't lazy-create hidden pages, so all ~74 controls instantiate when the form opens — expect a brief pause on a slower machine. What we *can* avoid is the expensive part, and do: no tab touches the database until you actually open it, and the shared cache means switching tabs doesn't re-read the shared file unless something invalidated it.

**Status changes still only flow through `LogCallAndUpdateStatus`.** The Search & Edit tab can correct site, provider, query and creation date, but not status or attempt count — same reasoning as before, those three have to move together or `Attempt` stops meaning anything.


## 7a. Styling and positioning page controls in code

You place the controls in the VBE designer; **`clsFormStyler`** snaps them into an aligned grid and applies consistent fonts, colours and borders at run time. Position and size in the designer don't matter — drop them roughly anywhere.

**`clsPageAddClaim.cls` is the worked example.**

### Why arrange in code rather than the designer

- Every tab lines up identically, because they share one engine. Hand-placed controls drift a few points and it shows.
- Re-space the whole app by changing `LabelWidth` or `RowGap` once, instead of nudging 70 controls.
- The layout is readable code, so it diffs in version control. A `.frm`'s companion `.frx` is binary.
- Styling stays consistent when someone adds a control later and forgets what font the others used.

### The three jobs, kept separate

| Method | Job |
|---|---|
| `BindControls` | Object references, so the class can talk to the controls. `WithEvents` ones get wired here |
| `ArrangeUI` | **The design** — where everything sits and how it looks |
| Event handlers | Behaviour. Untouched by layout changes |

That separation is the point: re-lay-out the whole tab by editing one method, without going near the save logic.

### What `ArrangeUI` looks like

```vb
Private Sub ArrangeUI()
    Dim ui As New clsFormStyler
    ui.Init m_pg, IPage_PageTitle, startLeft:=14, startTop:=12

    ui.LabelWidth = 110
    ui.FieldWidth = 250
    ui.RowHeight = 18
    ui.RowGap = 6

    ui.Header "lblAddHeader", "New Claim Details"

    ui.Row "lblCapAddClaimID", "txtAddClaimID", "Claim ID:"
    ui.Row "lblCapAddSite", "txtAddSite", "Claim Site:"
    ui.Row "lblCapAddProvider", "txtAddProvider", "Provider Name:"
    ui.Row "lblCapAddCreationDate", "txtAddCreationDate", "Creation Date:", fieldWidth:=110
    ui.Row "lblCapAddQuery", "txtAddQuery", "Claim Query:", fieldHeight:=60

    ui.Gap 4
    ui.BeginButtonRow
    ui.Button "cmdAddSave", "Save Claim", 90
    ui.Button "cmdAddClear", "Clear", 70
    ui.EndButtonRow

    ui.Gap
    ui.Note "lblAddInfo", "Status starts as Pending...", height:=30
End Sub
```

Reads top-to-bottom like the screen. Move a field by moving its line.

### Controls to place on the Add Claim page

| Name | Type | Referenced in |
|---|---|---|
| `txtAddClaimID` | TextBox | Bind + Arrange |
| `txtAddSite` | TextBox | Bind + Arrange |
| `txtAddProvider` | TextBox | Bind + Arrange |
| `txtAddQuery` | TextBox | Bind + Arrange |
| `txtAddCreationDate` | TextBox | Bind + Arrange |
| `cmdAddSave` | CommandButton | Bind + Arrange |
| `cmdAddClear` | CommandButton | Bind + Arrange |
| `lblAddHeader` | Label | Arrange only |
| `lblCapAddClaimID` | Label | Arrange only |
| `lblCapAddSite` | Label | Arrange only |
| `lblCapAddProvider` | Label | Arrange only |
| `lblCapAddQuery` | Label | Arrange only |
| `lblCapAddCreationDate` | Label | Arrange only |
| `lblAddInfo` | Label | Arrange only — optional |
| `lblAddHint` | Label | Arrange only — optional |

**Caption labels need names here.** That's the one cost of this approach: the styler positions them, so it has to find them. Use a consistent prefix (`lblCap...`) and they stay out of your way. Captions are set in `ArrangeUI`, so leave the designer Caption blank if you like.

Labels passed to `Note` are optional — if you don't create them, they're skipped silently rather than erroring.

### `clsFormStyler` reference

Tunables, set after `Init`: `LabelWidth`, `FieldWidth`, `RowHeight`, `RowGap`, `SectionGap`, `FontName`, `FontSize`, `HeaderColor`, `NoteColor`, `EditableBackColor`, `ReadOnlyBackColor`.

| Method | Does |
|---|---|
| `Init container, title, [startLeft], [startTop]` | Begins the layout |
| `Header name, [caption]` | Bold coloured section heading |
| `Row captionName, fieldName, [caption], [fieldWidth], [fieldHeight], [readOnlyLook]` | Caption label + input control on one row |
| `ValueRow captionName, valueLabelName, [caption], [w], [h]` | Caption + read-only value Label |
| `FullWidth name, width, height` | Control spanning the block, no caption — typically a ListBox |
| `Note name, [caption], [height], [width]` | Hint label. Silently skipped if the control doesn't exist |
| `BeginButtonRow` / `Button name, [caption], [width]` / `EndButtonRow` | Buttons left-to-right |
| `Gap [points]` | Vertical space (defaults to `SectionGap`) |
| `NewColumn left, [resetTop]` | Start a second column — e.g. a detail panel right of a list |
| `StyleOnly name, [readOnlyLook]` | Style without moving — for anything you positioned by hand |
| `CurrentTop` (get/let), `CurrentLeft` | Read or force the y-position |

`Row` and `FullWidth` style by control type — TextBox gets a border and white/grey background, ComboBox is forced to drop-down-list so typed junk can't get in, ListBox gets a border. Passing `readOnlyLook:=True` greys and locks a TextBox.

### Two-column layouts

For a list on the left and a detail panel on the right:

```vb
ui.FullWidth "lstViewClaims", 340, 220

ui.NewColumn 370
ui.Header "lblViewDetailHeader", "Claim Detail"
ui.ValueRow "lblCapViewClaimID", "lblViewClaimID", "Claim ID:"
ui.ValueRow "lblCapViewStatus", "lblViewStatus", "Status:"
```

### Styling the shell (frmMain itself)

The MultiPage, status bar and Close button are arranged the same way, by an `ArrangeShell` method in `frmMain`. Size tunables sit as constants at the top of the module:

```vb
Private Const FORM_WIDTH   As Single = 780
Private Const FORM_HEIGHT  As Single = 560
Private Const OUTER_MARGIN As Single = 8
Private Const TITLE_HEIGHT As Single = 22
Private Const FOOTER_HEIGHT As Single = 30
Private Const TAB_WIDTH    As Single = 96    ' 0 = auto-width tabs
```

```vb
Private Sub ArrangeShell()
    Dim ui As New clsFormStyler
    ui.Init Me, "Main Form"

    ui.SizeForm Me, FORM_WIDTH, FORM_HEIGHT, "CallTrail"
    ui.TitleBar "lblTitle", "CallTrail - Claim Calling Tracker", TITLE_HEIGHT, OUTER_MARGIN

    ui.Fill "MultiPage1", OUTER_MARGIN, FOOTER_HEIGHT, TitleReserve()
    ui.StyleMultiPage "MultiPage1", TAB_WIDTH

    ui.AnchorBottomLeft "lblStatusBar", FORM_WIDTH - 140, , OUTER_MARGIN
    ui.StyleStatusBar "lblStatusBar"
    ui.AnchorBottomRight "cmdClose", 80, , OUTER_MARGIN
End Sub
```

Change `FORM_WIDTH` and everything re-flows — the MultiPage refills, the status bar re-stretches, the Close button stays pinned bottom-right.

**Shell methods** (these size and anchor rather than stacking rows):

| Method | Does |
|---|---|
| `SizeForm frm, width, height, [caption]` | Sizes the form and centres it on screen |
| `Fill name, [margin], [reserveBottom], [reserveTop]` | Control fills the container, leaving room for a footer/title |
| `AnchorBottomLeft name, width, [height], [margin]` | Pins bottom-left — the status bar |
| `AnchorBottomRight name, width, [height], [margin]` | Pins bottom-right — the Close button |
| `TitleBar name, caption, [height], [margin]` | Banner across the top. Skipped silently if the label doesn't exist |
| `StyleMultiPage name, [tabFixedWidth]` | Font and tab appearance. Fixed-width tabs stop them jumping about as captions change length |
| `StyleStatusBar name` | Subdued styling so it reads as chrome, not content |

**Controls on frmMain itself:** `MultiPage1`, `lblStatusBar`, `cmdClose`, and optionally `lblTitle`. Their designer position and size don't matter. Leave out `lblTitle` and the banner is skipped and its space isn't reserved.

**`ArrangeShell` runs first in `UserForm_Initialize`**, before `RegisterPages`. That order matters: the MultiPage must be at its final size before any page arranges itself inside it, or a page sizing controls against the page width would lay out against the designer's size instead.

### Sizing page controls relative to the container

Rather than hardcoding widths in a page's `ArrangeUI`, ask the styler how much room it has — then a change to `FORM_WIDTH` flows through to the tabs instead of leaving a ListBox stranded at its old size:

```vb
ui.FullWidth "lstViewClaims", ui.AvailableWidth * 0.45, 220
```

| Property | Returns |
|---|---|
| `AvailableWidth([margin])` | Usable width of the container |
| `AvailableHeight([margin])` | Usable height |
| `RemainingHeight([bottomMargin])` | Room left below the current row — for a list that should stretch to the bottom |

### Honest limits

**No WYSIWYG.** You run the form to see the result, and the first pass usually needs a couple of spacing tweaks. Worth it for repeating label/field rows — which is most of this app. For a screen with genuinely bespoke placement, position it in the designer and call `StyleOnly` to keep the look consistent.

**Converting the other four tabs is optional and mechanical.** Add `m_arranged`, add an `ArrangeUI`, call it from `OnActivate`. Binding, event handlers and business logic don't change at all. Do them one at a time.

**UserForms don't resize.** `FORM_WIDTH`/`FORM_HEIGHT` set a fixed window — MSForms has no native resize grip, and adding one needs Windows API calls. Pick a size that fits your smallest user's screen. The layout code means changing that size later is a one-line edit, not a redesign.

## 7b. Making the form look the same on every machine

A UserForm that looks right on your PC often looks wrong on a colleague's. There are three separate causes, and they need different fixes — most advice online conflates them.

| Cause | Symptom | Fix |
|---|---|---|
| **Font substitution** | Text overflows boxes, captions clip, fields misalign | Runtime calibration — handled in code, below |
| **Display scaling (DPI)** | Everything too big or too small at 125%/150% | Mostly Office's job; `Zoom` as a fallback |
| **Screen resolution** | Form doesn't fit, buttons off-screen | Design for the smallest screen. No code fixes this |

### Font substitution — the one you can actually fix

If the font you specified isn't installed, Windows silently swaps in another with different metrics. Same point size, different real width — so your carefully aligned grid falls apart even though every coordinate in your code is identical.

`clsFormStyler.Calibrate` measures how wide a known string *actually* renders on this machine, compares it to the baseline measured on yours, and scales the layout by the difference.

The shell calibrates once and shares the result:

```vb
' In frmMain.ArrangeShell
ui.Calibrate
m_ctx.UIScale = ui.ScaleFactor      ' every tab reuses this
```

```vb
' In each page's ArrangeUI
ui.LabelWidth = 110
ui.FieldWidth = 250
ui.UseScale m_ctx.UIScale           ' AFTER the tunables
```

**Order matters and the code enforces it.** `UseScale` scales whatever the tunables currently hold, so setting `LabelWidth` afterwards would put an unscaled number back. Set tunables first, then `UseScale`, then arrange. A double-apply guard stops a second call compounding the scale.

**Font size is deliberately not scaled.** The scale factor is derived *from* measuring text at that size, so the right response to a wider font is wider boxes, not smaller text.

**Re-baselining:** `REF_WIDTH` in `clsFormStyler` is the reference measured on a standard 96 DPI machine with Segoe UI installed. If your development machine differs, run `Calibrate`, read `MeasuredReferenceWidth`, and put that number in `REF_WIDTH`.

The ratio is clamped to 0.8–1.4. A wilder number means something unexpected happened, and trusting it blindly would produce a worse layout than doing nothing. If calibration fails for any reason it falls back to unscaled rather than failing to draw the form.

### Display scaling (DPI)

Current Office builds scale UserForms for you; older ones don't, which is where the classic "everything is tiny at 150%" comes from. You can't detect this reliably from VBA, so the practical approach is to test at 100%, 125% and 150% on a representative machine.

If a specific user's display scaling is wrong, `UserForm.Zoom` scales the whole form and every control uniformly:

```vb
Me.Zoom = 125
Me.Width = Me.Width * 1.25
Me.Height = Me.Height * 1.25
```

Text gets slightly soft, so treat it as a per-site fallback rather than a default. Storing the value in the `Config` sheet lets you set it per environment without editing code.

### Practical rules that prevent most problems

- **Always set `Font.Name` explicitly**, never inherit. The styler does this for every control — an unstyled control inherits whatever the form's default is, which differs by Office version.
- **Use a font that's genuinely everywhere.** Segoe UI ships with Windows Vista and later. Tahoma goes back further and is the safest choice for mixed or older estates.
- **Leave 15–20% slack in field widths.** A box sized exactly to its expected text has no room when metrics shift.
- **Avoid `AutoSize` on labels in a grid.** It makes each label a different width and the column stops being a column. The styler sets explicit widths instead. (`Calibrate` uses `AutoSize` on a throwaway label purely to measure, then deletes it.)
- **Design for the smallest screen you must support.** At 1366×768 — still common on corporate laptops — keep the form under about 1150×620 points to clear the taskbar and title bar.
- **Test on one machine that isn't yours** before rolling out. This catches in five minutes what a week of guessing won't.

## 7c. Reopening a closed claim (Admin only)

Closing is normally final — `LogCallAndUpdateStatus` refuses further calls on a Closed claim. Admins can reverse it with `ReopenClaim`.

### On the Admin tab

Two steps, deliberately: **Check Status** first, then **Reopen**. The Reopen button stays disabled until a claim has been confirmed Closed, so an admin can't fire a status change at a typo'd ID or at a claim that was already open.

Controls to add to `pgAdmin`:

| Name | Type | Notes |
|---|---|---|
| `txtAdmReopenID` | TextBox | Claim ID |
| `cmdAdmCheckStatus` | CommandButton | "Check Status" |
| `lblAdmReopenStatus` | Label | Shows provider, site, status, call count, closed date |
| `txtAdmReopenReason` | TextBox | MultiLine — mandatory |
| `cmdAdmReopen` | CommandButton | "Reopen Claim" — starts disabled |

There's also `modMain.ReopenClaimPrompt` if you'd rather wire a single button without the tabbed shell.

### What it does, and why

| Field | Effect | Reason |
|---|---|---|
| `ClaimStatus` | → Pending | Callers can log calls again |
| `ClaimClosedDate` | **Cleared** | Left in place, the claim reads as closed to every dashboard and to `DaysOpen`, which would silently report a stale time-to-close |
| `Attempt` | **Unchanged** | That column counts *calls*. A reopen is an administrative act, not a call — incrementing it would overstate calling effort |
| History | **A row is written** | Unlike the detail edits, this is a status change. A status change with no trace is exactly what an audit needs to catch |
| `LastUpdatedDate` / `By` / `LastComment` | Stamped with the admin and reason | |

A reason is mandatory and is enforced in the data layer, not just the form — so it can't be bypassed by calling `ReopenClaim` from anywhere else.

### One consequence worth knowing

The reopen writes a row into the call history, prefixed **`[REOPENED BY ADMIN]`**. That's the right call for auditing, but it means **history row count is no longer identical to call count**. `Attempt` remains the authoritative number of calls; the history is the full event trail, which now includes administrative events as well as calls.

If you'd rather keep the history strictly to calls, the alternative is a separate `ClaimAudit` sheet for administrative events. That's cleaner conceptually but means two places to look when investigating a claim. I'd stay with the single trail unless your audit requirements push the other way.

## 7d. Archiving old claims (Admin only)

Moves old claims and their call history out of `Claims`/`History` into `ArchivedClaims`/`ArchivedHistory`, keeping the working sheets small. Every operation in the app reads the full working sheets, so this is what keeps the app fast as years of data accumulate.

### Database sheets

Two new sheets, mirroring the source columns plus `ArchivedDate` and `ArchivedBy`. They have no Excel Table on them deliberately — a Table auto-expands and fights bulk row appends from VBA.

### On the Admin tab

| Control | Type | Notes |
|---|---|---|
| `cboAdmArchiveAge` | ComboBox | Older than 2/3/4/5/6 months, or Any age |
| `chkAdmArchiveClosedOnly` | CheckBox | **Defaults to ticked** |
| `cmdAdmArchivePreview` | CommandButton | "Preview" |
| `lblAdmArchivePreview` | Label | Count, oldest claim, warnings |
| `lstAdmArchiveClaims` | ListBox | **The claims that would be archived.** 6 columns, set in code |
| `lstAdmArchiveHistory` | ListBox | Call history of the selected claim. 4 columns, set in code |
| `lblAdmArchiveHistCount` | Label | "N history row(s) for X would move with it" |
| `cmdAdmArchiveRun` | CommandButton | "Archive Now" — starts disabled |

**Preview is mandatory, and it shows the actual records.** Run stays disabled until a preview has been taken, and changing any criterion disables it again *and clears the list* — a stale list beside changed criteria is worse than an empty one, because it looks current.

Preview lists every matching claim (ID, site, provider, status, attempts, creation date). **Click any row** and the lower list shows exactly the history rows that would move with it, with a count. So before committing, an admin can see both what goes and what goes with it.

The preview list is capped at 5,000 rows — a UserForm ListBox degrades badly well before Excel's limits. Past the cap the status bar says how many actually matched and notes that **the archive still moves all of them**; the run re-evaluates the criteria against the full list rather than using the displayed slice.

### Design decisions that protect your data

**Write first, delete second, with a save in between.** The archive rows are written and saved *before* anything is removed from the working sheets. If the process dies halfway, the worst case is rows appearing in both places — annoying, and recoverable. The other order loses them outright. The error message says exactly this if it fires, so whoever hits it knows to check `ArchivedClaims` before retrying.

**Claims and history move together.** Archiving a claim and leaving its history behind produces orphan rows belonging to a claim that's no longer on the sheet.

**Age is measured from the closed date, falling back to creation.** A claim raised 8 months ago but closed last week is still recent work. Measuring from creation alone would archive it the moment it closed.

**Working sheets are rewritten, not row-deleted.** Deleting hundreds of non-contiguous rows is slow and reshuffles everything underneath on each delete. The service clears the data range and writes the survivors back in one block.

**Admin-only, enforced in `clsArchiveService`** — not just by the tab being hidden.

### Two things to decide for yourself

**"Less than N months" — I've implemented "older than."** Your message said "less than 2 months", but archiving *recent* records would clear out exactly what your callers are still working on. If you genuinely meant the other direction, it's a single comparison operator in `clsArchiveService.Matches`.

**Closed-only defaults to on, and I'd leave it on.** Archiving a Pending claim removes live work from callers' view — the claim still needs chasing, but nobody can see it. The checkbox lets you override, and the preview shows an explicit warning when you do, but there are few good reasons to.

### There is no un-archive

Moving rows back has to be done by hand in Excel. That's deliberate — a one-click restore would need to handle ID collisions with claims added since, and getting that subtly wrong is worse than a manual process an admin does rarely and carefully. The confirmation dialog says so before anything moves.

If you find yourself needing to un-archive regularly, that's a signal the age threshold is too aggressive rather than a reason to build the feature.

## 7e. Downloading a snapshot of the database

`modDownload.bas` pulls a full copy of the central database into **this** workbook — one local sheet per source sheet — for offline analysis, pivot tables, ad-hoc reporting, or taking a snapshot before something risky like an archive run.

| Macro | Does |
|---|---|
| `modMain.DownloadDatabase` | All four sheets — the main one |
| `modDownload.DownloadClaimsOnly` | Just claims |
| `modDownload.DownloadHistoryOnly` | Just call history |
| `modDownload.DownloadArchivesOnly` | Both archive sheets |

Creates or replaces: `DL_Claims`, `DL_History`, `DL_ArchivedClaims`, `DL_ArchivedHistory`, plus a `DL_Info` sheet.

Wire a worksheet button to `modMain.DownloadDatabase`. Optionally add a `cmdAdmDownload` button to the Admin tab — it's bound with `BindOptional`, so leaving it off doesn't break the page.

### Three implementation choices worth knowing

**Opens read-only.** A download only reads, so it takes no write lock. A read/write open would block callers trying to log a call for as long as the copy runs — and on a large database over a share drive, that's exactly when you don't want to be holding the file. (`OpenCentralDBReadOnly` is new in `modUtils.bas`.)

**One open for all four sheets.** Opening per sheet would mean four round trips to the network share for nothing.

**Bulk array transfer, not `Range.Copy`.** Copy/paste carries formatting, uses the clipboard (which the user can clobber mid-run), and is far slower. Each sheet moves in one read and one write, values only.

### The `DL_Info` sheet

Records when the snapshot was taken, by whom, and the row counts. This exists because a downloaded sheet sitting in a workbook for a week looks exactly like one pulled five minutes ago — and sooner or later someone reports from stale data believing it's current. The timestamp makes that visible.

### Two limits

**Snapshots don't refresh, and don't write back.** Editing `DL_Claims` changes nothing in the central database. It's a copy, not a connection.

**A sheet can't hold more rows than Excel allows (~1,048,576).** If a source sheet exceeds that, the download stops with a message telling you to archive older records — rather than failing with a cryptic subscript error. In practice, hitting this means the archive threshold is far too lax.

## 7f. Fixing blank rows / "key is already associated" error

If you ran an archive with an earlier build and now see:

> Could not read claims. This key is already associated with an element of this collection.

…along with blank rows left behind in `Claims`, this is that bug. Both halves are fixed; here's what happened and what to do.

### What went wrong

`RewriteSheet` used `ClearContents`, which blanks cells but **leaves the rows in place**. `Claims` and `History` both carry Excel Tables (`tblClaims`, `tblHistory`), so the Table kept its old size with empty rows inside it. `GetAllClaims` then read those blank rows, each produced an empty `ClaimID`, and the second empty key collided in the keyed Collection.

### What's fixed

| Fix | Where |
|---|---|
| Surplus rows are now **deleted**, not just cleared | `clsArchiveService.RewriteSheet` |
| The Excel Table is **resized** to the new data extent | `clsArchiveService.ResizeListObject` |
| Blank rows are **skipped** on read | `clsClaimRepository.GetAllClaims` / `GetHistory` |
| Duplicate Claim IDs no longer crash the read | `GetAllClaims` — loads the first, reports the rest |

The last two matter independently of the archive bug: a reporting read is the wrong place to fall over because someone hand-edited a sheet.

### Cleaning up a database already affected

Import `modRepair.bas` and run `modMain.RepairBlankRows` **once**. Admin-only. It removes rows whose key column is empty, deletes the surplus, resizes the Table, and reports the counts. Safe to run again — it does nothing if there's nothing to fix.

Then reload the app.

### Worth checking

If the failed archive ran partway, claims could exist in **both** `Claims` and `ArchivedClaims`. The archive deliberately writes and saves before deleting, so nothing is lost — but do check for duplicates. The new duplicate-ID warning in `GetAllClaims` will tell you if any are still in the working sheet.

## 8. Distribute
Save `CallTrail_App.xlsm`, digitally sign it or have IT trust the location, and send a copy to each caller. Everyone points at the same `DB_PATH`.

---

## Important platform-constraint notes (read before rollout)

**"HTML" UI ask, honestly assessed:** VBA's only native way to show real HTML is the `WebBrowser` control, which runs on the Internet Explorer engine — deprecated by Microsoft, and on newer Windows/Office builds it can be missing, blocked by group policy, or need the IE11 emulation registry key set per machine. It's workable in a controlled corporate image, but it's the least reliable part of this design and the first thing to break as machines get patched. Given your stated preference for accuracy over optimism on platform constraints, I built the primary app on styled UserForms instead — 100% supported, and looks clean if you keep it flat/one-accent-color as above. If you still want the HTML route as a stretch goal, the safer variant is an **HTA file launched via `Shell()`** that talks back to Excel through a COM automation object — separate from the WebBrowser-control approach, more control over styling, but is its own project. I can build this as an add-on later if you decide it's worth it.

**The real risk in this design is Excel-as-shared-database concurrency**, not the UI. A single `.xlsx` has no row-level locking: two people saving in the same second will occasionally cause a "file in use" error, and — worse — a Save from user B *after* user A opened but before user A wrote could silently overwrite user A's unrelated changes elsewhere in the sheet, because a full-workbook Save writes the whole file, not just the changed cells. The `OpenCentralDB`/`CloseCentralDB` retry pattern above minimizes the collision window (open → single write → save → close, every time) and will be fine for a small team (roughly under 10–15 concurrent users making occasional calls). If call volume or headcount grows, or if simultaneous-write errors start showing up, the next step is swapping `modDataAccess` for calls to a real database (Access, or SQL Server via ADO) — because all data access is isolated in that one module, the forms and business logic above don't need to change at all.

**CallingAttempted** currently counts every logged call (Pending or Closed) toward the total. If you instead want it to count only calls up to and including the closing call, that's a one-line change in `LogCallAndUpdateStatus` — let me know which you want.
