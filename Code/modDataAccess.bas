Attribute VB_Name = "modDataAccess"
Option Explicit

' =====================================================================
' modDataAccess
' Every read/write against the shared database goes through here.
' UI code (forms) should NEVER touch worksheet cells directly.
'
' All column access is by header NAME, never a hard-coded number.
' Single-row operations resolve names via a header map built once per
' operation (BuildHeaderMap + ColIdx); bulk reads pull the whole range
' into a Variant array in ONE call and filter in memory, rather than
' touching the worksheet cell by cell. This is what keeps the app fast
' as the sheets grow.
'
' To add a NEW column later (e.g. "Priority", "ClaimAmount"): add the
' header to row 1 of the Claims sheet, then add ONE line here that
' reads/writes it. Nothing else changes, and no existing column can
' break by shifting position.
' =====================================================================

' ---------------------------------------------------------------------
' Add a brand-new claim.
' Mandatory user inputs: claimID, claimSite, providerName, claimQuery, creationDate
' Auto-filled: ClaimStatus=Pending, Attempt=0, ClaimUpdatedSite=claimSite,
'              ClaimInsertionDate=Now, ClaimInsertedBy=current user
' ---------------------------------------------------------------------
Public Function AddClaim(ByVal claimID As String, ByVal claimSite As String, _
                          ByVal providerName As String, ByVal claimQuery As String, _
                          ByVal creationDate As Date) As Boolean
    Dim wb As Workbook, ws As Worksheet, newRow As Long, lastCol As Long
    Dim hmap As Object, rowArr() As Variant

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)

    If FindClaimRow(wb, claimID) > 0 Then
        MsgBox "Claim ID '" & claimID & "' already exists.", vbExclamation
        CloseCentralDB wb, False
        AddClaim = False
        Exit Function
    End If

    Set hmap = BuildHeaderMap(ws)          ' header row scanned ONCE per operation
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    newRow = ws.Cells(ws.Rows.Count, ColIdx(hmap, "ClaimID")).End(xlUp).Row + 1

    ' Assemble in memory, write once - not 10 separate per-cell writes.
    ReDim rowArr(1 To 1, 1 To lastCol)
    rowArr(1, ColIdx(hmap, "ClaimID")) = claimID
    rowArr(1, ColIdx(hmap, "ClaimSite")) = claimSite
    rowArr(1, ColIdx(hmap, "ClaimProviderName")) = providerName
    rowArr(1, ColIdx(hmap, "ClaimQuery")) = claimQuery
    rowArr(1, ColIdx(hmap, "ClaimCreationDate")) = creationDate
    rowArr(1, ColIdx(hmap, "ClaimStatus")) = STATUS_PENDING
    rowArr(1, ColIdx(hmap, "Attempt")) = 0
    rowArr(1, ColIdx(hmap, "ClaimUpdatedSite")) = claimSite
    rowArr(1, ColIdx(hmap, "ClaimInsertionDate")) = Now
    rowArr(1, ColIdx(hmap, "ClaimInsertedBy")) = GetWindowsUserName()
    ws.Range(ws.Cells(newRow, 1), ws.Cells(newRow, lastCol)).Value = rowArr
    ' LastUpdatedDate / LastUpdatedBy / LastComment / ClaimClosedDate stay
    ' blank until the first call is logged - see LogCallAndUpdateStatus.

    CloseCentralDB wb, True
    AddClaim = True
    Exit Function

Fail:
    MsgBox "Could not add claim: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    AddClaim = False
End Function

' ---------------------------------------------------------------------
' Any caller logs a call: appends a History row and refreshes the
' rollup fields on the Claims row (latest status, attempt count,
' last-touched audit fields, and closed date if applicable).
' ---------------------------------------------------------------------
Public Function LogCallAndUpdateStatus(ByVal claimID As String, ByVal comment As String, _
                                        ByVal newStatus As String) As Boolean
    Dim wb As Workbook, wsHist As Worksheet, wsClaims As Worksheet
    Dim claimRow As Long, histRow As Long, statusCol As Long, histLastCol As Long
    Dim hmapC As Object, hmapH As Object, histArr() As Variant
    Dim userName As String, stamp As Date

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set wsClaims = wb.Sheets(SHEET_CLAIMS)
    Set wsHist = wb.Sheets(SHEET_HISTORY)
    userName = GetWindowsUserName()
    stamp = Now

    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        LogCallAndUpdateStatus = False
        Exit Function
    End If

    Set hmapC = BuildHeaderMap(wsClaims)   ' each header row scanned ONCE
    Set hmapH = BuildHeaderMap(wsHist)
    statusCol = ColIdx(hmapC, "ClaimStatus")
    If LCase$(Trim$(wsClaims.Cells(claimRow, statusCol).Value)) = LCase$(STATUS_CLOSED) Then
        MsgBox "This claim is already Closed and cannot accept further calls.", vbExclamation
        CloseCentralDB wb, False
        LogCallAndUpdateStatus = False
        Exit Function
    End If

    ' -- append history row (one bulk write) --
    histLastCol = wsHist.Cells(1, wsHist.Columns.Count).End(xlToLeft).Column
    histRow = wsHist.Cells(wsHist.Rows.Count, ColIdx(hmapH, "HistoryID")).End(xlUp).Row + 1
    ReDim histArr(1 To 1, 1 To histLastCol)
    histArr(1, ColIdx(hmapH, "HistoryID")) = NextHistoryID(wb)
    histArr(1, ColIdx(hmapH, "ClaimID")) = claimID
    histArr(1, ColIdx(hmapH, "CallerName")) = userName
    histArr(1, ColIdx(hmapH, "CallDateTime")) = stamp
    histArr(1, ColIdx(hmapH, "CallerComment")) = comment
    histArr(1, ColIdx(hmapH, "CallerStatus")) = newStatus
    wsHist.Range(wsHist.Cells(histRow, 1), wsHist.Cells(histRow, histLastCol)).Value = histArr

    ' -- refresh Claims rollup / audit fields --
    wsClaims.Cells(claimRow, ColIdx(hmapC, "Attempt")).Value = _
        wsClaims.Cells(claimRow, ColIdx(hmapC, "Attempt")).Value + 1
    wsClaims.Cells(claimRow, statusCol).Value = newStatus
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastUpdatedDate")).Value = stamp
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastUpdatedBy")).Value = userName
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastComment")).Value = comment

    If LCase$(newStatus) = LCase$(STATUS_CLOSED) Then
        wsClaims.Cells(claimRow, ColIdx(hmapC, "ClaimClosedDate")).Value = stamp
    End If

    CloseCentralDB wb, True
    LogCallAndUpdateStatus = True
    Exit Function

Fail:
    MsgBox "Could not log call: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    LogCallAndUpdateStatus = False
End Function

' ---------------------------------------------------------------------
' Admin-only: change a claim's updated site.
' ---------------------------------------------------------------------
Public Function ChangeUpdatedSite(ByVal claimID As String, ByVal newSite As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, claimRow As Long
    Dim hmap As Object

    On Error GoTo Fail
    Set wb = OpenCentralDB()

    If Not IsCurrentUserAdmin(wb) Then
        MsgBox "Only Admin users can change the claim site.", vbExclamation
        CloseCentralDB wb, False
        ChangeUpdatedSite = False
        Exit Function
    End If

    Set ws = wb.Sheets(SHEET_CLAIMS)
    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        ChangeUpdatedSite = False
        Exit Function
    End If

    Set hmap = BuildHeaderMap(ws)
    ws.Cells(claimRow, ColIdx(hmap, "ClaimUpdatedSite")).Value = newSite
    ws.Cells(claimRow, ColIdx(hmap, "LastUpdatedDate")).Value = Now
    ws.Cells(claimRow, ColIdx(hmap, "LastUpdatedBy")).Value = GetWindowsUserName()

    CloseCentralDB wb, True
    ChangeUpdatedSite = True
    Exit Function

Fail:
    MsgBox "Could not change site: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    ChangeUpdatedSite = False
End Function

' ---------------------------------------------------------------------
' Updates the user-supplied detail fields on an existing claim -
' corrections to what was typed at insert time.
'
' Deliberately does NOT touch ClaimStatus, Attempt, ClaimClosedDate or
' any history. Those are maintained by LogCallAndUpdateStatus so that
' status, attempt count and the call trail can never drift apart.
' ---------------------------------------------------------------------
Public Function UpdateClaimDetails(ByVal claimID As String, ByVal claimSite As String, _
                                    ByVal providerName As String, ByVal claimQuery As String, _
                                    ByVal creationDate As Date, _
                                    Optional ByVal updatedSite As String = "", _
                                    Optional ByRef siteChangeDenied As Boolean = False) As Boolean
    Dim wb As Workbook, ws As Worksheet, claimRow As Long
    Dim hmap As Object

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)

    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        UpdateClaimDetails = False
        Exit Function
    End If

    Set hmap = BuildHeaderMap(ws)

    ws.Cells(claimRow, ColIdx(hmap, "ClaimSite")).Value = claimSite
    ws.Cells(claimRow, ColIdx(hmap, "ClaimProviderName")).Value = providerName
    ws.Cells(claimRow, ColIdx(hmap, "ClaimQuery")).Value = claimQuery
    ws.Cells(claimRow, ColIdx(hmap, "ClaimCreationDate")).Value = creationDate

    If Trim$(updatedSite) <> "" Then
        If IsCurrentUserAdmin(wb) Then
            ws.Cells(claimRow, ColIdx(hmap, "ClaimUpdatedSite")).Value = updatedSite
        Else
            siteChangeDenied = True
        End If
    End If

    ws.Cells(claimRow, ColIdx(hmap, "LastUpdatedDate")).Value = Now
    ws.Cells(claimRow, ColIdx(hmap, "LastUpdatedBy")).Value = GetWindowsUserName()

    CloseCentralDB wb, True
    UpdateClaimDetails = True
    Exit Function

Fail:
    MsgBox "Could not update claim: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    UpdateClaimDetails = False
End Function

' ---------------------------------------------------------------------
' Admin-only: reopen a Closed claim so calling can resume.
'
' Unlike UpdateClaimDetails, this DOES write a history row. Reopening is
' a status change, and a status change that leaves no trace is exactly
' the kind of thing an audit needs to see. A reason is mandatory.
'
' What it does NOT do is increment Attempt. That column counts CALLS,
' and a reopen is an administrative act, not a call to the customer.
' Incrementing it would overstate calling effort.
'
' ClaimClosedDate is cleared. If it were left in place the claim would
' read as closed to every dashboard and to clsClaim.DaysOpen, which
' would silently report a stale time-to-close.
' ---------------------------------------------------------------------
Public Function ReopenClaim(ByVal claimID As String, ByVal reason As String) As Boolean
    Dim wb As Workbook, wsClaims As Worksheet, wsHist As Worksheet
    Dim claimRow As Long, histRow As Long, histLastCol As Long
    Dim hmapC As Object, hmapH As Object, histArr() As Variant
    Dim userName As String, stamp As Date

    If Trim$(reason) = "" Then
        MsgBox "A reason is required to reopen a claim.", vbExclamation, "Reason Needed"
        ReopenClaim = False
        Exit Function
    End If

    On Error GoTo Fail
    Set wb = OpenCentralDB()

    If Not IsCurrentUserAdmin(wb) Then
        MsgBox "Only Admin users can reopen a closed claim.", vbExclamation, "Access Denied"
        CloseCentralDB wb, False
        ReopenClaim = False
        Exit Function
    End If

    Set wsClaims = wb.Sheets(SHEET_CLAIMS)
    Set wsHist = wb.Sheets(SHEET_HISTORY)
    Set hmapC = BuildHeaderMap(wsClaims)
    Set hmapH = BuildHeaderMap(wsHist)

    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        ReopenClaim = False
        Exit Function
    End If

    If LCase$(Trim$(wsClaims.Cells(claimRow, ColIdx(hmapC, "ClaimStatus")).Value)) _
       <> LCase$(STATUS_CLOSED) Then
        MsgBox "Claim '" & claimID & "' is not Closed, so there is nothing to reopen.", _
               vbExclamation, "Not Closed"
        CloseCentralDB wb, False
        ReopenClaim = False
        Exit Function
    End If

    userName = GetWindowsUserName()
    stamp = Now

    ' --- audit row in the history trail ---
    histLastCol = wsHist.Cells(1, wsHist.Columns.Count).End(xlToLeft).Column
    histRow = wsHist.Cells(wsHist.Rows.Count, ColIdx(hmapH, "HistoryID")).End(xlUp).Row + 1
    ReDim histArr(1 To 1, 1 To histLastCol)
    histArr(1, ColIdx(hmapH, "HistoryID")) = NextHistoryID(wb)
    histArr(1, ColIdx(hmapH, "ClaimID")) = claimID
    histArr(1, ColIdx(hmapH, "CallerName")) = userName
    histArr(1, ColIdx(hmapH, "CallDateTime")) = stamp
    ' The marker matters: without it this row is indistinguishable from a
    ' logged call when someone reads or counts the history.
    histArr(1, ColIdx(hmapH, "CallerComment")) = "[REOPENED BY ADMIN] " & Trim$(reason)
    histArr(1, ColIdx(hmapH, "CallerStatus")) = STATUS_PENDING
    wsHist.Range(wsHist.Cells(histRow, 1), wsHist.Cells(histRow, histLastCol)).Value = histArr

    ' --- flip the claim back to Pending ---
    wsClaims.Cells(claimRow, ColIdx(hmapC, "ClaimStatus")).Value = STATUS_PENDING
    wsClaims.Cells(claimRow, ColIdx(hmapC, "ClaimClosedDate")).ClearContents
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastUpdatedDate")).Value = stamp
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastUpdatedBy")).Value = userName
    wsClaims.Cells(claimRow, ColIdx(hmapC, "LastComment")).Value = _
        "[REOPENED BY ADMIN] " & Trim$(reason)
    ' Attempt intentionally left alone - see the note above.

    CloseCentralDB wb, True
    ReopenClaim = True
    Exit Function

Fail:
    MsgBox "Could not reopen claim: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    ReopenClaim = False
End Function

' ---------------------------------------------------------------------
' Read helpers for populating the UI / dashboards.
' Returns the full used range (headers included) so new columns
' automatically flow through - the consumer matches by header name.
' ---------------------------------------------------------------------
Public Function GetAllClaims() As Variant
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, lastCol As Long
    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        GetAllClaims = Empty
    Else
        GetAllClaims = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    End If
    CloseCentralDB wb, False
    Exit Function
Fail:
    MsgBox "Could not read claims: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    GetAllClaims = Empty
End Function

Public Function GetHistoryForClaim(ByVal claimID As String) As Variant
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim results() As Variant, matchCount As Long, r As Long
    Dim hmap As Object, dataArr As Variant, target As String
    Dim idCol As Long, nameCol As Long, dtCol As Long, commentCol As Long, statusCol As Long

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_HISTORY)
    Set hmap = BuildHeaderMap(ws)
    idCol = ColIdx(hmap, "ClaimID")
    nameCol = ColIdx(hmap, "CallerName")
    dtCol = ColIdx(hmap, "CallDateTime")
    commentCol = ColIdx(hmap, "CallerComment")
    statusCol = ColIdx(hmap, "CallerStatus")

    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    target = Trim$(claimID)

    If lastRow < 2 Then
        CloseCentralDB wb, False
        GetHistoryForClaim = Empty
        Exit Function
    End If

    ' ONE bulk read of the whole History block, then filter in memory.
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value

    matchCount = 0
    For i = 1 To UBound(dataArr, 1)
        If Trim$(CStr(dataArr(i, idCol))) = target Then matchCount = matchCount + 1
    Next i

    If matchCount = 0 Then
        CloseCentralDB wb, False
        GetHistoryForClaim = Empty
        Exit Function
    End If

    ReDim results(1 To matchCount, 1 To 4) ' CallerName, CallDateTime, Comment, Status
    r = 0
    For i = 1 To UBound(dataArr, 1)
        If Trim$(CStr(dataArr(i, idCol))) = target Then
            r = r + 1
            results(r, 1) = dataArr(i, nameCol)
            results(r, 2) = dataArr(i, dtCol)
            results(r, 3) = dataArr(i, commentCol)
            results(r, 4) = dataArr(i, statusCol)
        End If
    Next i

    CloseCentralDB wb, False
    GetHistoryForClaim = results
    Exit Function

Fail:
    MsgBox "Could not read history: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    GetHistoryForClaim = Empty
End Function
