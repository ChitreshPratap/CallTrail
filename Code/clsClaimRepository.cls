VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsClaimRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================================
' clsClaimRepository
' Object-oriented replacement for modDataAccess. Forms work with
' clsClaim / clsHistoryEntry objects instead of raw arrays and column
' numbers. Still built on modUtils (OpenCentralDB/CloseCentralDB/
' GetColIndex) so the file-locking/retry behaviour is unchanged -
' only the "shape" of what callers deal with is different.
'
' Usage from a form:
'   Dim repo As New clsClaimRepository
'   If repo.AddClaim("C123","Delhi","Acme","Query text",Date) Then ...
' =====================================================================

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

    ' Assemble the row in memory, then write it with a SINGLE range write
    ' instead of 10 separate per-cell writes.
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

    CloseCentralDB wb, True
    AddClaim = True
    Exit Function

Fail:
    MsgBox "Could not add claim: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    AddClaim = False
End Function

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
' Returns a Collection of clsClaim objects - ready for a dashboard loop,
' e.g.: For Each c In repo.GetAllClaims(): If c.IsClosed Then ...
' ---------------------------------------------------------------------
Public Function GetAllClaims() As Collection
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim result As New Collection
    Dim c As clsClaim
    Dim dataArr As Variant
    Dim hmap As Object

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)
    Set hmap = BuildHeaderMap(ws)  ' one bulk read of row 1, cached for the whole loop
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow >= 2 Then
        ' ONE bulk read of the whole data block - this replaces what was
        ' previously up to (rows x 14 columns) individual worksheet touches.
        dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value

        For i = 1 To UBound(dataArr, 1)
            Set c = New clsClaim
            c.claimID = dataArr(i, ColIdx(hmap, "ClaimID"))
            c.claimSite = dataArr(i, ColIdx(hmap, "ClaimSite"))
            c.ClaimProviderName = dataArr(i, ColIdx(hmap, "ClaimProviderName"))
            c.claimQuery = dataArr(i, ColIdx(hmap, "ClaimQuery"))
            c.ClaimCreationDate = dataArr(i, ColIdx(hmap, "ClaimCreationDate"))
            c.ClaimStatus = dataArr(i, ColIdx(hmap, "ClaimStatus"))
            c.attempt = dataArr(i, ColIdx(hmap, "Attempt"))
            c.ClaimUpdatedSite = dataArr(i, ColIdx(hmap, "ClaimUpdatedSite"))
            c.ClaimInsertionDate = dataArr(i, ColIdx(hmap, "ClaimInsertionDate"))
            c.ClaimInsertedBy = dataArr(i, ColIdx(hmap, "ClaimInsertedBy"))
            c.LastUpdatedDate = dataArr(i, ColIdx(hmap, "LastUpdatedDate"))
            c.LastUpdatedBy = dataArr(i, ColIdx(hmap, "LastUpdatedBy"))
            c.LastComment = dataArr(i, ColIdx(hmap, "LastComment"))
            c.ClaimClosedDate = dataArr(i, ColIdx(hmap, "ClaimClosedDate"))
            result.Add c, c.claimID
        Next i
    End If

    CloseCentralDB wb, False
    Set GetAllClaims = result
    Exit Function

Fail:
    MsgBox "Could not read claims: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    Set GetAllClaims = New Collection
End Function

' ---------------------------------------------------------------------
' Bulk insert. Takes a Collection of clsClaim objects (only the five
' mandatory user fields need to be set on each; the rest are auto-filled).
'
' Why this exists: calling AddClaim in a loop would open, save and close
' the shared workbook once PER CLAIM - 500 claims = 500 network round
' trips and 500 chances to collide with another user. This does the
' whole batch in ONE open/write/save/close, with a single bulk range
' write instead of per-cell writes.
'
' Returns the number of claims actually inserted; duplicates are skipped
' and reported back via the skippedIDs argument.
' ---------------------------------------------------------------------
Public Function AddClaimsBulk(claims As Collection, Optional ByRef skippedIDs As String) As Long
    Dim wb As Workbook, ws As Worksheet
    Dim hmap As Object, existing As Object
    Dim lastRow As Long, lastCol As Long, firstNewRow As Long
    Dim idArr As Variant, outArr() As Variant
    Dim c As clsClaim, i As Long, n As Long, inserted As Long
    Dim userName As String, stamp As Date

    On Error GoTo Fail
    If claims Is Nothing Then Exit Function
    If claims.Count = 0 Then Exit Function

    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)
    Set hmap = BuildHeaderMap(ws)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, ColIdx(hmap, "ClaimID")).End(xlUp).Row

    ' Build an in-memory set of existing IDs so the duplicate check is
    ' O(1) per claim instead of re-scanning the sheet for each one.
    Set existing = CreateObject("Scripting.Dictionary")
    If lastRow >= 2 Then
        idArr = ws.Range(ws.Cells(2, ColIdx(hmap, "ClaimID")), _
                          ws.Cells(lastRow, ColIdx(hmap, "ClaimID"))).Value
        ' LCase so "ABC123" and "abc123" are treated as the same claim -
        ' this must match the check in modBulkImport.ValidateStaging, or
        ' validation would flag a duplicate that the insert then allows.
        If lastRow = 2 Then
            existing(Trim$(LCase$(CStr(idArr)))) = True
        Else
            For i = 1 To UBound(idArr, 1)
                existing(Trim$(LCase$(CStr(idArr(i, 1))))) = True
            Next i
        End If
    End If

    ReDim outArr(1 To claims.Count, 1 To lastCol)
    userName = GetWindowsUserName()
    stamp = Now
    n = 0
    skippedIDs = ""

    For Each c In claims
        If existing.Exists(Trim$(LCase$(c.claimID))) Then
            skippedIDs = skippedIDs & IIf(skippedIDs = "", "", ", ") & c.claimID
        Else
            n = n + 1
            outArr(n, ColIdx(hmap, "ClaimID")) = c.claimID
            outArr(n, ColIdx(hmap, "ClaimSite")) = c.claimSite
            outArr(n, ColIdx(hmap, "ClaimProviderName")) = c.ClaimProviderName
            outArr(n, ColIdx(hmap, "ClaimQuery")) = c.claimQuery
            outArr(n, ColIdx(hmap, "ClaimCreationDate")) = c.ClaimCreationDate
            outArr(n, ColIdx(hmap, "ClaimStatus")) = STATUS_PENDING
            outArr(n, ColIdx(hmap, "Attempt")) = 0
            outArr(n, ColIdx(hmap, "ClaimUpdatedSite")) = c.claimSite
            outArr(n, ColIdx(hmap, "ClaimInsertionDate")) = stamp
            outArr(n, ColIdx(hmap, "ClaimInsertedBy")) = userName
            existing(Trim$(LCase$(c.claimID))) = True ' guard against dupes within the batch itself
        End If
    Next c

    If n > 0 Then
        ' If any were skipped, outArr has unused trailing rows. Trim it to
        ' exactly n rows so the array and the target range match exactly,
        ' rather than relying on Excel silently taking the top-left block.
        If n < claims.Count Then
            Dim trimmed() As Variant, r As Long, k As Long
            ReDim trimmed(1 To n, 1 To lastCol)
            For r = 1 To n
                For k = 1 To lastCol
                    trimmed(r, k) = outArr(r, k)
                Next k
            Next r
            outArr = trimmed
        End If

        firstNewRow = lastRow + 1
        ' ONE write for the entire batch
        ws.Range(ws.Cells(firstNewRow, 1), ws.Cells(firstNewRow + n - 1, lastCol)).Value = outArr
    End If

    CloseCentralDB wb, (n > 0)
    AddClaimsBulk = n
    Exit Function

Fail:
    MsgBox "Bulk insert failed: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    AddClaimsBulk = 0
End Function

' ---------------------------------------------------------------------
' Updates the user-supplied detail fields on an existing claim -
' corrections to what was typed at insert time.
'
' Deliberately does NOT touch ClaimStatus, Attempt, ClaimClosedDate or
' any history. Those are maintained by LogCallAndUpdateStatus so that
' status, attempt count and the call trail can never drift apart. A
' correction to a provider name is not a call and must not look like one.
'
' ClaimUpdatedSite is only changed when the current user is an Admin;
' a non-admin's value for it is ignored rather than rejected, so the
' rest of their corrections still save.
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
' Returns Nothing if not found - callers must check with `Is Nothing`.
' Reads ONLY the matching row, rather than loading every claim into
' memory. This is the single biggest win for forms that look up one
' claim at a time (the common case) once the sheet grows large.
Public Function FindClaim(ByVal claimID As String) As clsClaim
    Dim wb As Workbook, ws As Worksheet, claimRow As Long, lastCol As Long
    Dim rowArr As Variant, hmap As Object
    Dim c As clsClaim

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)

    claimRow = FindClaimRow(wb, claimID)  ' bulk-reads just the ID column
    If claimRow = 0 Then
        CloseCentralDB wb, False
        Set FindClaim = Nothing
        Exit Function
    End If

    Set hmap = BuildHeaderMap(ws)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    rowArr = ws.Range(ws.Cells(claimRow, 1), ws.Cells(claimRow, lastCol)).Value ' one row, one read

    Set c = New clsClaim
    c.claimID = rowArr(1, ColIdx(hmap, "ClaimID"))
    c.claimSite = rowArr(1, ColIdx(hmap, "ClaimSite"))
    c.ClaimProviderName = rowArr(1, ColIdx(hmap, "ClaimProviderName"))
    c.claimQuery = rowArr(1, ColIdx(hmap, "ClaimQuery"))
    c.ClaimCreationDate = rowArr(1, ColIdx(hmap, "ClaimCreationDate"))
    c.ClaimStatus = rowArr(1, ColIdx(hmap, "ClaimStatus"))
    c.attempt = rowArr(1, ColIdx(hmap, "Attempt"))
    c.ClaimUpdatedSite = rowArr(1, ColIdx(hmap, "ClaimUpdatedSite"))
    c.ClaimInsertionDate = rowArr(1, ColIdx(hmap, "ClaimInsertionDate"))
    c.ClaimInsertedBy = rowArr(1, ColIdx(hmap, "ClaimInsertedBy"))
    c.LastUpdatedDate = rowArr(1, ColIdx(hmap, "LastUpdatedDate"))
    c.LastUpdatedBy = rowArr(1, ColIdx(hmap, "LastUpdatedBy"))
    c.LastComment = rowArr(1, ColIdx(hmap, "LastComment"))
    c.ClaimClosedDate = rowArr(1, ColIdx(hmap, "ClaimClosedDate"))

    CloseCentralDB wb, False
    Set FindClaim = c
    Exit Function

Fail:
    MsgBox "Could not find claim: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    Set FindClaim = Nothing
End Function

' Returns a Collection of clsHistoryEntry for one claim, oldest first.
' Bulk-reads the whole History sheet once (single COM call), then
' filters in memory - fast even once History has many thousands of
' rows across all claims combined.
Public Function GetHistory(ByVal claimID As String) As Collection
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim result As New Collection
    Dim h As clsHistoryEntry
    Dim dataArr As Variant
    Dim hmap As Object
    Dim idColArr As Long, target As String

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_HISTORY)
    Set hmap = BuildHeaderMap(ws)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    target = Trim$(claimID)

    If lastRow >= 2 Then
        dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value ' one bulk read
        idColArr = ColIdx(hmap, "ClaimID")

        For i = 1 To UBound(dataArr, 1)
            If Trim$(CStr(dataArr(i, idColArr))) = target Then
                Set h = New clsHistoryEntry
                h.HistoryID = dataArr(i, ColIdx(hmap, "HistoryID"))
                h.claimID = dataArr(i, idColArr)
                h.CallerName = dataArr(i, ColIdx(hmap, "CallerName"))
                h.CallDateTime = dataArr(i, ColIdx(hmap, "CallDateTime"))
                h.CallerComment = dataArr(i, ColIdx(hmap, "CallerComment"))
                h.CallerStatus = dataArr(i, ColIdx(hmap, "CallerStatus"))
                result.Add h
            End If
        Next i
    End If

    CloseCentralDB wb, False
    Set GetHistory = result
    Exit Function

Fail:
    MsgBox "Could not read history: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    Set GetHistory = New Collection
End Function


