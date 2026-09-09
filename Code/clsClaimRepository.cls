VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClaimRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
    Dim wb As Workbook, ws As Worksheet, newRow As Long

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)

    If FindClaimRow(wb, claimID) > 0 Then
        MsgBox "Claim ID '" & claimID & "' already exists.", vbExclamation
        CloseCentralDB wb, False
        AddClaim = False
        Exit Function
    End If

    newRow = ws.Cells(ws.Rows.Count, GetColIndex(ws, "ClaimID")).End(xlUp).Row + 1
    ws.Cells(newRow, GetColIndex(ws, "ClaimID")).Value = claimID
    ws.Cells(newRow, GetColIndex(ws, "ClaimSite")).Value = claimSite
    ws.Cells(newRow, GetColIndex(ws, "ClaimProviderName")).Value = providerName
    ws.Cells(newRow, GetColIndex(ws, "ClaimQuery")).Value = claimQuery
    ws.Cells(newRow, GetColIndex(ws, "ClaimCreationDate")).Value = creationDate
    ws.Cells(newRow, GetColIndex(ws, "ClaimStatus")).Value = STATUS_PENDING
    ws.Cells(newRow, GetColIndex(ws, "Attempt")).Value = 0
    ws.Cells(newRow, GetColIndex(ws, "ClaimUpdatedSite")).Value = claimSite
    ws.Cells(newRow, GetColIndex(ws, "ClaimInsertionDate")).Value = Now
    ws.Cells(newRow, GetColIndex(ws, "ClaimInsertedBy")).Value = GetWindowsUserName()

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
    Dim claimRow As Long, histRow As Long, statusCol As Long

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set wsClaims = wb.Sheets(SHEET_CLAIMS)
    Set wsHist = wb.Sheets(SHEET_HISTORY)

    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        LogCallAndUpdateStatus = False
        Exit Function
    End If

    statusCol = GetColIndex(wsClaims, "ClaimStatus")
    If LCase$(Trim$(wsClaims.Cells(claimRow, statusCol).Value)) = LCase$(STATUS_CLOSED) Then
        MsgBox "This claim is already Closed and cannot accept further calls.", vbExclamation
        CloseCentralDB wb, False
        LogCallAndUpdateStatus = False
        Exit Function
    End If

    histRow = wsHist.Cells(wsHist.Rows.Count, GetColIndex(wsHist, "HistoryID")).End(xlUp).Row + 1
    wsHist.Cells(histRow, GetColIndex(wsHist, "HistoryID")).Value = NextHistoryID(wb)
    wsHist.Cells(histRow, GetColIndex(wsHist, "ClaimID")).Value = claimID
    wsHist.Cells(histRow, GetColIndex(wsHist, "CallerName")).Value = GetWindowsUserName()
    wsHist.Cells(histRow, GetColIndex(wsHist, "CallDateTime")).Value = Now
    wsHist.Cells(histRow, GetColIndex(wsHist, "CallerComment")).Value = comment
    wsHist.Cells(histRow, GetColIndex(wsHist, "CallerStatus")).Value = newStatus

    wsClaims.Cells(claimRow, GetColIndex(wsClaims, "Attempt")).Value = _
        wsClaims.Cells(claimRow, GetColIndex(wsClaims, "Attempt")).Value + 1
    wsClaims.Cells(claimRow, statusCol).Value = newStatus
    wsClaims.Cells(claimRow, GetColIndex(wsClaims, "LastUpdatedDate")).Value = Now
    wsClaims.Cells(claimRow, GetColIndex(wsClaims, "LastUpdatedBy")).Value = GetWindowsUserName()
    wsClaims.Cells(claimRow, GetColIndex(wsClaims, "LastComment")).Value = comment

    If LCase$(newStatus) = LCase$(STATUS_CLOSED) Then
        wsClaims.Cells(claimRow, GetColIndex(wsClaims, "ClaimClosedDate")).Value = Now
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

    ws.Cells(claimRow, GetColIndex(ws, "ClaimUpdatedSite")).Value = newSite
    ws.Cells(claimRow, GetColIndex(ws, "LastUpdatedDate")).Value = Now
    ws.Cells(claimRow, GetColIndex(ws, "LastUpdatedBy")).Value = GetWindowsUserName()

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
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, i As Long
    Dim result As New Collection
    Dim c As clsClaim

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)
    lastRow = ws.Cells(ws.Rows.Count, GetColIndex(ws, "ClaimID")).End(xlUp).Row

    For i = 2 To lastRow
        Set c = New clsClaim
        c.ClaimID = ws.Cells(i, GetColIndex(ws, "ClaimID")).Value
        c.ClaimSite = ws.Cells(i, GetColIndex(ws, "ClaimSite")).Value
        c.ClaimProviderName = ws.Cells(i, GetColIndex(ws, "ClaimProviderName")).Value
        c.ClaimQuery = ws.Cells(i, GetColIndex(ws, "ClaimQuery")).Value
        c.ClaimCreationDate = ws.Cells(i, GetColIndex(ws, "ClaimCreationDate")).Value
        c.ClaimStatus = ws.Cells(i, GetColIndex(ws, "ClaimStatus")).Value
        c.Attempt = ws.Cells(i, GetColIndex(ws, "Attempt")).Value
        c.ClaimUpdatedSite = ws.Cells(i, GetColIndex(ws, "ClaimUpdatedSite")).Value
        c.ClaimInsertionDate = ws.Cells(i, GetColIndex(ws, "ClaimInsertionDate")).Value
        c.ClaimInsertedBy = ws.Cells(i, GetColIndex(ws, "ClaimInsertedBy")).Value
        c.LastUpdatedDate = ws.Cells(i, GetColIndex(ws, "LastUpdatedDate")).Value
        c.LastUpdatedBy = ws.Cells(i, GetColIndex(ws, "LastUpdatedBy")).Value
        c.LastComment = ws.Cells(i, GetColIndex(ws, "LastComment")).Value
        c.ClaimClosedDate = ws.Cells(i, GetColIndex(ws, "ClaimClosedDate")).Value
        result.Add c, c.ClaimID
    Next i

    CloseCentralDB wb, False
    Set GetAllClaims = result
    Exit Function

Fail:
    MsgBox "Could not read claims: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    Set GetAllClaims = New Collection
End Function

' Returns Nothing if not found - callers must check.
Public Function FindClaim(ByVal claimID As String) As clsClaim
    Dim all As Collection
    Set all = GetAllClaims()
    On Error Resume Next
    Set FindClaim = all(claimID)
    On Error GoTo 0
End Function

' Returns a Collection of clsHistoryEntry for one claim, oldest first.
Public Function GetHistory(ByVal claimID As String) As Collection
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, i As Long
    Dim idCol As Long
    Dim result As New Collection
    Dim h As clsHistoryEntry

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_HISTORY)
    idCol = GetColIndex(ws, "ClaimID")
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row

    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, idCol).Value)) = Trim$(claimID) Then
            Set h = New clsHistoryEntry
            h.HistoryID = ws.Cells(i, GetColIndex(ws, "HistoryID")).Value
            h.ClaimID = ws.Cells(i, idCol).Value
            h.CallerName = ws.Cells(i, GetColIndex(ws, "CallerName")).Value
            h.CallDateTime = ws.Cells(i, GetColIndex(ws, "CallDateTime")).Value
            h.CallerComment = ws.Cells(i, GetColIndex(ws, "CallerComment")).Value
            h.CallerStatus = ws.Cells(i, GetColIndex(ws, "CallerStatus")).Value
            result.Add h
        End If
    Next i

    CloseCentralDB wb, False
    Set GetHistory = result
    Exit Function

Fail:
    MsgBox "Could not read history: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    Set GetHistory = New Collection
End Function
