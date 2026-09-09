Attribute VB_Name = "modDataAccess"
Option Explicit

' =====================================================================
' modDataAccess
' Every read/write against the shared database goes through here.
' UI code (forms) should NEVER touch worksheet cells directly -
' this is what makes the app modular: swap this module for an
' Access/SQL backend later and the forms don't change.
' =====================================================================

' ---------------------------------------------------------------------
' Add a brand-new claim. Sets ClaimInsertionDate/ClaimInsertedBy
' automatically - the UI form should not expose these as editable.
' ---------------------------------------------------------------------
Public Function AddClaim(ByVal claimID As String, ByVal callerLocation As String) As Boolean
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

    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = claimID
    ws.Cells(newRow, 2).Value = callerLocation
    ws.Cells(newRow, 3).Value = STATUS_PENDING     ' ClaimStatus starts Pending
    ws.Cells(newRow, 4).Value = 0                  ' CallingAttempted
    ws.Cells(newRow, 5).Value = callerLocation      ' UpdatedLocation = same initially
    ws.Cells(newRow, 6).Value = Now                ' ClaimInsertionDate
    ws.Cells(newRow, 7).Value = GetWindowsUserName() ' ClaimInsertedBy

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
' rollup fields on the Claims row (latest status + total attempts).
' This is the sub that "Pending -> other callers can keep calling"
' and "Closed -> stops" logic hinges on.
' ---------------------------------------------------------------------
Public Function LogCallAndUpdateStatus(ByVal claimID As String, ByVal comment As String, _
                                        ByVal newStatus As String) As Boolean
    Dim wb As Workbook, wsHist As Worksheet, wsClaims As Worksheet
    Dim claimRow As Long, histRow As Long, attempts As Long

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

    If LCase$(Trim$(wsClaims.Cells(claimRow, 3).Value)) = LCase$(STATUS_CLOSED) Then
        MsgBox "This claim is already Closed and cannot accept further calls.", vbExclamation
        CloseCentralDB wb, False
        LogCallAndUpdateStatus = False
        Exit Function
    End If

    ' -- append history row --
    histRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row + 1
    wsHist.Cells(histRow, 1).Value = NextHistoryID(wb)
    wsHist.Cells(histRow, 2).Value = claimID
    wsHist.Cells(histRow, 3).Value = GetWindowsUserName()
    wsHist.Cells(histRow, 4).Value = Now
    wsHist.Cells(histRow, 5).Value = comment
    wsHist.Cells(histRow, 6).Value = newStatus

    ' -- refresh Claims rollup --
    attempts = wsClaims.Cells(claimRow, 4).Value + 1
    wsClaims.Cells(claimRow, 4).Value = attempts       ' CallingAttempted
    wsClaims.Cells(claimRow, 3).Value = newStatus      ' ClaimStatus = latest status

    CloseCentralDB wb, True
    LogCallAndUpdateStatus = True
    Exit Function

Fail:
    MsgBox "Could not log call: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    LogCallAndUpdateStatus = False
End Function

' ---------------------------------------------------------------------
' Admin-only: change a claim's caller/updated location.
' ---------------------------------------------------------------------
Public Function ChangeCallerLocation(ByVal claimID As String, ByVal newLocation As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, claimRow As Long

    On Error GoTo Fail
    Set wb = OpenCentralDB()

    If Not IsCurrentUserAdmin(wb) Then
        MsgBox "Only Admin users can change the caller location.", vbExclamation
        CloseCentralDB wb, False
        ChangeCallerLocation = False
        Exit Function
    End If

    Set ws = wb.Sheets(SHEET_CLAIMS)
    claimRow = FindClaimRow(wb, claimID)
    If claimRow = 0 Then
        MsgBox "Claim ID '" & claimID & "' not found.", vbExclamation
        CloseCentralDB wb, False
        ChangeCallerLocation = False
        Exit Function
    End If

    ws.Cells(claimRow, 5).Value = newLocation ' UpdatedLocation

    CloseCentralDB wb, True
    ChangeCallerLocation = True
    Exit Function

Fail:
    MsgBox "Could not change location: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    ChangeCallerLocation = False
End Function

' ---------------------------------------------------------------------
' Read helpers for populating the UI (list of claims, history for one claim)
' Returned as Variant arrays so forms can bind to a ListBox/ListView directly.
' ---------------------------------------------------------------------
Public Function GetAllClaims() As Variant
    Dim wb As Workbook, ws As Worksheet, lastRow As Long
    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_CLAIMS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        GetAllClaims = Empty
    Else
        GetAllClaims = ws.Range("A2:G" & lastRow).Value
    End If
    CloseCentralDB wb, False
    Exit Function
Fail:
    MsgBox "Could not read claims: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
    GetAllClaims = Empty
End Function

Public Function GetHistoryForClaim(ByVal claimID As String) As Variant
    Dim wb As Workbook, ws As Worksheet, lastRow As Long, i As Long
    Dim results() As Variant, matchCount As Long, r As Long

    On Error GoTo Fail
    Set wb = OpenCentralDB()
    Set ws = wb.Sheets(SHEET_HISTORY)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        CloseCentralDB wb, False
        GetHistoryForClaim = Empty
        Exit Function
    End If

    matchCount = Application.WorksheetFunction.CountIf(ws.Range("B2:B" & lastRow), claimID)
    If matchCount = 0 Then
        CloseCentralDB wb, False
        GetHistoryForClaim = Empty
        Exit Function
    End If

    ReDim results(1 To matchCount, 1 To 4) ' CallerName, CallDateTime, Comment, Status
    r = 0
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 2).Value)) = Trim$(claimID) Then
            r = r + 1
            results(r, 1) = ws.Cells(i, 3).Value
            results(r, 2) = ws.Cells(i, 4).Value
            results(r, 3) = ws.Cells(i, 5).Value
            results(r, 4) = ws.Cells(i, 6).Value
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
