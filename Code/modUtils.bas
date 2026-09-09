Attribute VB_Name = "modUtils"
Option Explicit

' =====================================================================
' modUtils
' Generic helpers used across the app. Keep this file free of any
' Claims/History-specific logic so it stays reusable in other tools.
' =====================================================================

' Returns the Windows login name as a fallback identity.
' Prefer GetCurrentUserName() below, which maps this to the Users sheet.
Public Function GetWindowsUserName() As String
    GetWindowsUserName = Environ$("USERNAME")
End Function

' Looks up the caller's friendly name/role from the Users sheet in the
' central DB. dbWb must already be open.
Public Function GetCurrentUserRole(dbWb As Workbook) As String
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim winUser As String
    winUser = GetWindowsUserName()

    Set ws = dbWb.Sheets(SHEET_USERS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If LCase$(Trim$(ws.Cells(i, 1).Value)) = LCase$(Trim$(winUser)) Then
            GetCurrentUserRole = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetCurrentUserRole = "User" ' default role if not found in Users sheet
End Function

Public Function IsCurrentUserAdmin(dbWb As Workbook) As Boolean
    IsCurrentUserAdmin = (LCase$(GetCurrentUserRole(dbWb)) = "admin")
End Function

' ---------------------------------------------------------------------
' Opens the central DB workbook with retry logic, because two users
' saving at the exact same moment will otherwise throw a "file in use"
' or permission error. This is the single riskiest part of using a
' plain xlsx file as a shared database - see notes in chat.
' ---------------------------------------------------------------------
Public Function OpenCentralDB() As Workbook
    Dim attempt As Long
    Dim wb As Workbook

    On Error Resume Next
    Do
        attempt = attempt + 1
        Set wb = Nothing
        Set wb = Workbooks.Open(FileName:=DB_PATH, UpdateLinks:=0, ReadOnly:=False, Notify:=False)
        If Not wb Is Nothing Then Exit Do
        If attempt >= LOCK_MAX_RETRIES Then
            On Error GoTo 0
            Err.Raise vbObjectError + 1, "OpenCentralDB", _
                "Could not open the shared database after " & LOCK_MAX_RETRIES & _
                " attempts. Another user may be saving it right now. Please try again shortly."
        End If
        Application.Wait Now + TimeSerial(0, 0, LOCK_RETRY_WAIT_SEC)
    Loop
    On Error GoTo 0

    Set OpenCentralDB = wb
End Function

' Always close+save through this so every write path behaves the same
' way and releases the file lock as fast as possible.
Public Sub CloseCentralDB(wb As Workbook, ByVal saveChanges As Boolean)
    On Error Resume Next
    If saveChanges Then
        wb.Save
    End If
    wb.Close SaveChanges:=False ' already saved above; avoids double prompt
    On Error GoTo 0
End Sub

Public Function NextHistoryID(dbWb As Workbook) As Long
    Dim ws As Worksheet, lastRow As Long
    Set ws = dbWb.Sheets(SHEET_HISTORY)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        NextHistoryID = 1
    Else
        NextHistoryID = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow)) + 1
    End If
End Function

Public Function FindClaimRow(dbWb As Workbook, ByVal claimID As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = dbWb.Sheets(SHEET_CLAIMS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 1).Value)) = Trim$(claimID) Then
            FindClaimRow = i
            Exit Function
        End If
    Next i
    FindClaimRow = 0 ' not found
End Function
