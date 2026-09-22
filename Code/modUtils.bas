Attribute VB_Name = "modUtils"
Option Explicit

' =====================================================================
' modUtils
' Generic helpers used across the app. Keep this file free of any
' Claims/History-specific logic so it stays reusable in other tools.
' =====================================================================

' --- module-level state (VBA requires these above the first procedure) ---
Private savedScreenUpdating As Boolean
Private savedCalculation As XlCalculation
Private savedEnableEvents As Boolean

' Returns the Windows login name as a fallback identity.
' Prefer GetCurrentUserName() below, which maps this to the Users sheet.
Public Function GetWindowsUserName() As String
    GetWindowsUserName = Environ$("USERNAME")
End Function

' Looks up the caller's friendly name/role from the Users sheet in the
' central DB. dbWb must already be open.
' ---------------------------------------------------------------------
' Looks up the logged-in Windows user in the Users sheet.
'
' Returns True if they are registered. Role and defaultLocation come
' back ByRef. An UNREGISTERED user is not an error - they get Guest
' role and a locked-down app, which the home page explains.
'
' Note this replaces the old "default to User if not found" behaviour.
' Defaulting an unknown person to a working role meant anyone who opened
' the file could add and edit claims.
' ---------------------------------------------------------------------
Public Function GetUserRecord(dbWb As Workbook, ByRef role As String, _
                               ByRef defaultLocation As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim nameCol As Long, roleCol As Long, locCol As Long
    Dim winUser As String
    Dim data As Variant

    role = "Guest"
    defaultLocation = ""
    winUser = LCase$(Trim$(GetWindowsUserName()))

    On Error GoTo NotFound
    Set ws = dbWb.Sheets(SHEET_USERS)
    nameCol = GetColIndex(ws, "UserName")
    roleCol = GetColIndex(ws, "Role")

    ' DefaultLocation is optional - an older Users sheet may not have it
    locCol = 0
    On Error Resume Next
    locCol = GetColIndex(ws, "DefaultLocation")
    On Error GoTo NotFound

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    If lastRow < 2 Then GoTo NotFound

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).Value

    For i = 1 To UBound(data, 1)
        If LCase$(Trim$(CStr(data(i, nameCol) & ""))) = winUser Then
            role = Trim$(CStr(data(i, roleCol) & ""))
            If role = "" Then role = "User"
            If locCol > 0 Then defaultLocation = Trim$(CStr(data(i, locCol) & ""))
            GetUserRecord = True
            Exit Function
        End If
    Next i

NotFound:
    GetUserRecord = False
End Function

Public Function GetCurrentUserRole(dbWb As Workbook) As String
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim winUser As String, nameCol As Long, roleCol As Long
    winUser = GetWindowsUserName()

    Set ws = dbWb.Sheets(SHEET_USERS)
    nameCol = GetColIndex(ws, "UserName")
    roleCol = GetColIndex(ws, "Role")
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row

    For i = 2 To lastRow
        If LCase$(Trim$(ws.Cells(i, nameCol).Value)) = LCase$(Trim$(winUser)) Then
            GetCurrentUserRole = ws.Cells(i, roleCol).Value
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
'
' Also suspends screen updating / auto-calc / events for the duration -
' this alone is a real speedup once the DB has any volume of data,
' and costs nothing since these are always restored in CloseCentralDB.
' ---------------------------------------------------------------------

Public Function OpenCentralDB() As Workbook
    Set OpenCentralDB = OpenDBCore(False)
End Function

' Always close+save through this so every write path behaves the same
' way, releases the file lock as fast as possible, and restores the
' app settings OpenCentralDB suspended.
Public Sub CloseCentralDB(wb As Workbook, ByVal saveChanges As Boolean)
    On Error Resume Next
    If saveChanges Then
        wb.Save
    End If
    wb.Close SaveChanges:=False ' already saved above; avoids double prompt
    On Error GoTo 0
    RestoreAppSettings
End Sub

Private Sub RestoreAppSettings()
    Application.ScreenUpdating = savedScreenUpdating
    Application.EnableEvents = savedEnableEvents
    Application.Calculation = savedCalculation
End Sub

' ---------------------------------------------------------------------
' Read-only open, for operations that only READ the central DB.
'
' Use this instead of OpenCentralDB for exports, reports and downloads.
' A read/write open takes a lock on the shared file, so anyone else
' trying to log a call while a large download runs gets blocked or
' bounced to read-only themselves. A read-only open takes no such lock.
' ---------------------------------------------------------------------
Public Function OpenCentralDBReadOnly() As Workbook
    Set OpenCentralDBReadOnly = OpenDBCore(True)
End Function

' ---------------------------------------------------------------------
' The ONE place the database file is opened. Both public openers route
' through here, so the password, the retry loop and the lock checks
' can't drift apart between read and write paths.
'
' Three failure cases are told apart, because they need different
' handling:
'
'  * WRONG PASSWORD - never retried. Retrying a wrong password five times
'    just wastes 15 seconds before saying the same thing. It means the
'    macro file and the database are out of step (see modSecurity).
'
'  * FILE IN USE - retried. Someone else is mid-save.
'
'  * OPENED READ-ONLY WHEN WE WANTED TO WRITE - retried. With Notify off,
'    Excel quietly opens a file another user holds as read-only instead
'    of failing. Every later Save would then fail, or worse, look like it
'    worked. So a write open checks wb.ReadOnly and treats True as "busy".
' ---------------------------------------------------------------------
Private Function OpenDBCore(ByVal readOnly As Boolean) As Workbook
    Dim attempt As Long
    Dim wb As Workbook
    Dim errMsg As String

    savedScreenUpdating = Application.ScreenUpdating
    savedCalculation = Application.Calculation
    savedEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Do
        attempt = attempt + 1
        Set wb = Nothing
        errMsg = ""

        On Error Resume Next
        Set wb = Workbooks.Open(FileName:=DB_PATH, UpdateLinks:=0, _
                                ReadOnly:=readOnly, Notify:=False, _
                                Password:=DbPassword(), IgnoreReadOnlyRecommended:=True)
        If Err.Number <> 0 Then errMsg = Err.Description
        On Error GoTo 0

        ' --- wrong password: stop immediately ---
        If wb Is Nothing And InStr(1, LCase$(errMsg), "password") > 0 Then
            RestoreAppSettings
            Err.Raise vbObjectError + 5, "OpenCentralDB", _
                "The database password in this macro file does not match the database." & vbCrLf & vbCrLf & _
                "Your copy of the CallTrail macro file is probably out of date. " & _
                "Ask your administrator for the current version."
        End If

        ' --- opened, but read-only when we need to write: someone holds it ---
        If Not wb Is Nothing And Not readOnly Then
            If wb.ReadOnly Then
                wb.Close SaveChanges:=False
                Set wb = Nothing
            End If
        End If

        If Not wb Is Nothing Then Exit Do

        If attempt >= LOCK_MAX_RETRIES Then
            RestoreAppSettings
            Err.Raise vbObjectError + 1, "OpenCentralDB", _
                "Could not open the shared database after " & LOCK_MAX_RETRIES & _
                " attempts. Another user may be saving it right now. Please try again shortly." & _
                IIf(errMsg <> "", vbCrLf & vbCrLf & "Last error: " & errMsg, "")
        End If
        Application.Wait Now + TimeSerial(0, 0, LOCK_RETRY_WAIT_SEC)
    Loop

    Set OpenDBCore = wb
End Function

Public Function NextHistoryID(dbWb As Workbook) As Long
    Dim ws As Worksheet, lastRow As Long, idCol As Long
    Set ws = dbWb.Sheets(SHEET_HISTORY)
    idCol = GetColIndex(ws, "HistoryID")
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    If lastRow < 2 Then
        NextHistoryID = 1
    Else
        NextHistoryID = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, idCol), ws.Cells(lastRow, idCol))) + 1
    End If
End Function

' ---------------------------------------------------------------------
' Finds a claim's row by ID. Reads the whole ID column in ONE call
' into memory, then searches in-memory - not one worksheet touch per
' row. At a few thousand rows this is the difference between a
' noticeable pause and instant.
' ---------------------------------------------------------------------
Public Function FindClaimRow(dbWb As Workbook, ByVal claimID As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long, idCol As Long
    Dim idArr As Variant
    Dim target As String

    Set ws = dbWb.Sheets(SHEET_CLAIMS)
    idCol = GetColIndex(ws, "ClaimID")
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    If lastRow < 2 Then
        FindClaimRow = 0
        Exit Function
    End If

    target = Trim$(claimID)
    idArr = ws.Range(ws.Cells(2, idCol), ws.Cells(lastRow, idCol)).Value ' single bulk read

    ' A single-row result comes back as a plain value, not an array -
    ' handle that edge case (lastRow = 2) explicitly.
    If lastRow = 2 Then
        If Trim$(CStr(idArr)) = target Then FindClaimRow = 2
        Exit Function
    End If

    For i = 1 To UBound(idArr, 1)
        If Trim$(CStr(idArr(i, 1))) = target Then
            FindClaimRow = i + 1 ' +1 because the array started at sheet row 2
            Exit Function
        End If
    Next i
    FindClaimRow = 0 ' not found
End Function

' ---------------------------------------------------------------------
' Looks up a column's position by its header text in row 1, instead of
' a hard-coded number. This is what lets you add new columns to the
' Claims or History sheet later (e.g. Priority, ClaimAmount) without
' touching a single line of modDataAccess - only sheets that actually
' read/write that new column need a one-line addition.
' Raises a clear error if the header is missing/misspelled, so a typo
' fails loudly at the point of use rather than silently writing to the
' wrong cell.
'
' Cheap to call for single lookups (AddClaim, ChangeUpdatedSite - one
' row at a time). NEVER call this inside a per-row loop over many rows
' - use BuildHeaderMap once instead and index into the cached map. See
' clsClaimRepository for the pattern.
' ---------------------------------------------------------------------
Public Function GetColIndex(ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, i As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If Trim$(LCase$(CStr(ws.Cells(1, i).Value))) = Trim$(LCase$(headerName)) Then
            GetColIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, "GetColIndex", _
        "Column header '" & headerName & "' not found on sheet '" & ws.Name & "'. " & _
        "Check the header spelling in row 1."
End Function

' ---------------------------------------------------------------------
' Builds a header-name -> column-index map ONCE (one bulk read of row 1),
' for use when you're about to loop many rows and need several columns
' per row. Pass the returned Dictionary to ColIdx() instead of calling
' GetColIndex repeatedly inside the loop.
' ---------------------------------------------------------------------
Public Function BuildHeaderMap(ws As Worksheet) As Object
    Dim map As Object, lastCol As Long, i As Long
    Dim headerRow As Variant

    Set map = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value ' 1 bulk read

    If lastCol = 1 Then
        map(Trim$(LCase$(CStr(headerRow)))) = 1
    Else
        For i = 1 To lastCol
            map(Trim$(LCase$(CStr(headerRow(1, i))))) = i
        Next i
    End If
    Set BuildHeaderMap = map
End Function

Public Function ColIdx(headerMap As Object, ByVal headerName As String) As Long
    Dim key As String
    key = Trim$(LCase$(headerName))
    If Not headerMap.Exists(key) Then
        Err.Raise vbObjectError + 3, "ColIdx", _
            "Column header '" & headerName & "' not found in cached header map."
    End If
    ColIdx = headerMap(key)
End Function
