Attribute VB_Name = "modDownload"
Option Explicit

' =====================================================================
' modDownload
' Pulls a full copy of the central database into THIS workbook (the
' local macro file), one sheet per source sheet.
'
' What it's for: offline analysis, pivot tables, ad-hoc reporting, and
' taking a snapshot before a risky operation like an archive run.
'
' Three things worth knowing about how it works:
'
'  1. READ-ONLY OPEN. A download only reads, so it opens the shared file
'     without taking a write lock. A read/write open would block callers
'     trying to log a call for as long as the copy takes - which, on a
'     large database over a share drive, is exactly when you don't want
'     to be holding the file.
'
'  2. ONE OPEN for all four sheets. Opening per sheet would mean four
'     round trips to the network share for no benefit.
'
'  3. BULK ARRAY TRANSFER, not copy/paste. Range.Copy carries formatting,
'     uses the clipboard (which the user can clobber mid-run) and is far
'     slower. Reading a range into a Variant and writing it back moves
'     values only, in two operations per sheet.
'
' The copies are a SNAPSHOT. They do not refresh and they are not
' written back - editing them changes nothing in the central database.
' =====================================================================

' Local sheet names. Prefixed so they can't be confused with the live
' sheets if someone opens both workbooks side by side.
Private Const LOCAL_CLAIMS As String = "DL_Claims"
Private Const LOCAL_HISTORY As String = "DL_History"
Private Const LOCAL_ARC_CLAIMS As String = "DL_ArchivedClaims"
Private Const LOCAL_ARC_HISTORY As String = "DL_ArchivedHistory"
Private Const LOCAL_INFO As String = "DL_Info"

' =====================================================================
' MAIN ENTRY POINT - wire a button to this
' =====================================================================
Public Sub DownloadAllData()
    Dim wb As Workbook
    Dim rowsClaims As Long, rowsHist As Long
    Dim rowsArcClaims As Long, rowsArcHist As Long
    Dim startedAt As Date
    Dim resp As VbMsgBoxResult

    resp = MsgBox("Download a full copy of the central database into this workbook?" & vbCrLf & vbCrLf & _
                  "Four sheets will be created or REPLACED:" & vbCrLf & _
                  "  " & LOCAL_CLAIMS & vbCrLf & _
                  "  " & LOCAL_HISTORY & vbCrLf & _
                  "  " & LOCAL_ARC_CLAIMS & vbCrLf & _
                  "  " & LOCAL_ARC_HISTORY & vbCrLf & vbCrLf & _
                  "Any changes you made to those sheets will be lost.", _
                  vbYesNo + vbQuestion, "Download Database")
    If resp = vbNo Then Exit Sub

    On Error GoTo Fail
    startedAt = Now

    Application.ScreenUpdating = False
    Application.StatusBar = "Opening the central database..."

    ' Read-only: no write lock, so callers aren't blocked while this runs
    Set wb = OpenCentralDBReadOnly()

    Application.StatusBar = "Copying claims..."
    rowsClaims = CopySheet(wb, SHEET_CLAIMS, LOCAL_CLAIMS)

    Application.StatusBar = "Copying call history..."
    rowsHist = CopySheet(wb, SHEET_HISTORY, LOCAL_HISTORY)

    Application.StatusBar = "Copying archived claims..."
    rowsArcClaims = CopySheet(wb, SHEET_ARC_CLAIMS, LOCAL_ARC_CLAIMS)

    Application.StatusBar = "Copying archived history..."
    rowsArcHist = CopySheet(wb, SHEET_ARC_HISTORY, LOCAL_ARC_HISTORY)

    CloseCentralDB wb, False        ' never save - we only read

    WriteInfoSheet startedAt, rowsClaims, rowsHist, rowsArcClaims, rowsArcHist

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ThisWorkbook.Sheets(LOCAL_INFO).Activate

    MsgBox "Download complete." & vbCrLf & vbCrLf & _
           LOCAL_CLAIMS & ": " & Format$(rowsClaims, "#,##0") & " row(s)" & vbCrLf & _
           LOCAL_HISTORY & ": " & Format$(rowsHist, "#,##0") & " row(s)" & vbCrLf & _
           LOCAL_ARC_CLAIMS & ": " & Format$(rowsArcClaims, "#,##0") & " row(s)" & vbCrLf & _
           LOCAL_ARC_HISTORY & ": " & Format$(rowsArcHist, "#,##0") & " row(s)" & vbCrLf & vbCrLf & _
           "Took " & Format$((Now - startedAt) * 86400, "0") & " second(s)." & vbCrLf & vbCrLf & _
           "This is a snapshot. It won't refresh, and editing it does not " & _
           "change the central database.", _
           vbInformation, "Download Complete"
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wb Is Nothing Then CloseCentralDB wb, False
    On Error GoTo 0
    MsgBox "Download failed: " & Err.Description, vbCritical, "Download Error"
End Sub

' =====================================================================
' Download just one sheet - useful when you only need claims, and the
' history sheet is the big one.
' =====================================================================
Public Sub DownloadClaimsOnly()
    DownloadSingle SHEET_CLAIMS, LOCAL_CLAIMS
End Sub

Public Sub DownloadHistoryOnly()
    DownloadSingle SHEET_HISTORY, LOCAL_HISTORY
End Sub

Public Sub DownloadArchivesOnly()
    Dim wb As Workbook
    Dim n1 As Long, n2 As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wb = OpenCentralDBReadOnly()
    n1 = CopySheet(wb, SHEET_ARC_CLAIMS, LOCAL_ARC_CLAIMS)
    n2 = CopySheet(wb, SHEET_ARC_HISTORY, LOCAL_ARC_HISTORY)
    CloseCentralDB wb, False
    Application.ScreenUpdating = True
    MsgBox "Archived claims: " & Format$(n1, "#,##0") & " row(s)" & vbCrLf & _
           "Archived history: " & Format$(n2, "#,##0") & " row(s)", _
           vbInformation, "Download Complete"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wb Is Nothing Then CloseCentralDB wb, False
    On Error GoTo 0
    MsgBox "Download failed: " & Err.Description, vbCritical
End Sub

Private Sub DownloadSingle(ByVal sourceSheet As String, ByVal targetSheet As String)
    Dim wb As Workbook
    Dim n As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wb = OpenCentralDBReadOnly()
    n = CopySheet(wb, sourceSheet, targetSheet)
    CloseCentralDB wb, False
    Application.ScreenUpdating = True

    ThisWorkbook.Sheets(targetSheet).Activate
    MsgBox targetSheet & ": " & Format$(n, "#,##0") & " row(s) downloaded.", _
           vbInformation, "Download Complete"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wb Is Nothing Then CloseCentralDB wb, False
    On Error GoTo 0
    MsgBox "Download failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' The actual copy. Returns the number of DATA rows copied (excluding
' the header).
' =====================================================================
Private Function CopySheet(srcWb As Workbook, ByVal srcName As String, _
                            ByVal destName As String) As Long
    Dim src As Worksheet, dest As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant

    On Error GoTo NoSource
    Set src = srcWb.Sheets(srcName)
    On Error GoTo 0

    Set dest = GetOrCreateSheet(destName)
    dest.Cells.Clear

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column

    If lastCol = 0 Then Exit Function

    ' Excel can't hold more rows than its own limit. Hitting this means
    ' the archive threshold is far too lax - say so rather than failing
    ' with a cryptic subscript error.
    If lastRow > dest.Rows.Count Then
        Err.Raise vbObjectError + 30, "CopySheet", _
            "'" & srcName & "' has " & Format$(lastRow, "#,##0") & " rows, which is more " & _
            "than a worksheet can hold. Archive older records first."
    End If

    ' ONE read, ONE write - no clipboard, no per-cell loop
    data = src.Range(src.Cells(1, 1), src.Cells(lastRow, lastCol)).Value
    dest.Range(dest.Cells(1, 1), dest.Cells(lastRow, lastCol)).Value = data

    FormatHeader dest, lastCol
    dest.Rows(1).AutoFilter
    dest.Cells.EntireColumn.AutoFit

    CopySheet = IIf(lastRow > 1, lastRow - 1, 0)
    Exit Function

NoSource:
    ' A missing archive sheet is normal on an older database file, so
    ' don't fail the whole download over it.
    Err.Clear
    CopySheet = 0
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function

Private Sub FormatHeader(ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).AutoFilter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' =====================================================================
' A small sheet recording WHEN this snapshot was taken and by whom.
' Without it, a downloaded sheet sitting in a workbook for a week looks
' exactly like one pulled five minutes ago - and someone will report
' from it believing it's current.
' =====================================================================
Private Sub WriteInfoSheet(ByVal takenAt As Date, ByVal nClaims As Long, ByVal nHist As Long, _
                            ByVal nArcClaims As Long, ByVal nArcHist As Long)
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(LOCAL_INFO)
    ws.Cells.Clear

    ws.Range("A1").Value = "Database Snapshot"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "Downloaded"
    ws.Range("B3").Value = Format$(takenAt, "dd-mmm-yyyy hh:nn:ss")
    ws.Range("A4").Value = "Downloaded by"
    ws.Range("B4").Value = GetWindowsUserName()
    ws.Range("A5").Value = "Source"
    ws.Range("B5").Value = DB_PATH

    ws.Range("A7").Value = "Sheet"
    ws.Range("B7").Value = "Rows"
    ws.Range("A7:B7").Font.Bold = True

    ws.Range("A8").Value = LOCAL_CLAIMS:       ws.Range("B8").Value = nClaims
    ws.Range("A9").Value = LOCAL_HISTORY:      ws.Range("B9").Value = nHist
    ws.Range("A10").Value = LOCAL_ARC_CLAIMS:  ws.Range("B10").Value = nArcClaims
    ws.Range("A11").Value = LOCAL_ARC_HISTORY: ws.Range("B11").Value = nArcHist

    ws.Range("A13").Value = "This is a point-in-time snapshot. It does not refresh, and " & _
                            "editing these sheets does not change the central database."
    ws.Range("A13").Font.Italic = True

    ws.Columns("A:B").AutoFit
End Sub
