Attribute VB_Name = "modRepair"
Option Explicit

' =====================================================================
' modRepair
' One-off cleanup tools for a database left in a bad state.
'
' RemoveBlankRows exists because an earlier version of the archive
' routine cleared row contents without deleting the rows. That left
' blank rows inside the Claims and History sheets, which then produced:
'
'   "Could not read claims. This key is already associated with an
'    element of this collection."
'
' - because every blank row yields an empty ClaimID, and two empty keys
' collide in a keyed Collection.
'
' The archive routine no longer does this, and the readers now skip
' blank rows regardless. This macro cleans up databases that were
' already affected. Run it once; it is safe to run again.
' =====================================================================

Public Sub RemoveBlankRows()
    Dim wb As Workbook
    Dim nClaims As Long, nHist As Long
    Dim resp As VbMsgBoxResult

    resp = MsgBox("Scan the central database for blank rows in Claims and History " & _
                  "and remove them?" & vbCrLf & vbCrLf & _
                  "A row counts as blank when its key column (ClaimID / HistoryID) " & _
                  "is empty. Rows with real data are not touched.", _
                  vbYesNo + vbQuestion, "Remove Blank Rows")
    If resp = vbNo Then Exit Sub

    On Error GoTo Fail
    Set wb = OpenCentralDB()

    If Not IsCurrentUserAdmin(wb) Then
        MsgBox "Only Admin users can run database repairs.", vbExclamation, "Access Denied"
        CloseCentralDB wb, False
        Exit Sub
    End If

    nClaims = StripBlanks(wb.Sheets(SHEET_CLAIMS), "ClaimID")
    nHist = StripBlanks(wb.Sheets(SHEET_HISTORY), "HistoryID")

    CloseCentralDB wb, (nClaims + nHist > 0)

    MsgBox "Blank rows removed:" & vbCrLf & vbCrLf & _
           SHEET_CLAIMS & ": " & nClaims & vbCrLf & _
           SHEET_HISTORY & ": " & nHist & vbCrLf & vbCrLf & _
           IIf(nClaims + nHist = 0, "Nothing needed fixing.", "Reload the app to pick up the change."), _
           vbInformation, "Repair Complete"
    Exit Sub

Fail:
    MsgBox "Repair failed: " & Err.Description, vbCritical
    If Not wb Is Nothing Then CloseCentralDB wb, False
End Sub

' Rewrites the sheet with only the rows that have a non-empty key,
' then deletes the surplus and resizes any Table. Same approach the
' archive service now uses.
Private Function StripBlanks(ws As Worksheet, ByVal keyHeader As String) As Long
    Dim hmap As Object
    Dim lastRow As Long, lastCol As Long, keyCol As Long
    Dim data As Variant, kept() As Variant
    Dim i As Long, k As Long, n As Long, removed As Long

    Set hmap = BuildHeaderMap(ws)
    keyCol = ColIdx(hmap, keyHeader)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).value
    ReDim kept(1 To UBound(data, 1), 1 To lastCol)

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, keyCol) & "")) <> "" Then
            n = n + 1
            For k = 1 To lastCol
                kept(n, k) = data(i, k)
            Next k
        Else
            removed = removed + 1
        End If
    Next i

    If removed = 0 Then Exit Function

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
    If n > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, lastCol)).value = kept
    End If
    If n + 2 <= lastRow Then
        ws.Rows((n + 2) & ":" & lastRow).Delete Shift:=xlUp
    End If

    ResizeTable ws, n, lastCol
    StripBlanks = removed
End Function

Private Sub ResizeTable(ws As Worksheet, ByVal n As Long, ByVal cols As Long)
    Dim lo As ListObject
    Dim newLast As Long

    On Error Resume Next
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If lo Is Nothing Then Exit Sub

    newLast = n + 1
    If newLast < 2 Then newLast = 2      ' a Table needs one body row
    lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(newLast, cols))
    On Error GoTo 0
End Sub
