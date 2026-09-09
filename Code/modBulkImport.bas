Attribute VB_Name = "modBulkImport"
Option Explicit

' =====================================================================
' modBulkImport
' Bulk claim entry via the "bulkClaimAdd" sheet in the APP workbook.
'
' The flow the user sees is ONE action:
'   ProcessBulkClaims -> every row is validated
'                        valid rows   -> written to the central DB,
'                                        then DELETED from the sheet
'                        invalid rows -> LEFT in place, with the reason
'                                        written in the ValidationError
'                                        column so they can be fixed
'                                        and re-run later
'
' SetupBulkSheet (creates the sheet) and CheckBulkClaims (dry-run
' validation, changes nothing) are available as separate helpers.
'
' Design note: although the user experiences this as "check each row,
' move it, delete it", internally the valid rows are written to the
' shared workbook in ONE bulk operation. Writing row-by-row would mean
' opening and saving the shared file once per claim - 500 claims = 500
' network round trips and 500 chances to collide with another user.
' The visible result is identical; it's just far faster and safer.
' =====================================================================

Public Const BULK_SHEET As String = "bulkClaimAdd"

' Columns the user fills (the 5 mandatory fields).
' Everything else on the Claims sheet is auto-filled at insert time.
Private Const COL_ID As Long = 1
Private Const COL_SITE As Long = 2
Private Const COL_PROVIDER As Long = 3
Private Const COL_QUERY As Long = 4
Private Const COL_CREATED As Long = 5
Private Const COL_ERROR As Long = 6    ' validation output, not user input

' =====================================================================
' MAIN ENTRY POINT - validate, import the good rows, keep the bad ones
' =====================================================================
Public Sub ProcessBulkClaims()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim dataArr As Variant, errorArr() As Variant
    Dim repo As New clsClaimRepository
    Dim batch As New Collection
    Dim c As clsClaim
    Dim rowsToDelete As Range
    Dim problems As String
    Dim validCount As Long, invalidCount As Long, insertedCount As Long
    Dim skipped As String, resp As VbMsgBoxResult
    Dim seen As Object, existing As Object, idKey As String

    Set ws = GetBulkSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No claims to process. Enter your claims in the " & BULK_SHEET & _
               " sheet starting at row 2.", vbExclamation, "Nothing to Process"
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' --- one bulk read of everything the user entered ---
    dataArr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastRow, COL_CREATED)).Value
    ReDim errorArr(1 To UBound(dataArr, 1), 1 To 1)

    ' --- pull existing claim IDs from the DB ONCE, not once per row ---
    Set existing = CreateObject("Scripting.Dictionary")
    For Each c In repo.GetAllClaims()
        existing(Trim$(LCase$(c.ClaimID))) = True
    Next c
    Set seen = CreateObject("Scripting.Dictionary")

    ' --- validate every row, collecting the good ones into a batch ---
    For i = 1 To UBound(dataArr, 1)
        problems = ValidateRow(dataArr, i, seen, existing)

        If problems = "" Then
            Set c = New clsClaim
            c.ClaimID = Trim$(CStr(dataArr(i, COL_ID)))
            c.ClaimSite = Trim$(CStr(dataArr(i, COL_SITE)))
            c.ClaimProviderName = Trim$(CStr(dataArr(i, COL_PROVIDER)))
            c.ClaimQuery = Trim$(CStr(dataArr(i, COL_QUERY)))
            c.ClaimCreationDate = CDate(dataArr(i, COL_CREATED))
            batch.Add c

            idKey = Trim$(LCase$(c.ClaimID))
            seen(idKey) = True    ' so a later duplicate in the same sheet is caught

            errorArr(i, 1) = ""
            validCount = validCount + 1

            ' mark this sheet row for deletion once the insert succeeds
            If rowsToDelete Is Nothing Then
                Set rowsToDelete = ws.Rows(i + 1)
            Else
                Set rowsToDelete = Union(rowsToDelete, ws.Rows(i + 1))
            End If
        Else
            errorArr(i, 1) = problems
            invalidCount = invalidCount + 1
        End If
    Next i

    ' --- write all the error messages back in one go ---
    ws.Range(ws.Cells(2, COL_ERROR), ws.Cells(lastRow, COL_ERROR)).Value = errorArr
    ColorErrorColumn ws, lastRow
    Application.ScreenUpdating = True

    If validCount = 0 Then
        MsgBox "No valid rows to add." & vbCrLf & vbCrLf & _
               invalidCount & " row(s) have problems - see the ValidationError column.", _
               vbExclamation, "Nothing Imported"
        Exit Sub
    End If

    resp = MsgBox(validCount & " valid claim(s) will be added to the database and removed from this sheet." & vbCrLf & _
                  IIf(invalidCount > 0, invalidCount & " row(s) with errors will stay here for you to correct." & vbCrLf, "") & _
                  vbCrLf & "Continue?", vbYesNo + vbQuestion, "Confirm Bulk Add")
    If resp = vbNo Then Exit Sub

    ' --- single write of the whole batch to the central DB ---
    insertedCount = repo.AddClaimsBulk(batch, skipped)

    ' Only delete the sheet rows if the insert actually succeeded, and
    ' only if every claim we sent was accepted. If anything was skipped
    ' server-side, leave the rows alone rather than risk deleting a claim
    ' that never made it into the database.
    If insertedCount > 0 And skipped = "" Then
        Application.ScreenUpdating = False
        If Not rowsToDelete Is Nothing Then rowsToDelete.Delete
        Application.ScreenUpdating = True
    End If

    MsgBox insertedCount & " claim(s) added to the database." & vbCrLf & _
           IIf(skipped <> "", vbCrLf & "Skipped as already present: " & skipped & vbCrLf & _
                              "(rows left in the sheet - please review)" & vbCrLf, "") & _
           IIf(invalidCount > 0, vbCrLf & invalidCount & " row(s) with errors remain in " & BULK_SHEET & _
                                 " - correct them and run again.", ""), _
           vbInformation, "Bulk Add Complete"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Bulk add failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' Dry run - validate and flag everything, but change nothing.
' Useful before committing, and for checking corrections.
' =====================================================================
Public Sub CheckBulkClaims()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim dataArr As Variant, errorArr() As Variant
    Dim repo As New clsClaimRepository
    Dim c As clsClaim
    Dim seen As Object, existing As Object
    Dim problems As String
    Dim validCount As Long, invalidCount As Long

    Set ws = GetBulkSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No rows to check.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    dataArr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastRow, COL_CREATED)).Value
    ReDim errorArr(1 To UBound(dataArr, 1), 1 To 1)

    Set existing = CreateObject("Scripting.Dictionary")
    For Each c In repo.GetAllClaims()
        existing(Trim$(LCase$(c.ClaimID))) = True
    Next c
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(dataArr, 1)
        problems = ValidateRow(dataArr, i, seen, existing)
        If problems = "" Then
            errorArr(i, 1) = ""
            seen(Trim$(LCase$(CStr(dataArr(i, COL_ID))))) = True
            validCount = validCount + 1
        Else
            errorArr(i, 1) = problems
            invalidCount = invalidCount + 1
        End If
    Next i

    ws.Range(ws.Cells(2, COL_ERROR), ws.Cells(lastRow, COL_ERROR)).Value = errorArr
    ColorErrorColumn ws, lastRow
    Application.ScreenUpdating = True

    MsgBox validCount & " row(s) ready to add." & vbCrLf & _
           invalidCount & " row(s) with problems." & vbCrLf & vbCrLf & _
           IIf(invalidCount > 0, "See the ValidationError column for details.", "Run 'Add Claims to Database' to import them."), _
           IIf(invalidCount > 0, vbExclamation, vbInformation), "Check Complete"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Check failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' Creates the bulkClaimAdd sheet, or clears it if it already exists.
' Only ever touches the LOCAL app workbook, never the shared DB.
' =====================================================================
Public Sub SetupBulkSheet()
    Dim ws As Worksheet
    Dim resp As VbMsgBoxResult

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BULK_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BULK_SHEET
    Else
        If Application.WorksheetFunction.CountA(ws.Cells) > COL_ERROR Then
            resp = MsgBox("The " & BULK_SHEET & " sheet already has data. Clear it?" & vbCrLf & vbCrLf & _
                          "Any uncorrected rows will be lost.", _
                          vbYesNo + vbQuestion, "Reset Bulk Sheet")
            If resp = vbNo Then
                ws.Activate
                Exit Sub
            End If
        End If
        ws.Cells.Clear
    End If

    Application.ScreenUpdating = False

    ws.Cells(1, COL_ID).Value = "ClaimID"
    ws.Cells(1, COL_SITE).Value = "ClaimSite"
    ws.Cells(1, COL_PROVIDER).Value = "ClaimProviderName"
    ws.Cells(1, COL_QUERY).Value = "ClaimQuery"
    ws.Cells(1, COL_CREATED).Value = "ClaimCreationDate"
    ws.Cells(1, COL_ERROR).Value = "ValidationError"

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_ERROR))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(COL_ID).ColumnWidth = 16
    ws.Columns(COL_SITE).ColumnWidth = 18
    ws.Columns(COL_PROVIDER).ColumnWidth = 26
    ws.Columns(COL_QUERY).ColumnWidth = 40
    ws.Columns(COL_CREATED).ColumnWidth = 18
    ws.Columns(COL_CREATED).NumberFormat = "dd-mmm-yyyy"
    ws.Columns(COL_ERROR).ColumnWidth = 38
    ws.Rows(1).AutoFilter
    ws.Range("A2").Select
    ws.Activate

    Application.ScreenUpdating = True

    MsgBox "Enter or paste your claims starting at row 2 (columns A to E)." & vbCrLf & vbCrLf & _
           "All five fields are required." & vbCrLf & vbCrLf & _
           "Then click 'Add Claims to Database'. Valid rows are imported and " & _
           "removed from this sheet; rows with problems stay here with the reason " & _
           "shown in the ValidationError column.", vbInformation, "Bulk Sheet Ready"
End Sub

' =====================================================================
' Optional: load a CSV/xlsx of claims into bulkClaimAdd. Appends below
' whatever is already there, so multiple files can be stacked and any
' uncorrected rows aren't lost.
'
' This only LOADS - the rows still go through the same validation on
' import, so a bad file can't write junk into the central database.
' =====================================================================
Public Sub LoadClaimsFromFile()
    Dim ws As Worksheet, srcWb As Workbook, srcWs As Worksheet
    Dim filePath As Variant
    Dim lastRow As Long, srcLast As Long, pasteRow As Long

    Set ws = GetBulkSheet()
    If ws Is Nothing Then Exit Sub

    filePath = Application.GetOpenFilename( _
        "Claim files (*.csv;*.xlsx;*.xls),*.csv;*.xlsx;*.xls", , "Select a file of claims")
    If filePath = False Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set srcWb = Workbooks.Open(FileName:=CStr(filePath), ReadOnly:=True, UpdateLinks:=0)
    Set srcWs = srcWb.Sheets(1)
    srcLast = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    If srcLast < 2 Then
        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "That file has no data rows below the header.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    pasteRow = IIf(lastRow < 2, 2, lastRow + 1)

    ' single bulk transfer of columns A:E, skipping the source header row
    ws.Range(ws.Cells(pasteRow, COL_ID), ws.Cells(pasteRow + srcLast - 2, COL_CREATED)).Value = _
        srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(srcLast, 5)).Value

    srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ws.Activate
    MsgBox (srcLast - 1) & " row(s) loaded into " & BULK_SHEET & "." & vbCrLf & vbCrLf & _
           "Click 'Add Claims to Database' to validate and import them.", _
           vbInformation, "File Loaded"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Could not load that file: " & Err.Description, vbCritical
End Sub

' =====================================================================
' Validation rules for a single staged row.
' Returns "" if the row is good, or a semicolon-separated list of
' problems. Add new rules here - one place, used by both the dry-run
' check and the real import, so the two can never disagree.
' =====================================================================
Private Function ValidateRow(dataArr As Variant, ByVal i As Long, _
                              seen As Object, existing As Object) As String
    Dim problems As String, idKey As String

    If Trim$(CStr(dataArr(i, COL_ID))) = "" Then problems = AddIssue(problems, "ClaimID missing")
    If Trim$(CStr(dataArr(i, COL_SITE))) = "" Then problems = AddIssue(problems, "ClaimSite missing")
    If Trim$(CStr(dataArr(i, COL_PROVIDER))) = "" Then problems = AddIssue(problems, "ProviderName missing")
    If Trim$(CStr(dataArr(i, COL_QUERY))) = "" Then problems = AddIssue(problems, "ClaimQuery missing")

    If Trim$(CStr(dataArr(i, COL_CREATED))) = "" Then
        problems = AddIssue(problems, "CreationDate missing")
    ElseIf Not IsDate(dataArr(i, COL_CREATED)) Then
        problems = AddIssue(problems, "CreationDate is not a valid date")
    ElseIf CDate(dataArr(i, COL_CREATED)) > Date Then
        problems = AddIssue(problems, "CreationDate is in the future")
    End If

    idKey = Trim$(LCase$(CStr(dataArr(i, COL_ID))))
    If idKey <> "" Then
        If seen.Exists(idKey) Then problems = AddIssue(problems, "Duplicate ClaimID in this sheet")
        If existing.Exists(idKey) Then problems = AddIssue(problems, "ClaimID already in database")
    End If

    ValidateRow = problems
End Function

' =====================================================================
' Helpers
' =====================================================================
Private Function GetBulkSheet() As Worksheet
    On Error Resume Next
    Set GetBulkSheet = ThisWorkbook.Sheets(BULK_SHEET)
    On Error GoTo 0
    If GetBulkSheet Is Nothing Then
        MsgBox "The " & BULK_SHEET & " sheet doesn't exist yet." & vbCrLf & vbCrLf & _
               "Run 'Setup Bulk Sheet' to create it.", vbExclamation, "Sheet Not Found"
    End If
End Function

Private Function AddIssue(ByVal soFar As String, ByVal issue As String) As String
    AddIssue = soFar & IIf(soFar = "", "", "; ") & issue
End Function

Private Sub ColorErrorColumn(ws As Worksheet, ByVal lastRow As Long)
    Dim i As Long
    With ws.Range(ws.Cells(2, COL_ERROR), ws.Cells(lastRow, COL_ERROR))
        .Interior.ColorIndex = xlNone
        .Font.Color = RGB(150, 0, 0)
    End With
    For i = 2 To lastRow
        If ws.Cells(i, COL_ERROR).Value <> "" Then
            ws.Cells(i, COL_ERROR).Interior.Color = RGB(255, 214, 214)  ' soft red
        End If
    Next i
End Sub
