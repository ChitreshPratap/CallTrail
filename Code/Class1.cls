VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False




Attribute VB_Name = "MatchCustomerOwner"
Option Explicit

'=====================================================================
' MatchCustomerToOwner
' Matches "Customer name" (sheet "Excel data") against "Owner Name"
' (sheet "Brass data") and writes the matched owner name(s) into a
' "Matched Owner Name" column on "Excel data".
'
' Rules:
'   - First name + last name must match (order can be reversed)
'   - Middle name optional; if both sides have one, initials must agree
'   - Extra spaces, periods, apostrophes, case are ignored
'   - Commas understood as "Last, First Middle"
'   - Suffixes/titles ignored: JR SR II III IV ETAL MR MRS MS DR
'   - Joint owners ("A & B", "A AND B", "A / B") are checked separately
'   - Multiple matches are joined with " | "
'=====================================================================

Private Const SRC_SHEET As String = "Excel data"
Private Const SRC_HDR As String = "Customer name"
Private Const LKP_SHEET As String = "Brass data"
Private Const LKP_HDR As String = "Owner Name"
Private Const RESULT_HDR As String = "Matched Owner Name"
Private Const NO_MATCH As String = "No Match"

Public Sub MatchCustomerToOwner()
    Dim wb As Workbook, wsC As Worksheet, wsO As Worksheet
    Dim cCol As Long, oCol As Long, rCol As Long
    Dim lastC As Long, lastO As Long
    Dim cVals As Variant, oVals As Variant, res() As Variant
    Dim idx As Object, found As Object, entries As Collection
    Dim parts As Variant, keys As Variant, kv As Variant, e As Variant
    Dim i As Long, p As Long, k As Long, matched As Long
    Dim nm As String, pairK As String, midI As String

    Set wb = ThisWorkbook   ' change to ActiveWorkbook if macro is stored elsewhere
    On Error Resume Next
    Set wsC = wb.Worksheets(SRC_SHEET)
    Set wsO = wb.Worksheets(LKP_SHEET)
    On Error GoTo 0
    If wsC Is Nothing Or wsO Is Nothing Then
        MsgBox "Sheet '" & SRC_SHEET & "' or '" & LKP_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    cCol = FindHeader(wsC, SRC_HDR)
    oCol = FindHeader(wsO, LKP_HDR)
    If cCol = 0 Or oCol = 0 Then
        MsgBox "Header '" & SRC_HDR & "' or '" & LKP_HDR & "' not found in row 1.", vbExclamation
        Exit Sub
    End If

    lastC = wsC.Cells(wsC.Rows.Count, cCol).End(xlUp).Row
    lastO = wsO.Cells(wsO.Rows.Count, oCol).End(xlUp).Row
    If lastC < 2 Or lastO < 2 Then
        MsgBox "No data rows found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- 1. Build index of owner names: pairKey -> list of (row, middleInitial)
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbBinaryCompare
    oVals = ColValues(wsO, oCol, lastO)

    For i = 1 To UBound(oVals, 1)
        If Not IsError(oVals(i, 1)) Then
            nm = Trim$(CStr(oVals(i, 1)))
            If Len(nm) > 0 Then
                parts = SplitParts(nm)
                For p = 0 To UBound(parts)
                    keys = NameKeys(CStr(parts(p)))
                    For k = 0 To UBound(keys)
                        kv = Split(keys(k), "#")
                        If Not idx.Exists(kv(0)) Then idx.Add kv(0), New Collection
                        idx(kv(0)).Add Array(i, kv(1))
                    Next k
                Next p
            End If
        End If
    Next i

    '--- 2. Match each customer
    cVals = ColValues(wsC, cCol, lastC)
    ReDim res(1 To UBound(cVals, 1), 1 To 1)

    For i = 1 To UBound(cVals, 1)
        res(i, 1) = ""
        If Not IsError(cVals(i, 1)) Then
            nm = Trim$(CStr(cVals(i, 1)))
            If Len(nm) > 0 Then
                Set found = CreateObject("Scripting.Dictionary")
                found.CompareMode = vbTextCompare
                parts = SplitParts(nm)
                For p = 0 To UBound(parts)
                    keys = NameKeys(CStr(parts(p)))
                    For k = 0 To UBound(keys)
                        kv = Split(keys(k), "#")
                        pairK = kv(0): midI = kv(1)
                        If idx.Exists(pairK) Then
                            Set entries = idx(pairK)
                            For Each e In entries
                                If MiddleOK(midI, CStr(e(1))) Then
                                    nm = Trim$(CStr(oVals(e(0), 1)))
                                    If Not found.Exists(nm) Then found.Add nm, True
                                End If
                            Next e
                        End If
                    Next k
                Next p
                If found.Count > 0 Then
                    res(i, 1) = Join(found.keys, " | ")
                    matched = matched + 1
                Else
                    res(i, 1) = NO_MATCH
                End If
            End If
        End If
    Next i

    '--- 3. Write results
    rCol = FindHeader(wsC, RESULT_HDR)
    If rCol = 0 Then
        rCol = wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column + 1
        wsC.Cells(1, rCol).Value = RESULT_HDR
        wsC.Cells(1, rCol).Font.Bold = True
    End If
    wsC.Range(wsC.Cells(2, rCol), wsC.Cells(lastC, rCol)).Value = res

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done. " & matched & " of " & (lastC - 1) & " customers matched.", vbInformation
End Sub

'--- Returns 2D (1 To n, 1 To 1) array of a column from row 2 to lastRow
Private Function ColValues(ws As Worksheet, col As Long, lastRow As Long) As Variant
    Dim v As Variant
    If lastRow = 2 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = ws.Cells(2, col).Value2
    Else
        v = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).Value2
    End If
    ColValues = v
End Function

'--- Finds header in row 1 (case/space insensitive)
Private Function FindHeader(ws As Worksheet, hdr As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Application.WorksheetFunction.Trim(CStr(ws.Cells(1, c).Value)), _
                   hdr, vbTextCompare) = 0 Then
            FindHeader = c
            Exit Function
        End If
    Next c
End Function

'--- Splits joint names: "A & B", "A AND B", "A / B", "A; B"
Private Function SplitParts(ByVal nm As String) As Variant
    Dim s As String
    s = " " & UCase$(Replace(nm, Chr(160), " ")) & " "
    s = Replace(s, " AND ", "&")
    s = Replace(s, "/", "&")
    s = Replace(s, ";", "&")
    s = Replace(s, "+", "&")
    SplitParts = Split(s, "&")
End Function

'--- Builds match keys "FIRST|LAST#MiddleInitial" (pair sorted so order doesn't matter)
Private Function NameKeys(ByVal nm As String) As Variant
    Dim s As String, outS As String, m As String
    Dim lastT As Variant, firstT As Variant, t As Variant
    Dim pos() As Long, n As Long, i As Long, fPos As Long
    Dim cPos As Long

    s = CleanName(nm)
    cPos = InStr(s, ",")

    '--- "Last, First Middle" format
    If cPos > 0 Then
        lastT = Tokens(Left$(s, cPos - 1))
        firstT = Tokens(Replace(Mid$(s, cPos + 1), ",", " "))
        If UBound(lastT) >= 0 And UBound(firstT) >= 0 Then
            fPos = 0
            For i = 0 To UBound(firstT)
                If Len(firstT(i)) > 1 Then fPos = i: Exit For
            Next i
            m = ""
            If fPos < UBound(firstT) Then m = Left$(firstT(fPos + 1), 1)
            NameKeys = Split(PairKey(CStr(firstT(fPos)), CStr(lastT(UBound(lastT)))) & "#" & m, ";")
            Exit Function
        End If
        s = Replace(s, ",", " ")
    End If

    '--- No comma: could be "First Middle Last" or "Last First Middle"
    t = Tokens(s)
    If UBound(t) < 1 Then NameKeys = Split("", ";"): Exit Function

    ReDim pos(0 To UBound(t))
    n = 0
    For i = 0 To UBound(t)
        If Len(t(i)) > 1 Then pos(n) = i: n = n + 1
    Next i
    If n < 2 Then NameKeys = Split("", ";"): Exit Function

    ' Candidate A: First [Middle] Last
    m = ""
    If pos(n - 1) - pos(0) > 1 Then m = Left$(t(pos(0) + 1), 1)
    outS = PairKey(CStr(t(pos(0))), CStr(t(pos(n - 1)))) & "#" & m

    ' Candidate B: Last First [Middle]
    m = ""
    If pos(1) < UBound(t) Then m = Left$(t(pos(1) + 1), 1)
    outS = outS & ";" & PairKey(CStr(t(pos(0))), CStr(t(pos(1)))) & "#" & m

    NameKeys = Split(outS, ";")
End Function

'--- Uppercase, keep letters/comma/hyphen, drop . and ', everything else -> space
Private Function CleanName(ByVal nm As String) As String
    Dim i As Long, ch As String, s As String, outS As String
    s = UCase$(nm)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "A" To "Z", ",", "-": outS = outS & ch
            Case ".", "'", "`", Chr(146): ' drop
            Case Else: outS = outS & " "
        End Select
    Next i
    CleanName = outS
End Function

'--- Splits on spaces (extra spaces collapsed) and removes suffixes/titles
Private Function Tokens(ByVal s As String) As Variant
    Dim raw As Variant, i As Long, outS As String
    raw = Split(Application.WorksheetFunction.Trim(s), " ")
    For i = 0 To UBound(raw)
        If Len(raw(i)) > 0 And Not IsNoise(CStr(raw(i))) Then
            If Len(outS) > 0 Then outS = outS & " "
            outS = outS & raw(i)
        End If
    Next i
    Tokens = Split(outS, " ")
End Function

Private Function IsNoise(ByVal w As String) As Boolean
    Select Case w
        Case "JR", "SR", "II", "III", "IV", "ETAL", "ET", "AL", "MR", "MRS", "MS", "DR"
            IsNoise = True
    End Select
End Function

Private Function PairKey(ByVal a As String, ByVal b As String) As String
    If StrComp(a, b, vbBinaryCompare) <= 0 Then
        PairKey = a & "|" & b
    Else
        PairKey = b & "|" & a
    End If
End Function

'--- Middle is optional; if both present, initials must match
Private Function MiddleOK(ByVal a As String, ByVal b As String) As Boolean
    MiddleOK = (Len(a) = 0 Or Len(b) = 0 Or a = b)
End Function
