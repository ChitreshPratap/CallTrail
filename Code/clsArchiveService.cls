VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArchiveService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsArchiveService
' Moves old claims - and their call history - out of the working sheets
' into ArchivedClaims / ArchivedHistory, so Claims and History stay
' small and fast.
'
' THIS IS THE MOST DESTRUCTIVE OPERATION IN THE APP. It deletes rows
' from the live sheets. The design reflects that:
'
'  * WRITE FIRST, DELETE SECOND, with a save in between. If the process
'    dies halfway, you are left with rows duplicated into the archive -
'    annoying, recoverable. The other order would lose them outright.
'
'  * Claims and their history move TOGETHER. Archiving a claim and
'    leaving its history behind produces orphan rows that belong to a
'    claim no longer on the sheet.
'
'  * Preview before commit. CountMatching answers "what would this do"
'    without touching anything, so the admin sees the number before
'    agreeing to it.
'
'  * Admin only, enforced here rather than in the form.
'
'  * Closed-only by DEFAULT. Archiving a Pending claim hides live work
'    from the callers who are meant to be chasing it. Archiving one is
'    possible but has to be asked for explicitly.
' =====================================================================

' ---------------------------------------------------------------------
' Criteria. Both filters are optional and combine with AND.
'   olderThanMonths = 0   -> no age filter
'   closedOnly = True     -> only Closed claims (the safe default)
' ---------------------------------------------------------------------
Private m_olderThanMonths As Long
Private m_closedOnly As Boolean

Private Sub Class_Initialize()
    m_olderThanMonths = 6
    m_closedOnly = True
End Sub

Public Property Let OlderThanMonths(ByVal v As Long)
    m_olderThanMonths = v
End Property
Public Property Get OlderThanMonths() As Long
    OlderThanMonths = m_olderThanMonths
End Property

Public Property Let ClosedOnly(ByVal v As Boolean)
    m_closedOnly = v
End Property
Public Property Get ClosedOnly() As Boolean
    ClosedOnly = m_closedOnly
End Property

' ---------------------------------------------------------------------
' Does a claim match the current criteria?
'
' Age is measured from the CLOSED date where there is one, falling back
' to the creation date. That matters: a claim raised 8 months ago but
' only closed last week is still recent work, and measuring from
' creation alone would archive it the moment it closed.
' ---------------------------------------------------------------------
Public Function Matches(c As clsClaim) As Boolean
    Dim cutoff As Date
    Dim refDate As Date

    Matches = False

    If m_closedOnly Then
        If Not c.IsClosed Then Exit Function
    End If

    If m_olderThanMonths > 0 Then
        cutoff = DateAdd("m", -m_olderThanMonths, Date)

        If c.IsClosed And Not IsEmpty(c.ClaimClosedDate) Then
            If IsDate(c.ClaimClosedDate) Then
                refDate = CDate(c.ClaimClosedDate)
            Else
                refDate = c.ClaimCreationDate
            End If
        Else
            refDate = c.ClaimCreationDate
        End If

        If refDate > cutoff Then Exit Function
    End If

    Matches = True
End Function

' ---------------------------------------------------------------------
' Preview: how many claims would be archived, without changing anything.
' ---------------------------------------------------------------------
Public Function CountMatching(claims As Collection, Optional ByRef oldestDate As Variant) As Long
    Dim c As clsClaim
    Dim n As Long
    Dim oldest As Date

    If claims Is Nothing Then Exit Function

    For Each c In claims
        If Matches(c) Then
            n = n + 1
            If oldest = 0 Or c.ClaimCreationDate < oldest Then oldest = c.ClaimCreationDate
        End If
    Next c

    If n > 0 Then oldestDate = oldest
    CountMatching = n
End Function

' ---------------------------------------------------------------------
' Preview: the actual claims that would be archived, so the admin can
' see and inspect them rather than trusting a bare count.
'
' Returns full clsClaim objects, so the caller can show any column and
' can look up each claim's history.
' ---------------------------------------------------------------------
Public Function GetMatching(claims As Collection) As Collection
    Dim result As New Collection
    Dim c As clsClaim

    Set GetMatching = result
    If claims Is Nothing Then Exit Function

    For Each c In claims
        If Matches(c) Then result.Add c
    Next c

    Set GetMatching = result
End Function

' ---------------------------------------------------------------------
' Does the move. Returns the number of claims archived.
' historyMoved reports how many history rows went with them.
' ---------------------------------------------------------------------
Public Function ArchiveMatching(Optional ByRef historyMoved As Long) As Long
    Dim wb As Workbook
    Dim wsClaims As Worksheet, wsHist As Worksheet
    Dim wsArcClaims As Worksheet, wsArcHist As Worksheet
    Dim hmapC As Object, hmapH As Object
    Dim toArchive As Object                  ' claim IDs, for O(1) lookup
    Dim claimsData As Variant, histData As Variant
    Dim arcClaims() As Variant, keepClaims() As Variant
    Dim arcHist() As Variant, keepHist() As Variant
    Dim lastRowC As Long, lastColC As Long, lastRowH As Long, lastColH As Long
    Dim i As Long, k As Long
    Dim nArc As Long, nKeep As Long, nArcH As Long, nKeepH As Long
    Dim stamp As Date, userName As String
    Dim c As clsClaim
    Dim repo As New ClsClaimRepository

    On Error GoTo Fail
    Set wb = OpenCentralDB()

    If Not IsCurrentUserAdmin(wb) Then
        MsgBox "Only Admin users can archive claims.", vbExclamation, "Access Denied"
        CloseCentralDB wb, False
        Exit Function
    End If

    Set wsClaims = wb.Sheets(SHEET_CLAIMS)
    Set wsHist = wb.Sheets(SHEET_HISTORY)
    Set wsArcClaims = wb.Sheets(SHEET_ARC_CLAIMS)
    Set wsArcHist = wb.Sheets(SHEET_ARC_HISTORY)
    Set hmapC = BuildHeaderMap(wsClaims)
    Set hmapH = BuildHeaderMap(wsHist)

    lastRowC = wsClaims.Cells(wsClaims.Rows.Count, ColIdx(hmapC, "ClaimID")).End(xlUp).Row
    lastColC = wsClaims.Cells(1, wsClaims.Columns.Count).End(xlToLeft).Column
    If lastRowC < 2 Then
        CloseCentralDB wb, False
        Exit Function
    End If

    stamp = Now
    userName = GetWindowsUserName()

    ' --- decide what moves, using the in-memory claim objects ---
    Set toArchive = CreateObject("Scripting.Dictionary")
    claimsData = wsClaims.Range(wsClaims.Cells(2, 1), wsClaims.Cells(lastRowC, lastColC)).value

    ReDim arcClaims(1 To UBound(claimsData, 1), 1 To lastColC + 2)   ' +2 for archive stamp
    ReDim keepClaims(1 To UBound(claimsData, 1), 1 To lastColC)

    For i = 1 To UBound(claimsData, 1)
        Set c = ClaimFromRow(claimsData, i, hmapC)

        If Matches(c) Then
            nArc = nArc + 1
            For k = 1 To lastColC
                arcClaims(nArc, k) = claimsData(i, k)
            Next k
            arcClaims(nArc, lastColC + 1) = stamp
            arcClaims(nArc, lastColC + 2) = userName
            toArchive(Trim$(LCase$(CStr(c.claimID)))) = True
        Else
            nKeep = nKeep + 1
            For k = 1 To lastColC
                keepClaims(nKeep, k) = claimsData(i, k)
            Next k
        End If
    Next i

    If nArc = 0 Then
        CloseCentralDB wb, False
        ArchiveMatching = 0
        Exit Function
    End If

    ' --- split history the same way ---
    lastRowH = wsHist.Cells(wsHist.Rows.Count, ColIdx(hmapH, "HistoryID")).End(xlUp).Row
    lastColH = wsHist.Cells(1, wsHist.Columns.Count).End(xlToLeft).Column

    If lastRowH >= 2 Then
        histData = wsHist.Range(wsHist.Cells(2, 1), wsHist.Cells(lastRowH, lastColH)).value
        ReDim arcHist(1 To UBound(histData, 1), 1 To lastColH + 2)
        ReDim keepHist(1 To UBound(histData, 1), 1 To lastColH)

        For i = 1 To UBound(histData, 1)
            If toArchive.Exists(Trim$(LCase$(CStr(histData(i, ColIdx(hmapH, "ClaimID")))))) Then
                nArcH = nArcH + 1
                For k = 1 To lastColH
                    arcHist(nArcH, k) = histData(i, k)
                Next k
                arcHist(nArcH, lastColH + 1) = stamp
                arcHist(nArcH, lastColH + 2) = userName
            Else
                nKeepH = nKeepH + 1
                For k = 1 To lastColH
                    keepHist(nKeepH, k) = histData(i, k)
                Next k
            End If
        Next i
    End If

    ' =================================================================
    ' PHASE 1 - write to the archive sheets and SAVE.
    ' Nothing is deleted yet. If anything fails from here on, the worst
    ' case is rows present in both places, which is recoverable.
    ' =================================================================
    AppendBlock wsArcClaims, arcClaims, nArc, lastColC + 2
    If nArcH > 0 Then AppendBlock wsArcHist, arcHist, nArcH, lastColH + 2
    wb.save

    ' =================================================================
    ' PHASE 2 - rewrite the working sheets with only the kept rows.
    '
    ' Rewriting beats deleting rows one by one: deleting hundreds of
    ' non-contiguous rows is slow and shifts everything underneath each
    ' time. This clears the data range and writes the survivors back in
    ' one block.
    ' =================================================================
    RewriteSheet wsClaims, keepClaims, nKeep, lastColC, lastRowC
    If lastRowH >= 2 Then RewriteSheet wsHist, keepHist, nKeepH, lastColH, lastRowH

    CloseCentralDB wb, True

    historyMoved = nArcH
    ArchiveMatching = nArc
    Exit Function

Fail:
    MsgBox "Archive failed: " & Err.Description & vbCrLf & vbCrLf & _
           "If this happened after the archive was written, some claims may now " & _
           "appear in BOTH the working and archive sheets. Nothing has been lost - " & _
           "check ArchivedClaims before retrying.", vbCritical, "Archive Error"
    If Not wb Is Nothing Then CloseCentralDB wb, False
    ArchiveMatching = 0
End Function

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------

' Builds a claim object from one row of an already-read array, so the
' matching rules can use clsClaim.IsClosed etc. rather than re-deriving
' them from raw cells.
Private Function ClaimFromRow(data As Variant, ByVal i As Long, hmap As Object) As clsClaim
    Dim c As New clsClaim
    On Error Resume Next
    c.claimID = data(i, ColIdx(hmap, "ClaimID"))
    c.ClaimStatus = data(i, ColIdx(hmap, "ClaimStatus"))
    c.ClaimCreationDate = data(i, ColIdx(hmap, "ClaimCreationDate"))
    c.ClaimClosedDate = data(i, ColIdx(hmap, "ClaimClosedDate"))
    On Error GoTo 0
    Set ClaimFromRow = c
End Function

' Appends n rows to the bottom of a sheet in one write.
Private Sub AppendBlock(ws As Worksheet, data As Variant, ByVal n As Long, ByVal cols As Long)
    Dim firstRow As Long
    Dim trimmed() As Variant
    Dim r As Long, k As Long

    If n = 0 Then Exit Sub

    ' Trim to exactly n rows so the array and the target range match,
    ' rather than relying on Excel taking the top-left block.
    ReDim trimmed(1 To n, 1 To cols)
    For r = 1 To n
        For k = 1 To cols
            trimmed(r, k) = data(r, k)
        Next k
    Next r

    firstRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If firstRow < 2 Then firstRow = 2
    ws.Range(ws.Cells(firstRow, 1), ws.Cells(firstRow + n - 1, cols)).value = trimmed
End Sub

' Replaces the sheet's data with only the kept rows.
'
' ClearContents alone is NOT enough, and that was a real bug: it blanks
' the cells but leaves the rows there. On a sheet carrying an Excel Table
' (Claims and History both do) the Table keeps its old size with empty
' rows inside it, and every later read then picks up blank rows. A blank
' row yields an empty ClaimID, and two of those collide when claims are
' keyed by ID - producing "This key is already associated with an element
' of this collection".
'
' So: write the survivors, then DELETE the surplus rows outright, then
' resize any ListObject to match.
Private Sub RewriteSheet(ws As Worksheet, data As Variant, ByVal n As Long, _
                          ByVal cols As Long, ByVal oldLastRow As Long)
    Dim trimmed() As Variant
    Dim r As Long, k As Long
    Dim firstSurplus As Long

    If oldLastRow < 2 Then Exit Sub

    ' --- write the kept rows back at the top ---
    If n > 0 Then
        ReDim trimmed(1 To n, 1 To cols)
        For r = 1 To n
            For k = 1 To cols
                trimmed(r, k) = data(r, k)
            Next k
        Next r
        ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, cols)).value = trimmed
    End If

    ' --- delete whatever is left below them ---
    firstSurplus = n + 2
    If firstSurplus <= oldLastRow Then
        ws.Rows(firstSurplus & ":" & oldLastRow).Delete Shift:=xlUp
    End If

    ResizeListObject ws, n, cols
End Sub

' If the sheet has an Excel Table, point it at the new data extent.
' A Table left spanning deleted rows reports phantom empty rows to
' anything that reads it through the ListObject.
Private Sub ResizeListObject(ws As Worksheet, ByVal n As Long, ByVal cols As Long)
    Dim lo As ListObject
    Dim newLast As Long

    On Error Resume Next
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If lo Is Nothing Then Exit Sub

    ' A Table must keep at least one body row, so never shrink past row 2
    newLast = n + 1
    If newLast < 2 Then newLast = 2

    lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(newLast, cols))
    On Error GoTo 0
End Sub

