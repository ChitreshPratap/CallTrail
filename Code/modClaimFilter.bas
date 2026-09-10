Attribute VB_Name = "modClaimFilter"
Option Explicit

' =====================================================================
' modClaimFilter
' Pure filtering / shaping logic for claim collections. No UI, no
' worksheet access - it works on a Collection of clsClaim that someone
' else already loaded.
'
' Kept separate from the form on purpose: the dashboards you want to
' build later need exactly this same filtering, and they shouldn't have
' to open a form to get it.
' =====================================================================

' Which date field a filter applies to. Claims carry several dates and
' "filter by date" means different things depending on the question -
' when the claim arose, when it was keyed in, when it was last touched.
Public Enum ClaimDateField
    cdfCreationDate = 0
    cdfInsertionDate = 1
    cdfLastUpdated = 2
    cdfClosedDate = 3
End Enum

' ---------------------------------------------------------------------
' Filters a collection of clsClaim.
'
' Every filter is optional - pass "" / "All" for status and 0 for the
' dates to skip that criterion. With nothing supplied, everything comes
' back, which is the "no filter = show all records" behaviour.
'
' dateFrom / dateTo are INCLUSIVE. dateTo is compared at end-of-day so
' a claim stamped 14:30 on the To date isn't silently excluded - that
' off-by-one is easy to miss and makes users think records are missing.
' ---------------------------------------------------------------------
Public Function FilterClaims(Source As Collection, _
                              Optional ByVal statusFilter As String = "", _
                              Optional ByVal dateField As ClaimDateField = cdfCreationDate, _
                              Optional ByVal dateFrom As Date = 0, _
                              Optional ByVal dateTo As Date = 0, _
                              Optional ByVal searchText As String = "") As Collection
    Dim result As New Collection
    Dim c As clsClaim
    Dim keep As Boolean
    Dim v As Variant
    Dim d As Date
    Dim needle As String

    Set FilterClaims = result
    If Source Is Nothing Then Exit Function

    statusFilter = Trim$(statusFilter)
    needle = Trim$(LCase$(searchText))

    For Each c In Source
        keep = True

        ' --- status ---
        If keep And statusFilter <> "" And LCase$(statusFilter) <> "all" Then
            If LCase$(Trim$(c.ClaimStatus)) <> LCase$(statusFilter) Then keep = False
        End If

        ' --- date range ---
        If keep And (dateFrom > 0 Or dateTo > 0) Then
            v = GetDateValue(c, dateField)
            If IsEmpty(v) Or Not IsDate(v) Then
                ' A blank date can't satisfy a date range. This matters for
                ' ClaimClosedDate / LastUpdated, which are empty until the
                ' claim is actually worked - those rows drop out rather
                ' than appearing as if they matched.
                keep = False
            Else
                d = CDate(v)
                If dateFrom > 0 Then If d < dateFrom Then keep = False
                If keep And dateTo > 0 Then
                    ' +1 day minus a moment = inclusive of the whole To date
                    If d >= (dateTo + 1) Then keep = False
                End If
            End If
        End If

        ' --- free-text search across the fields a user would look in ---
        If keep And needle <> "" Then
            If InStr(1, LCase$(c.claimID), needle) = 0 _
               And InStr(1, LCase$(c.claimSite), needle) = 0 _
               And InStr(1, LCase$(c.ClaimProviderName), needle) = 0 _
               And InStr(1, LCase$(c.claimQuery), needle) = 0 Then
                keep = False
            End If
        End If

        If keep Then result.Add c
    Next c

    Set FilterClaims = result
End Function

Private Function GetDateValue(c As clsClaim, ByVal f As ClaimDateField) As Variant
    Select Case f
        Case cdfCreationDate:  GetDateValue = c.ClaimCreationDate
        Case cdfInsertionDate: GetDateValue = c.ClaimInsertionDate
        Case cdfLastUpdated:   GetDateValue = c.LastUpdatedDate
        Case cdfClosedDate:    GetDateValue = c.ClaimClosedDate
        Case Else:             GetDateValue = c.ClaimCreationDate
    End Select
End Function

' ---------------------------------------------------------------------
' Converts a claim collection into a 2D array shaped for a ListBox.
' Columns: ClaimID | Site | Provider | Status | Attempt | CreationDate
'
' Assigning ListBox.List = <array> in ONE go is dramatically faster than
' looping AddItem, which repaints the control on every single row.
' ---------------------------------------------------------------------
Public Function ClaimsToListArray(Source As Collection) As Variant
    Dim arr() As Variant
    Dim c As clsClaim
    Dim i As Long

    If Source Is Nothing Then Exit Function
    If Source.Count = 0 Then Exit Function

    ReDim arr(0 To Source.Count - 1, 0 To 5)  ' 0-based: ListBox expects this
    i = 0
    For Each c In Source
        arr(i, 0) = c.claimID
        arr(i, 1) = c.claimSite
        arr(i, 2) = c.ClaimProviderName
        arr(i, 3) = c.ClaimStatus
        arr(i, 4) = c.attempt
        arr(i, 5) = Format$(c.ClaimCreationDate, "dd-mmm-yyyy")
        i = i + 1
    Next c

    ClaimsToListArray = arr
End Function

' ---------------------------------------------------------------------
' Converts a history collection into a 2D array for a ListBox.
' Columns: CallerName | CallDateTime | Status | Comment
' ---------------------------------------------------------------------
Public Function HistoryToListArray(Source As Collection) As Variant
    Dim arr() As Variant
    Dim h As clsHistoryEntry
    Dim i As Long

    If Source Is Nothing Then Exit Function
    If Source.Count = 0 Then Exit Function

    ReDim arr(0 To Source.Count - 1, 0 To 3)
    i = 0
    For Each h In Source
        arr(i, 0) = h.CallerName
        arr(i, 1) = Format$(h.CallDateTime, "dd-mmm-yyyy hh:nn")
        arr(i, 2) = h.CallerStatus
        arr(i, 3) = h.CallerComment
        i = i + 1
    Next h

    HistoryToListArray = arr
End Function

' Small helper so the form doesn't repeat this parsing everywhere.
' Returns 0 for blank/invalid input, which the filter reads as "no limit".
Public Function ParseDateOrZero(ByVal s As String) As Date
    s = Trim$(s)
    If s = "" Then Exit Function
    If Not IsDate(s) Then Exit Function
    ParseDateOrZero = CDate(s)
End Function
