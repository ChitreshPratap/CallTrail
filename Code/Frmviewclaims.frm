VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frmviewclaims 
   Caption         =   "UserForm1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14850
   OleObjectBlob   =   "Frmviewclaims.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frmviewclaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================================
' frmViewClaims - code-behind
'
' Paste ALL of this into the code module of a UserForm named
' frmViewClaims. Control names must match the layout table in the
' Setup Guide, section 6.
'
' Performance approach: the claims are loaded from the shared database
' ONCE (into m_allClaims) when the form opens. Every filter change then
' works on that in-memory copy - no re-reading the shared file, so
' filtering is instant and the shared workbook isn't repeatedly locked
' while someone browses. Refresh re-reads deliberately.
' =====================================================================

Option Explicit

Private m_allClaims As Collection      ' everything loaded from the DB
Private m_shownClaims As Collection    ' what's currently in the list
Private m_loading As Boolean           ' suppresses events during setup

' Guard rail: a UserForm ListBox becomes sluggish and eventually fails
' well before Excel's row limit. Rather than hang, we cap what's shown
' and tell the user to narrow the filter.
Private Const MAX_LIST_ROWS As Long = 10000

' =====================================================================
' Form setup
' =====================================================================

Private Sub UserForm_Initialize()
    
    m_loading = True

    Me.caption = "CallTrail - View Claims"

    ' Status filter: blank/All means no status filtering
    cboStatus.List = Array("All", "Pending", "Closed")
    cboStatus.value = "All"

    ' Which date the range applies to
    cboDateField.List = Array("Creation Date", "Insertion Date", "Last Updated", "Closed Date")
    cboDateField.value = "Creation Date"

    
    ' Claims list: 6 columns
    With lstClaims
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "70;80;140;60;50;80"
        .MultiSelect = fmMultiSelectSingle
    End With

    ' History list for the selected claim: 4 columns
    With lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "80;110;55;220"
    End With

    ClearDetailPanel

    m_loading = False

    LoadClaimsFromDatabase
    
    ApplyFilters
    
End Sub

' =====================================================================
' Data loading
' =====================================================================
Private Sub LoadClaimsFromDatabase()
    
    Dim repo As ClsClaimRepository
    Set repo = New ClsClaimRepository

    On Error GoTo Fail
    Me.MousePointer = fmMousePointerHourGlass
    lblStatusBar.caption = "Loading claims..."
    DoEvents

    Set m_allClaims = repo.GetAllClaims()

    Me.MousePointer = fmMousePointerDefault
    Exit Sub

Fail:
    Me.MousePointer = fmMousePointerDefault
    Set m_allClaims = New Collection
    lblStatusBar.caption = "Could not load claims: " & Err.Description
End Sub

' =====================================================================
' Filtering - runs entirely in memory
' =====================================================================
Private Sub ApplyFilters()
    Dim dFrom As Date, dTo As Date
    Dim listData As Variant
    Dim shownCount As Long

    If m_loading Then Exit Sub
    If m_allClaims Is Nothing Then Exit Sub

    dFrom = ParseDateOrZero(txtDateFrom.value)
    dTo = ParseDateOrZero(txtDateTo.value)

    ' Warn on unparseable input rather than silently ignoring it - a user
    ' who typed a date and sees it quietly do nothing assumes the filter
    ' is broken.
    If Trim$(txtDateFrom.value) <> "" And dFrom = 0 Then
        lblStatusBar.caption = "'From' date not recognised - ignoring it."
    ElseIf Trim$(txtDateTo.value) <> "" And dTo = 0 Then
        lblStatusBar.caption = "'To' date not recognised - ignoring it."
    End If

    If dFrom > 0 And dTo > 0 And dFrom > dTo Then
        MsgBox "The 'From' date is after the 'To' date.", vbExclamation, "Check Dates"
        Exit Sub
    End If

    Set m_shownClaims = FilterClaims(m_allClaims, _
                                      CStr(cboStatus.value), _
                                      SelectedDateField(), _
                                      dFrom, dTo, _
                                      CStr(txtSearch.value))

    ClearDetailPanel
    lstClaims.Clear

    shownCount = m_shownClaims.Count

    If shownCount = 0 Then
        lblStatusBar.caption = "No claims match the current filter."
        lblCount.caption = "0 of " & m_allClaims.Count
        Exit Sub
    End If

    If shownCount > MAX_LIST_ROWS Then
        lblStatusBar.caption = shownCount & " matches - too many to display. " & _
                               "Showing the first " & MAX_LIST_ROWS & "; narrow the filter to see the rest."
        Set m_shownClaims = TakeFirst(m_shownClaims, MAX_LIST_ROWS)
    Else
        lblStatusBar.caption = "Ready."
    End If

    listData = ClaimsToListArray(m_shownClaims)
    If Not IsEmpty(listData) Then lstClaims.List = listData   ' single assignment
        
    lblCount.caption = m_shownClaims.Count & " of " & m_allClaims.Count
    
    
    Dim myHeaders As Variant
    
    ' 3. Define your headers in an Array
    myHeaders = Array("Emp ID", "Full Name", "Department")
    
    ' 4. Call the helper function
    CreateListBoxHeaders lstClaims, myHeaders
    
End Sub

Private Function SelectedDateField() As ClaimDateField

    Select Case CStr(cboDateField.value)
        Case "Insertion Date": SelectedDateField = cdfInsertionDate
        Case "Last Updated":   SelectedDateField = cdfLastUpdated
        Case "Closed Date":    SelectedDateField = cdfClosedDate
        Case Else:             SelectedDateField = cdfCreationDate
    End Select
    
End Function

Private Function TakeFirst(source As Collection, ByVal n As Long) As Collection
    Dim result As New Collection
    Dim i As Long, limit As Long

    limit = n
    If source.Count < limit Then limit = source.Count

    For i = 1 To limit
        result.Add source(i)
    Next i
    Set TakeFirst = result
End Function

' =====================================================================
' Selection - show full detail for the highlighted claim
' =====================================================================
Private Sub lstClaims_Click()
    Dim idx As Long
    Dim c As clsClaim

    idx = lstClaims.ListIndex
    If idx < 0 Then
        ClearDetailPanel
        Exit Sub
    End If

    ' ListIndex is 0-based; the Collection is 1-based
    Set c = m_shownClaims(idx + 1)
    ShowClaimDetail c
End Sub

Private Sub ShowClaimDetail(c As clsClaim)
    
    lblClaimID.caption = c.claimID
    lblSite.caption = c.claimSite
    lblProvider.caption = c.ClaimProviderName
    lblQuery.caption = c.claimQuery
    lblStatus.caption = c.ClaimStatus
    lblAttempt.caption = CStr(c.attempt)
    lblCreated.caption = Format$(c.ClaimCreationDate, "dd-mmm-yyyy")
    lblUpdatedSite.caption = c.ClaimUpdatedSite
    lblInsertedOn.caption = Format$(c.ClaimInsertionDate, "dd-mmm-yyyy hh:nn")
    lblInsertedBy.caption = c.ClaimInsertedBy
    lblLastUpdated.caption = FormatDateOrBlank(c.LastUpdatedDate)
    lblLastUpdatedBy.caption = c.LastUpdatedBy
    lblLastComment.caption = c.LastComment
    lblClosedOn.caption = FormatDateOrBlank(c.ClaimClosedDate)
    lblDaysOpen.caption = CStr(c.DaysOpen) & IIf(c.IsClosed, " (to close)", " (open)")
    LoadHistoryFor c.claimID
    
End Sub

Private Sub LoadHistoryFor(ByVal claimID As String)
    Dim repo As New ClsClaimRepository
    Dim histData As Variant

    lstHistory.Clear

    On Error GoTo Fail
    Me.MousePointer = fmMousePointerHourGlass

    histData = HistoryToListArray(repo.GetHistory(claimID))
    If Not IsEmpty(histData) Then lstHistory.List = histData

    Me.MousePointer = fmMousePointerDefault
    Exit Sub

Fail:
    Me.MousePointer = fmMousePointerDefault
    lblStatusBar.caption = "Could not load call history: " & Err.Description
End Sub

Private Sub ClearDetailPanel()
    
    lblClaimID.caption = ""
    lblSite.caption = ""
    lblProvider.caption = ""
    lblQuery.caption = ""
    lblStatus.caption = ""
    lblAttempt.caption = ""
    lblCreated.caption = ""
    lblUpdatedSite.caption = ""
    lblInsertedOn.caption = ""
    lblInsertedBy.caption = ""
    lblLastUpdated.caption = ""
    lblLastUpdatedBy.caption = ""
    lblLastComment.caption = ""
    lblClosedOn.caption = ""
    lblDaysOpen.caption = ""
    lstHistory.Clear
    
End Sub


Private Function FormatDateOrBlank(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    FormatDateOrBlank = Format$(CDate(v), "dd-mmm-yyyy hh:nn")
End Function

' =====================================================================
' Buttons
' =====================================================================
Private Sub cmdApply_Click()
    ApplyFilters
End Sub

Private Sub cmdClear_Click()
    m_loading = True
    cboStatus.value = "All"
    cboDateField.value = "Creation Date"
    txtDateFrom.value = ""
    txtDateTo.value = ""
    txtSearch.value = ""
    m_loading = False
    ApplyFilters
End Sub

' Re-reads from the shared database. Needed because the in-memory copy
' won't show claims other callers added since this form was opened.
Private Sub cmdRefresh_Click()
    LoadClaimsFromDatabase
    ApplyFilters
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' Live-filter as the user types in the search box
Private Sub txtSearch_Change()
    ApplyFilters
End Sub

' Enter key in either date box applies the filter
Private Sub txtDateFrom_KeyDown(ByVal KeyCode As Msforms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplyFilters
End Sub

Private Sub txtDateTo_KeyDown(ByVal KeyCode As Msforms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplyFilters
End Sub

Private Sub cboStatus_Change()
    ApplyFilters
End Sub

Private Sub cboDateField_Change()
    ApplyFilters
End Sub

Private Sub CreateListBoxHeaders(lst As Msforms.ListBox, headerNames As Variant)
    Dim i As Integer
    Dim lbl As Msforms.Label
    Dim currentLeft As Single
    Dim colWidths() As String
    Dim singleWidth As Single
    Dim ctrl As control
    
    ' 1. Delete any existing dynamic headers (prevents duplicates if code runs twice)
    For Each ctrl In Me.Controls
        If Left(ctrl.Name, 10) = "dynHeader_" Then Me.Controls.Remove ctrl.Name
    Next ctrl

    ' 2. Start positioning at the left edge of the ListBox
    currentLeft = lst.Left
    
    ' 3. Parse the ColumnWidths property (e.g., "50;100;75")
    If lst.ColumnWidths <> "" Then
        colWidths = Split(lst.ColumnWidths, ";")
    End If

    ' 4. Create a label for each header
    For i = LBound(headerNames) To UBound(headerNames)
        
        ' Determine the width for this specific column
        If lst.ColumnWidths <> "" And i <= UBound(colWidths) Then
            singleWidth = Val(colWidths(i)) ' Val ignores the " pt" text if present
        Else
            ' Fallback: evenly divide the ListBox width if ColumnWidths aren't set
            singleWidth = lst.width / (UBound(headerNames) - LBound(headerNames) + 1)
        End If

        ' Add the label control to the UserForm dynamically
        Set lbl = Me.Controls.Add("Forms.Label.1", "dynHeader_" & i, True)
        
        ' Format and position the label
        With lbl
            .caption = " " & headerNames(i)
            .Left = currentLeft
            .Top = lst.Top - 15          ' Place it 15 points above the ListBox
            .width = singleWidth
            .height = 15
            .BackColor = &H8000000F      ' Standard grey button-face color
            .SpecialEffect = fmSpecialEffectSunken
            .Font.Bold = True
            .Font.Size = 9
        End With

        ' Move the starting position for the next label
        currentLeft = currentLeft + singleWidth
    Next i
End Sub
