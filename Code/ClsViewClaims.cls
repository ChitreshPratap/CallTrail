VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsViewClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private var_frm As UserFormMain

'Initialization of Tab - 'View Claims'
Private var_m_allClaims As Collection      ' everything loaded from the DB
Private var_m_shownClaims As Collection    ' what's currently in the list
Private var_m_loading As Boolean           ' suppresses events during setup

' Guard rail: a UserForm ListBox becomes sluggish and eventually fails
' well before Excel's row limit. Rather than hang, we cap what's shown
' and tell the user to narrow the filter.
Private Const MAX_LIST_ROWS As Long = 10000



Public Property Set setForm(uform As UserFormMain)
    
    Set var_frm = uform
    
End Property

Sub initialize_viewClaimsTab()
    
    Dim uform As UserFormMain
    Set uform = var_frm
    
    var_m_loading = True

'    Me.Caption = "CallTrail - View Claims"

    ' Status filter: blank/All means no status filtering
    uform.cboStatus.List = Array("All", "Pending", "Closed")
    uform.cboStatus.Value = "All"

    ' Which date the range applies to
    uform.cboDateField.List = Array("Creation Date", "Insertion Date", "Last Updated", "Closed Date")
    uform.cboDateField.Value = "Creation Date"

    
    ' Claims list: 6 columns
    With uform.lstClaims
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "70;80;140;60;50;80"
        .MultiSelect = fmMultiSelectSingle
    End With

    ' History list for the selected claim: 4 columns
    With uform.lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "80;110;55;220"
    End With
    
    'Uncomment
'    ClearDetailPanel

    var_m_loading = False

    Me.LoadClaimsFromDatabase
    
    ApplyFilters

End Sub




Public Sub LoadClaimsFromDatabase()
    
    Dim uform As UserFormMain
    Set uform = var_frm
    
    Dim repo As ClsClaimRepository
    Set repo = New ClsClaimRepository

    On Error GoTo Fail
    'Me.MousePointer = fmMousePointerHourGlass
    uform.MousePointer = fmMousePointerHourGlass
    uform.lblStatusBar.Caption = "Loading claims..."
    DoEvents

    Set var_m_allClaims = repo.GetAllClaims()

    uform.MousePointer = fmMousePointerDefault
    Exit Sub

Fail:
    uform.MousePointer = fmMousePointerDefault
    Set var_m_allClaims = New Collection
    uform.lblStatusBar.Caption = "Could not load claims: " & Err.Description
End Sub


' =====================================================================
' Filtering - runs entirely in memory
' =====================================================================
Private Sub ApplyFilters()
        
    Dim uform As UserFormMain
    Set uform = var_frm
    
    Dim dFrom As Date, dTo As Date
    Dim listData As Variant
    Dim shownCount As Long

    If var_m_loading Then Exit Sub
    If var_m_allClaims Is Nothing Then Exit Sub

    dFrom = ParseDateOrZero(uform.txtDateFrom.Value)
    dTo = ParseDateOrZero(uform.txtDateTo.Value)

    ' Warn on unparseable input rather than silently ignoring it - a user
    ' who typed a date and sees it quietly do nothing assumes the filter
    ' is broken.
    If Trim$(uform.txtDateFrom.Value) <> "" And dFrom = 0 Then
        uform.lblStatusBar.Caption = "'From' date not recognised - ignoring it."
    ElseIf Trim$(uform.txtDateTo.Value) <> "" And dTo = 0 Then
        uform.lblStatusBar.Caption = "'To' date not recognised - ignoring it."
    End If

    If dFrom > 0 And dTo > 0 And dFrom > dTo Then
        MsgBox "The 'From' date is after the 'To' date.", vbExclamation, "Check Dates"
        Exit Sub
    End If

    Set var_m_shownClaims = FilterClaims(var_m_allClaims, _
                                      CStr(uform.cboStatus.Value), _
                                      SelectedDateField(), _
                                      dFrom, dTo, _
                                      CStr(uform.txtSearch.Value))
    'Uncomment
'    ClearDetailPanel
    uform.lstClaims.Clear

    shownCount = var_m_shownClaims.Count

    If shownCount = 0 Then
        uform.lblStatusBar.Caption = "No claims match the current filter."
        uform.lblCount.Caption = "0 of " & var_m_allClaims.Count
        Exit Sub
    End If

    If shownCount > MAX_LIST_ROWS Then
        uform.lblStatusBar.Caption = shownCount & " matches - too many to display. " & _
                               "Showing the first " & MAX_LIST_ROWS & "; narrow the filter to see the rest."
        Set var_m_shownClaims = TakeFirst(var_m_shownClaims, MAX_LIST_ROWS)
    Else
        uform.lblStatusBar.Caption = "Ready."
    End If

    listData = ClaimsToListArray(var_m_shownClaims)
    If Not IsEmpty(listData) Then uform.lstClaims.List = listData   ' single assignment
        
    uform.lblCount.Caption = var_m_shownClaims.Count & " of " & var_m_allClaims.Count
    
    
    Dim myHeaders As Variant
    
    ' 3. Define your headers in an Array
    myHeaders = Array("Emp ID", "Full Name", "Department")
    
    ' 4. Call the helper function
    CreateListBoxHeaders uform.lstClaims, myHeaders
    
End Sub

Private Function SelectedDateField() As ClaimDateField
    
    Dim uform As UserForm
    Set uform = var_frm
    
    Select Case CStr(uform.cboDateField.Value)
        Case "Insertion Date": SelectedDateField = cdfInsertionDate
        Case "Last Updated":   SelectedDateField = cdfLastUpdated
        Case "Closed Date":    SelectedDateField = cdfClosedDate
        Case Else:             SelectedDateField = cdfCreationDate
    End Select
    
End Function

Private Function TakeFirst(Source As Collection, ByVal n As Long) As Collection
    Dim result As New Collection
    Dim i As Long, limit As Long

    limit = n
    If Source.Count < limit Then limit = Source.Count

    For i = 1 To limit
        result.Add Source(i)
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
    Set c = var_m_shownClaims(idx + 1)
    ShowClaimDetail c
End Sub

Private Sub ShowClaimDetail(c As clsClaim)
    
    lblClaimID.Caption = c.claimID
    lblSite.Caption = c.claimSite
    lblProvider.Caption = c.ClaimProviderName
    lblQuery.Caption = c.claimQuery
    lblStatus.Caption = c.ClaimStatus
    lblAttempt.Caption = CStr(c.attempt)
    lblCreated.Caption = Format$(c.ClaimCreationDate, "dd-mmm-yyyy")
    lblUpdatedSite.Caption = c.ClaimUpdatedSite
    lblInsertedOn.Caption = Format$(c.ClaimInsertionDate, "dd-mmm-yyyy hh:nn")
    lblInsertedBy.Caption = c.ClaimInsertedBy
    lblLastUpdated.Caption = FormatDateOrBlank(c.LastUpdatedDate)
    lblLastUpdatedBy.Caption = c.LastUpdatedBy
    lblLastComment.Caption = c.LastComment
    lblClosedOn.Caption = FormatDateOrBlank(c.ClaimClosedDate)
    lblDaysOpen.Caption = CStr(c.DaysOpen) & IIf(c.IsClosed, " (to close)", " (open)")
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
    lblStatusBar.Caption = "Could not load call history: " & Err.Description
End Sub

Private Sub ClearDetailPanel()
    
    lblClaimID.Caption = ""
    lblSite.Caption = ""
    lblProvider.Caption = ""
    lblQuery.Caption = ""
    lblStatus.Caption = ""
    lblAttempt.Caption = ""
    lblCreated.Caption = ""
    lblUpdatedSite.Caption = ""
    lblInsertedOn.Caption = ""
    lblInsertedBy.Caption = ""
    lblLastUpdated.Caption = ""
    lblLastUpdatedBy.Caption = ""
    lblLastComment.Caption = ""
    lblClosedOn.Caption = ""
    lblDaysOpen.Caption = ""
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
    var_m_loading = True
    cboStatus.Value = "All"
    cboDateField.Value = "Creation Date"
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    txtSearch.Value = ""
    var_m_loading = False
    ApplyFilters
End Sub

' Re-reads from the shared database. Needed because the in-memory copy
' won't show claims other callers added since this form was opened.
Private Sub cmdRefresh_Click()
    LoadClaimsFromDatabase
    ApplyFilters
End Sub

'Private Sub cmdClose_Click()
'    Unload Me
'End Sub

' Live-filter as the user types in the search box
Private Sub txtSearch_Change()
    ApplyFilters
End Sub




' Enter key in either date box applies the filter
Private Sub txtDateFrom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplyFilters
End Sub

Private Sub txtDateTo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplyFilters
End Sub

Private Sub cboStatus_Change()
    ApplyFilters
End Sub

Private Sub cboDateField_Change()
    ApplyFilters
End Sub

Private Sub CreateListBoxHeaders(lst As MSForms.ListBox, headerNames As Variant)
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim currentLeft As Single
    Dim colWidths() As String
    Dim singleWidth As Single
    Dim ctrl As control
    
    Dim uform As UserFormMain
    Set uform = UserFormMain
    
    ' 1. Delete any existing dynamic headers (prevents duplicates if code runs twice)
    For Each ctrl In uform.Controls
        If Left(ctrl.Name, 10) = "dynHeader_" Then uform.Controls.Remove ctrl.Name
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
            singleWidth = lst.Width / (UBound(headerNames) - LBound(headerNames) + 1)
        End If

        ' Add the label control to the UserForm dynamically
        Set lbl = uform.Controls.Add("Forms.Label.1", "dynHeader_" & i, True)
        
        ' Format and position the label
        With lbl
            .Caption = " " & headerNames(i)
            .Left = currentLeft
            .Top = lst.Top - 15          ' Place it 15 points above the ListBox
            .Width = singleWidth
            .Height = 15
            .BackColor = &H8000000F      ' Standard grey button-face color
            .SpecialEffect = fmSpecialEffectSunken
            .Font.Bold = True
            .Font.Size = 9
        End With

        ' Move the starting position for the next label
        currentLeft = currentLeft + singleWidth
    Next i
End Sub




