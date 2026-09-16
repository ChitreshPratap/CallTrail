VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPageViewRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageView - controller for the "View Claims" tab.

' Browse and filter everything in the database. Filtering runs entirely
' in memory against the shared cache in clsAppContext, so changing a
' filter never touches the network.
' =====================================================================
Implements IPage

Private m_ctx As clsAppContext
Private m_pg As MSForms.Page
Private m_shown As Collection
Private m_shown_history As Collection
Private m_loading As Boolean
Private m_arranged As Boolean

Private Const MAX_LIST_ROWS As Long = 10000

Private WithEvents cboStatus As MSForms.ComboBox
Attribute cboStatus.VB_VarHelpID = -1
'Attribute cboStatus.VB_VarHelpID = -1
Private WithEvents cboDateField As MSForms.ComboBox
Attribute cboDateField.VB_VarHelpID = -1
'Attribute cboDateField.VB_VarHelpID = -1
Private WithEvents txtSearch As MSForms.TextBox
Attribute txtSearch.VB_VarHelpID = -1
'Attribute txtSearch.VB_VarHelpID = -1

Private WithEvents cmdApply As MSForms.Label
Attribute cmdApply.VB_VarHelpID = -1
'Attribute cmdApply.VB_VarHelpID = -1
Private WithEvents cmdClear As MSForms.Label
Attribute cmdClear.VB_VarHelpID = -1
'Attribute cmdClear.VB_VarHelpID = -1
Private WithEvents cmdRefresh As MSForms.Label
Attribute cmdRefresh.VB_VarHelpID = -1
'Attribute cmdRefresh.VB_VarHelpID = -1
Private WithEvents lstClaims As MSForms.ListBox
Attribute lstClaims.VB_VarHelpID = -1
'Attribute lstClaims.VB_VarHelpID = -1
Private WithEvents lstHistory As MSForms.ListBox
Attribute lstHistory.VB_VarHelpID = -1
'Attribute lstClaims.VB_VarHelpID = -1

Private txtDateFrom As MSForms.TextBox
Private txtDateTo As MSForms.TextBox
'Private lstHistory As MSForms.ListBox
Private lblCount As MSForms.Label

' Detail labels
Private frameDetails As MSForms.Frame
Private multiPageDetails As MSForms.MultiPage
Private lstViewHistory As MSForms.ListBox
Private lblClaimID As MSForms.Label
Private lblClaimSite As MSForms.Label
Private lblProviderName As MSForms.Label
Private lblQuery As MSForms.Label
Private lblStatus As MSForms.Label
Private lblAttempt As MSForms.Label
Private lblCreated As MSForms.Label
Private lblUpdatedClaimSite As MSForms.Label
Private lblInsertedOn As MSForms.Label
Private lblInsertedBy As MSForms.Label
Private lblLastUpdated As MSForms.Label
Private lblLastUpdatedBy As MSForms.Label
Private lblLastComment As MSForms.Label
Private lblClosedOn As MSForms.Label
Private lblDaysOpen As MSForms.Label


' Detail history
Private lblHistoryId As MSForms.Label
Private lblHistoryClaimId As MSForms.Label
Private lblHistoryCallerName As MSForms.Label
Private lblHistoryCallDate As MSForms.Label
Private lblHistoryCallComment As MSForms.Label
Private lblHistoryCallStatus As MSForms.Label

Private collectionStyledControls As Collection

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "View Claims"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As MSForms.Page)
    
    Set m_ctx = ctx
    Set m_pg = pg
    m_loading = True
    
    Set collectionStyledControls = New Collection

    Set cboStatus = BindControl(pg, "cboViewStatus", IPage_PageTitle)
    Set cboDateField = BindControl(pg, "cboViewDateField", IPage_PageTitle)
    Set txtDateFrom = BindControl(pg, "txtViewDateFrom", IPage_PageTitle)
    Set txtDateTo = BindControl(pg, "txtViewDateTo", IPage_PageTitle)
    Set txtSearch = BindControl(pg, "txtViewSearch", IPage_PageTitle)
    Set cmdApply = BindControl(pg, "cmdViewApply", IPage_PageTitle)
    Set cmdClear = BindControl(pg, "cmdViewClear", IPage_PageTitle)
    Set cmdRefresh = BindControl(pg, "cmdViewRefresh", IPage_PageTitle)
    Set lstClaims = BindControl(pg, "lstViewClaims", IPage_PageTitle)
    Set lstHistory = BindControl(pg, "lstViewHistory", IPage_PageTitle)
    Set lblCount = BindControl(pg, "lblViewCount", IPage_PageTitle)

    Set lblClaimID = BindControl(pg, "lblViewClaimID", IPage_PageTitle)
    Set lblClaimSite = BindControl(pg, "lblViewClaimSite", IPage_PageTitle)
    Set lblProviderName = BindControl(pg, "lblViewProviderName", IPage_PageTitle)
    Set lblQuery = BindControl(pg, "lblViewQuery", IPage_PageTitle)
    Set lblStatus = BindControl(pg, "lblViewStatus", IPage_PageTitle)
    Set lblAttempt = BindControl(pg, "lblViewAttempt", IPage_PageTitle)
    Set lblCreated = BindControl(pg, "lblViewCreated", IPage_PageTitle)
    Set lblUpdatedClaimSite = BindControl(pg, "lblViewUpdatedClaimSite", IPage_PageTitle)
    Set lblInsertedOn = BindControl(pg, "lblViewInsertedOn", IPage_PageTitle)
    Set lblInsertedBy = BindControl(pg, "lblViewInsertedBy", IPage_PageTitle)
    Set lblLastUpdated = BindControl(pg, "lblViewLastUpdated", IPage_PageTitle)
    Set lblLastUpdatedBy = BindControl(pg, "lblViewLastUpdatedBy", IPage_PageTitle)
    Set lblLastComment = BindControl(pg, "lblViewLastComment", IPage_PageTitle)
    Set lblClosedOn = BindControl(pg, "lblViewClosedOn", IPage_PageTitle)
    Set lblDaysOpen = BindControl(pg, "lblViewDaysOpen", IPage_PageTitle)


    Set lblHistoryId = BindControl(pg, "lblHistoryId", IPage_PageTitle)
    Set lblHistoryClaimId = BindControl(pg, "lblHistoryClaimId", IPage_PageTitle)
    Set lblHistoryCallerName = BindControl(pg, "lblHistoryCallerName", IPage_PageTitle)
    Set lblHistoryCallDate = BindControl(pg, "lblHistoryCallDate", IPage_PageTitle)
    Set lblHistoryCallComment = BindControl(pg, "lblHistoryCallComment", IPage_PageTitle)
    Set lblHistoryCallStatus = BindControl(pg, "lblHistoryCallStatus", IPage_PageTitle)
    Set multiPageDetails = BindControl(pg, "MultiPageDetailsViewRecords", IPage_PageTitle)
        
    cboStatus.List = Array("All", STATUS_PENDING, STATUS_CLOSED)
    cboStatus.value = "All"
    
    cboDateField.List = Array("Creation Date", "Insertion Date", "Last Updated", "Closed Date")
    cboDateField.value = "Creation Date"

    ConfigureClaimList lstClaims
    ConfigureHistoryList lstHistory

    ClearDetail
    clearDetailHistory
    m_loading = False
    
End Sub

Private Sub ConfigureHistoryList(lst As MSForms.ListBox)

    With lst
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "90;120;60;10"
    End With
    
End Sub


Private Sub ConfigureClaimList(lst As MSForms.ListBox)
      
    ' 3. Define your headers in an Array
    With lst
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "70;80;140;60;50;80"
        .MultiSelect = fmMultiSelectSingle
    End With
    
End Sub


'' Reloads only if another tab invalidated the cache - switching tabs
'' back and forth doesn't re-hit the shared file.

'Private Sub IPage_OnActivate()
'    ApplyFilters
'
'
'End Sub

Private Sub IPage_OnActivate()
    
    ' Arranged once, on first visit. Doing it here rather than in
    ' InitPage means tabs the user never opens are never laid out - a
    ' small saving, and it keeps startup snappy.
    arrangeUI
    If Not m_arranged Then
         arrangeUI
         m_arranged = True
     End If
     
     ApplyFilters
     
     m_ctx.SetStatus "View Records"
    
'    If Not lblInfo Is Nothing Then
'        lblInfo.caption = "Status starts as Pending and Attempt at 0. " & _
'                          "Insertion date and your username are recorded automatically."
'    End If
    
End Sub


Private Sub ApplyFilters()
    
    Dim dFrom As Date, dTo As Date
    Dim listData As Variant
    Dim allClaims As Collection
    Dim shownCount As Long

    If m_loading Then Exit Sub

    Set allClaims = m_ctx.claims()     ' cached; loads only if stale

    dFrom = ParseDateOrZero(txtDateFrom.value)
    dTo = ParseDateOrZero(txtDateTo.value)

    If Trim$(txtDateFrom.value) <> "" And dFrom = 0 Then
        m_ctx.SetStatus "'From' date not recognised - ignoring it."
    ElseIf Trim$(txtDateTo.value) <> "" And dTo = 0 Then
        m_ctx.SetStatus "'To' date not recognised - ignoring it."
    End If

    If dFrom > 0 And dTo > 0 And dFrom > dTo Then
        MsgBox "The 'From' date is after the 'To' date.", vbExclamation, "Check Dates"
        Exit Sub
    End If

    Set m_shown = FilterClaims(allClaims, CStr(cboStatus.value), SelectedDateField(), _
                                dFrom, dTo, CStr(txtSearch.value))

    ClearDetail
    clearDetailHistory
    
    lstClaims.Clear
    lstHistory.Clear
    
    shownCount = m_shown.Count

    If shownCount = 0 Then
        lblCount.caption = "0 of " & allClaims.Count
        m_ctx.SetStatus "No claims match the current filter."
        Exit Sub
    End If

    If shownCount > MAX_LIST_ROWS Then
        m_ctx.SetStatus shownCount & " matches - showing first " & MAX_LIST_ROWS & _
                        ". Narrow the filter to see the rest."
        Set m_shown = TakeFirst(m_shown, MAX_LIST_ROWS)
    Else
        m_ctx.SetStatus "Ready."
    End If

    listData = ClaimsToListArray(m_shown)
    If Not IsEmpty(listData) Then lstClaims.List = listData
    lblCount.caption = m_shown.Count & " of " & allClaims.Count
    
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

Private Sub lstHistory_click()

    Dim idx As Long
    idx = lstHistory.ListIndex
    If idx < 0 Then
        clearDetailHistory
        Exit Sub
    End If
    
    showDetailHistory m_shown_history(idx + 1)      ' ListIndex 0-based, Collection 1-based
    
End Sub

Private Sub lstClaims_Click()
    
    Dim idx As Long
    idx = lstClaims.ListIndex
    If idx < 0 Then
        ClearDetail
        Exit Sub
    End If
    ShowDetail m_shown(idx + 1)     ' ListIndex 0-based, Collection 1-based
    
End Sub

Private Sub showDetailHistory(c As clsHistoryEntry)
       
    lblHistoryId.caption = c.historyId
    lblHistoryClaimId.caption = c.claimID
    lblHistoryCallStatus.caption = c.CallerStatus
    lblHistoryCallerName.caption = c.callerName
    lblHistoryCallDate.caption = c.CallDateTime
    lblHistoryCallComment.caption = c.CallerComment
    
End Sub

Private Sub ShowDetail(c As clsClaim)
    
    Dim histData As Variant

    lblClaimID.caption = c.claimID
    lblClaimSite.caption = c.claimSite
    lblProviderName.caption = c.ClaimProviderName
    lblQuery.caption = c.claimQuery
    lblStatus.caption = c.ClaimStatus
    lblAttempt.caption = CStr(c.attempt)
    lblCreated.caption = Format$(c.ClaimCreationDate, "dd-mmm-yyyy")
    lblUpdatedClaimSite.caption = c.ClaimUpdatedSite
    lblInsertedOn.caption = Format$(c.ClaimInsertionDate, "dd-mmm-yyyy hh:nn")
    lblInsertedBy.caption = c.ClaimInsertedBy
    lblLastUpdated.caption = FormatDateOrBlank(c.LastUpdatedDate)
    lblLastUpdatedBy.caption = c.LastUpdatedBy
    lblLastComment.caption = c.LastComment
    lblClosedOn.caption = FormatDateOrBlank(c.ClaimClosedDate)
    lblDaysOpen.caption = CStr(c.DaysOpen) & IIf(c.IsClosed, " (to close)", " (open)")

    ' History is fetched per selection, not preloaded - preloading every
    ' claim's calls would mean pulling the whole History sheet to show one.
    lstHistory.Clear
    Set m_shown_history = m_ctx.repo.GetHistory(c.claimID)
    histData = HistoryToListArray(m_shown_history)
    If Not IsEmpty(histData) Then lstHistory.List = histData
    
End Sub

Private Sub clearDetailHistory()

    lblHistoryId.caption = ""
    lblHistoryClaimId.caption = ""
    lblHistoryCallStatus.caption = ""
    lblHistoryCallerName.caption = ""
    lblHistoryCallDate.caption = ""
    lblHistoryCallComment.caption = ""

End Sub


Private Sub ClearDetail()

    lblClaimID.caption = ""
    lblClaimSite.caption = ""
    lblProviderName.caption = ""
    lblQuery.caption = ""
    lblStatus.caption = ""
    lblAttempt.caption = ""
    lblCreated.caption = ""
    lblUpdatedClaimSite.caption = ""
    lblInsertedOn.caption = ""
    lblInsertedBy.caption = ""
    lblLastUpdated.caption = ""
    lblLastUpdatedBy.caption = ""
    lblLastComment.caption = ""
    lblClosedOn.caption = ""
    lblDaysOpen.caption = ""
    lstHistory.Clear
    
End Sub

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

Private Sub cmdRefresh_Click()
    m_ctx.InvalidateClaims
    ApplyFilters
End Sub

Private Sub cboStatus_Change()
    ApplyFilters
End Sub

Private Sub cboDateField_Change()
    ApplyFilters
End Sub

Private Sub txtSearch_Change()
    ApplyFilters
End Sub


Private Sub arrangeUI()
    
    'Dim m_pg As UserFormMain
    Dim leftMargin As Double
    
    leftMargin = 25
    
    With m_pg.framePageViewRecords
        
        .width = m_pg.Parent.width + 1
        .height = m_pg.Parent.height + 1
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = -1
        .Top = -1
                
        'Styling frame details
        .frameDetailsViewRecords.Top = -1
        .frameDetailsViewRecords.Left = .width - .frameDetailsViewRecords.width
        .frameDetailsViewRecords.height = .height
                        
                        
        .lblHeadingViewRecords.Top = 4.5
        .lblHeadingViewRecords.Left = 9
        
        .lblHeadingULineViewRecords.Top = .lblHeadingViewRecords.Top + .lblHeadingViewRecords.height + 2
        .lblHeadingULineViewRecords.Left = 9
        .lblHeadingULineViewRecords.width = .frameDetailsViewRecords.Left - 2 * .lblHeadingViewRecords.Left
        .lblHeadingULineViewRecords.height = 2
                                        
        multiPageDetails.Left = 0
        multiPageDetails.Top = 0
        multiPageDetails.height = .frameDetailsViewRecords.height
        multiPageDetails.width = .frameDetailsViewRecords.width
                                
                                
        .frameFilterViewClaims.Left = 0
        .frameFilterViewClaims.height = multiPageDetails.height
        .frameFilterViewClaims.Top = -2
        .frameFilterViewClaims.BorderStyle = fmBorderStyleNone
        .frameFilterViewClaims.ZOrder 0
        
        .frameViewRecordDetails.width = multiPageDetails.width
        .frameViewRecordDetails.Left = 0
        .frameViewRecordDetails.height = multiPageDetails.height
        .frameViewRecordDetails.Top = -2
        .frameViewRecordDetails.BorderStyle = fmBorderStyleNone
        .frameViewRecordDetails.ZOrder 1
        
        .frameViewHistoryDetail.width = multiPageDetails.width
        .frameViewHistoryDetail.Left = 0
        .frameViewHistoryDetail.height = multiPageDetails.height
        .frameViewHistoryDetail.Top = -2
        .frameViewHistoryDetail.BorderStyle = fmBorderStyleNone
        .frameViewHistoryDetail.ZOrder 1
                        
'        .frameViewRecords.Left = 10
'        .frameViewRecords.width = .width - multiPageDetails.width - 2 * .frameViewRecords.Left
'        .frameViewRecords.Top = 50
'        .frameViewRecords.height = .height / 2
'        .frameViewRecords.ZOrder 1
        
        
        .lstViewClaims.Left = leftMargin + 5
        .lstViewClaims.Top = .lblHeadingULineViewRecords.Top + 30
        .lstViewClaims.height = ((.height - .lstViewClaims.Top) / 2) - 30
        .lstViewClaims.width = .frameDetailsViewRecords.Left - 2 * .lstViewClaims.Left + 10
        .lstViewClaims.BorderStyle = fmBorderStyleSingle
        .lstViewClaims.Font.Name = "Segoe UI"
        .lstViewClaims.Font.Size = 9
        
                        
'        .frameViewHistory.Left = 10
'        .frameViewHistory.width = .width - .frameFilterViewClaims.width - 2 * .frameViewHistory.Left
'        .frameViewHistory.Top = .frameViewRecords.height + .frameViewRecords.Top + 10
'        .frameViewHistory.height = .height / 3
'        .frameViewHistory.ZOrder 1
        
        .lstViewHistory.Left = leftMargin + 5
        .lstViewHistory.Top = .lstViewClaims.Top + .lstViewClaims.height + 30
        .lstViewHistory.height = .height - .lstViewHistory.Top - 30
        .lstViewHistory.width = .frameDetailsViewRecords.Left - 2 * .lstViewHistory.Left + 10
        .lstViewHistory.BorderStyle = fmBorderStyleSingle
        .lstViewHistory.Font.Name = "Segoe UI"
        .lstViewHistory.Font.Size = 10
        
        
        
'        .lstViewClaims.Left = 10
'        .lstViewClaims.width = .width - 2 * m_pg.lstViewClaims.Left
'        '.lstViewClaims.ZOrder = 1
'
'        .lstViewHistory.Left = 10
'        .lstViewHistory.width = .width - 2 * .lstViewHistory.Left
'        '.ZOrder = 1
'
        
    End With
 
     'Style Page Add Claim
     
     
    collectionStyledControls.Add Styler.getStyledBoxULine(txtSearch, "Claim Id*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtDateFrom, "Date From")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtDateTo, "Date To")
    collectionStyledControls.Add Styler.getStyleButton(cmdApply)
    collectionStyledControls.Add Styler.getStyleButton(cmdClear)
    collectionStyledControls.Add Styler.getStyleButton(cmdRefresh)
    'collectionStyledControls.Add Styler.getStyledBoxULine(cboDateField, "Creation Date*")

    'Adding headers to listboxClaims
    Dim headersClaimList As Variant
    Dim headersHistoryList As Variant
    headersClaimList = Array("Claim ID", "Site", "Provider Name", "Status", "Attempt", "Creation Date")
    headersHistoryList = Array("Caller Name", "Call Date", "Status", "Comment")
    
    ' 4. Call the helper function
    ListBoxUtil.generateListBoxHeaders lstClaims, headersClaimList
    DoEvents
    ListBoxUtil.generateListBoxHeaders lstHistory, headersHistoryList
    

 End Sub


