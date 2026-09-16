VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPageSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageSearch - controller for the "Search & Edit" tab.
'
' Look up one claim, see everything including its call history, and
' correct the detail fields that were keyed in wrong.
'
' Editable:  ClaimSite, ClaimProviderName, ClaimQuery, ClaimCreationDate
'            ClaimUpdatedSite (Admin only)
' Read only: ClaimStatus, Attempt, ClaimClosedDate - maintained by
'            LogCallAndUpdateStatus so status, attempt count and the
'            call trail stay in step. Editing status here would let
'            someone close a claim with no call behind it.
' Read only: ClaimID and all audit fields. The call history is
'            append-only and never editable.
' =====================================================================

Implements IPage

Private m_ctx As clsAppContext
Private m_pg As MSForms.Page
Private m_claim As clsClaim
Private m_dirty As Boolean
Private m_loading As Boolean
Private m_arranged As Boolean

'Private WithEvents cmdSearch As MSForms.CommandButton
Private WithEvents cmdSearch As MSForms.Label
Attribute cmdSearch.VB_VarHelpID = -1
'Attribute cmdSearch.VB_VarHelpID = -1
'Private WithEvents cmdSave As MSForms.CommandButton
Private WithEvents cmdSave As MSForms.Label
Attribute cmdSave.VB_VarHelpID = -1
'Attribute cmdSave.VB_VarHelpID = -1
'Private WithEvents cmdRevert As MSForms.CommandButton
Private WithEvents cmdRevert As MSForms.Label
Attribute cmdRevert.VB_VarHelpID = -1
'Attribute cmdRevert.VB_VarHelpID = -1
'Private WithEvents cmdSrcReset As MSForms.CommandButton
Private WithEvents cmdSrcReset As MSForms.Label
Attribute cmdSrcReset.VB_VarHelpID = -1
'Attribute cmdRevert.VB_VarHelpID = -1

Private WithEvents txtSearchID As MSForms.TextBox
Attribute txtSearchID.VB_VarHelpID = -1
'Attribute txtSearchID.VB_VarHelpID = -1
Private WithEvents txtSite As MSForms.TextBox
Attribute txtSite.VB_VarHelpID = -1
'Attribute txtSite.VB_VarHelpID = -1
Private WithEvents txtProvider As MSForms.TextBox
Attribute txtProvider.VB_VarHelpID = -1
'Attribute txtProvider.VB_VarHelpID = -1
Private WithEvents datePickerCreatedOn As MSForms.Label
Attribute datePickerCreatedOn.VB_VarHelpID = -1

Private WithEvents txtQuery As MSForms.TextBox
Attribute txtQuery.VB_VarHelpID = -1
'Attribute txtQuery.VB_VarHelpID = -1
Private WithEvents txtCreationDate As MSForms.TextBox
Attribute txtCreationDate.VB_VarHelpID = -1
'Attribute txtCreationDate.VB_VarHelpID = -1
Private WithEvents txtUpdatedSite As MSForms.TextBox
Attribute txtUpdatedSite.VB_VarHelpID = -1
'Attribute txtUpdatedSite.VB_VarHelpID = -1


Private lstHistory As MSForms.ListBox
Private lblClaimID As MSForms.Label
Private lblSiteInitial As MSForms.Label
Private lblSiteUpdated As MSForms.Label
Private lblStatus As MSForms.Label
Private lblCreatedOn As MSForms.Label


Private lblAttempt As MSForms.Label
Private lblInsertedOn As MSForms.Label
Private lblInsertedBy As MSForms.Label
Private lblLastUpdated As MSForms.Label
Private lblLastUpdatedBy As MSForms.Label
Private lblClosedOn As MSForms.Label
Private lblDaysOpen As MSForms.Label
Private lblHistoryCount As MSForms.Label
Private lblAdminNote As MSForms.Label

Private lblProviderName As MSForms.Label
Private lblClaimQuery As MSForms.Label

Private collectionStyledControls As Collection

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Search & Edit"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As MSForms.Page)
    
    Set m_ctx = ctx
    Set m_pg = pg
    m_loading = True

    Set txtSearchID = BindControl(pg, "txtSrchID", IPage_PageTitle)
    Set cmdSearch = BindControl(pg, "cmdSrchFind", IPage_PageTitle)
    Set txtSite = BindControl(pg, "txtSrchSite", IPage_PageTitle)
    Set txtProvider = BindControl(pg, "txtSrchProvider", IPage_PageTitle)
    Set txtQuery = BindControl(pg, "txtSrchQuery", IPage_PageTitle)
    Set txtCreationDate = BindControl(pg, "txtSrchCreationDate", IPage_PageTitle)
    Set txtUpdatedSite = BindControl(pg, "txtSrchUpdatedSite", IPage_PageTitle)
    Set cmdSave = BindControl(pg, "cmdSrchSave", IPage_PageTitle)
    Set cmdRevert = BindControl(pg, "cmdSrchRevert", IPage_PageTitle)
    Set cmdSrcReset = BindControl(pg, "cmdSrcReset", IPage_PageTitle)
    Set lstHistory = BindControl(pg, "lstSrchHistory", IPage_PageTitle)

    Set lblClaimID = BindControl(pg, "lblSrchClaimID", IPage_PageTitle)
    Set lblSiteInitial = BindControl(pg, "lblSearchSiteInitial", IPage_PageTitle)
    Set lblSiteUpdated = BindControl(pg, "lblSearchSiteUpdated", IPage_PageTitle)
    Set lblStatus = BindControl(pg, "lblSrchStatus", IPage_PageTitle)
    Set lblCreatedOn = BindControl(pg, "lblSearchCreatedOn", IPage_PageTitle)
    
    Set datePickerCreatedOn = BindControl(pg, "lblSearchDatePickerCreatedOn", IPage_PageTitle)
    
    Set lblClosedOn = BindControl(pg, "lblSrchClosedOn", IPage_PageTitle)
    Set lblInsertedOn = BindControl(pg, "lblSrchInsertedOn", IPage_PageTitle)
    Set lblInsertedBy = BindControl(pg, "lblSrchInsertedBy", IPage_PageTitle)
    Set lblLastUpdated = BindControl(pg, "lblSrchLastUpdated", IPage_PageTitle)
    Set lblLastUpdatedBy = BindControl(pg, "lblSrchLastUpdatedBy", IPage_PageTitle)
    Set lblAttempt = BindControl(pg, "lblSrchAttempt", IPage_PageTitle)
    
    Set lblDaysOpen = BindControl(pg, "lblSrchDaysOpen", IPage_PageTitle)
    
    Set lblProviderName = BindControl(pg, "lblSearchProviderName", IPage_PageTitle)
    Set lblClaimQuery = BindControl(pg, "lblSearchClaimQuery", IPage_PageTitle)
    
    Set lblHistoryCount = BindControl(pg, "lblSrchHistoryCount", IPage_PageTitle)
    Set lblAdminNote = BindOptional(pg, "lblSrchAdminNote")
    
    Set collectionStyledControls = New Collection
    txtQuery.Multiline = True
    ConfigureHistoryList lstHistory

    ClearAll
    m_loading = False
    
End Sub


Public Sub ConfigureHistoryList(lst As MSForms.ListBox)

    With lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "90;120;60;10"
    End With
    
End Sub




Private Sub IPage_OnActivate()

    m_ctx.SetStatus "Search a Claim ID to view and correct its details."
    
    arrangeUI
    
    If Not m_arranged Then
         arrangeUI
         m_arranged = True
     End If
    
End Sub

' ---------------------------------------------------------------------
' Search
' ---------------------------------------------------------------------
Private Sub txtSearchID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then DoSearch
End Sub

Private Sub cmdSearch_Click()

    DoSearch
    
End Sub

Private Sub cmdSrcReset_Click()

    ClearAll
    
End Sub

Private Sub DoSearch()

    Dim idToFind As String

    idToFind = Trim$(txtSearchID.value)
    If idToFind = "" Then
        MsgBox "Enter a Claim ID to search for.", vbExclamation, "Nothing to Search"
        Exit Sub
    End If

    If Not ConfirmDiscardIfDirty() Then Exit Sub

    m_ctx.SetStatus "Searching..."
    Set m_claim = m_ctx.repo.FindClaim(idToFind)

    If m_claim Is Nothing Then
        ClearAll
        m_ctx.SetStatus "No claim found with ID '" & idToFind & "'."
        MsgBox "No claim found with ID '" & idToFind & "'.", vbExclamation, "Search Found"
        Exit Sub
    End If

    Populate
    m_ctx.SetStatus "Claim loaded."
    
End Sub

Private Sub Populate()
    
    Dim histData As Variant
    Dim hist As Collection

    m_loading = True
    lblClaimID.caption = m_claim.claimID
    lblSiteInitial.caption = m_claim.claimSite
    lblSiteUpdated.caption = m_claim.ClaimUpdatedSite
    lblStatus.caption = m_claim.ClaimStatus
    lblCreatedOn.caption = Format$(m_claim.ClaimCreationDate, "dd-mmm-yyyy hh:nn")
    lblClosedOn.caption = FormatDateOrBlank(m_claim.ClaimClosedDate)
    lblInsertedOn.caption = Format$(m_claim.ClaimInsertionDate, "dd-mmm-yyyy hh:nn")
    lblInsertedBy.caption = m_claim.ClaimInsertedBy
    lblLastUpdated.caption = FormatDateOrBlank(m_claim.LastUpdatedDate)
    lblLastUpdatedBy.caption = m_claim.LastUpdatedBy
    lblAttempt.caption = CStr(m_claim.attempt)
    'lblDaysOpen.caption = CStr(m_claim.DaysOpen) & IIf(m_claim.IsClosed, " (to close)", " (open)")
    lblProviderName.caption = m_claim.ClaimProviderName
    lblClaimQuery.caption = m_claim.claimQuery
    
    txtSite.value = m_claim.claimSite
    txtProvider.value = m_claim.ClaimProviderName
    txtQuery.value = m_claim.claimQuery
    txtCreationDate.value = Format$(m_claim.ClaimCreationDate, "dd-mmm-yyyy")
    txtUpdatedSite.value = m_claim.ClaimUpdatedSite

    Set hist = m_ctx.repo.GetHistory(m_claim.claimID)
    lstHistory.Clear
    histData = HistoryToListArray(hist)
    If Not IsEmpty(histData) Then lstHistory.List = histData
    lblHistoryCount.caption = hist.Count & " call(s) logged"

    SetEditingEnabled True
    ApplyPermissions

    m_dirty = False
    cmdSave.enabled = False
    m_loading = False
    
End Sub

' Non-admins see the Updated Site value but can't type in it - clearer
' than hiding the field and leaving them wondering where it went.
Private Sub ApplyPermissions()
    
    Dim isAdmin As Boolean
    isAdmin = m_ctx.isAdmin          ' cached in the context, not re-checked
    txtUpdatedSite.enabled = isAdmin
    txtUpdatedSite.BackColor = IIf(isAdmin, &H80000005, &H8000000F)
    If Not lblAdminNote Is Nothing Then lblAdminNote.Visible = Not isAdmin

End Sub

' ---------------------------------------------------------------------
' Editing
' ---------------------------------------------------------------------
Private Sub txtSite_Change()
    MarkDirty
End Sub
Private Sub txtProvider_Change()
    MarkDirty
End Sub
Private Sub txtQuery_Change()
    MarkDirty
End Sub
Private Sub txtCreationDate_Change()
    MarkDirty
End Sub
Private Sub txtUpdatedSite_Change()
    MarkDirty
End Sub

Private Sub datePickerCreatedOn_click()
    
    Dim dtPick As Date
    Dim dtPickStr As String
    
    dtPick = CalendarForm.GetDate()
    dtPickStr = Format(dtPick, "MMM-dd-yyyy")
    txtCreationDate.value = dtPickStr
    
End Sub


Private Sub MarkDirty()
    
    If m_loading Then Exit Sub
    If m_claim Is Nothing Then Exit Sub
    m_dirty = True
    cmdSave.enabled = True
    m_ctx.SetStatus "Unsaved changes."

End Sub

Private Sub cmdSave_Click()
    Dim siteDenied As Boolean
    Dim creationDate As Date

    If m_claim Is Nothing Then Exit Sub

    If Trim$(txtSite.value) = "" Then
        FailField "Claim Site cannot be blank.", txtSite
        Exit Sub
    End If
    If Trim$(txtProvider.value) = "" Then
        FailField "Provider Name cannot be blank.", txtProvider
        Exit Sub
    End If
    If Trim$(txtQuery.value) = "" Then
        FailField "Claim Query cannot be blank.", txtQuery
        Exit Sub
    End If
    If Not IsDate(txtCreationDate.value) Then
        FailField "Creation Date is not a valid date.", txtCreationDate
        Exit Sub
    End If
    creationDate = CDate(txtCreationDate.value)
    If creationDate > Date Then
        FailField "Creation Date cannot be in the future.", txtCreationDate
        Exit Sub
    End If

    If m_ctx.repo.UpdateClaimDetails(m_claim.claimID, Trim$(txtSite.value), _
                                      Trim$(txtProvider.value), Trim$(txtQuery.value), _
                                      creationDate, Trim$(txtUpdatedSite.value), _
                                      siteDenied) Then
        m_ctx.InvalidateClaims

        ' Re-read rather than trusting our copy, so the audit fields on
        ' screen show what was actually written.
        Set m_claim = m_ctx.repo.FindClaim(m_claim.claimID)
        If Not m_claim Is Nothing Then Populate

        m_ctx.SetStatus "Changes saved." & _
            IIf(siteDenied, " (Updated Site not changed - Admin only.)", "")
    Else
        m_ctx.SetStatus "Save failed."
    End If
End Sub

Private Sub cmdRevert_Click()
    If m_claim Is Nothing Then Exit Sub
    If Not m_dirty Then Exit Sub

    If MsgBox("Discard your unsaved changes to this claim?", _
              vbYesNo + vbQuestion, "Discard Changes") = vbNo Then Exit Sub

    Set m_claim = m_ctx.repo.FindClaim(m_claim.claimID)
    If m_claim Is Nothing Then
        ClearAll
        m_ctx.SetStatus "This claim no longer exists in the database."
        Exit Sub
    End If

    Populate
    m_ctx.SetStatus "Changes discarded."
End Sub

' Called by the shell before closing or switching away, so unsaved work
' isn't silently lost.
Public Function ConfirmDiscardIfDirty() As Boolean
    ConfirmDiscardIfDirty = True
    If Not m_dirty Then Exit Function
    If m_claim Is Nothing Then Exit Function

    If MsgBox("You have unsaved changes to claim " & m_claim.claimID & "." & vbCrLf & vbCrLf & _
              "Discard them?", vbYesNo + vbExclamation, "Unsaved Changes") = vbNo Then
        ConfirmDiscardIfDirty = False
    End If
End Function

' Lets the View tab hand a claim straight over to this tab.
Public Sub LoadClaimByID(ByVal claimID As String)
    txtSearchID.value = claimID
    DoSearch
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------
Private Sub ClearAll()
    
    m_loading = True
    Set m_claim = Nothing
    m_dirty = False

    lblClaimID.caption = ""
    lblSiteInitial.caption = ""
    lblSiteUpdated.caption = ""
    lblStatus.caption = ""
    lblCreatedOn.caption = ""
    lblClosedOn.caption = ""
    
    lblInsertedBy.caption = ""
    lblInsertedOn.caption = ""
    lblLastUpdated.caption = ""
    lblLastUpdatedBy.caption = ""
    
    lblAttempt.caption = ""
    lblProviderName.caption = ""
    lblClaimQuery.caption = ""
    
    lblDaysOpen.caption = ""
    lblHistoryCount.caption = ""
    lblAdminNote.caption = ""
    
    txtSite.value = ""
    txtProvider.value = ""
    txtQuery.value = ""
    txtCreationDate.value = ""
    txtUpdatedSite.value = ""
    lstHistory.Clear

    SetEditingEnabled False
    cmdSave.enabled = False
    m_loading = False
    
End Sub

' Editable boxes stay locked until a claim is loaded, so nobody types
' into a blank form and wonders why Save does nothing.
Private Sub SetEditingEnabled(ByVal enabled As Boolean)
    
    txtSite.enabled = enabled
    txtProvider.enabled = enabled
    txtQuery.enabled = enabled
    txtCreationDate.enabled = enabled
    cmdRevert.enabled = enabled
    If Not enabled Then txtUpdatedSite.enabled = False
    
End Sub

Private Sub FailField(ByVal msg As String, ctrl As MSForms.control)
    
    MsgBox msg, vbExclamation, "Check Details"
    On Error Resume Next
    ctrl.SetFocus
    On Error GoTo 0
    
End Sub


Private Sub arrangeUI()
    
    'Dim m_pg As UserFormMain
    Dim endingMark As Double
    Dim leftMargin As Double
    Dim btnControlLeftK As Double
    Dim editingWidthK As Double
    Dim rightMargin As Double
    Dim widgetSpace As Double
    
    
    leftMargin = 25
    widgetSpace = 20
    
    With m_pg.framePageSearchEdit
        
        .width = m_pg.Parent.width + 1
        .height = m_pg.Parent.height + 1
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = -1
        .Top = -1
                        
        'Styling frame src details
        .frameSrcDetails.Top = -1
        .frameSrcDetails.Left = .width - .frameSrcDetails.width
        .frameSrcDetails.height = .height
                                
        .lblHeadingSearchFind.Top = 4.5
        .lblHeadingSearchFind.Left = 9
        .lblHeadingULineSearchEdit.Top = .lblHeadingSearchFind.Top + .lblHeadingSearchFind.height + 2
        .lblHeadingULineSearchEdit.Left = 9
        .lblHeadingULineSearchEdit.width = .frameSrcDetails.Left - 2 * .lblHeadingSearchFind.Left
        .lblHeadingULineSearchEdit.height = 2
        
        btnControlLeftK = .lblHeadingULineSearchEdit.Left + .lblHeadingULineSearchEdit.width
        .cmdSrchFind.Left = btnControlLeftK - .cmdSrchFind.width - 10
        .cmdSrchRevert.Left = btnControlLeftK - .cmdSrchRevert.width - 10
        .cmdSrchSave.Left = btnControlLeftK - .cmdSrchSave.width - 10
        .cmdSrcReset.Left = btnControlLeftK - .cmdSrcReset.width - 10
                        
        endingMark = .cmdSrchFind.Left
        .txtSrchQuery.Left = leftMargin
        .txtSrchQuery.width = endingMark - 2 * .txtSrchQuery.Left
        
        .lstSrchHistory.Left = leftMargin + 5
        .lstSrchHistory.width = .frameSrcDetails.Left - 2 * .lstSrchHistory.Left + 10
        .lstSrchHistory.height = .height - .lstSrchHistory.Top - 30
        
        
        editingWidthK = .cmdSrchFind.Left - 2 * leftMargin - widgetSpace
        
        .txtSrchID.Left = leftMargin
        .txtSrchID.width = editingWidthK / 2
        .txtSrchProvider.Left = .txtSrchID.Left + .txtSrchID.width + widgetSpace
        .txtSrchProvider.width = editingWidthK / 2
                
        editingWidthK = .cmdSrchFind.Left - 2 * leftMargin - 2 * widgetSpace
        
        .txtSrchSite.Left = leftMargin
        .txtSrchSite.width = editingWidthK / 3
        .txtSrchUpdatedSite.Left = .txtSrchSite.Left + .txtSrchSite.width + widgetSpace
        .txtSrchUpdatedSite.width = editingWidthK / 3
        .txtSrchCreationDate.Left = .txtSrchUpdatedSite.Left + .txtSrchUpdatedSite.width + widgetSpace
        .txtSrchCreationDate.width = editingWidthK / 3
        .lblSearchDatePickerCreatedOn.Left = .txtSrchCreationDate.Left + .txtSrchCreationDate.width - .lblSearchDatePickerCreatedOn.width + 4.8
        
        
        '.BackColor = vbYellow
        
'        .MultiPageDetails.Left = .width - .MultiPageDetails.width
'        .MultiPageDetails.height = .height
'        .MultiPageDetails.Top = -2
'
'        .frameFilterViewClaims.Left = 0
'        .frameFilterViewClaims.height = .MultiPageDetails.height
'        .frameFilterViewClaims.Top = -2
'        .frameFilterViewClaims.BorderStyle = fmBorderStyleNone
'        .frameFilterViewClaims.ZOrder 0
'
'        .frameViewRecordDetails.width = .MultiPageDetails.width
'        .frameViewRecordDetails.Left = 0
'        .frameViewRecordDetails.height = .MultiPageDetails.height
'        .frameViewRecordDetails.Top = -2
'        .frameViewRecordDetails.BorderStyle = fmBorderStyleNone
'        .frameViewRecordDetails.ZOrder 1
'
'        .frameViewHistoryDetail.width = .MultiPageDetails.width
'        .frameViewHistoryDetail.Left = 0
'        .frameViewHistoryDetail.height = .MultiPageDetails.height
'        .frameViewHistoryDetail.Top = -2
'        .frameViewHistoryDetail.BorderStyle = fmBorderStyleNone
'        .frameViewHistoryDetail.ZOrder 1
'
'        .frameViewRecords.Left = 10
'        .frameViewRecords.width = .width - .MultiPageDetails.width - 2 * .frameViewRecords.Left
'        .frameViewRecords.Top = 50
'        .frameViewRecords.height = .height / 2
'        .frameViewRecords.ZOrder 1
'
'
'        .lstViewClaims.Left = 5
'        .lstViewClaims.Top = 30
'        .lstViewClaims.height = .frameViewRecords.height - .lstViewClaims.Top - 10
'        .lstViewClaims.width = .frameViewRecords.width - 2 * .lstViewClaims.Left
'        .lstViewClaims.BorderStyle = fmBorderStyleSingle
'        .lstViewClaims.Font.Name = "Segoe UI"
'        .lstViewClaims.Font.Size = 10
'
'
'        .frameViewHistory.Left = 10
'        .frameViewHistory.width = .width - .frameFilterViewClaims.width - 2 * .frameViewHistory.Left
'        .frameViewHistory.Top = .frameViewRecords.height + .frameViewRecords.Top + 10
'        .frameViewHistory.height = .height / 3
'        .frameViewHistory.ZOrder 1
'
'        .lstViewHistory.Left = 5
'        .lstViewHistory.Top = 30
'        .lstViewHistory.height = .frameViewHistory.height - .lstViewHistory.Top - 10
'        .lstViewHistory.width = .frameViewHistory.width - 2 * .lstViewHistory.Left
'        .lstViewHistory.BorderStyle = fmBorderStyleSingle
'        .lstViewHistory.Font.Name = "Segoe UI"
'        .lstViewHistory.Font.Size = 10
'
'
'
'        .lstViewClaims.Left = 10
'        .lstViewClaims.width = .width - 2 * m_pg.lstViewClaims.Left
'        '.lstViewClaims.ZOrder = 1
'
'        .lstViewHistory.Left = 10
'        .lstViewHistory.width = .width - 2 * .lstViewHistory.Left
'        '.ZOrder = 1
'
        
    End With
 
'     'Style Page Add Claim
'    Dim m_pg As UserFormMain

    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchID, "Search Claim Id")
    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchCreationDate, "Creation Date")
    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchProvider, "Provider Name")
    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchQuery, "Query")
    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchUpdatedSite, "Site")
    collectionStyledControls.Add Styler.getStyledBoxULine(m_pg.txtSrchSite, "Initial Site")
    collectionStyledControls.Add Styler.getStyleButton(cmdSearch)
    collectionStyledControls.Add Styler.getStyleButton(cmdRevert)
    collectionStyledControls.Add Styler.getStyleButton(cmdSave)
    collectionStyledControls.Add Styler.getStyleButton(cmdSrcReset)
    


    'Adding headers to listboxClaims
    Dim headersHistoryList As Variant
    headersHistoryList = Array("Caller Name", "Call Date", "Status", "Comment")

    'Call the helper function
    'ListBoxUtil.CreateListBoxHeaders m_pg.framePageViewRecords.lstViewClaims, headersClaimList
    'ListBoxUtil.CreateListBoxHeaders lstHistory, headersHistoryList
    ListBoxUtil.generateListBoxHeaders lstHistory, headersHistoryList, RGB(17, 49, 97), RGB(17, 69, 143)
    
    

 End Sub


