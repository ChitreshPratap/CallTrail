VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPageLogCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageLogCall - controller for the "Log Call" tab.
'
' This is the screen callers use all day: look up a claim, see where it
' stands and what previous callers said, log the outcome of this call.
'
' Status changes flow through LogCallAndUpdateStatus ONLY - that is what
' keeps ClaimStatus, Attempt and the call history in step with each
' other. Nothing on this tab writes status directly.
' =====================================================================

Implements IPage

Private m_ctx As clsAppContext
Private m_pg As Msforms.Page
Private m_claim As clsClaim
Private m_arranged As Boolean

'Private WithEvents cmdFind As MSForms.CommandButton
Private WithEvents cmdFind As Msforms.Label
Attribute cmdFind.VB_VarHelpID = -1
'Attribute cmdFind.VB_VarHelpID = -1
'Private WithEvents cmdLog As MSForms.CommandButton
Private WithEvents cmdLog As Msforms.Label
Attribute cmdLog.VB_VarHelpID = -1
'Attribute cmdLog.VB_VarHelpID = -1
Private WithEvents txtClaimID As Msforms.TextBox
Attribute txtClaimID.VB_VarHelpID = -1
'Attribute txtClaimID.VB_VarHelpID = -1

Private cboStatus As Msforms.ComboBox
Private txtComment As Msforms.TextBox
Private lstHistory As Msforms.ListBox
Private lblClaimInfo As Msforms.Label
Private lblCurrentStatus As Msforms.Label
Private lblAttempt As Msforms.Label
Private lblLastComment As Msforms.Label

Private collectionStyledControls As Collection

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Log Call"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As Msforms.Page)
    
    Set m_ctx = ctx
    Set m_pg = pg

    Set txtClaimID = BindControl(pg, "txtLogClaimID", IPage_PageTitle)
    Set cmdFind = BindControl(pg, "cmdLogFind", IPage_PageTitle)
    Set cboStatus = BindControl(pg, "cboLogStatus", IPage_PageTitle)
    Set txtComment = BindControl(pg, "txtLogComment", IPage_PageTitle)
    Set cmdLog = BindControl(pg, "cmdLogSubmit", IPage_PageTitle)
    Set lstHistory = BindControl(pg, "lstLogHistory", IPage_PageTitle)
    Set lblClaimInfo = BindControl(pg, "lblLogClaimInfo", IPage_PageTitle)
    Set lblCurrentStatus = BindControl(pg, "lblLogCurrentStatus", IPage_PageTitle)
    Set lblAttempt = BindControl(pg, "lblLogAttempt", IPage_PageTitle)
    Set lblLastComment = BindOptional(pg, "lblLogLastComment")

    cboStatus.List = Array(STATUS_PENDING, STATUS_CLOSED)
    txtComment.Multiline = True
    ConfigureHistoryList lstHistory

    Set collectionStyledControls = New Collection
    ConfigureHistoryList lstHistory
    
    ClearClaim
End Sub

Private Sub IPage_OnActivate()
    
    m_ctx.SetStatus "Find a claim, then log the outcome of your call."
'    ArrangeUI
    If Not m_arranged Then
         ArrangeUI
         m_arranged = True
     End If
    
End Sub

Public Sub ConfigureHistoryList(lst As Msforms.ListBox)

    With lst
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "90;120;60;10"
    End With
    
End Sub


Private Sub txtClaimID_KeyDown(ByVal KeyCode As Msforms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then FindClaim
End Sub

Private Sub cmdFind_Click()
    FindClaim
End Sub

Private Sub FindClaim()
    Dim idToFind As String

    idToFind = Trim$(txtClaimID.value)
    If idToFind = "" Then
        MsgBox "Enter a Claim ID.", vbExclamation, "Nothing to Find"
        Exit Sub
    End If

    m_ctx.SetStatus "Looking up " & idToFind & "..."
    Set m_claim = m_ctx.repo.FindClaim(idToFind)

    If m_claim Is Nothing Then
        ClearClaim
        m_ctx.SetStatus "No claim found with ID '" & idToFind & "'."
        Exit Sub
    End If

    ShowClaim
End Sub

Private Sub ShowClaim()
    lblClaimInfo.caption = m_claim.ClaimProviderName & "  |  " & m_claim.claimSite & _
                           "  |  raised " & Format$(m_claim.ClaimCreationDate, "dd-mmm-yyyy")
    lblCurrentStatus.caption = m_claim.ClaimStatus
    lblAttempt.caption = "Calls so far: " & m_claim.attempt
    If Not lblLastComment Is Nothing Then
        lblLastComment.caption = m_claim.LastComment
    End If

    LoadHistory m_claim.claimID

    ' A closed claim takes no further calls. Disabling the controls says
    ' so up front instead of letting someone type a comment and only then
    ' discover it's rejected.
    If m_claim.IsClosed Then
        cboStatus.enabled = False
        txtComment.enabled = False
        cmdLog.enabled = False
        m_ctx.SetStatus "This claim is already Closed - no further calls can be logged."
    Else
        cboStatus.enabled = True
        txtComment.enabled = True
        cmdLog.enabled = True
        cboStatus.value = STATUS_PENDING
        txtComment.value = ""
        m_ctx.SetStatus "Claim loaded. Enter your comment and status."
        txtComment.SetFocus
    End If
End Sub

Private Sub LoadHistory(ByVal claimID As String)
    Dim histData As Variant
    lstHistory.Clear
    histData = HistoryToListArray(m_ctx.repo.GetHistory(claimID))
    If Not IsEmpty(histData) Then lstHistory.List = histData
End Sub

Private Sub cmdLog_Click()
    If m_claim Is Nothing Then
        MsgBox "Find a claim first.", vbExclamation, "No Claim Loaded"
        Exit Sub
    End If
    If Trim$(txtComment.value) = "" Then
        MsgBox "A comment is required for every logged call.", vbExclamation, "Comment Needed"
        txtComment.SetFocus
        Exit Sub
    End If
    If Trim$(CStr(cboStatus.value)) = "" Then
        MsgBox "Select a status.", vbExclamation, "Status Needed"
        cboStatus.SetFocus
        Exit Sub
    End If

    ' Closing is not reversible through the app, so confirm it explicitly.
    If LCase$(CStr(cboStatus.value)) = LCase$(STATUS_CLOSED) Then
        If MsgBox("Close claim " & m_claim.claimID & "?" & vbCrLf & vbCrLf & _
                  "No further calls can be logged against it.", _
                  vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
    End If

    If m_ctx.repo.LogCallAndUpdateStatus(m_claim.claimID, Trim$(txtComment.value), _
                                          CStr(cboStatus.value)) Then
        m_ctx.InvalidateClaims

        ' Re-read so the attempt count and status shown are what was
        ' actually written, not what we assume was written.
        Set m_claim = m_ctx.repo.FindClaim(m_claim.claimID)
        If Not m_claim Is Nothing Then ShowClaim

        m_ctx.SetStatus "Call logged."
    End If
End Sub

Private Sub ClearClaim()
    Set m_claim = Nothing
    lblClaimInfo.caption = ""
    lblCurrentStatus.caption = ""
    lblAttempt.caption = ""
    If Not lblLastComment Is Nothing Then lblLastComment.caption = ""
    lstHistory.Clear
    txtComment.value = ""
    cboStatus.enabled = False
    txtComment.enabled = False
    cmdLog.enabled = False
End Sub


Private Sub ArrangeUI()
    
    Dim btnControlLeftK As Double
    Dim endingMark As Double
    Dim leftMargin As Double
    Dim rightMargin As Double
    Dim editingWidthK As Double
    Dim widgetSpace As Double
    
    'Dim m_pg As UserFormMain
    
    leftMargin = 25
    widgetSpace = 20
        
    With m_pg.framePageLogCall
        
        .width = m_pg.Parent.width + 1
        .height = m_pg.Parent.height + 1
        .Left = -1
        .Top = -1
        '.BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        Styler.stylePageFrame m_pg.framePageLogCall
        
        'Style Frame
            
        .frameLogCallHistoryDetail.Top = -1
        .frameLogCallHistoryDetail.Left = .width - .frameLogCallHistoryDetail.width
        .frameLogCallHistoryDetail.height = .height
        .frameLogCallHistoryDetail.BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.96)
        
        .lblHeadingLogCall.Top = 4.5
        .lblHeadingLogCall.Left = 9
        Styler.getStyledHeading .lblHeadingLogCall
        
        .lblHeadingULineLogCall.Top = .lblHeadingLogCall.Top + .lblHeadingLogCall.height + 2
        .lblHeadingULineLogCall.Left = 9
        .lblHeadingULineLogCall.width = .frameLogCallHistoryDetail.Left - 2 * .lblHeadingULineLogCall.Left
        .lblHeadingULineLogCall.height = 2
        Styler.getStyledHeadingULine .lblHeadingULineLogCall
        
        
        btnControlLeftK = .lblHeadingULineLogCall.Left + .lblHeadingULineLogCall.width
        .cmdLogFind.Left = btnControlLeftK - .cmdLogFind.width - 10
        .cmdLogSubmit.Left = btnControlLeftK - .cmdLogSubmit.width - 10
        
        endingMark = .cmdLogFind.Left
        .txtLogComment.Left = leftMargin
        .txtLogComment.width = endingMark - 2 * .txtLogComment.Left
        
        .lstLogHistory.Left = leftMargin
        .lstLogHistory.width = .frameLogCallHistoryDetail.Left - 2 * .lstLogHistory.Left + 10
        .lstLogHistory.height = .height - .lstLogHistory.Top - 30
        .lstLogHistory.BorderStyle = fmBorderStyleSingle
        .lstLogHistory.Font.Name = "Segoe UI"
        .lstLogHistory.Font.Size = 10
        
        editingWidthK = .cmdLogFind.Left - 2 * leftMargin - widgetSpace
        .txtLogClaimID.Left = leftMargin
        .txtLogClaimID.width = editingWidthK / 2
        
        .cboLogStatus.Left = .txtLogClaimID.Left + .txtLogClaimID.width + widgetSpace
        .cboLogStatus.width = editingWidthK / 2
        
    End With
 
    'Style Page Call Log
    collectionStyledControls.Add Styler.getStyledBoxULine(txtClaimID, "Claim Id*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtComment, "Comment")
    collectionStyledControls.Add Styler.getStyleButton(cmdFind)
    collectionStyledControls.Add Styler.getStyleButton(cmdLog)

    'Adding headers to listboxClaims
    Dim headersHistoryList As Variant
    headersHistoryList = Array("Caller Name", "Call Date", "Status", "Comment")
    ListBoxUtil.generateListBoxHeaders lstHistory, headersHistoryList
    

 End Sub

