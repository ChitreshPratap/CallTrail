VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPageAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageAdmin - controller for the "Admin" tab.
'
' The shell hides this whole tab for non-admins, but every action here
' ALSO re-checks the role server-side via the repository. Hiding a tab
' is a convenience, not a security control - anyone who can open the VBE
' can show it again. The real enforcement lives in the data layer.
'
' Bulk claim entry is launched from here but runs on the bulkClaimAdd
' WORKSHEET, not in this form. Pasting hundreds of rows is a worksheet
' job; cramming it into a form control would be worse at the thing the
' user actually needs to do.
' =====================================================================

Implements IPage

Private m_ctx As clsAppContext
Private m_pg As MSForms.Page

Private WithEvents cmdChangeSite As MSForms.CommandButton
Attribute cmdChangeSite.VB_VarHelpID = -1
'Attribute cmdChangeSite.VB_VarHelpID = -1
Private WithEvents cmdBulkSetup As MSForms.CommandButton
Attribute cmdBulkSetup.VB_VarHelpID = -1
'Attribute cmdBulkSetup.VB_VarHelpID = -1
Private WithEvents cmdBulkLoadFile As MSForms.CommandButton
Attribute cmdBulkLoadFile.VB_VarHelpID = -1
'Attribute cmdBulkLoadFile.VB_VarHelpID = -1
Private WithEvents cmdBulkCheck As MSForms.CommandButton
Attribute cmdBulkCheck.VB_VarHelpID = -1
'Attribute cmdBulkCheck.VB_VarHelpID = -1
Private WithEvents cmdBulkProcess As MSForms.CommandButton
Attribute cmdBulkProcess.VB_VarHelpID = -1
'Attribute cmdBulkProcess.VB_VarHelpID = -1
Private WithEvents cmdReopen As MSForms.CommandButton
Attribute cmdReopen.VB_VarHelpID = -1
'Attribute cmdReopen.VB_VarHelpID = -1
Private WithEvents cmdCheckStatus As MSForms.CommandButton
Attribute cmdCheckStatus.VB_VarHelpID = -1
'Attribute cmdCheckStatus.VB_VarHelpID = -1



Private txtSiteClaimID As MSForms.TextBox
Private txtNewSite As MSForms.TextBox
Private txtReopenClaimID As MSForms.TextBox
Private txtReopenReason As MSForms.TextBox
Private lblReopenStatus As MSForms.Label
Private lblAdminUser As MSForms.Label
Private lblBulkInfo As MSForms.Label

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Admin"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As MSForms.Page)
    Set m_ctx = ctx
    Set m_pg = pg
                                            
    Set txtSiteClaimID = BindControl(pg, "txtAdmClaimID", IPage_PageTitle)
    Set txtNewSite = BindControl(pg, "txtAdmNewSite", IPage_PageTitle)
    Set cmdChangeSite = BindControl(pg, "cmdAdmChangeSite", IPage_PageTitle)
    Set cmdBulkSetup = BindControl(pg, "cmdAdmBulkSetup", IPage_PageTitle)
    Set cmdBulkLoadFile = BindControl(pg, "cmdAdmBulkLoad", IPage_PageTitle)
    Set cmdBulkCheck = BindControl(pg, "cmdAdmBulkCheck", IPage_PageTitle)
    Set cmdBulkProcess = BindControl(pg, "cmdAdmBulkProcess", IPage_PageTitle)
     Set txtReopenClaimID = BindControl(pg, "txtAdmReopenID", IPage_PageTitle)
    Set txtReopenReason = BindControl(pg, "txtAdmReopenReason", IPage_PageTitle)
    Set cmdCheckStatus = BindControl(pg, "cmdAdmCheckStatus", IPage_PageTitle)
    Set cmdReopen = BindControl(pg, "cmdAdmReopen", IPage_PageTitle)
    Set lblReopenStatus = BindControl(pg, "lblAdmReopenStatus", IPage_PageTitle)

    txtReopenReason.Multiline = True
    cmdReopen.enabled = False        ' until a Closed claim is confirmed

    Set lblAdminUser = BindOptional(pg, "lblAdmUser")
    Set lblBulkInfo = BindOptional(pg, "lblAdmBulkInfo")
End Sub

Private Sub IPage_OnActivate()
    If Not lblAdminUser Is Nothing Then
        lblAdminUser.caption = "Signed in as: " & m_ctx.userName & _
                               IIf(m_ctx.isAdmin, " (Admin)", " (User)")
    End If
    If Not lblBulkInfo Is Nothing Then
        lblBulkInfo.caption = "Bulk entry uses the '" & BULK_SHEET & "' sheet. " & _
                              "Valid rows are imported and removed; rows with problems stay " & _
                              "with the reason in the ValidationError column."
    End If
    m_ctx.SetStatus "Admin tools."
End Sub

Private Sub cmdChangeSite_Click()
    If Trim$(txtSiteClaimID.Value) = "" Or Trim$(txtNewSite.Value) = "" Then
        MsgBox "Enter both a Claim ID and the new site.", vbExclamation, "Missing Details"
        Exit Sub
    End If

    ' ChangeUpdatedSite re-checks admin rights itself - this tab being
    ' visible is not what authorises the change.
    If m_ctx.repo.ChangeUpdatedSite(Trim$(txtSiteClaimID.Value), Trim$(txtNewSite.Value)) Then
        m_ctx.InvalidateClaims
        m_ctx.SetStatus "Site updated for " & Trim$(txtSiteClaimID.Value) & "."
        txtSiteClaimID.Value = ""
        txtNewSite.Value = ""
    End If
End Sub


 '---------------------------------------------------------------------
' Reopen a closed claim.
'
' Two steps on purpose: check first, then reopen. The Reopen button
' stays disabled until a claim has been confirmed Closed, so an admin
' can't fire a status change at a typo'd ID or at a claim that was
' already open.
' ---------------------------------------------------------------------
Private Sub cmdCheckStatus_Click()
    Dim c As clsClaim

    cmdReopen.enabled = False
    lblReopenStatus.caption = ""

    If Trim$(txtReopenClaimID.Value) = "" Then
        MsgBox "Enter a Claim ID to check.", vbExclamation, "Nothing to Check"
        Exit Sub
    End If

    Set c = m_ctx.repo.FindClaim(Trim$(txtReopenClaimID.Value))

    If c Is Nothing Then
        lblReopenStatus.caption = "No claim found with that ID."
        Exit Sub
    End If

    lblReopenStatus.caption = c.ClaimProviderName & " | " & c.claimSite & _
                              " | Status: " & c.ClaimStatus & _
                              " | Calls: " & c.attempt & _
                              IIf(c.IsClosed, " | Closed " & FormatDateOrBlank(c.ClaimClosedDate), "")

    If c.IsClosed Then
        cmdReopen.enabled = True
        m_ctx.SetStatus "Claim is Closed - enter a reason and reopen."
        txtReopenReason.SetFocus
    Else
        m_ctx.SetStatus "Claim is already " & c.ClaimStatus & " - nothing to reopen."
    End If
End Sub

Private Sub cmdReopen_Click()
    Dim claimID As String

    claimID = Trim$(txtReopenClaimID.Value)
    If claimID = "" Then Exit Sub

    If Trim$(txtReopenReason.Value) = "" Then
        MsgBox "A reason is required. It is written to the claim history as a " & _
               "permanent record of why this claim was reopened.", _
               vbExclamation, "Reason Needed"
        txtReopenReason.SetFocus
        Exit Sub
    End If

    If MsgBox("Reopen claim " & claimID & "?" & vbCrLf & vbCrLf & _
              "It returns to Pending and callers can log calls against it again. " & _
              "The reason is recorded in the claim history against your name.", _
              vbYesNo + vbQuestion, "Confirm Reopen") = vbNo Then Exit Sub

    If m_ctx.repo.ReopenClaim(claimID, Trim$(txtReopenReason.Value)) Then
        m_ctx.InvalidateClaims
        m_ctx.SetStatus "Claim " & claimID & " reopened and set back to Pending."
        txtReopenClaimID.Value = ""
        txtReopenReason.Value = ""
        lblReopenStatus.caption = ""
        cmdReopen.enabled = False
    End If
End Sub

' --- Bulk entry: these run against the bulkClaimAdd worksheet ---

Private Sub cmdBulkSetup_Click()
    SetupBulkSheet
End Sub

Private Sub cmdBulkLoadFile_Click()
    LoadClaimsFromFile
End Sub

Private Sub cmdBulkCheck_Click()
    CheckBulkClaims
End Sub

Private Sub cmdBulkProcess_Click()
    ProcessBulkClaims
    ' Claims were very likely added, so every other tab's cached list
    ' is now stale.
    m_ctx.InvalidateClaims
    m_ctx.SetStatus "Bulk processing finished - see the " & BULK_SHEET & " sheet."
End Sub

