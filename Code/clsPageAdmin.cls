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
Private WithEvents cmdBulkSetup As MSForms.CommandButton
Attribute cmdBulkSetup.VB_VarHelpID = -1
Private WithEvents cmdBulkLoadFile As MSForms.CommandButton
Attribute cmdBulkLoadFile.VB_VarHelpID = -1
Private WithEvents cmdBulkCheck As MSForms.CommandButton
Attribute cmdBulkCheck.VB_VarHelpID = -1
Private WithEvents cmdBulkProcess As MSForms.CommandButton
Attribute cmdBulkProcess.VB_VarHelpID = -1
Private WithEvents cmdDownload As MSForms.CommandButton
Attribute cmdDownload.VB_VarHelpID = -1
Private WithEvents cmdReopen As MSForms.CommandButton
Attribute cmdReopen.VB_VarHelpID = -1
Private WithEvents cmdCheckStatus As MSForms.CommandButton
Attribute cmdCheckStatus.VB_VarHelpID = -1
Private WithEvents cmdArchivePreview As MSForms.CommandButton
Attribute cmdArchivePreview.VB_VarHelpID = -1
Private WithEvents cmdArchiveRun As MSForms.CommandButton
Attribute cmdArchiveRun.VB_VarHelpID = -1
Private WithEvents cboArchiveAge As MSForms.ComboBox
Attribute cboArchiveAge.VB_VarHelpID = -1
Private WithEvents chkArchiveClosedOnly As MSForms.CheckBox
Attribute chkArchiveClosedOnly.VB_VarHelpID = -1
Private WithEvents lstArchiveClaims As MSForms.ListBox
Attribute lstArchiveClaims.VB_VarHelpID = -1

Private txtSiteClaimID As MSForms.TextBox
Private txtNewSite As MSForms.TextBox
Private txtReopenClaimID As MSForms.TextBox
Private txtReopenReason As MSForms.TextBox
Private lblReopenStatus As MSForms.Label
Private lblArchivePreview As MSForms.Label
Private lstArchiveHistory As MSForms.ListBox
Private lblArchiveHistCount As MSForms.Label

' The claims currently shown in the preview list. Held so a click can
' map a list row back to a claim without re-running the filter.
Private m_previewClaims As Collection

' A UserForm ListBox degrades badly long before Excel's row limits, and
' an archive run can easily match thousands of claims. Past this we show
' a slice and say so rather than freezing the form.
Private Const MAX_PREVIEW_ROWS As Long = 5000
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

    txtReopenReason.MultiLine = True
    cmdReopen.Enabled = False        ' until a Closed claim is confirmed

    Set cboArchiveAge = BindControl(pg, "cboAdmArchiveAge", IPage_PageTitle)
    Set chkArchiveClosedOnly = BindControl(pg, "chkAdmArchiveClosedOnly", IPage_PageTitle)
    Set cmdArchivePreview = BindControl(pg, "cmdAdmArchivePreview", IPage_PageTitle)
    Set cmdArchiveRun = BindControl(pg, "cmdAdmArchiveRun", IPage_PageTitle)
    Set lblArchivePreview = BindControl(pg, "lblAdmArchivePreview", IPage_PageTitle)
    Set lstArchiveClaims = BindControl(pg, "lstAdmArchiveClaims", IPage_PageTitle)
    Set lstArchiveHistory = BindControl(pg, "lstAdmArchiveHistory", IPage_PageTitle)
    Set lblArchiveHistCount = BindControl(pg, "lblAdmArchiveHistCount", IPage_PageTitle)

    ConfigureClaimList lstArchiveClaims
    ConfigureHistoryList lstArchiveHistory

    cboArchiveAge.List = Array("Older than 2 months", "Older than 3 months", _
                                "Older than 4 months", "Older than 5 months", _
                                "Older than 6 months", "Any age")
    cboArchiveAge.Value = "Older than 6 months"
    chkArchiveClosedOnly.Value = True       ' the safe default
    cmdArchiveRun.Enabled = False            ' until a preview has been run

    Set cmdDownload = BindOptional(pg, "cmdAdmDownload")

    Set lblAdminUser = BindOptional(pg, "lblAdmUser")
    Set lblBulkInfo = BindOptional(pg, "lblAdmBulkInfo")
End Sub

Private Sub IPage_OnActivate()
    If Not lblAdminUser Is Nothing Then
        lblAdminUser.Caption = "Signed in as: " & m_ctx.UserName & _
                               IIf(m_ctx.IsAdmin, " (Admin)", " (User)")
    End If
    If Not lblBulkInfo Is Nothing Then
        lblBulkInfo.Caption = "Bulk entry uses the '" & BULK_SHEET & "' sheet. " & _
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
    If m_ctx.Repo.ChangeUpdatedSite(Trim$(txtSiteClaimID.Value), Trim$(txtNewSite.Value)) Then
        m_ctx.InvalidateClaims
        m_ctx.SetStatus "Site updated for " & Trim$(txtSiteClaimID.Value) & "."
        txtSiteClaimID.Value = ""
        txtNewSite.Value = ""
    End If
End Sub

' ---------------------------------------------------------------------
' Reopen a closed claim.
'
' Two steps on purpose: check first, then reopen. The Reopen button
' stays disabled until a claim has been confirmed Closed, so an admin
' can't fire a status change at a typo'd ID or at a claim that was
' already open.
' ---------------------------------------------------------------------
Private Sub cmdCheckStatus_Click()
    Dim c As clsClaim

    cmdReopen.Enabled = False
    lblReopenStatus.Caption = ""

    If Trim$(txtReopenClaimID.Value) = "" Then
        MsgBox "Enter a Claim ID to check.", vbExclamation, "Nothing to Check"
        Exit Sub
    End If

    Set c = m_ctx.Repo.FindClaim(Trim$(txtReopenClaimID.Value))

    If c Is Nothing Then
        lblReopenStatus.Caption = "No claim found with that ID."
        Exit Sub
    End If

    lblReopenStatus.Caption = c.ClaimProviderName & " | " & c.ClaimSite & _
                              " | Status: " & c.ClaimStatus & _
                              " | Calls: " & c.Attempt & _
                              IIf(c.IsClosed, " | Closed " & FormatDateOrBlank(c.ClaimClosedDate), "")

    If c.IsClosed Then
        cmdReopen.Enabled = True
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

    If m_ctx.Repo.ReopenClaim(claimID, Trim$(txtReopenReason.Value)) Then
        m_ctx.InvalidateClaims
        m_ctx.SetStatus "Claim " & claimID & " reopened and set back to Pending."
        txtReopenClaimID.Value = ""
        txtReopenReason.Value = ""
        lblReopenStatus.Caption = ""
        cmdReopen.Enabled = False
    End If
End Sub

' ---------------------------------------------------------------------
' Archiving.
'
' Preview is mandatory: the Run button stays disabled until a preview
' has been taken, and any change to the criteria disables it again. An
' admin should never be one stray click away from moving hundreds of
' rows they haven't seen a count for.
' ---------------------------------------------------------------------
Private Function BuildArchiveService() As clsArchiveService
    Dim svc As New clsArchiveService
    svc.OlderThanMonths = MonthsFromCombo()
    svc.ClosedOnly = CBool(chkArchiveClosedOnly.Value)
    Set BuildArchiveService = svc
End Function

Private Function MonthsFromCombo() As Long
    Select Case CStr(cboArchiveAge.Value)
        Case "Older than 2 months": MonthsFromCombo = 2
        Case "Older than 3 months": MonthsFromCombo = 3
        Case "Older than 4 months": MonthsFromCombo = 4
        Case "Older than 5 months": MonthsFromCombo = 5
        Case "Older than 6 months": MonthsFromCombo = 6
        Case Else: MonthsFromCombo = 0      ' "Any age" - no date filter
    End Select
End Function

Private Sub cboArchiveAge_Change()
    InvalidatePreview
End Sub

Private Sub chkArchiveClosedOnly_Click()
    InvalidatePreview
End Sub

Private Sub InvalidatePreview()
    cmdArchiveRun.Enabled = False
    lblArchivePreview.Caption = "Criteria changed - preview again before archiving."
    ' Clear the list too: leaving a stale list beside changed criteria is
    ' worse than showing nothing, because it looks current.
    ClearPreviewLists
End Sub

Private Sub cmdArchivePreview_Click()
    Dim svc As clsArchiveService
    Dim n As Long
    Dim oldest As Variant

    Set svc = BuildArchiveService()
    n = svc.CountMatching(m_ctx.Claims(), oldest)

    If n = 0 Then
        lblArchivePreview.Caption = "No claims match these criteria - nothing to archive."
        ClearPreviewLists
        cmdArchiveRun.Enabled = False
        Exit Sub
    End If

    ' Show the claims themselves, not just a number. An admin about to
    ' move rows out of the working sheets should be able to look at what
    ' is going, and drill into any one of them.
    Set m_previewClaims = svc.GetMatching(m_ctx.Claims())
    PopulatePreviewList

    lblArchivePreview.Caption = n & " claim(s) would be archived, with all their call history." & _
        IIf(IsDate(oldest), "  Oldest raised " & Format$(oldest, "dd-mmm-yyyy") & ".", "") & _
        IIf(Not CBool(chkArchiveClosedOnly.Value), _
            "  WARNING: Pending claims are included - callers will lose sight of live work.", "")

    cmdArchiveRun.Enabled = True
    m_ctx.SetStatus "Preview ready - " & n & " claim(s) match. Click any row to see its call history."
End Sub

' ---------------------------------------------------------------------
' Preview list handling
' ---------------------------------------------------------------------
Private Sub PopulatePreviewList()
    Dim listData As Variant
    Dim shown As Collection

    lstArchiveClaims.Clear
    lstArchiveHistory.Clear
    lblArchiveHistCount.Caption = ""

    If m_previewClaims Is Nothing Then Exit Sub
    If m_previewClaims.Count = 0 Then Exit Sub

    If m_previewClaims.Count > MAX_PREVIEW_ROWS Then
        Set shown = TakeFirst(m_previewClaims, MAX_PREVIEW_ROWS)
        m_ctx.SetStatus m_previewClaims.Count & " claims match - listing the first " & _
                        MAX_PREVIEW_ROWS & ". The archive still moves ALL of them."
        ' The displayed collection is truncated so the click handler can
        ' index it directly. The archive run does NOT use this collection -
        ' it re-evaluates the criteria against the full list, so every
        ' matching claim still moves even though only a slice is listed.
        Set m_previewClaims = shown
    End If

    listData = ClaimsToListArray(m_previewClaims)
    If Not IsEmpty(listData) Then lstArchiveClaims.List = listData
End Sub

' Clicking a claim shows the exact history rows that would move with it.
Private Sub lstArchiveClaims_Click()
    Dim idx As Long
    Dim c As clsClaim
    Dim hist As Collection
    Dim histData As Variant

    lstArchiveHistory.Clear
    lblArchiveHistCount.Caption = ""

    idx = lstArchiveClaims.ListIndex
    If idx < 0 Then Exit Sub
    If m_previewClaims Is Nothing Then Exit Sub
    If idx + 1 > m_previewClaims.Count Then Exit Sub

    Set c = m_previewClaims(idx + 1)     ' ListIndex 0-based, Collection 1-based
    Set hist = m_ctx.Repo.GetHistory(c.ClaimID)

    histData = HistoryToListArray(hist)
    If Not IsEmpty(histData) Then lstArchiveHistory.List = histData

    lblArchiveHistCount.Caption = hist.Count & " history row(s) for " & c.ClaimID & _
                                   " would move with it"
End Sub

Private Sub ClearPreviewLists()
    Set m_previewClaims = Nothing
    lstArchiveClaims.Clear
    lstArchiveHistory.Clear
    lblArchiveHistCount.Caption = ""
End Sub

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

Private Sub cmdArchiveRun_Click()
    Dim svc As clsArchiveService
    Dim n As Long, moved As Long, histMoved As Long
    Dim oldest As Variant
    Dim warn As String

    Set svc = BuildArchiveService()
    n = svc.CountMatching(m_ctx.Claims(), oldest)

    If n = 0 Then
        lblArchivePreview.Caption = "Nothing matches any more - re-run the preview."
        cmdArchiveRun.Enabled = False
        Exit Sub
    End If

    If Not CBool(chkArchiveClosedOnly.Value) Then
        warn = vbCrLf & vbCrLf & "This includes PENDING claims. Callers will no longer see them " & _
               "in the working sheets."
    End If

    If MsgBox(n & " claim(s) and all their call history will be moved to " & _
              SHEET_ARC_CLAIMS & " / " & SHEET_ARC_HISTORY & _
              " and removed from the working sheets." & warn & vbCrLf & vbCrLf & _
              "This cannot be undone from the app - the rows would have to be moved back " & _
              "by hand." & vbCrLf & vbCrLf & "Continue?", _
              vbYesNo + vbExclamation, "Confirm Archive") = vbNo Then Exit Sub

    m_ctx.SetStatus "Archiving - this may take a moment..."
    moved = svc.ArchiveMatching(histMoved)

    If moved > 0 Then
        m_ctx.InvalidateClaims
        lblArchivePreview.Caption = moved & " claim(s) and " & histMoved & _
                                     " history row(s) archived."
        ClearPreviewLists
        m_ctx.SetStatus "Archive complete."
    Else
        m_ctx.SetStatus "Nothing was archived."
    End If

    cmdArchiveRun.Enabled = False
End Sub

' Snapshot the whole database into local sheets of this workbook.
' Optional control - BindOptional, so leaving the button off the tab
' doesn't break the page. The macro is also callable directly as
' modMain.DownloadDatabase.
Private Sub cmdDownload_Click()
    DownloadAllData
    m_ctx.SetStatus "Snapshot downloaded into this workbook - see the DL_ sheets."
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
