VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchClaim 
   Caption         =   "UserForm1"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13155
   OleObjectBlob   =   "frmSearchClaim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================================
' frmSearchClaim - code-behind
'
' Paste ALL of this into the code module of a UserForm named
' frmSearchClaim. Control names must match the layout table in the
' Setup Guide, section 5a.
'
' Purpose: look up ONE claim, see everything about it including its
' full call history, and correct the detail fields if they were keyed
' in wrong.
'
' What is editable and what is not - and why:
'   EDITABLE   ClaimSite, ClaimProviderName, ClaimQuery, ClaimCreationDate
'              (what a user typed at insert time, so a user can fix it)
'              ClaimUpdatedSite - Admin only
'   READ ONLY  ClaimStatus, Attempt, ClaimClosedDate
'              These are maintained by LogCallAndUpdateStatus so status,
'              attempt count and the call trail stay in sync. Editing
'              status here would let someone close a claim with no call
'              behind it and make Attempt meaningless.
'   READ ONLY  ClaimID, ClaimInsertionDate, ClaimInsertedBy,
'              LastUpdatedDate, LastUpdatedBy - identity and audit fields
'   READ ONLY  The entire call history. It is an append-only trail;
'              rows are added by logging a call, never edited.
' =====================================================================
Option Explicit

Private m_claim As clsClaim      ' the claim currently loaded, or Nothing
Private m_dirty As Boolean       ' has the user changed an editable field?
Private m_loading As Boolean     ' suppress Change events while populating

' =====================================================================
' Setup
' =====================================================================
Private Sub UserForm_Initialize()
    m_loading = True

    Me.caption = "CallTrail - Search & Edit Claim"

    With lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "90;120;60;260"
    End With

    ClearAll
    m_loading = False
End Sub

' =====================================================================
' Search
' =====================================================================
Private Sub cmdSearch_Click()
    DoSearch
End Sub

' Enter key in the search box searches
Private Sub txtSearchID_KeyDown(ByVal KeyCode As Msforms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then DoSearch
End Sub

Private Sub DoSearch()
    
    Dim repo As New ClsClaimRepository
    Dim idToFind As String

    idToFind = Trim$(txtSearchID.value)
    If idToFind = "" Then
        MsgBox "Enter a Claim ID to search for.", vbExclamation, "Nothing to Search"
        Exit Sub
    End If

    If Not ConfirmDiscardIfDirty() Then Exit Sub

    On Error GoTo Fail
    Me.MousePointer = fmMousePointerHourGlass
    lblStatusBar.caption = "Searching..."
    DoEvents

    Set m_claim = repo.FindClaim(idToFind)

    If m_claim Is Nothing Then
        ClearAll
        Me.MousePointer = fmMousePointerDefault
        lblStatusBar.caption = "No claim found with ID '" & idToFind & "'."
        Exit Sub
    End If

    PopulateForm m_claim
    LoadHistory m_claim.claimID

    Me.MousePointer = fmMousePointerDefault
    lblStatusBar.caption = "Claim loaded."
    Exit Sub

Fail:
    Me.MousePointer = fmMousePointerDefault
    lblStatusBar.caption = "Search failed: " & Err.Description
End Sub

' =====================================================================
' Populate
' =====================================================================
Private Sub PopulateForm(c As clsClaim)
    m_loading = True

    ' --- read-only identity / audit fields ---
    lblClaimID.caption = c.claimID
    lblStatus.caption = c.ClaimStatus
    lblAttempt.caption = CStr(c.attempt)
    lblInsertedOn.caption = Format$(c.ClaimInsertionDate, "dd-mmm-yyyy hh:nn")
    lblInsertedBy.caption = c.ClaimInsertedBy
    lblLastUpdated.caption = FormatDateOrBlank(c.LastUpdatedDate)
    lblLastUpdatedBy.caption = c.LastUpdatedBy
    lblLastComment.caption = c.LastComment
    lblClosedOn.caption = FormatDateOrBlank(c.ClaimClosedDate)
    lblDaysOpen.caption = CStr(c.DaysOpen) & IIf(c.IsClosed, " (to close)", " (open)")

    ' --- editable fields ---
    txtSite.value = c.claimSite
    txtProvider.value = c.ClaimProviderName
    txtQuery.value = c.claimQuery
    txtCreationDate.value = Format$(c.ClaimCreationDate, "dd-mmm-yyyy")
    txtUpdatedSite.value = c.ClaimUpdatedSite

    ApplyPermissions
    SetEditingEnabled True

    m_dirty = False
    cmdSave.enabled = False
    m_loading = False
End Sub

' Only Admins may change the updated-site field; everyone else sees the
' value but can't type in it, which is clearer than hiding it entirely.
Private Sub ApplyPermissions()
    Dim wb As Workbook
    Dim IsAdmin As Boolean

    On Error Resume Next
    Set wb = OpenCentralDB()
    IsAdmin = IsCurrentUserAdmin(wb)
    CloseCentralDB wb, False
    On Error GoTo 0

    txtUpdatedSite.enabled = IsAdmin
    txtUpdatedSite.BackColor = IIf(IsAdmin, &H80000005, &H8000000F)  ' white / grey
    lblAdminNote.Visible = Not IsAdmin
End Sub

Private Sub LoadHistory(ByVal claimID As String)
    Dim repo As New ClsClaimRepository
    Dim histData As Variant
    Dim hist As Collection

    lstHistory.Clear

    On Error GoTo Fail
    Set hist = repo.GetHistory(claimID)
    histData = HistoryToListArray(hist)
    If Not IsEmpty(histData) Then lstHistory.List = histData

    lblHistoryCount.caption = hist.Count & " call(s) logged"
    Exit Sub

Fail:
    lblStatusBar.caption = "Could not load history: " & Err.Description
End Sub

' =====================================================================
' Editing
' =====================================================================
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

Private Sub MarkDirty()
    If m_loading Then Exit Sub
    If m_claim Is Nothing Then Exit Sub
    m_dirty = True
    cmdSave.enabled = True
    lblStatusBar.caption = "Unsaved changes."
End Sub

Private Sub cmdSave_Click()
    Dim repo As New ClsClaimRepository
    Dim siteDenied As Boolean
    Dim creationDate As Date

    If m_claim Is Nothing Then Exit Sub

    ' --- validate before writing anything ---
    If Trim$(txtSite.value) = "" Then
        MsgBox "Claim Site cannot be blank.", vbExclamation, "Check Details"
        txtSite.SetFocus
        Exit Sub
    End If
    If Trim$(txtProvider.value) = "" Then
        MsgBox "Provider Name cannot be blank.", vbExclamation, "Check Details"
        txtProvider.SetFocus
        Exit Sub
    End If
    If Trim$(txtQuery.value) = "" Then
        MsgBox "Claim Query cannot be blank.", vbExclamation, "Check Details"
        txtQuery.SetFocus
        Exit Sub
    End If
    If Not IsDate(txtCreationDate.value) Then
        MsgBox "Creation Date is not a valid date.", vbExclamation, "Check Details"
        txtCreationDate.SetFocus
        Exit Sub
    End If
    creationDate = CDate(txtCreationDate.value)
    If creationDate > Date Then
        MsgBox "Creation Date cannot be in the future.", vbExclamation, "Check Details"
        txtCreationDate.SetFocus
        Exit Sub
    End If

    On Error GoTo Fail
    Me.MousePointer = fmMousePointerHourGlass

    If repo.UpdateClaimDetails(m_claim.claimID, _
                                Trim$(txtSite.value), _
                                Trim$(txtProvider.value), _
                                Trim$(txtQuery.value), _
                                creationDate, _
                                Trim$(txtUpdatedSite.value), _
                                siteDenied) Then

        ' Reload from the database rather than trusting the form's copy,
        ' so the audit fields shown (LastUpdated etc.) are what was
        ' actually written.
        Set m_claim = repo.FindClaim(m_claim.claimID)
        If Not m_claim Is Nothing Then PopulateForm m_claim

        Me.MousePointer = fmMousePointerDefault
        lblStatusBar.caption = "Changes saved." & _
            IIf(siteDenied, " (Updated Site not changed - Admin only.)", "")
    Else
        Me.MousePointer = fmMousePointerDefault
        lblStatusBar.caption = "Save failed."
    End If
    Exit Sub

Fail:
    Me.MousePointer = fmMousePointerDefault
    MsgBox "Could not save: " & Err.Description, vbCritical
End Sub

' Throws away edits and re-reads the claim from the database
Private Sub cmdRevert_Click()
    Dim repo As New ClsClaimRepository

    If m_claim Is Nothing Then Exit Sub
    If Not m_dirty Then Exit Sub

    If MsgBox("Discard your unsaved changes to this claim?", _
              vbYesNo + vbQuestion, "Discard Changes") = vbNo Then Exit Sub

    Set m_claim = repo.FindClaim(m_claim.claimID)
    If m_claim Is Nothing Then
        ClearAll
        lblStatusBar.caption = "This claim no longer exists in the database."
        Exit Sub
    End If

    PopulateForm m_claim
    lblStatusBar.caption = "Changes discarded."
End Sub

Private Sub cmdClose_Click()
    If Not ConfirmDiscardIfDirty() Then Exit Sub
    Unload Me
End Sub

' Also catches the X button in the title bar
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If Not ConfirmDiscardIfDirty() Then Cancel = True
    End If
End Sub

' Returns False if the user wants to stay and keep editing
Private Function ConfirmDiscardIfDirty() As Boolean
    ConfirmDiscardIfDirty = True
    If Not m_dirty Then Exit Function
    If m_claim Is Nothing Then Exit Function

    If MsgBox("You have unsaved changes to claim " & m_claim.claimID & "." & vbCrLf & vbCrLf & _
              "Discard them?", vbYesNo + vbExclamation, "Unsaved Changes") = vbNo Then
        ConfirmDiscardIfDirty = False
    End If
End Function

' =====================================================================
' Helpers
' =====================================================================
Private Sub ClearAll()
    m_loading = True

    Set m_claim = Nothing
    m_dirty = False

    lblClaimID.caption = ""
    lblStatus.caption = ""
    lblAttempt.caption = ""
    lblInsertedOn.caption = ""
    lblInsertedBy.caption = ""
    lblLastUpdated.caption = ""
    lblLastUpdatedBy.caption = ""
    lblLastComment.caption = ""
    lblClosedOn.caption = ""
    lblDaysOpen.caption = ""
    lblHistoryCount.caption = ""

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

' Editable fields stay locked until a claim is actually loaded, so a
' user can't type into a blank form and wonder why Save does nothing.
Private Sub SetEditingEnabled(ByVal enabled As Boolean)
    txtSite.enabled = enabled
    txtProvider.enabled = enabled
    txtQuery.enabled = enabled
    txtCreationDate.enabled = enabled
    cmdRevert.enabled = enabled
    ' txtUpdatedSite is governed by ApplyPermissions, not this
    If Not enabled Then txtUpdatedSite.enabled = False
End Sub

Private Function FormatDateOrBlank(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    FormatDateOrBlank = Format$(CDate(v), "dd-mmm-yyyy hh:nn")
End Function

