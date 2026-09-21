VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPageAddClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageAddClaim - controller for the "Add Claim" tab.
'
' *** REFERENCE EXAMPLE for styling and positioning controls that
'     already exist on the page. ***
'
' You place the controls in the VBE designer. Position and size don't
' matter there - drop them anywhere, roughly. ArrangeUI below snaps them
' into an aligned grid and applies consistent fonts, colours and borders
' at run time.
'
' The class has three clearly separated jobs:
'
'   BindControls - get object references, so the rest of the class can
'                  talk to the controls. WithEvents ones get their
'                  event handlers wired up here too.
'   ArrangeUI    - WHERE everything sits and HOW it looks. This is the
'                  design. Reads top-to-bottom like the screen.
'   Behaviour    - validation, saving. Untouched by layout changes.
'
' Keeping them separate is the point: you can re-lay-out the whole tab
' by editing one method, without going anywhere near the save logic.
' =====================================================================

Implements IPage

Private m_ctx As clsAppContext
Private m_pg As MSForms.Page
Private m_arranged As Boolean

' --- controls that raise events MUST be WithEvents ---
Private WithEvents cmdSave As MSForms.CommandButton
Attribute cmdSave.VB_VarHelpID = -1
Private WithEvents cmdClear As MSForms.CommandButton
Attribute cmdClear.VB_VarHelpID = -1
Private WithEvents txtClaimID As MSForms.TextBox
Attribute txtClaimID.VB_VarHelpID = -1

' --- controls we only read or write need no events ---
Private txtSite As MSForms.TextBox
Private txtProvider As MSForms.TextBox
Private txtQuery As MSForms.TextBox
Private txtCreationDate As MSForms.TextBox

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Add Claim"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As MSForms.Page)
    Set m_ctx = ctx
    Set m_pg = pg

    BindControls
    ClearFields
End Sub

Private Sub IPage_OnActivate()
    ' Arranged once, on first visit. Doing it here rather than in
    ' InitPage means tabs the user never opens are never laid out - a
    ' small saving, and it keeps startup snappy.
    If Not m_arranged Then
        ArrangeUI
        m_arranged = True
    End If

    ' Refresh the default site on each visit, so a user registered
    ' mid-session (via the Home tab's re-check) picks it up.
    If Trim$(txtSite.Value) = "" Then txtSite.Value = m_ctx.DefaultLocation

    m_ctx.SetStatus "Add a new claim. All five fields are required." & _
        IIf(m_ctx.DefaultLocation <> "", "  Site defaults to " & m_ctx.DefaultLocation & ".", "")
End Sub

' =====================================================================
' 1. BINDING - object references for the controls on the page
' =====================================================================
Private Sub BindControls()
    Set txtClaimID = BindControl(m_pg, "txtAddClaimID", IPage_PageTitle)
    Set txtSite = BindControl(m_pg, "txtAddSite", IPage_PageTitle)
    Set txtProvider = BindControl(m_pg, "txtAddProvider", IPage_PageTitle)
    Set txtQuery = BindControl(m_pg, "txtAddQuery", IPage_PageTitle)
    Set txtCreationDate = BindControl(m_pg, "txtAddCreationDate", IPage_PageTitle)
    Set cmdSave = BindControl(m_pg, "cmdAddSave", IPage_PageTitle)
    Set cmdClear = BindControl(m_pg, "cmdAddClear", IPage_PageTitle)
End Sub

' =====================================================================
' 2. THE DESIGN - where things sit and how they look.
'
' Every control is referenced BY NAME, so this reads as a description of
' the screen from top to bottom. To move a field, move its line. To
' re-space the whole tab, change the tunables - not twenty controls.
'
' Controls named here that only appear in this method (caption labels,
' the header, the notes) don't need to be bound above - nothing else in
' the class talks to them.
' =====================================================================
Private Sub ArrangeUI()
    Dim ui As New clsFormStyler

    ui.Init m_pg, IPage_PageTitle, startLeft:=14, startTop:=12

    ' --- tunables for this tab ---
    ui.LabelWidth = 110
    ui.FieldWidth = 250
    ui.RowHeight = 18
    ui.RowGap = 6

    ' Reuse the scale the shell measured, so this tab matches the rest of
    ' the app on machines where the font renders differently.
    ' MUST come after the tunables above - it scales their current values.
    ui.UseScale m_ctx.UIScale

    ' --- section: the five mandatory fields ---
    ui.Header "lblAddHeader", "New Claim Details"

    ui.Row "lblCapAddClaimID", "txtAddClaimID", "Claim ID:"
    ui.Row "lblCapAddSite", "txtAddSite", "Claim Site:"
    ui.Row "lblCapAddProvider", "txtAddProvider", "Provider Name:"
    ui.Row "lblCapAddCreationDate", "txtAddCreationDate", "Creation Date:", fieldWidth:=110
    ui.Row "lblCapAddQuery", "txtAddQuery", "Claim Query:", fieldHeight:=60

    ' --- section: buttons ---
    ui.Gap 4
    ui.BeginButtonRow
    ui.Button "cmdAddSave", "Save Claim", 90
    ui.Button "cmdAddClear", "Clear", 70
    ui.EndButtonRow

    ' --- section: notes (optional labels - skipped silently if absent) ---
    ui.Gap
    ui.Note "lblAddInfo", _
        "Status starts as Pending and Attempt at 0. The insertion date and " & _
        "your username are recorded automatically - you don't enter those.", _
        height:=30

    ui.Note "lblAddHint", _
        "Adding many claims at once? Use the Admin tab's bulk entry instead.", _
        height:=16
End Sub

' =====================================================================
' 3. BEHAVIOUR - untouched by anything above
' =====================================================================

' Enter in the ID box moves to the next field
Private Sub txtClaimID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSite.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim creationDate As Date

    If Trim$(txtClaimID.Value) = "" Then
        Fail "Claim ID is required.", txtClaimID
        Exit Sub
    End If
    If Trim$(txtSite.Value) = "" Then
        Fail "Claim Site is required.", txtSite
        Exit Sub
    End If
    If Trim$(txtProvider.Value) = "" Then
        Fail "Provider Name is required.", txtProvider
        Exit Sub
    End If
    If Trim$(txtQuery.Value) = "" Then
        Fail "Claim Query is required.", txtQuery
        Exit Sub
    End If
    If Not IsDate(txtCreationDate.Value) Then
        Fail "Creation Date is not a valid date.", txtCreationDate
        Exit Sub
    End If

    creationDate = CDate(txtCreationDate.Value)
    If creationDate > Date Then
        Fail "Creation Date cannot be in the future.", txtCreationDate
        Exit Sub
    End If

    If m_ctx.Repo.AddClaim(Trim$(txtClaimID.Value), Trim$(txtSite.Value), _
                            Trim$(txtProvider.Value), Trim$(txtQuery.Value), _
                            creationDate) Then

        ' Other tabs' cached claim list is now stale. They reload lazily,
        ' on next activation - not right now.
        m_ctx.InvalidateClaims

        m_ctx.SetStatus "Claim " & Trim$(txtClaimID.Value) & " added."
        ClearFields
        txtClaimID.SetFocus
    End If
End Sub

Private Sub cmdClear_Click()
    ClearFields
    m_ctx.SetStatus "Cleared."
End Sub

Private Sub ClearFields()
    txtClaimID.Value = ""
    txtProvider.Value = ""
    txtQuery.Value = ""
    txtCreationDate.Value = Format$(Date, "dd-mmm-yyyy")

    ' Site is pre-filled from the signed-in user's DefaultLocation in the
    ' Users sheet. It stays EDITABLE on purpose - a caller can legitimately
    ' enter a claim for another site, and locking the field would force
    ' them to the Admin tab for something routine.
    txtSite.Value = m_ctx.DefaultLocation
End Sub

Private Sub Fail(ByVal msg As String, ctrl As MSForms.Control)
    MsgBox msg, vbExclamation, "Check Details"
    On Error Resume Next
    ctrl.SetFocus
    On Error GoTo 0
End Sub
