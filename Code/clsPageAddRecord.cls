VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPageAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
' =====================================================================
' clsPageAddClaim - controller for the "Add Claim" tab.
'
' WithEvents is what makes this work: a class module can hold event
' handlers for MSForms controls, so all the code for this tab lives in
' this file rather than in one enormous form module.
' =====================================================================
 
Implements IPage
 
Private m_ctx As clsAppContext
Private m_pg As Msforms.Page
Private m_arranged As Boolean

Private WithEvents cmdSave As Msforms.Label
Attribute cmdSave.VB_VarHelpID = -1
'Attribute cmdSave.VB_VarHelpID = -1
Private WithEvents cmdClear As Msforms.Label
Attribute cmdClear.VB_VarHelpID = -1
'Attribute cmdClear.VB_VarHelpID = -1
Private WithEvents cmdCancel As Msforms.Label
Attribute cmdCancel.VB_VarHelpID = -1
Private WithEvents datePickerCreatedOn As Msforms.Label
Attribute datePickerCreatedOn.VB_VarHelpID = -1

Private txtClaimID As Msforms.TextBox
Private txtClaimSite As Msforms.TextBox
Private txtProviderName As Msforms.TextBox
Private txtQuery As Msforms.TextBox
Private txtCreationDate As Msforms.TextBox
Private lblInfo As Msforms.Label


Private collectionStyledControls As Collection
 
Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Add Claim"
End Property



 
 
Private Sub IPage_InitPage(ctx As clsAppContext, pg As Msforms.Page)
    
    Set m_ctx = ctx
    Set m_pg = pg
 
    Set txtClaimID = BindControl(pg, "txtAddClaimID", IPage_PageTitle)
    Set txtClaimSite = BindControl(pg, "txtAddClaimSite", IPage_PageTitle)
    Set txtProviderName = BindControl(pg, "txtAddProviderName", IPage_PageTitle)
    Set txtQuery = BindControl(pg, "txtAddClaimQuery", IPage_PageTitle)
    Set txtCreationDate = BindControl(pg, "dtAddCreationDate", IPage_PageTitle)
    Set datePickerCreatedOn = BindControl(pg, "dtPickerCreatedOn", IPage_PageTitle)
    Set cmdSave = BindControl(pg, "cmdAddSave", IPage_PageTitle)
    Set cmdClear = BindControl(pg, "cmdAddReset", IPage_PageTitle)
    Set cmdCancel = BindControl(pg, "cmdAddCancel", IPage_PageTitle)
    Set lblInfo = BindOptional(pg, "lblAddInfo")
    
    Set collectionStyledControls = New Collection
    txtQuery.Multiline = True
    
    ClearFields
    
End Sub
 
Private Sub IPage_OnActivate()
    ' Arranged once, on first visit. Doing it here rather than in
    ' InitPage means tabs the user never opens are never laid out - a
    ' small saving, and it keeps startup snappy.
'    ArrangeUI
   If Not m_arranged Then
        ArrangeUI
        m_arranged = True
    End If
    If Trim$(txtClaimSite.value) = "" Then txtClaimSite.value = m_ctx.DefaultLocation
    
    m_ctx.SetStatus "Add a new claim. All five fields are required."
         
    If Not lblInfo Is Nothing Then
        lblInfo.caption = "Status starts as Pending and Attempt at 0. " & _
                          "Insertion date and your username are recorded automatically."
    End If
    
End Sub

Private Sub datePickerCreatedOn_click()
    
    Dim dtPick As Date
    Dim dtPickStr As String
    
    dtPick = CalendarForm.GetDate()
    dtPickStr = Format(dtPick, "MMM-dd-yyyy")
    txtCreationDate.value = dtPickStr
    
End Sub

 
Private Sub cmdSave_Click()
        
    Dim tClaimId As String
    Dim tClaimSite As String
    Dim tProviderName As String
    Dim tDtCreationDate As String
    Dim tClaimQuery As String
    
    Dim txtStatus As String
    
    tClaimId = Trim(txtClaimID.value)
    tClaimSite = Trim(txtClaimSite.value)
    tProviderName = Trim(txtProviderName.value)
    tClaimQuery = Trim(txtQuery.value)
    tDtCreationDate = Trim(txtCreationDate.value)

    
    Dim creationDate As Date
 
    If tClaimId = "" Then
        Fail "'Claim ID' is required.", txtClaimID
        Exit Sub
    End If
    If tClaimSite = "" Then
        Fail "'Claim Site' is required.", txtClaimSite
        Exit Sub
    End If
    If tProviderName = "" Then
        Fail "'Provider Name' is required.", txtProviderName
        Exit Sub
    End If
    If tClaimQuery = "" Then
        Fail "'Claim Query' is required.", txtQuery
        Exit Sub
    End If
    
    If Not IsDate(tDtCreationDate) Then
        Fail "'Creation Date' is not a valid date.", txtCreationDate
        Exit Sub
    End If
 
    creationDate = CDate(txtCreationDate.value)
    
'    If creationDate > Date Then
'        Fail "Creation Date cannot be in the future.", txtCreationDate
'        Exit Sub
'    End If
    
    
    If m_ctx.repo.AddClaim(tClaimId, tClaimSite, _
                            tProviderName, tClaimQuery, _
                            creationDate) Then
 
        ' Tell every other tab their claim list is out of date. They
        ' reload lazily, on next activation - not right now.
        m_ctx.InvalidateClaims
        txtStatus = "Claim Id : '" & tClaimId & "' added."
        m_ctx.SetStatus txtStatus
        MsgBox txtStatus, vbOKOnly + vbInformation, "Claim Add Status"
        ClearFields
        txtClaimID.SetFocus
    End If
    
End Sub
 
Private Sub cmdClear_Click()

    ClearFields
    m_ctx.SetStatus "Cleared."
    'MsgBox "Form is cleared", vbOKOnly + vbInformation, "Clear Status"
    
End Sub
 
Private Sub ClearFields()
    
    txtClaimID.value = ""
    txtClaimSite.value = ""
    txtProviderName.value = ""
    txtQuery.value = ""
    'txtCreationDate.Value = Format$(Date, "dd-mmm-yyyy")
    txtCreationDate.value = ""

End Sub
 
Private Sub Fail(ByVal msg As String, ctrl As Msforms.control)
    MsgBox msg, vbExclamation, "Check Details"
    On Error Resume Next
    ctrl.SetFocus
    On Error GoTo 0
End Sub
 
Private Sub ArrangeUI()
 
            
    'Dim m_pg As UserFormMain
            
    With m_pg.framePageAddRecord
        
        .width = m_pg.Parent.width + 1
        .height = m_pg.Parent.height + 1
        '.BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        Styler.stylePageFrame m_pg.framePageAddRecord
        
'        .BackColor = vbYellow
        .Left = -1
        .Top = -1
        '.height = mainWindowHeight - lblCloseDashboard.height - lblCloseDashboard.Top
        
        .lblHeadingAddClaim.Top = 4.5
        .lblHeadingAddClaim.Left = 9
        Styler.getStyledHeading .lblHeadingAddClaim
        .lblHeadingULineAddClaim.Top = .lblHeadingAddClaim.Top + .lblHeadingAddClaim.height + 2
        .lblHeadingULineAddClaim.Left = 9
        Styler.getStyledHeadingULine .lblHeadingULineAddClaim
        .lblHeadingULineAddClaim.width = .width - 2 * .lblHeadingAddClaim.Left
        .lblHeadingULineAddClaim.height = 2
        
'        Me.frameFilterViewClaims.Left = .width - Me.frameFilterViewClaims.width
'        Me.frameFilterViewClaims.height = .height
'        Me.frameFilterViewClaims.Top = -2
        
        
    End With
 
    'Style Page Add Claim
    collectionStyledControls.Add Styler.getStyledBoxULine(txtClaimID, "Claim Id*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtClaimSite, "Claim Site*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtProviderName, "Provider Name*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtCreationDate, "Creation Date*")
    collectionStyledControls.Add Styler.getStyledBoxULine(txtQuery, "ClaimQuery*")
    collectionStyledControls.Add Styler.getStyleButton(cmdSave)
    collectionStyledControls.Add Styler.getStyleButton(cmdCancel)
    collectionStyledControls.Add Styler.getStyleButton(cmdClear)
    collectionStyledControls.Add Styler.getStyleButton(datePickerCreatedOn)
    

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
Private Sub arrangeUI_Back()
    
    Dim ui As New clsFormStyler
     
    '
 
    ui.Init m_pg, IPage_PageTitle, startLeft:=14, startTop:=12
 
    ' --- tunables for this tab ---
    ui.LabelWidth = 110
    ui.fieldWidth = 250
    ui.RowHeight = 18
    ui.RowGap = 6
 
    ' Reuse the scale the shell measured, so this tab matches the rest of
    ' the app on machines where the font renders differently.
    ' MUST come after the tunables above - it scales their current values.
    'ui.UseScale m_ctx.UIScale
 
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
 

' Enter in the ID box moves to the next field
Private Sub txtClaimID_KeyDown(ByVal KeyCode As Msforms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtClaimSite.SetFocus
End Sub

